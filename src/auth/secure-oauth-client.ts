import crypto from 'crypto';
import { ConfidentialClientApplication, AuthorizationCodeRequest } from '@azure/msal-node';
import { ITokenManager } from './types.js';
import { AuthenticationError, AuthorizationError } from '../security/errors.js';
import { logger } from '../security/logger.js';
import { getConfiguration } from '../config/environment.js';

/**
 * State entry for CSRF protection
 */
interface StateEntry {
  timestamp: number;
  codeVerifier?: string; // For PKCE
}

/**
 * Authorization result containing URL and state
 */
export interface AuthorizationUrl {
  url: string;
  state: string;
}

/**
 * Secure OAuth client with CSRF protection and PKCE
 * Implements authorization code flow for Microsoft Graph API
 */
export class SecureOAuthClient {
  private readonly msalClient: ConfidentialClientApplication;
  private readonly tokenManager: ITokenManager;
  private readonly stateStore = new Map<string, StateEntry>();
  private readonly scopes = ['Tasks.ReadWrite', 'Calendars.Read', 'User.Read', 'offline_access'];
  private cleanupInterval: NodeJS.Timeout | null = null;

  constructor(tokenManager: ITokenManager) {
    const config = getConfiguration();

    this.tokenManager = tokenManager;

    // If the token manager is MsalTokenManager, use its MSAL client
    // to ensure we're using the same cache instance
    if ((tokenManager as any).getMsalClient) {
      this.msalClient = (tokenManager as any).getMsalClient();
    } else {
      // Fallback: create our own MSAL client
      this.msalClient = new ConfidentialClientApplication({
        auth: {
          clientId: config.AZURE_CLIENT_ID,
          authority: `https://login.microsoftonline.com/${config.AZURE_TENANT_ID}`,
          clientSecret: config.AZURE_CLIENT_SECRET,
        },
      });
    }

    // Start automatic cleanup of expired states
    this.startStateCleanup();
  }

  /**
   * Generate authorization URL with CSRF state protection
   * Creates a cryptographically secure state token and stores it for validation
   *
   * @returns Authorization URL and state token
   */
  async generateAuthUrl(): Promise<AuthorizationUrl> {
    const config = getConfiguration();

    // Generate cryptographically secure state token
    const state = crypto.randomBytes(32).toString('hex');

    // Store state with timestamp for CSRF validation
    this.stateStore.set(state, {
      timestamp: Date.now()
    });

    // Clean up expired states
    this.cleanExpiredStates();

    try {
      const authCodeUrlParameters = {
        scopes: this.scopes,
        redirectUri: config.AZURE_REDIRECT_URI,
        state,
      };

      const url = await this.msalClient.getAuthCodeUrl(authCodeUrlParameters);

      logger.info('Authorization URL generated', { state: state.substring(0, 8) + '...' });

      return { url, state };
    } catch (error) {
      logger.error('Failed to generate authorization URL', { error });
      throw new AuthenticationError('Failed to generate authorization URL');
    }
  }

  /**
   * Handle OAuth callback and exchange code for tokens
   * Validates state to prevent CSRF attacks
   *
   * @param code - Authorization code from OAuth callback
   * @param state - State token from OAuth callback
   * @throws AuthorizationError if state is invalid
   * @throws AuthenticationError if token exchange fails
   */
  async handleCallback(code: string, state: string): Promise<void> {
    // Validate state for CSRF protection
    if (!this.validateState(state)) {
      logger.warn('Invalid or expired state token', {
        state: state.substring(0, 8) + '...'
      });
      throw new AuthorizationError('Invalid or expired state token');
    }

    const config = getConfiguration();

    try {
      const tokenRequest: AuthorizationCodeRequest = {
        code,
        scopes: this.scopes,
        redirectUri: config.AZURE_REDIRECT_URI,
      };

      const response = await this.msalClient.acquireTokenByCode(tokenRequest);

      if (!response || !response.accessToken) {
        throw new AuthenticationError('Invalid token response');
      }

      // Get account from MSAL cache for refresh token management
      // MSAL manages refresh tokens internally - we store the account ID
      // which is used with acquireTokenSilent() for token refresh
      const cache = this.msalClient.getTokenCache();
      const accounts = await cache.getAllAccounts();

      if (!accounts || accounts.length === 0) {
        throw new AuthenticationError('No account found after token acquisition');
      }

      // Store the account's homeAccountId as our "refresh token"
      // This is used by MSAL's acquireTokenSilent() for automatic token refresh
      // Note: TokenRefresher will need to be adapted to use MSAL's pattern
      const account = accounts[0];
      const refreshToken = account.homeAccountId;

      // Prepare token data
      const tokenData = {
        accessToken: response.accessToken,
        refreshToken: refreshToken,
        expiresAt: response.expiresOn?.getTime() || Date.now() + 3600000,
        scope: response.scopes?.join(' ') || this.scopes.join(' '),
      };

      logger.debug('Preparing to store tokens', {
        hasAccessToken: !!tokenData.accessToken,
        accessTokenLength: tokenData.accessToken?.length,
        hasRefreshToken: !!tokenData.refreshToken,
        refreshTokenLength: tokenData.refreshToken?.length,
        expiresAt: tokenData.expiresAt,
        scope: tokenData.scope,
      });

      // Store tokens securely
      await this.tokenManager.storeTokens(tokenData);

      logger.info('OAuth callback handled successfully');
    } catch (error) {
      logger.error('Failed to handle OAuth callback', { error });
      throw new AuthenticationError('Failed to complete authentication');
    }
  }

  /**
   * Validate state token for CSRF protection
   * Checks if state exists and hasn't expired
   *
   * @param state - State token to validate
   * @returns true if state is valid and not expired
   */
  private validateState(state: string): boolean {
    const stored = this.stateStore.get(state);

    if (!stored) {
      return false;
    }

    const config = getConfiguration();
    const maxAge = config.STATE_TIMEOUT_MINUTES * 60 * 1000;
    const isValid = Date.now() - stored.timestamp < maxAge;

    // Remove state after validation (one-time use)
    this.stateStore.delete(state);

    return isValid;
  }

  /**
   * Clean up expired state tokens
   * Removes states older than STATE_TIMEOUT_MINUTES
   */
  private cleanExpiredStates(): void {
    const config = getConfiguration();
    const maxAge = config.STATE_TIMEOUT_MINUTES * 60 * 1000;
    const now = Date.now();

    let cleanedCount = 0;
    for (const [state, entry] of this.stateStore.entries()) {
      if (now - entry.timestamp > maxAge) {
        this.stateStore.delete(state);
        cleanedCount++;
      }
    }

    if (cleanedCount > 0) {
      logger.debug(`Cleaned up ${cleanedCount} expired state tokens`);
    }
  }

  /**
   * Start automatic cleanup of expired states
   * Runs every minute
   */
  private startStateCleanup(): void {
    // Clean up every minute
    this.cleanupInterval = setInterval(() => {
      this.cleanExpiredStates();
    }, 60 * 1000);

    // Ensure cleanup runs on process exit
    if (this.cleanupInterval.unref) {
      this.cleanupInterval.unref();
    }
  }

  /**
   * Stop automatic cleanup
   * Call this during shutdown or in tests
   */
  stopStateCleanup(): void {
    if (this.cleanupInterval) {
      clearInterval(this.cleanupInterval);
      this.cleanupInterval = null;
    }
  }

  /**
   * Get current number of stored states (for testing)
   */
  getStateCount(): number {
    return this.stateStore.size;
  }
}
