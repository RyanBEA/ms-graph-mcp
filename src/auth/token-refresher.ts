import { ConfidentialClientApplication } from '@azure/msal-node';
import { ITokenManager } from './types.js';
import { AuthenticationError } from '../security/errors.js';
import { logger } from '../security/logger.js';
import { getConfiguration } from '../config/environment.js';

/**
 * Handles automatic token refresh with retry logic
 * Ensures only one refresh operation happens at a time
 */
export class TokenRefresher {
  private msalClient: ConfidentialClientApplication;
  private tokenManager: ITokenManager;
  private refreshPromise: Promise<string> | null = null;

  constructor(tokenManager: ITokenManager) {
    const config = getConfiguration();

    this.tokenManager = tokenManager;
    this.msalClient = new ConfidentialClientApplication({
      auth: {
        clientId: config.AZURE_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${config.AZURE_TENANT_ID}`,
        clientSecret: config.AZURE_CLIENT_SECRET,
      },
    });
  }

  /**
   * Get a valid access token, refreshing if necessary
   * Uses a 5-minute buffer before expiry to prevent race conditions
   *
   * @returns Valid access token
   * @throws AuthenticationError if refresh fails or tokens expired
   */
  async getValidAccessToken(): Promise<string> {
    const tokens = await this.tokenManager.getTokens();

    // Token is still valid (with 5-minute buffer)
    const expiryBuffer = 5 * 60 * 1000; // 5 minutes
    if (tokens.expiresAt - Date.now() > expiryBuffer) {
      return tokens.accessToken;
    }

    // Token needs refresh - ensure only one refresh happens
    if (this.refreshPromise) {
      return this.refreshPromise;
    }

    this.refreshPromise = this.refreshAccessToken(tokens.refreshToken);

    try {
      const newAccessToken = await this.refreshPromise;
      return newAccessToken;
    } finally {
      this.refreshPromise = null;
    }
  }

  /**
   * Refresh the access token using the refresh token
   * Clears all tokens if refresh fails (forces re-authentication)
   *
   * @param refreshToken - Current refresh token
   * @returns New access token
   * @throws AuthenticationError if refresh fails
   */
  private async refreshAccessToken(refreshToken: string): Promise<string> {
    logger.info('Refreshing access token');

    try {
      const response = await this.msalClient.acquireTokenByRefreshToken({
        refreshToken,
        scopes: ['Tasks.Read', 'User.Read'],
      });

      if (!response) {
        throw new AuthenticationError('Token refresh failed');
      }

      // Update stored tokens
      await this.tokenManager.updateAccessToken(
        response.accessToken,
        response.expiresOn?.getTime() || Date.now() + 3600000
      );

      logger.info('Access token refreshed successfully');
      return response.accessToken;

    } catch (error) {
      logger.error('Token refresh failed', { error });

      // If refresh token is invalid, clear all tokens
      await this.tokenManager.clearTokens();

      throw new AuthenticationError('Session expired. Please re-authenticate.');
    }
  }
}
