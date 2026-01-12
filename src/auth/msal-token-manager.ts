import { ConfidentialClientApplication } from '@azure/msal-node';
import { promises as fs } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { ITokenManager, TokenSet } from './types.js';
import { TokenStorageError } from '../security/errors.js';
import { logger } from '../security/logger.js';
import { getConfiguration } from '../config/environment.js';

// Get the directory where this module is located, then go up to project root
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = path.resolve(__dirname, '..', '..');

const CACHE_FILE = process.env.MSAL_CACHE_FILE || path.join(PROJECT_ROOT, '.msal-cache.json');
const ACCOUNT_FILE = process.env.MSAL_ACCOUNT_FILE || path.join(PROJECT_ROOT, '.msal-account.json');

/**
 * Token manager using MSAL's built-in cache with file persistence.
 * This avoids the Windows Credential Manager size limitations.
 * MSAL handles refresh tokens internally and stores them in its cache.
 */
export class MsalTokenManager implements ITokenManager {
  private msalClient: ConfidentialClientApplication;
  private accountId: string | null = null;
  private cacheLoaded: boolean = false;

  constructor() {
    const config = getConfiguration();

    this.msalClient = new ConfidentialClientApplication({
      auth: {
        clientId: config.AZURE_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${config.AZURE_TENANT_ID}`,
        clientSecret: config.AZURE_CLIENT_SECRET,
      },
    });
  }

  /**
   * Ensure cache is loaded before operations
   */
  private async ensureCacheLoaded(): Promise<void> {
    if (!this.cacheLoaded) {
      await this.loadCache();
      this.cacheLoaded = true;
    }
  }

  /**
   * Load cache from file
   */
  private async loadCache(): Promise<void> {
    try {
      const data = await fs.readFile(CACHE_FILE, 'utf-8');
      const cache = this.msalClient.getTokenCache();
      cache.deserialize(data);
      logger.debug('MSAL cache loaded from file');
    } catch (error) {
      // Cache file doesn't exist yet, that's okay
      if ((error as NodeJS.ErrnoException).code !== 'ENOENT') {
        logger.debug('Failed to read MSAL cache file', { error });
      }
    }
  }

  /**
   * Save cache to file
   */
  private async saveCache(): Promise<void> {
    try {
      const cache = this.msalClient.getTokenCache();
      const data = cache.serialize();
      await fs.writeFile(CACHE_FILE, data);
      logger.debug('MSAL cache saved to file');
    } catch (error) {
      logger.error('Failed to write MSAL cache file', { error });
    }
  }

  /**
   * Store tokens by setting the account ID.
   * MSAL stores the actual tokens in its cache automatically.
   */
  async storeTokens(tokens: TokenSet): Promise<void> {
    try {
      // The refreshToken we receive is actually the account's homeAccountId
      // MSAL manages the actual refresh tokens internally
      this.accountId = tokens.refreshToken;

      // Persist the account ID to a file so it survives process restarts
      await fs.writeFile(ACCOUNT_FILE, JSON.stringify({ accountId: this.accountId }));

      // Save the MSAL cache to file
      await this.saveCache();

      logger.info('Token reference stored (using MSAL cache)');
    } catch (error) {
      logger.error('Failed to store token reference', { error });
      throw new TokenStorageError('Failed to store tokens');
    }
  }

  /**
   * Get tokens using MSAL's acquireTokenSilent with the stored account ID.
   */
  async getTokens(): Promise<TokenSet> {
    try {
      // Ensure cache is loaded first
      await this.ensureCacheLoaded();

      // Load account ID from file if not in memory
      if (!this.accountId) {
        try {
          const data = await fs.readFile(ACCOUNT_FILE, 'utf-8');
          const { accountId } = JSON.parse(data);
          this.accountId = accountId;
        } catch (error) {
          throw new TokenStorageError('No account ID stored');
        }
      }

      if (!this.accountId) {
        throw new TokenStorageError('No account ID stored');
      }

      const cache = this.msalClient.getTokenCache();
      const accounts = await cache.getAllAccounts();

      const account = accounts.find(acc => acc.homeAccountId === this.accountId);

      if (!account) {
        throw new TokenStorageError('Account not found in MSAL cache');
      }

      // Use MSAL's acquireTokenSilent to get a fresh token
      const response = await this.msalClient.acquireTokenSilent({
        account,
        scopes: ['Tasks.Read', 'Calendars.Read', 'User.Read', 'offline_access'],
      });

      if (!response || !response.accessToken) {
        throw new TokenStorageError('Failed to acquire token silently');
      }

      return {
        accessToken: response.accessToken,
        refreshToken: account.homeAccountId,
        expiresAt: response.expiresOn?.getTime() || Date.now() + 3600000,
        scope: response.scopes?.join(' ') || 'Tasks.Read User.Read',
      };
    } catch (error) {
      logger.error('Failed to retrieve tokens from MSAL cache', { error });
      throw new TokenStorageError('No tokens found');
    }
  }

  /**
   * Check if we have valid tokens by trying to get them.
   */
  async hasValidTokens(): Promise<boolean> {
    try {
      await this.ensureCacheLoaded();
      const tokens = await this.getTokens();
      return tokens.expiresAt > Date.now();
    } catch {
      return false;
    }
  }

  /**
   * Clear tokens by removing the account from MSAL cache.
   */
  async clearTokens(): Promise<void> {
    try {
      if (this.accountId) {
        const cache = this.msalClient.getTokenCache();
        const accounts = await cache.getAllAccounts();

        const account = accounts.find(acc => acc.homeAccountId === this.accountId);

        if (account) {
          await cache.removeAccount(account);
          logger.info('Account removed from MSAL cache');
        }

        this.accountId = null;
      }
    } catch (error) {
      logger.error('Failed to clear tokens from MSAL cache', { error });
      throw new TokenStorageError('Failed to clear tokens');
    }
  }

  /**
   * Update access token - with MSAL cache, we don't need to do anything
   * as MSAL will automatically use the cached tokens.
   */
  async updateAccessToken(_accessToken: string, _expiresAt: number): Promise<void> {
    // MSAL manages this automatically, nothing to do
    logger.debug('Token update requested - MSAL handles this automatically');
  }

  /**
   * Get the MSAL client instance (for use by OAuth client).
   */
  getMsalClient(): ConfidentialClientApplication {
    return this.msalClient;
  }

  /**
   * Get the stored account ID.
   */
  getAccountId(): string | null {
    return this.accountId;
  }
}
