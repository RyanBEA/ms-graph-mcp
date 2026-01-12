import { ITokenManager, TokenSet } from './types.js';
import { TokenStorageError } from '../security/errors.js';
import { logger } from '../security/logger.js';

/**
 * Token manager using 1Password SDK
 * Stores tokens in a 1Password vault as API credentials
 *
 * NOTE: This is a stub implementation. The actual 1Password SDK integration
 * will be completed in a future phase once we have proper access to test.
 * For now, use KeytarTokenManager for development.
 */
export class OnePasswordTokenManager implements ITokenManager {
  /**
   * Initialize the 1Password client
   * Must be called before using any other methods
   */
  async initialize(): Promise<void> {
    logger.warn('1Password token manager is not yet fully implemented');
    throw new TokenStorageError(
      '1Password integration not yet implemented. Please use TOKEN_STORAGE=keytar'
    );
  }

  async storeTokens(_tokens: TokenSet): Promise<void> {
    throw new TokenStorageError('1Password not yet implemented');
  }

  async getTokens(): Promise<TokenSet> {
    throw new TokenStorageError('1Password not yet implemented');
  }

  async hasValidTokens(): Promise<boolean> {
    return false;
  }

  async clearTokens(): Promise<void> {
    logger.info('1Password clearTokens called (stub implementation)');
  }

  async updateAccessToken(_accessToken: string, _expiresAt: number): Promise<void> {
    throw new TokenStorageError('1Password not yet implemented');
  }
}
