import { ITokenManager } from './types.js';
import { MsalTokenManager } from './msal-token-manager.js';
import { logger } from '../security/logger.js';

let tokenManagerInstance: ITokenManager | null = null;

/**
 * Get or create a singleton token manager instance.
 * Uses MsalTokenManager which leverages MSAL's built-in cache
 * and acquireTokenSilent for automatic token refresh.
 *
 * @returns Configured token manager instance
 */
export async function getTokenManager(): Promise<ITokenManager> {
  if (tokenManagerInstance) {
    return tokenManagerInstance;
  }

  logger.info('Using MSAL token storage');
  tokenManagerInstance = new MsalTokenManager();

  return tokenManagerInstance;
}

/**
 * Reset the token manager singleton
 * Primarily used for testing to allow fresh instances
 */
export function resetTokenManager(): void {
  tokenManagerInstance = null;
}
