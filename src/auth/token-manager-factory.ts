import { ITokenManager } from './types.js';
import { MsalTokenManager } from './msal-token-manager.js';
import { FileTokenManager } from './file-token-manager.js';
import { logger } from '../security/logger.js';
import { getConfiguration } from '../config/environment.js';

let tokenManagerInstance: ITokenManager | null = null;

/**
 * Get or create a singleton token manager instance.
 * Respects TOKEN_STORAGE config: 'file' uses FileTokenManager,
 * others use MsalTokenManager.
 *
 * @returns Configured token manager instance
 */
export async function getTokenManager(): Promise<ITokenManager> {
  if (tokenManagerInstance) {
    return tokenManagerInstance;
  }

  const config = getConfiguration();

  if (config.TOKEN_STORAGE === 'file') {
    logger.info('Using file-based token storage');
    tokenManagerInstance = new FileTokenManager();
  } else {
    logger.info('Using MSAL token storage');
    tokenManagerInstance = new MsalTokenManager();
  }

  return tokenManagerInstance;
}

/**
 * Reset the token manager singleton
 * Primarily used for testing to allow fresh instances
 */
export function resetTokenManager(): void {
  tokenManagerInstance = null;
}
