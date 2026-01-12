import { ITokenManager } from './types.js';
import { FileTokenManager } from './file-token-manager.js';
import { logger } from '../security/logger.js';

let tokenManagerInstance: ITokenManager | null = null;

/**
 * Get or create a singleton token manager instance
 * Automatically selects token storage based on configuration:
 * - file: Simple JSON file storage (default, simplest)
 * - msal: MSAL's built-in cache (most reliable for production)
 * - keytar: Windows Credential Manager (may have size limits)
 * - 1password: 1Password SDK (requires service account)
 *
 * @returns Configured token manager instance
 */
export async function getTokenManager(): Promise<ITokenManager> {
  if (tokenManagerInstance) {
    return tokenManagerInstance;
  }

  // HARDCODED: Always use file storage for simplicity
  logger.info('Using file-based token storage (hardcoded)');
  tokenManagerInstance = new FileTokenManager();

  return tokenManagerInstance;
}

/**
 * Reset the token manager singleton
 * Primarily used for testing to allow fresh instances
 */
export function resetTokenManager(): void {
  tokenManagerInstance = null;
}
