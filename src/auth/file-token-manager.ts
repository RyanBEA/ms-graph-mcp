import { promises as fs } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { ITokenManager, TokenSet } from './types.js';
import { TokenStorageError } from '../security/errors.js';
import { logger } from '../security/logger.js';

// Use absolute path relative to the package root, not cwd
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const TOKEN_FILE = path.join(__dirname, '..', '..', '.tokens.json');

/**
 * Simple file-based token manager.
 * Stores tokens in a JSON file - simple and reliable.
 */
export class FileTokenManager implements ITokenManager {
  async storeTokens(tokens: TokenSet): Promise<void> {
    try {
      logger.debug('Storing tokens to file', { path: TOKEN_FILE });
      await fs.writeFile(TOKEN_FILE, JSON.stringify(tokens, null, 2));
      logger.info('Tokens stored in file', { path: TOKEN_FILE });
    } catch (error) {
      logger.error('Failed to store tokens', { path: TOKEN_FILE, error });
      throw new TokenStorageError('Failed to store tokens');
    }
  }

  async getTokens(): Promise<TokenSet> {
    try {
      logger.debug('Reading tokens from file', { path: TOKEN_FILE });
      const data = await fs.readFile(TOKEN_FILE, 'utf-8');
      const tokens = JSON.parse(data);
      logger.debug('Tokens loaded successfully');
      return tokens;
    } catch (error) {
      logger.error('Failed to retrieve tokens', { path: TOKEN_FILE, error });
      throw new TokenStorageError('No tokens found');
    }
  }

  async hasValidTokens(): Promise<boolean> {
    try {
      const tokens = await this.getTokens();
      return tokens.expiresAt > Date.now();
    } catch {
      return false;
    }
  }

  async clearTokens(): Promise<void> {
    try {
      await fs.unlink(TOKEN_FILE);
      logger.info('Tokens cleared');
    } catch (error) {
      logger.error('Failed to clear tokens', { error });
      throw new TokenStorageError('Failed to clear tokens');
    }
  }

  async updateAccessToken(accessToken: string, expiresAt: number): Promise<void> {
    const tokens = await this.getTokens();
    await this.storeTokens({
      ...tokens,
      accessToken,
      expiresAt,
    });
  }
}
