import keytar from 'keytar';
import { ITokenManager, TokenSet } from './types.js';
import { TokenStorageError } from '../security/errors.js';
import { logger } from '../security/logger.js';

/**
 * Token manager using keytar (Windows Credential Manager)
 * Stores each token field as a separate credential entry
 */
export class KeytarTokenManager implements ITokenManager {
  private readonly SERVICE_NAME = 'toto-mcp-server';

  async storeTokens(tokens: TokenSet): Promise<void> {
    try {
      logger.debug('Storing tokens in keytar', {
        hasAccessToken: !!tokens.accessToken,
        hasRefreshToken: !!tokens.refreshToken,
        hasExpiresAt: !!tokens.expiresAt,
        hasScope: !!tokens.scope,
      });

      await keytar.setPassword(
        this.SERVICE_NAME,
        'access_token',
        tokens.accessToken
      );
      await keytar.setPassword(
        this.SERVICE_NAME,
        'refresh_token',
        tokens.refreshToken
      );
      await keytar.setPassword(
        this.SERVICE_NAME,
        'expires_at',
        tokens.expiresAt.toString()
      );
      await keytar.setPassword(
        this.SERVICE_NAME,
        'scope',
        tokens.scope
      );

      logger.info('Tokens stored in Windows Credential Manager');
    } catch (error) {
      logger.error('Failed to store tokens in keytar', {
        error,
        errorMessage: error instanceof Error ? error.message : 'Unknown error',
        errorStack: error instanceof Error ? error.stack : undefined,
      });
      throw new TokenStorageError('Failed to store tokens');
    }
  }

  async getTokens(): Promise<TokenSet> {
    try {
      const accessToken = await keytar.getPassword(
        this.SERVICE_NAME,
        'access_token'
      );
      const refreshToken = await keytar.getPassword(
        this.SERVICE_NAME,
        'refresh_token'
      );
      const expiresAt = await keytar.getPassword(
        this.SERVICE_NAME,
        'expires_at'
      );
      const scope = await keytar.getPassword(
        this.SERVICE_NAME,
        'scope'
      );

      if (!accessToken || !refreshToken || !expiresAt || !scope) {
        throw new TokenStorageError('Incomplete token set');
      }

      return {
        accessToken,
        refreshToken,
        expiresAt: parseInt(expiresAt, 10),
        scope,
      };
    } catch (error) {
      logger.error('Failed to retrieve tokens from keytar');
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
      await keytar.deletePassword(this.SERVICE_NAME, 'access_token');
      await keytar.deletePassword(this.SERVICE_NAME, 'refresh_token');
      await keytar.deletePassword(this.SERVICE_NAME, 'expires_at');
      await keytar.deletePassword(this.SERVICE_NAME, 'scope');

      logger.info('Tokens cleared from Windows Credential Manager');
    } catch (error) {
      logger.error('Failed to clear tokens from keytar', { error });
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
