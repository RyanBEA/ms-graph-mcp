import { describe, it, expect, vi, beforeEach } from 'vitest';
import { TokenRefresher } from '../../src/auth/token-refresher.js';
import { ITokenManager, TokenSet } from '../../src/auth/types.js';
import { AuthenticationError } from '../../src/security/errors.js';

// Mock environment configuration
vi.mock('../../src/config/environment.js', () => ({
  getConfiguration: vi.fn(() => ({
    AZURE_CLIENT_ID: 'test-client-id',
    AZURE_TENANT_ID: 'test-tenant-id',
    AZURE_CLIENT_SECRET: 'test-secret',
    AZURE_REDIRECT_URI: 'http://localhost:3000/callback',
    TOKEN_STORAGE: 'msal',
    NODE_ENV: 'test',
    LOG_LEVEL: 'info',
    RATE_LIMIT_PER_MINUTE: 60,
    STATE_TIMEOUT_MINUTES: 5,
  })),
}));

// Mock token manager implementation
class MockTokenManager implements ITokenManager {
  private tokens: TokenSet | null = null;

  async storeTokens(tokens: TokenSet): Promise<void> {
    this.tokens = tokens;
  }

  async getTokens(): Promise<TokenSet> {
    if (!this.tokens) {
      throw new AuthenticationError('No tokens stored');
    }
    return this.tokens;
  }

  async hasValidTokens(): Promise<boolean> {
    return this.tokens !== null && this.tokens.expiresAt > Date.now();
  }

  async clearTokens(): Promise<void> {
    this.tokens = null;
  }

  async updateAccessToken(accessToken: string, expiresAt: number): Promise<void> {
    if (this.tokens) {
      this.tokens.accessToken = accessToken;
      this.tokens.expiresAt = expiresAt;
    }
  }
}

describe('TokenRefresher', () => {
  let manager: MockTokenManager;
  let refresher: TokenRefresher;

  beforeEach(() => {
    manager = new MockTokenManager();
    refresher = new TokenRefresher(manager);
    vi.clearAllMocks();
  });

  describe('getValidAccessToken', () => {
    it('should return access token from token manager', async () => {
      await manager.storeTokens({
        accessToken: 'test-token',
        refreshToken: 'refresh-token',
        expiresAt: Date.now() + 3600000, // 1 hour from now
        scope: 'Tasks.ReadWrite User.Read',
      });

      const token = await refresher.getValidAccessToken();

      expect(token).toBe('test-token');
    });

    it('should throw AuthenticationError when no tokens available', async () => {
      // Don't store any tokens
      await expect(refresher.getValidAccessToken()).rejects.toThrow(
        AuthenticationError
      );
    });

    it('should delegate token retrieval to token manager', async () => {
      const getTokensSpy = vi.spyOn(manager, 'getTokens');

      await manager.storeTokens({
        accessToken: 'valid-token',
        refreshToken: 'refresh-token',
        expiresAt: Date.now() + 3600000,
        scope: 'Tasks.ReadWrite User.Read',
      });

      await refresher.getValidAccessToken();

      expect(getTokensSpy).toHaveBeenCalledOnce();
    });

    it('should handle concurrent requests', async () => {
      await manager.storeTokens({
        accessToken: 'concurrent-token',
        refreshToken: 'refresh-token',
        expiresAt: Date.now() + 3600000,
        scope: 'Tasks.ReadWrite User.Read',
      });

      // Trigger multiple concurrent requests
      const promises = [
        refresher.getValidAccessToken(),
        refresher.getValidAccessToken(),
        refresher.getValidAccessToken(),
      ];

      const results = await Promise.all(promises);

      // All should get the same token
      expect(results).toEqual(['concurrent-token', 'concurrent-token', 'concurrent-token']);
    });

    it('should propagate errors from token manager', async () => {
      // Mock getTokens to throw an error
      const error = new Error('Token retrieval failed');
      vi.spyOn(manager, 'getTokens').mockRejectedValueOnce(error);

      await expect(refresher.getValidAccessToken()).rejects.toThrow(
        'Token retrieval failed'
      );
    });
  });
});
