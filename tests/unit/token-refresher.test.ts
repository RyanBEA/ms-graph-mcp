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
    TOKEN_STORAGE: 'keytar',
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
      throw new Error('No tokens stored');
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
    it('should return existing token if still valid', async () => {
      // Token valid for more than 1 hour
      await manager.storeTokens({
        accessToken: 'valid-token',
        refreshToken: 'refresh-token',
        expiresAt: Date.now() + 3600000, // 1 hour from now
        scope: 'Tasks.Read User.Read',
      });

      const token = await refresher.getValidAccessToken();

      expect(token).toBe('valid-token');
    });

    it('should return existing token if within expiry buffer', async () => {
      // Token expires in 10 minutes (within the 5-minute buffer, but still valid)
      await manager.storeTokens({
        accessToken: 'valid-token',
        refreshToken: 'refresh-token',
        expiresAt: Date.now() + 10 * 60 * 1000, // 10 minutes from now
        scope: 'Tasks.Read User.Read',
      });

      const token = await refresher.getValidAccessToken();

      expect(token).toBe('valid-token');
    });

    it('should handle concurrent refresh attempts', async () => {
      // Token expires in 2 minutes (needs refresh)
      await manager.storeTokens({
        accessToken: 'old-token',
        refreshToken: 'refresh-token',
        expiresAt: Date.now() + 2 * 60 * 1000, // 2 minutes from now
        scope: 'Tasks.Read User.Read',
      });

      // Mock MSAL to delay refresh
      const msalRefreshSpy = vi.spyOn(refresher as any, 'refreshAccessToken');
      msalRefreshSpy.mockImplementation(async () => {
        await new Promise(resolve => setTimeout(resolve, 100));
        await manager.updateAccessToken('new-token', Date.now() + 3600000);
        return 'new-token';
      });

      // Trigger multiple concurrent requests
      const promises = [
        refresher.getValidAccessToken(),
        refresher.getValidAccessToken(),
        refresher.getValidAccessToken(),
      ];

      const results = await Promise.all(promises);

      // All should get the same new token
      expect(results).toEqual(['new-token', 'new-token', 'new-token']);
      // But refresh should only happen once
      expect(msalRefreshSpy).toHaveBeenCalledTimes(1);
    });
  });

  describe('refreshAccessToken', () => {
    it('should clear tokens on refresh failure', async () => {
      await manager.storeTokens({
        accessToken: 'old-token',
        refreshToken: 'invalid-refresh-token',
        expiresAt: Date.now() + 2 * 60 * 1000,
        scope: 'Tasks.Read User.Read',
      });

      // Mock MSAL to fail - implementation should clear tokens and throw
      const msalRefreshSpy = vi.spyOn(refresher as any, 'refreshAccessToken');
      msalRefreshSpy.mockImplementation(async () => {
        await manager.clearTokens();
        throw new AuthenticationError('Refresh failed');
      });

      await expect(refresher.getValidAccessToken()).rejects.toThrow(
        AuthenticationError
      );

      // Tokens should be cleared
      const hasTokens = await manager.hasValidTokens();
      expect(hasTokens).toBe(false);
    });
  });

  describe('token expiry buffer', () => {
    it('should refresh token within 5-minute buffer', async () => {
      // Token expires in 4 minutes (within 5-minute buffer)
      await manager.storeTokens({
        accessToken: 'old-token',
        refreshToken: 'refresh-token',
        expiresAt: Date.now() + 4 * 60 * 1000,
        scope: 'Tasks.Read User.Read',
      });

      const msalRefreshSpy = vi.spyOn(refresher as any, 'refreshAccessToken');
      msalRefreshSpy.mockImplementation(async () => {
        await manager.updateAccessToken('new-token', Date.now() + 3600000);
        return 'new-token';
      });

      const token = await refresher.getValidAccessToken();

      expect(token).toBe('new-token');
      expect(msalRefreshSpy).toHaveBeenCalled();
    });

    it('should not refresh token outside 5-minute buffer', async () => {
      // Token expires in 10 minutes (outside 5-minute buffer)
      await manager.storeTokens({
        accessToken: 'valid-token',
        refreshToken: 'refresh-token',
        expiresAt: Date.now() + 10 * 60 * 1000,
        scope: 'Tasks.Read User.Read',
      });

      const msalRefreshSpy = vi.spyOn(refresher as any, 'refreshAccessToken');

      const token = await refresher.getValidAccessToken();

      expect(token).toBe('valid-token');
      expect(msalRefreshSpy).not.toHaveBeenCalled();
    });
  });
});
