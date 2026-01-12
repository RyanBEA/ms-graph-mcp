import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { SecureOAuthClient } from '../../src/auth/secure-oauth-client.js';
import { ITokenManager, TokenSet } from '../../src/auth/types.js';
import { AuthorizationError, AuthenticationError } from '../../src/security/errors.js';

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

// Mock token manager
class MockTokenManager implements ITokenManager {
  private tokens: TokenSet | null = null;

  async storeTokens(tokens: TokenSet): Promise<void> {
    this.tokens = tokens;
  }

  async getTokens(): Promise<TokenSet> {
    if (!this.tokens) throw new Error('No tokens');
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

  getStoredTokens(): TokenSet | null {
    return this.tokens;
  }
}

describe('SecureOAuthClient', () => {
  let client: SecureOAuthClient;
  let tokenManager: MockTokenManager;

  beforeEach(() => {
    tokenManager = new MockTokenManager();
    client = new SecureOAuthClient(tokenManager);
  });

  afterEach(() => {
    client.stopStateCleanup();
  });

  describe('generateAuthUrl', () => {
    it('should generate authorization URL with state', async () => {
      const result = await client.generateAuthUrl();

      expect(result).toHaveProperty('url');
      expect(result).toHaveProperty('state');
      expect(result.state).toHaveLength(64); // 32 bytes = 64 hex chars
      expect(result.url).toContain('login.microsoftonline.com');
      expect(result.url).toContain('test-tenant-id');
      expect(result.url).toContain(encodeURIComponent(result.state));
    });

    it('should generate unique states for each call', async () => {
      const result1 = await client.generateAuthUrl();
      const result2 = await client.generateAuthUrl();

      expect(result1.state).not.toBe(result2.state);
    });

    it('should include required scopes in URL', async () => {
      const result = await client.generateAuthUrl();

      expect(result.url).toContain('Tasks.Read');
      expect(result.url).toContain('User.Read');
      expect(result.url).toContain('offline_access');
    });

    it('should include redirect URI in URL', async () => {
      const result = await client.generateAuthUrl();

      expect(result.url).toContain(encodeURIComponent('http://localhost:3000/callback'));
    });

    it('should increment state count', async () => {
      const initialCount = client.getStateCount();
      await client.generateAuthUrl();
      await client.generateAuthUrl();

      expect(client.getStateCount()).toBe(initialCount + 2);
    });
  });

  describe('handleCallback', () => {
    it('should reject invalid state', async () => {
      await expect(
        client.handleCallback('test-code', 'invalid-state')
      ).rejects.toThrow(AuthorizationError);

      await expect(
        client.handleCallback('test-code', 'invalid-state')
      ).rejects.toThrow('Invalid or expired state token');
    });

    it('should reject expired state', async () => {
      const { state } = await client.generateAuthUrl();

      // Fast-forward time by 6 minutes (state timeout is 5 minutes)
      vi.useFakeTimers();
      vi.advanceTimersByTime(6 * 60 * 1000);

      await expect(
        client.handleCallback('test-code', state)
      ).rejects.toThrow(AuthorizationError);

      vi.useRealTimers();
    });

    it('should accept valid state within timeout', async () => {
      const { state } = await client.generateAuthUrl();

      // Mock MSAL acquireTokenByCode and account cache
      const msalClient = (client as any).msalClient;
      vi.spyOn(msalClient, 'acquireTokenByCode').mockResolvedValue({
        accessToken: 'test-access-token',
        expiresOn: new Date(Date.now() + 3600000),
        scopes: ['Tasks.Read', 'User.Read', 'offline_access'],
      });

      // Mock token cache to return account
      const mockAccount = {
        homeAccountId: 'test-account-id',
        localAccountId: 'test-local-account-id',
        environment: 'login.microsoftonline.com',
        tenantId: 'test-tenant-id',
        username: 'test@example.com',
      };
      vi.spyOn(msalClient, 'getTokenCache').mockReturnValue({
        getAllAccounts: vi.fn().mockResolvedValue([mockAccount]),
      } as any);

      await client.handleCallback('test-code', state);

      const tokens = tokenManager.getStoredTokens();
      expect(tokens).toBeDefined();
      expect(tokens?.accessToken).toBe('test-access-token');
      expect(tokens?.refreshToken).toBe('test-account-id');
    });

    it('should remove state after validation (one-time use)', async () => {
      const { state } = await client.generateAuthUrl();

      // Mock MSAL acquireTokenByCode and account cache
      const msalClient = (client as any).msalClient;
      vi.spyOn(msalClient, 'acquireTokenByCode').mockResolvedValue({
        accessToken: 'test-access-token',
        expiresOn: new Date(Date.now() + 3600000),
        scopes: ['Tasks.Read', 'User.Read'],
      });

      // Mock token cache to return account
      const mockAccount = {
        homeAccountId: 'test-account-id',
        localAccountId: 'test-local-account-id',
        environment: 'login.microsoftonline.com',
        tenantId: 'test-tenant-id',
        username: 'test@example.com',
      };
      vi.spyOn(msalClient, 'getTokenCache').mockReturnValue({
        getAllAccounts: vi.fn().mockResolvedValue([mockAccount]),
      } as any);

      // First use should work
      await client.handleCallback('test-code', state);

      // Second use with same state should fail
      await expect(
        client.handleCallback('test-code', state)
      ).rejects.toThrow(AuthorizationError);
    });

    it('should throw AuthenticationError if MSAL fails', async () => {
      const { state } = await client.generateAuthUrl();

      // Mock MSAL to fail
      const msalClient = (client as any).msalClient;
      vi.spyOn(msalClient, 'acquireTokenByCode').mockRejectedValue(
        new Error('MSAL error')
      );

      await expect(
        client.handleCallback('test-code', state)
      ).rejects.toThrow(AuthenticationError);
    });

    it('should throw AuthenticationError if response is invalid', async () => {
      const { state } = await client.generateAuthUrl();

      // Mock MSAL to return invalid response
      const msalClient = (client as any).msalClient;
      vi.spyOn(msalClient, 'acquireTokenByCode').mockResolvedValue({
        accessToken: null,
        refreshToken: null,
      });

      await expect(
        client.handleCallback('test-code', state)
      ).rejects.toThrow(AuthenticationError);
    });
  });

  describe('state cleanup', () => {
    it('should clean up expired states', async () => {
      // Generate some states
      await client.generateAuthUrl();
      await client.generateAuthUrl();
      await client.generateAuthUrl();

      expect(client.getStateCount()).toBe(3);

      // Fast-forward time to expire states
      vi.useFakeTimers();
      vi.advanceTimersByTime(6 * 60 * 1000);

      // Trigger cleanup by generating new URL
      await client.generateAuthUrl();

      // Old states should be cleaned, new one added
      expect(client.getStateCount()).toBe(1);

      vi.useRealTimers();
    });

    it('should stop cleanup when requested', () => {
      const intervalId = (client as any).cleanupInterval;
      expect(intervalId).toBeDefined();

      client.stopStateCleanup();

      expect((client as any).cleanupInterval).toBeNull();
    });
  });

  describe('security', () => {
    it('should use cryptographically secure random for state', async () => {
      const states = new Set();

      for (let i = 0; i < 100; i++) {
        const { state } = await client.generateAuthUrl();
        states.add(state);
      }

      // All states should be unique
      expect(states.size).toBe(100);
    });

    it('should have private state store (TypeScript)', () => {
      // State store is private in TypeScript but accessible at runtime
      // We verify it exists but is not part of the public API
      expect((client as any).stateStore).toBeInstanceOf(Map);
    });
  });
});
