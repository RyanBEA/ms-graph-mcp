import { describe, it, expect, beforeEach, vi } from 'vitest';
import { KeytarTokenManager } from '../../src/auth/keytar-token-manager.js';
import { TokenSet } from '../../src/auth/types.js';
import { TokenStorageError } from '../../src/security/errors.js';
import keytar from 'keytar';

// Mock keytar
vi.mock('keytar', () => ({
  default: {
    setPassword: vi.fn(),
    getPassword: vi.fn(),
    deletePassword: vi.fn(),
  },
}));

describe('KeytarTokenManager', () => {
  let manager: KeytarTokenManager;
  const mockTokens: TokenSet = {
    accessToken: 'test-access-token',
    refreshToken: 'test-refresh-token',
    expiresAt: Date.now() + 3600000, // 1 hour from now
    scope: 'Tasks.Read User.Read',
  };

  beforeEach(() => {
    manager = new KeytarTokenManager();
    vi.clearAllMocks();
  });

  describe('storeTokens', () => {
    it('should store all token fields in keytar', async () => {
      await manager.storeTokens(mockTokens);

      expect(keytar.setPassword).toHaveBeenCalledWith(
        'toto-mcp-server',
        'access_token',
        mockTokens.accessToken
      );
      expect(keytar.setPassword).toHaveBeenCalledWith(
        'toto-mcp-server',
        'refresh_token',
        mockTokens.refreshToken
      );
      expect(keytar.setPassword).toHaveBeenCalledWith(
        'toto-mcp-server',
        'expires_at',
        mockTokens.expiresAt.toString()
      );
      expect(keytar.setPassword).toHaveBeenCalledWith(
        'toto-mcp-server',
        'scope',
        mockTokens.scope
      );
      expect(keytar.setPassword).toHaveBeenCalledTimes(4);
    });

    it('should throw TokenStorageError on failure', async () => {
      vi.mocked(keytar.setPassword).mockRejectedValueOnce(
        new Error('Keytar error')
      );

      await expect(manager.storeTokens(mockTokens)).rejects.toThrow(
        TokenStorageError
      );
    });
  });

  describe('getTokens', () => {
    it('should retrieve all token fields from keytar', async () => {
      vi.mocked(keytar.getPassword).mockImplementation(async (service, account) => {
        if (account === 'access_token') return mockTokens.accessToken;
        if (account === 'refresh_token') return mockTokens.refreshToken;
        if (account === 'expires_at') return mockTokens.expiresAt.toString();
        if (account === 'scope') return mockTokens.scope;
        return null;
      });

      const tokens = await manager.getTokens();

      expect(tokens).toEqual(mockTokens);
      expect(keytar.getPassword).toHaveBeenCalledTimes(4);
    });

    it('should throw TokenStorageError if any field is missing', async () => {
      vi.mocked(keytar.getPassword).mockResolvedValue(null);

      await expect(manager.getTokens()).rejects.toThrow(TokenStorageError);
    });

    it('should parse expiresAt as number', async () => {
      const expiresAt = Date.now() + 3600000;
      vi.mocked(keytar.getPassword).mockImplementation(async (service, account) => {
        if (account === 'access_token') return 'token1';
        if (account === 'refresh_token') return 'token2';
        if (account === 'expires_at') return expiresAt.toString();
        if (account === 'scope') return 'scope';
        return null;
      });

      const tokens = await manager.getTokens();

      expect(typeof tokens.expiresAt).toBe('number');
      expect(tokens.expiresAt).toBe(expiresAt);
    });
  });

  describe('hasValidTokens', () => {
    it('should return true for non-expired tokens', async () => {
      vi.mocked(keytar.getPassword).mockImplementation(async (service, account) => {
        if (account === 'access_token') return mockTokens.accessToken;
        if (account === 'refresh_token') return mockTokens.refreshToken;
        if (account === 'expires_at') return (Date.now() + 3600000).toString();
        if (account === 'scope') return mockTokens.scope;
        return null;
      });

      const result = await manager.hasValidTokens();

      expect(result).toBe(true);
    });

    it('should return false for expired tokens', async () => {
      vi.mocked(keytar.getPassword).mockImplementation(async (service, account) => {
        if (account === 'access_token') return mockTokens.accessToken;
        if (account === 'refresh_token') return mockTokens.refreshToken;
        if (account === 'expires_at') return (Date.now() - 1000).toString(); // Expired
        if (account === 'scope') return mockTokens.scope;
        return null;
      });

      const result = await manager.hasValidTokens();

      expect(result).toBe(false);
    });

    it('should return false if tokens do not exist', async () => {
      vi.mocked(keytar.getPassword).mockResolvedValue(null);

      const result = await manager.hasValidTokens();

      expect(result).toBe(false);
    });
  });

  describe('clearTokens', () => {
    it('should delete all token fields from keytar', async () => {
      await manager.clearTokens();

      expect(keytar.deletePassword).toHaveBeenCalledWith(
        'toto-mcp-server',
        'access_token'
      );
      expect(keytar.deletePassword).toHaveBeenCalledWith(
        'toto-mcp-server',
        'refresh_token'
      );
      expect(keytar.deletePassword).toHaveBeenCalledWith(
        'toto-mcp-server',
        'expires_at'
      );
      expect(keytar.deletePassword).toHaveBeenCalledWith(
        'toto-mcp-server',
        'scope'
      );
      expect(keytar.deletePassword).toHaveBeenCalledTimes(4);
    });

    it('should throw TokenStorageError on failure', async () => {
      vi.mocked(keytar.deletePassword).mockRejectedValueOnce(
        new Error('Delete failed')
      );

      await expect(manager.clearTokens()).rejects.toThrow(TokenStorageError);
    });
  });

  describe('updateAccessToken', () => {
    it('should update only access token and expiry', async () => {
      // Mock getTokens to return existing tokens
      vi.mocked(keytar.getPassword).mockImplementation(async (service, account) => {
        if (account === 'access_token') return mockTokens.accessToken;
        if (account === 'refresh_token') return mockTokens.refreshToken;
        if (account === 'expires_at') return mockTokens.expiresAt.toString();
        if (account === 'scope') return mockTokens.scope;
        return null;
      });

      const newAccessToken = 'new-access-token';
      const newExpiresAt = Date.now() + 7200000;

      await manager.updateAccessToken(newAccessToken, newExpiresAt);

      // Should call setPassword with updated values
      expect(keytar.setPassword).toHaveBeenCalledWith(
        'toto-mcp-server',
        'access_token',
        newAccessToken
      );
      expect(keytar.setPassword).toHaveBeenCalledWith(
        'toto-mcp-server',
        'expires_at',
        newExpiresAt.toString()
      );
      // Refresh token should remain unchanged
      expect(keytar.setPassword).toHaveBeenCalledWith(
        'toto-mcp-server',
        'refresh_token',
        mockTokens.refreshToken
      );
    });
  });
});
