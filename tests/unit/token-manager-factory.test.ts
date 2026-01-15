import { describe, it, expect, beforeEach, vi, afterEach } from 'vitest';
import { getTokenManager, resetTokenManager } from '../../src/auth/token-manager-factory.js';
import { MsalTokenManager } from '../../src/auth/msal-token-manager.js';
import { FileTokenManager } from '../../src/auth/file-token-manager.js';
import { getConfiguration } from '../../src/config/environment.js';

// Mock the environment configuration
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

describe('Token Manager Factory', () => {
  beforeEach(() => {
    resetTokenManager();
    vi.clearAllMocks();
  });

  afterEach(() => {
    resetTokenManager();
  });

  it('should create MsalTokenManager when TOKEN_STORAGE is msal', async () => {
    vi.mocked(getConfiguration).mockReturnValue({
      AZURE_CLIENT_ID: 'test-client-id',
      AZURE_TENANT_ID: 'test-tenant-id',
      AZURE_CLIENT_SECRET: 'test-secret',
      AZURE_REDIRECT_URI: 'http://localhost:3000/callback',
      TOKEN_STORAGE: 'msal',
      NODE_ENV: 'test',
      LOG_LEVEL: 'info',
      RATE_LIMIT_PER_MINUTE: 60,
      STATE_TIMEOUT_MINUTES: 5,
    });

    const manager = await getTokenManager();

    expect(manager).toBeInstanceOf(MsalTokenManager);
  });

  it('should create FileTokenManager when TOKEN_STORAGE is file', async () => {
    vi.mocked(getConfiguration).mockReturnValue({
      AZURE_CLIENT_ID: 'test-client-id',
      AZURE_TENANT_ID: 'test-tenant-id',
      AZURE_CLIENT_SECRET: 'test-secret',
      AZURE_REDIRECT_URI: 'http://localhost:3000/callback',
      TOKEN_STORAGE: 'file',
      NODE_ENV: 'test',
      LOG_LEVEL: 'info',
      RATE_LIMIT_PER_MINUTE: 60,
      STATE_TIMEOUT_MINUTES: 5,
    });

    const manager = await getTokenManager();

    expect(manager).toBeInstanceOf(FileTokenManager);
  });

  it('should return same instance on subsequent calls (singleton)', async () => {
    vi.mocked(getConfiguration).mockReturnValue({
      AZURE_CLIENT_ID: 'test-client-id',
      AZURE_TENANT_ID: 'test-tenant-id',
      AZURE_CLIENT_SECRET: 'test-secret',
      AZURE_REDIRECT_URI: 'http://localhost:3000/callback',
      TOKEN_STORAGE: 'msal',
      NODE_ENV: 'test',
      LOG_LEVEL: 'info',
      RATE_LIMIT_PER_MINUTE: 60,
      STATE_TIMEOUT_MINUTES: 5,
    });

    const manager1 = await getTokenManager();
    const manager2 = await getTokenManager();

    expect(manager1).toBe(manager2);
  });

  it('should create new instance after reset', async () => {
    vi.mocked(getConfiguration).mockReturnValue({
      AZURE_CLIENT_ID: 'test-client-id',
      AZURE_TENANT_ID: 'test-tenant-id',
      AZURE_CLIENT_SECRET: 'test-secret',
      AZURE_REDIRECT_URI: 'http://localhost:3000/callback',
      TOKEN_STORAGE: 'msal',
      NODE_ENV: 'test',
      LOG_LEVEL: 'info',
      RATE_LIMIT_PER_MINUTE: 60,
      STATE_TIMEOUT_MINUTES: 5,
    });

    const manager1 = await getTokenManager();
    resetTokenManager();
    const manager2 = await getTokenManager();

    expect(manager1).not.toBe(manager2);
    expect(manager1).toBeInstanceOf(MsalTokenManager);
    expect(manager2).toBeInstanceOf(MsalTokenManager);
  });
});
