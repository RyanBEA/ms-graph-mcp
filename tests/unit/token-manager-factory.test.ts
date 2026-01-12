import { describe, it, expect, beforeEach, vi, afterEach } from 'vitest';
import { getTokenManager, resetTokenManager } from '../../src/auth/token-manager-factory.js';
import { KeytarTokenManager } from '../../src/auth/keytar-token-manager.js';
import { OnePasswordTokenManager } from '../../src/auth/one-password-token-manager.js';
import { getConfiguration } from '../../src/config/environment.js';

// Mock the environment configuration
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

// Mock keytar
vi.mock('keytar', () => ({
  default: {
    setPassword: vi.fn(),
    getPassword: vi.fn(),
    deletePassword: vi.fn(),
  },
}));

// Mock 1Password SDK
vi.mock('@1password/sdk', () => ({
  createClient: vi.fn(() => ({
    items: {
      list: vi.fn(),
      create: vi.fn(),
      update: vi.fn(),
      delete: vi.fn(),
    },
    secrets: {
      resolve: vi.fn(),
    },
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

  it('should create KeytarTokenManager when TOKEN_STORAGE is keytar', async () => {
    vi.mocked(getConfiguration).mockReturnValue({
      AZURE_CLIENT_ID: 'test-client-id',
      AZURE_TENANT_ID: 'test-tenant-id',
      AZURE_CLIENT_SECRET: 'test-secret',
      AZURE_REDIRECT_URI: 'http://localhost:3000/callback',
      TOKEN_STORAGE: 'keytar',
      NODE_ENV: 'test',
      LOG_LEVEL: 'info',
      RATE_LIMIT_PER_MINUTE: 60,
      STATE_TIMEOUT_MINUTES: 5,
    });

    const manager = await getTokenManager();

    expect(manager).toBeInstanceOf(KeytarTokenManager);
  });

  it('should throw error for 1Password (not yet implemented)', async () => {
    vi.mocked(getConfiguration).mockReturnValue({
      AZURE_CLIENT_ID: 'test-client-id',
      AZURE_TENANT_ID: 'test-tenant-id',
      AZURE_CLIENT_SECRET: 'test-secret',
      AZURE_REDIRECT_URI: 'http://localhost:3000/callback',
      TOKEN_STORAGE: '1password',
      OP_SERVICE_ACCOUNT_TOKEN: 'ops_test_token',
      NODE_ENV: 'test',
      LOG_LEVEL: 'info',
      RATE_LIMIT_PER_MINUTE: 60,
      STATE_TIMEOUT_MINUTES: 5,
    });

    // 1Password is not yet implemented, should throw during initialize
    await expect(getTokenManager()).rejects.toThrow('1Password integration not yet implemented');
  });

  it('should return same instance on subsequent calls (singleton)', async () => {
    vi.mocked(getConfiguration).mockReturnValue({
      AZURE_CLIENT_ID: 'test-client-id',
      AZURE_TENANT_ID: 'test-tenant-id',
      AZURE_CLIENT_SECRET: 'test-secret',
      AZURE_REDIRECT_URI: 'http://localhost:3000/callback',
      TOKEN_STORAGE: 'keytar',
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
      TOKEN_STORAGE: 'keytar',
      NODE_ENV: 'test',
      LOG_LEVEL: 'info',
      RATE_LIMIT_PER_MINUTE: 60,
      STATE_TIMEOUT_MINUTES: 5,
    });

    const manager1 = await getTokenManager();
    resetTokenManager();
    const manager2 = await getTokenManager();

    expect(manager1).not.toBe(manager2);
    expect(manager1).toBeInstanceOf(KeytarTokenManager);
    expect(manager2).toBeInstanceOf(KeytarTokenManager);
  });
});
