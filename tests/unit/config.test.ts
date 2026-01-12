import { describe, it, expect, beforeEach } from 'vitest';
import { loadConfiguration } from '../../src/config/environment.js';

describe('Configuration', () => {
  const originalEnv = process.env;

  beforeEach(() => {
    // Reset env for each test
    process.env = { ...originalEnv };
  });

  afterEach(() => {
    process.env = originalEnv;
  });

  it('should load valid configuration', () => {
    process.env.AZURE_CLIENT_ID = 'test-client-id';
    process.env.AZURE_TENANT_ID = 'test-tenant-id';
    process.env.AZURE_CLIENT_SECRET = 'test-secret';
    process.env.AZURE_REDIRECT_URI = 'http://localhost:3000/callback';
    process.env.TOKEN_STORAGE = 'keytar';

    const config = loadConfiguration();
    expect(config.AZURE_CLIENT_ID).toBe('test-client-id');
    expect(config.AZURE_TENANT_ID).toBe('test-tenant-id');
    expect(config.TOKEN_STORAGE).toBe('keytar');
  });

  it('should use default values for optional fields', () => {
    process.env.AZURE_CLIENT_ID = 'test-client-id';
    process.env.AZURE_TENANT_ID = 'test-tenant-id';
    process.env.AZURE_CLIENT_SECRET = 'test-secret';
    process.env.AZURE_REDIRECT_URI = 'http://localhost:3000/callback';
    // Clear NODE_ENV to test default
    delete process.env.NODE_ENV;

    const config = loadConfiguration();
    expect(config.TOKEN_STORAGE).toBe('keytar');
    expect(config.NODE_ENV).toBe('development');
    expect(config.LOG_LEVEL).toBe('info');
    expect(config.RATE_LIMIT_PER_MINUTE).toBe(60);
    expect(config.STATE_TIMEOUT_MINUTES).toBe(5);
  });

  it('should reject missing required fields', () => {
    expect(() => loadConfiguration()).toThrow();
  });

  it('should reject invalid URL for redirect URI', () => {
    process.env.AZURE_CLIENT_ID = 'test-client-id';
    process.env.AZURE_TENANT_ID = 'test-tenant-id';
    process.env.AZURE_CLIENT_SECRET = 'test-secret';
    process.env.AZURE_REDIRECT_URI = 'not-a-url';

    expect(() => loadConfiguration()).toThrow(/Must be valid URL/);
  });

  it('should enforce 1Password token when selected', () => {
    process.env.AZURE_CLIENT_ID = 'test-client-id';
    process.env.AZURE_TENANT_ID = 'test-tenant-id';
    process.env.AZURE_CLIENT_SECRET = 'test-secret';
    process.env.AZURE_REDIRECT_URI = 'http://localhost:3000/callback';
    process.env.TOKEN_STORAGE = '1password';
    // Missing OP_SERVICE_ACCOUNT_TOKEN

    expect(() => loadConfiguration()).toThrow(/OP_SERVICE_ACCOUNT_TOKEN/);
  });

  it('should accept valid 1Password configuration', () => {
    process.env.AZURE_CLIENT_ID = 'test-client-id';
    process.env.AZURE_TENANT_ID = 'test-tenant-id';
    process.env.AZURE_CLIENT_SECRET = 'test-secret';
    process.env.AZURE_REDIRECT_URI = 'http://localhost:3000/callback';
    process.env.TOKEN_STORAGE = '1password';
    process.env.OP_SERVICE_ACCOUNT_TOKEN = 'ops_test_token';

    const config = loadConfiguration();
    expect(config.TOKEN_STORAGE).toBe('1password');
    expect(config.OP_SERVICE_ACCOUNT_TOKEN).toBe('ops_test_token');
  });
});
