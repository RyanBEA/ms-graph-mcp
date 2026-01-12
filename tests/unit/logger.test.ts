import { describe, it, expect } from 'vitest';
import { redactSensitiveData } from '../../src/security/logger.js';

describe('Logger Security - Redaction', () => {
  it('should redact access_token', () => {
    const data = { access_token: 'secret123' };
    const redacted = redactSensitiveData(data);
    expect(redacted.access_token).toBe('[REDACTED]');
  });

  it('should redact refresh_token', () => {
    const data = { refresh_token: 'secret456' };
    const redacted = redactSensitiveData(data);
    expect(redacted.refresh_token).toBe('[REDACTED]');
  });

  it('should redact client_secret', () => {
    const data = { client_secret: 'secret789' };
    const redacted = redactSensitiveData(data);
    expect(redacted.client_secret).toBe('[REDACTED]');
  });

  it('should redact password fields', () => {
    const data = { password: 'mypassword' };
    const redacted = redactSensitiveData(data);
    expect(redacted.password).toBe('[REDACTED]');
  });

  it('should redact authorization headers', () => {
    const data = { authorization: 'Bearer token123' };
    const redacted = redactSensitiveData(data);
    expect(redacted.authorization).toBe('[REDACTED]');
  });

  it('should handle nested objects', () => {
    const data = {
      user: {
        name: 'John',
        access_token: 'secret789'
      }
    };
    const redacted = redactSensitiveData(data);
    expect(redacted.user.name).toBe('John');
    expect(redacted.user.access_token).toBe('[REDACTED]');
  });

  it('should not redact non-sensitive fields', () => {
    const data = { userId: '12345', action: 'login' };
    const redacted = redactSensitiveData(data);
    expect(redacted.userId).toBe('12345');
    expect(redacted.action).toBe('login');
  });

  it('should redact fields case-insensitively', () => {
    const data = {
      Access_Token: 'secret1',
      CLIENT_SECRET: 'secret2',
      Password: 'secret3'
    };
    const redacted = redactSensitiveData(data);
    expect(redacted.Access_Token).toBe('[REDACTED]');
    expect(redacted.CLIENT_SECRET).toBe('[REDACTED]');
    expect(redacted.Password).toBe('[REDACTED]');
  });

  it('should not mutate the original object', () => {
    const data = { access_token: 'secret123', userId: '456' };
    const redacted = redactSensitiveData(data);

    // Original should be unchanged
    expect(data.access_token).toBe('secret123');
    // Redacted should be redacted
    expect(redacted.access_token).toBe('[REDACTED]');
    // Non-sensitive should be same in both
    expect(data.userId).toBe('456');
    expect(redacted.userId).toBe('456');
  });
});
