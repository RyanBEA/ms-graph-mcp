import { describe, it, expect } from 'vitest';
import {
  AppError,
  AuthenticationError,
  AuthorizationError,
  ValidationError,
  RateLimitError,
  GraphAPIError,
  TokenStorageError,
  ConfigurationError,
} from '../../src/security/errors.js';

describe('Error Classes', () => {
  it('should create AuthenticationError with correct properties', () => {
    const error = new AuthenticationError('Custom auth error');

    expect(error).toBeInstanceOf(AppError);
    expect(error).toBeInstanceOf(AuthenticationError);
    expect(error.message).toBe('Custom auth error');
    expect(error.statusCode).toBe(401);
    expect(error.isOperational).toBe(true);
    expect(error.name).toBe('AuthenticationError');
  });

  it('should use default message for AuthenticationError', () => {
    const error = new AuthenticationError();
    expect(error.message).toBe('Authentication required');
  });

  it('should create AuthorizationError', () => {
    const error = new AuthorizationError('Access denied');

    expect(error).toBeInstanceOf(AppError);
    expect(error.statusCode).toBe(403);
    expect(error.isOperational).toBe(true);
  });

  it('should create ValidationError', () => {
    const error = new ValidationError('Invalid data');

    expect(error).toBeInstanceOf(AppError);
    expect(error.statusCode).toBe(400);
    expect(error.isOperational).toBe(true);
  });

  it('should create RateLimitError', () => {
    const error = new RateLimitError('Too many requests');

    expect(error).toBeInstanceOf(AppError);
    expect(error.statusCode).toBe(429);
    expect(error.isOperational).toBe(true);
  });

  it('should create GraphAPIError as non-operational', () => {
    const error = new GraphAPIError('Graph API failed');

    expect(error).toBeInstanceOf(AppError);
    expect(error.statusCode).toBe(502);
    expect(error.isOperational).toBe(false); // External service errors are not operational
  });

  it('should create TokenStorageError', () => {
    const error = new TokenStorageError('Storage failed');

    expect(error).toBeInstanceOf(AppError);
    expect(error.statusCode).toBe(500);
    expect(error.isOperational).toBe(false);
  });

  it('should create ConfigurationError', () => {
    const error = new ConfigurationError('Invalid config');

    expect(error).toBeInstanceOf(AppError);
    expect(error.statusCode).toBe(500);
    expect(error.isOperational).toBe(false);
  });

  it('should sanitize errors via toJSON', () => {
    const error = new AuthenticationError('Sensitive data in error');
    const json = error.toJSON();

    expect(json.name).toBe('AuthenticationError');
    expect(json.message).toBe('Sensitive data in error');
    expect(json).not.toHaveProperty('stack'); // Stack traces never exposed
    expect(json).not.toHaveProperty('statusCode'); // Implementation details not exposed
  });

  it('should have stack trace for debugging', () => {
    const error = new ValidationError('Test error');
    expect(error.stack).toBeDefined();
    expect(error.stack).toContain('ValidationError');
  });
});
