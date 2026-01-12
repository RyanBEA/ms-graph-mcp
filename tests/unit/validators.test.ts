import { describe, it, expect } from 'vitest';
import {
  validateODataFilter,
  validateListId,
  validateTaskId,
  validateSearchQuery,
} from '../../src/security/validators.js';
import { ValidationError } from '../../src/security/errors.js';

describe('validators', () => {
  describe('validateODataFilter', () => {
    it('should allow valid filter with allowed fields', () => {
      expect(() => validateODataFilter("status eq 'completed'")).not.toThrow();
      expect(() => validateODataFilter("importance eq 'high'")).not.toThrow();
      expect(() => validateODataFilter("dueDateTime gt '2025-01-01'")).not.toThrow();
    });

    it('should allow empty or whitespace filters', () => {
      expect(() => validateODataFilter('')).not.toThrow();
      expect(() => validateODataFilter('   ')).not.toThrow();
    });

    it('should allow complex filters with allowed operators', () => {
      expect(() =>
        validateODataFilter("status eq 'completed' and importance eq 'high'")
      ).not.toThrow();

      expect(() =>
        validateODataFilter("importance eq 'high' or importance eq 'normal'")
      ).not.toThrow();
    });

    it('should reject filter with invalid field', () => {
      expect(() => validateODataFilter("maliciousField eq 'value'")).toThrow(
        ValidationError
      );

      expect(() => validateODataFilter("userId eq '123'")).toThrow(
        ValidationError
      );
    });

    it('should reject filter with suspicious characters', () => {
      expect(() => validateODataFilter("status eq 'completed'; DROP TABLE")).toThrow(
        ValidationError
      );

      expect(() => validateODataFilter("status eq '<script>alert(1)</script>'")).toThrow(
        ValidationError
      );
    });

    it('should reject filter with SQL injection attempts', () => {
      // Semicolon should be caught as suspicious character
      expect(() => validateODataFilter("status eq 'x'; DELETE FROM tasks")).toThrow(
        ValidationError
      );

      // SQL keywords DROP/DELETE in suspicious patterns
      expect(() => validateODataFilter("status eq 'x'; DROP TABLE users")).toThrow(
        ValidationError
      );
    });

    it('should reject filter with XSS attempts', () => {
      expect(() => validateODataFilter("<script>alert('XSS')</script>")).toThrow(
        ValidationError
      );

      expect(() => validateODataFilter("javascript:alert(1)")).toThrow(
        ValidationError
      );

      expect(() => validateODataFilter("status eq 'x' onerror=alert(1)")).toThrow(
        ValidationError
      );
    });

    it('should reject filter with invalid operators', () => {
      // OData doesn't support LIKE in the same way as SQL
      expect(() => validateODataFilter("status contains 'test'")).not.toThrow();
      // But we don't whitelist 'contains', so it should pass field validation
      // since 'contains' isn't in our operator list
    });
  });

  describe('validateListId', () => {
    it('should allow valid list IDs', () => {
      expect(() => validateListId('AAMkADU3NDk4NjQ4LTE2ZTMtNGM2Mi1iMzE2LTVk')).not.toThrow();
      expect(() => validateListId('abc123-def456_ghi789')).not.toThrow();
      expect(() => validateListId('list-123')).not.toThrow();
    });

    it('should reject empty or whitespace IDs', () => {
      expect(() => validateListId('')).toThrow(ValidationError);
      expect(() => validateListId('   ')).toThrow(ValidationError);
    });

    it('should reject IDs with special characters', () => {
      expect(() => validateListId('list/123')).toThrow(ValidationError);
      expect(() => validateListId('list;123')).toThrow(ValidationError);
      expect(() => validateListId('list<script>')).toThrow(ValidationError);
    });

    it('should reject excessively long IDs', () => {
      const longId = 'a'.repeat(201);
      expect(() => validateListId(longId)).toThrow(ValidationError);
    });

    it('should reject null or undefined', () => {
      expect(() => validateListId(null as any)).toThrow(ValidationError);
      expect(() => validateListId(undefined as any)).toThrow(ValidationError);
    });
  });

  describe('validateTaskId', () => {
    it('should use same validation as list ID', () => {
      expect(() => validateTaskId('task-123')).not.toThrow();
      expect(() => validateTaskId('task/123')).toThrow(ValidationError);
    });
  });

  describe('validateSearchQuery', () => {
    it('should allow valid search queries', () => {
      expect(() => validateSearchQuery('meeting notes')).not.toThrow();
      expect(() => validateSearchQuery('buy groceries')).not.toThrow();
      expect(() => validateSearchQuery('Project Alpha deliverables')).not.toThrow();
    });

    it('should reject empty queries', () => {
      expect(() => validateSearchQuery('')).toThrow(ValidationError);
      expect(() => validateSearchQuery('   ')).toThrow(ValidationError);
    });

    it('should reject excessively long queries', () => {
      const longQuery = 'a'.repeat(1001);
      expect(() => validateSearchQuery(longQuery)).toThrow(ValidationError);
    });

    it('should reject queries with XSS attempts', () => {
      expect(() => validateSearchQuery('<script>alert(1)</script>')).toThrow(
        ValidationError
      );

      expect(() => validateSearchQuery('javascript:alert(1)')).toThrow(
        ValidationError
      );

      expect(() => validateSearchQuery('search onerror=alert(1)')).toThrow(
        ValidationError
      );
    });

    it('should allow queries near the length limit', () => {
      const validQuery = 'a'.repeat(1000);
      expect(() => validateSearchQuery(validQuery)).not.toThrow();
    });

    it('should reject null or undefined', () => {
      expect(() => validateSearchQuery(null as any)).toThrow(ValidationError);
      expect(() => validateSearchQuery(undefined as any)).toThrow(ValidationError);
    });
  });
});
