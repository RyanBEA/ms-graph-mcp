import { describe, it, expect, beforeEach, vi } from 'vitest';
import { RateLimiter } from '../../src/graph/rate-limiter.js';
import { RateLimitError } from '../../src/security/errors.js';

describe('RateLimiter', () => {
  beforeEach(() => {
    vi.useFakeTimers();
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  describe('constructor', () => {
    it('should initialize with correct token count', () => {
      const limiter = new RateLimiter({ maxRequestsPerMinute: 60 });
      expect(limiter.getAvailableTokens()).toBe(60);
    });

    it('should respect custom burst allowance', () => {
      const limiter = new RateLimiter({
        maxRequestsPerMinute: 60,
        burstAllowance: 30,
      });
      expect(limiter.getAvailableTokens()).toBe(30);
    });

    it('should use maxRequestsPerMinute as default burst allowance', () => {
      const limiter = new RateLimiter({ maxRequestsPerMinute: 100 });
      expect(limiter.getAvailableTokens()).toBe(100);
    });
  });

  describe('checkLimit', () => {
    it('should allow requests within rate limit', async () => {
      const limiter = new RateLimiter({ maxRequestsPerMinute: 60 });

      await expect(limiter.checkLimit()).resolves.toBeUndefined();
      await expect(limiter.checkLimit()).resolves.toBeUndefined();
      await expect(limiter.checkLimit()).resolves.toBeUndefined();
    });

    it('should consume tokens on each request', async () => {
      const limiter = new RateLimiter({ maxRequestsPerMinute: 60 });

      const initialTokens = limiter.getAvailableTokens();
      await limiter.checkLimit();
      expect(limiter.getAvailableTokens()).toBe(initialTokens - 1);
    });

    it('should throw RateLimitError when limit exceeded', async () => {
      const limiter = new RateLimiter({
        maxRequestsPerMinute: 2,
        burstAllowance: 2,
      });

      // Use up tokens
      await limiter.checkLimit();
      await limiter.checkLimit();

      // Next request should fail
      await expect(limiter.checkLimit()).rejects.toThrow(RateLimitError);
    });

    it('should include retry time in error', async () => {
      const limiter = new RateLimiter({
        maxRequestsPerMinute: 1,
        burstAllowance: 1,
      });

      await limiter.checkLimit();

      try {
        await limiter.checkLimit();
        expect.fail('Should have thrown RateLimitError');
      } catch (error) {
        expect(error).toBeInstanceOf(RateLimitError);
        expect((error as RateLimitError).retryAfterMs).toBeGreaterThan(0);
      }
    });
  });

  describe('token refill', () => {
    it('should refill tokens over time', async () => {
      const limiter = new RateLimiter({
        maxRequestsPerMinute: 60,
        burstAllowance: 60,
      });

      // Use up some tokens
      await limiter.checkLimit();
      await limiter.checkLimit();
      expect(limiter.getAvailableTokens()).toBe(58);

      // Advance time by 1 second (should refill 1 token)
      vi.advanceTimersByTime(1000);
      expect(Math.floor(limiter.getAvailableTokens())).toBe(59);

      // Advance time by 1 more second (should refill 1 more token)
      vi.advanceTimersByTime(1000);
      expect(Math.floor(limiter.getAvailableTokens())).toBe(60);
    });

    it('should not exceed burst allowance when refilling', async () => {
      const limiter = new RateLimiter({
        maxRequestsPerMinute: 60,
        burstAllowance: 30,
      });

      // Wait for a long time
      vi.advanceTimersByTime(60000); // 1 minute

      // Should not exceed burst allowance
      expect(limiter.getAvailableTokens()).toBeLessThanOrEqual(30);
    });

    it('should allow requests after refill', async () => {
      const limiter = new RateLimiter({
        maxRequestsPerMinute: 60,
        burstAllowance: 2,
      });

      // Use up all tokens
      await limiter.checkLimit();
      await limiter.checkLimit();

      // Should fail
      await expect(limiter.checkLimit()).rejects.toThrow(RateLimitError);

      // Advance time by 2 seconds (should refill 2 tokens)
      vi.advanceTimersByTime(2000);

      // Should succeed now
      await expect(limiter.checkLimit()).resolves.toBeUndefined();
    });
  });

  describe('reset', () => {
    it('should reset tokens to full capacity', async () => {
      const limiter = new RateLimiter({
        maxRequestsPerMinute: 60,
        burstAllowance: 30,
      });

      // Use some tokens
      await limiter.checkLimit();
      await limiter.checkLimit();
      await limiter.checkLimit();

      expect(limiter.getAvailableTokens()).toBe(27);

      // Reset
      limiter.reset();

      expect(limiter.getAvailableTokens()).toBe(30);
    });
  });

  describe('concurrent requests', () => {
    it('should handle concurrent requests correctly', async () => {
      const limiter = new RateLimiter({
        maxRequestsPerMinute: 10,
        burstAllowance: 10,
      });

      // Make 10 concurrent requests
      const requests = Array(10)
        .fill(null)
        .map(() => limiter.checkLimit());

      await expect(Promise.all(requests)).resolves.toBeDefined();

      // 11th request should fail
      await expect(limiter.checkLimit()).rejects.toThrow(RateLimitError);
    });
  });

  describe('different rate limits', () => {
    it('should work with low rate limit (1 req/min)', async () => {
      const limiter = new RateLimiter({
        maxRequestsPerMinute: 1,
        burstAllowance: 1,
      });

      await limiter.checkLimit();
      await expect(limiter.checkLimit()).rejects.toThrow(RateLimitError);

      // Wait 60 seconds
      vi.advanceTimersByTime(60000);

      await expect(limiter.checkLimit()).resolves.toBeUndefined();
    });

    it('should work with high rate limit (1000 req/min)', async () => {
      const limiter = new RateLimiter({
        maxRequestsPerMinute: 1000,
        burstAllowance: 1000,
      });

      // Make 1000 requests
      const requests = Array(1000)
        .fill(null)
        .map(() => limiter.checkLimit());

      await expect(Promise.all(requests)).resolves.toBeDefined();

      // 1001st should fail
      await expect(limiter.checkLimit()).rejects.toThrow(RateLimitError);
    });
  });
});
