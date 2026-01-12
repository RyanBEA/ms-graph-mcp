/**
 * Rate limiter using token bucket algorithm.
 * Enforces Microsoft Graph API rate limits (default: 60 requests/minute).
 */

import { logger } from '../security/logger.js';
import { RateLimitError } from '../security/errors.js';

export interface RateLimiterConfig {
  maxRequestsPerMinute: number;
  burstAllowance?: number; // Allow bursts up to this many requests
}

export class RateLimiter {
  private tokens: number;
  private readonly refillRate: number; // Tokens per millisecond
  private lastRefillTime: number;
  private readonly burstAllowance: number;

  constructor(config: RateLimiterConfig) {
    this.burstAllowance = config.burstAllowance ?? config.maxRequestsPerMinute;
    this.tokens = this.burstAllowance; // Start with full burst allowance
    this.refillRate = config.maxRequestsPerMinute / 60000; // Tokens per ms
    this.lastRefillTime = Date.now();

    logger.info('Rate limiter initialized', {
      maxRequestsPerMinute: config.maxRequestsPerMinute,
      burstAllowance: this.burstAllowance,
      refillRate: this.refillRate,
    });
  }

  /**
   * Check if a request can proceed. Throws RateLimitError if limit exceeded.
   * @throws {RateLimitError} When rate limit is exceeded
   */
  async checkLimit(): Promise<void> {
    this.refillTokens();

    if (this.tokens < 1) {
      const waitTimeMs = Math.ceil((1 - this.tokens) / this.refillRate);
      logger.warn('Rate limit exceeded', {
        availableTokens: this.tokens,
        waitTimeMs,
      });
      throw new RateLimitError(
        `Rate limit exceeded. Please wait ${Math.ceil(waitTimeMs / 1000)} seconds.`,
        waitTimeMs
      );
    }

    // Consume one token
    this.tokens -= 1;
    logger.debug('Rate limit check passed', {
      remainingTokens: this.tokens,
    });
  }

  /**
   * Refill tokens based on elapsed time since last refill.
   */
  private refillTokens(): void {
    const now = Date.now();
    const elapsedMs = now - this.lastRefillTime;
    const tokensToAdd = elapsedMs * this.refillRate;

    if (tokensToAdd > 0) {
      this.tokens = Math.min(this.burstAllowance, this.tokens + tokensToAdd);
      this.lastRefillTime = now;
    }
  }

  /**
   * Get current token count (for testing/monitoring).
   */
  getAvailableTokens(): number {
    this.refillTokens();
    return this.tokens;
  }

  /**
   * Reset rate limiter to full capacity (for testing).
   */
  reset(): void {
    this.tokens = this.burstAllowance;
    this.lastRefillTime = Date.now();
    logger.debug('Rate limiter reset');
  }
}
