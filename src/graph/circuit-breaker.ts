/**
 * Circuit breaker pattern to protect against cascading failures.
 * Opens circuit after sustained failures, preventing futile requests.
 */

import { logger } from '../security/logger.js';
import { GraphAPIError } from '../security/errors.js';

export enum CircuitState {
  CLOSED = 'CLOSED',     // Normal operation
  OPEN = 'OPEN',         // Too many failures, reject requests
  HALF_OPEN = 'HALF_OPEN' // Testing recovery
}

export interface CircuitBreakerConfig {
  failureThreshold: number;    // Number of failures before opening
  successThreshold: number;    // Successes needed to close from half-open
  timeout: number;             // Milliseconds before trying half-open
}

export class CircuitBreaker {
  private state: CircuitState = CircuitState.CLOSED;
  private failureCount: number = 0;
  private successCount: number = 0;
  private nextAttemptTime: number = 0;

  private readonly failureThreshold: number;
  private readonly successThreshold: number;
  private readonly timeout: number;

  constructor(config: CircuitBreakerConfig) {
    this.failureThreshold = config.failureThreshold;
    this.successThreshold = config.successThreshold;
    this.timeout = config.timeout;

    logger.info('Circuit breaker initialized', {
      failureThreshold: this.failureThreshold,
      successThreshold: this.successThreshold,
      timeoutMs: this.timeout,
    });
  }

  /**
   * Execute a function with circuit breaker protection.
   * @throws {GraphAPIError} When circuit is open
   */
  async execute<T>(fn: () => Promise<T>): Promise<T> {
    if (this.state === CircuitState.OPEN) {
      if (Date.now() < this.nextAttemptTime) {
        const waitSeconds = Math.ceil((this.nextAttemptTime - Date.now()) / 1000);
        logger.warn('Circuit breaker is OPEN', {
          state: this.state,
          failureCount: this.failureCount,
          waitSeconds,
        });
        throw new GraphAPIError(
          `Service temporarily unavailable. Circuit breaker is open. Retry in ${waitSeconds}s.`
        );
      }

      // Timeout expired, try half-open
      this.state = CircuitState.HALF_OPEN;
      this.successCount = 0;
      logger.info('Circuit breaker transitioning to HALF_OPEN');
    }

    try {
      const result = await fn();
      this.onSuccess();
      return result;
    } catch (error) {
      this.onFailure();
      throw error;
    }
  }

  /**
   * Record a successful request.
   */
  private onSuccess(): void {
    this.failureCount = 0;

    if (this.state === CircuitState.HALF_OPEN) {
      this.successCount++;
      logger.debug('Circuit breaker success in HALF_OPEN', {
        successCount: this.successCount,
        threshold: this.successThreshold,
      });

      if (this.successCount >= this.successThreshold) {
        this.state = CircuitState.CLOSED;
        this.successCount = 0;
        logger.info('Circuit breaker CLOSED after recovery');
      }
    }
  }

  /**
   * Record a failed request.
   */
  private onFailure(): void {
    this.failureCount++;
    this.successCount = 0;

    logger.debug('Circuit breaker failure recorded', {
      state: this.state,
      failureCount: this.failureCount,
      threshold: this.failureThreshold,
    });

    if (this.failureCount >= this.failureThreshold) {
      this.state = CircuitState.OPEN;
      this.nextAttemptTime = Date.now() + this.timeout;

      logger.warn('Circuit breaker OPENED', {
        failureCount: this.failureCount,
        nextAttemptIn: Math.ceil(this.timeout / 1000) + 's',
      });
    }
  }

  /**
   * Get current circuit state (for monitoring/testing).
   */
  getState(): CircuitState {
    if (this.state === CircuitState.OPEN && Date.now() >= this.nextAttemptTime) {
      return CircuitState.HALF_OPEN;
    }
    return this.state;
  }

  /**
   * Reset circuit breaker (for testing).
   */
  reset(): void {
    this.state = CircuitState.CLOSED;
    this.failureCount = 0;
    this.successCount = 0;
    this.nextAttemptTime = 0;
    logger.debug('Circuit breaker reset to CLOSED');
  }

  /**
   * Get current failure count (for monitoring/testing).
   */
  getFailureCount(): number {
    return this.failureCount;
  }
}
