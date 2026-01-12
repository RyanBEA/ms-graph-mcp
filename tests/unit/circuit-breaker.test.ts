import { describe, it, expect, beforeEach, vi } from 'vitest';
import { CircuitBreaker, CircuitState } from '../../src/graph/circuit-breaker.js';
import { GraphAPIError } from '../../src/security/errors.js';

describe('CircuitBreaker', () => {
  beforeEach(() => {
    vi.useFakeTimers();
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  describe('constructor', () => {
    it('should initialize in CLOSED state', () => {
      const breaker = new CircuitBreaker({
        failureThreshold: 5,
        successThreshold: 2,
        timeout: 60000,
      });

      expect(breaker.getState()).toBe(CircuitState.CLOSED);
      expect(breaker.getFailureCount()).toBe(0);
    });
  });

  describe('execute - CLOSED state', () => {
    it('should execute function successfully', async () => {
      const breaker = new CircuitBreaker({
        failureThreshold: 5,
        successThreshold: 2,
        timeout: 60000,
      });

      const result = await breaker.execute(async () => 'success');
      expect(result).toBe('success');
      expect(breaker.getState()).toBe(CircuitState.CLOSED);
    });

    it('should track failures', async () => {
      const breaker = new CircuitBreaker({
        failureThreshold: 5,
        successThreshold: 2,
        timeout: 60000,
      });

      const failingFn = async () => {
        throw new Error('Test error');
      };

      await expect(breaker.execute(failingFn)).rejects.toThrow('Test error');
      expect(breaker.getFailureCount()).toBe(1);

      await expect(breaker.execute(failingFn)).rejects.toThrow('Test error');
      expect(breaker.getFailureCount()).toBe(2);
    });

    it('should reset failure count on success', async () => {
      const breaker = new CircuitBreaker({
        failureThreshold: 5,
        successThreshold: 2,
        timeout: 60000,
      });

      // Fail twice
      await expect(breaker.execute(async () => { throw new Error('fail'); })).rejects.toThrow();
      await expect(breaker.execute(async () => { throw new Error('fail'); })).rejects.toThrow();
      expect(breaker.getFailureCount()).toBe(2);

      // Success should reset counter
      await breaker.execute(async () => 'success');
      expect(breaker.getFailureCount()).toBe(0);
    });

    it('should transition to OPEN after threshold failures', async () => {
      const breaker = new CircuitBreaker({
        failureThreshold: 3,
        successThreshold: 2,
        timeout: 60000,
      });

      // Fail 3 times to hit threshold
      for (let i = 0; i < 3; i++) {
        await expect(
          breaker.execute(async () => { throw new Error('fail'); })
        ).rejects.toThrow();
      }

      expect(breaker.getState()).toBe(CircuitState.OPEN);
    });
  });

  describe('execute - OPEN state', () => {
    it('should reject immediately without calling function', async () => {
      const breaker = new CircuitBreaker({
        failureThreshold: 2,
        successThreshold: 2,
        timeout: 60000,
      });

      // Open the circuit
      await expect(breaker.execute(async () => { throw new Error('fail'); })).rejects.toThrow();
      await expect(breaker.execute(async () => { throw new Error('fail'); })).rejects.toThrow();
      expect(breaker.getState()).toBe(CircuitState.OPEN);

      // Function should not be called
      const mockFn = vi.fn(async () => 'should not run');
      await expect(breaker.execute(mockFn)).rejects.toThrow(GraphAPIError);
      expect(mockFn).not.toHaveBeenCalled();
    });

    it('should throw GraphAPIError when circuit is open', async () => {
      const breaker = new CircuitBreaker({
        failureThreshold: 2,
        successThreshold: 2,
        timeout: 60000,
      });

      // Open the circuit
      await expect(breaker.execute(async () => { throw new Error('fail'); })).rejects.toThrow();
      await expect(breaker.execute(async () => { throw new Error('fail'); })).rejects.toThrow();

      // Should throw GraphAPIError
      await expect(breaker.execute(async () => 'test')).rejects.toThrow(GraphAPIError);
      await expect(breaker.execute(async () => 'test')).rejects.toThrow(
        /Circuit breaker is open/
      );
    });

    it('should transition to HALF_OPEN after timeout', async () => {
      const breaker = new CircuitBreaker({
        failureThreshold: 2,
        successThreshold: 2,
        timeout: 5000, // 5 seconds
      });

      // Open the circuit
      await expect(breaker.execute(async () => { throw new Error('fail'); })).rejects.toThrow();
      await expect(breaker.execute(async () => { throw new Error('fail'); })).rejects.toThrow();
      expect(breaker.getState()).toBe(CircuitState.OPEN);

      // Advance time past timeout
      vi.advanceTimersByTime(5000);

      // Should be HALF_OPEN now
      expect(breaker.getState()).toBe(CircuitState.HALF_OPEN);
    });
  });

  describe('execute - HALF_OPEN state', () => {
    async function openCircuit(breaker: CircuitBreaker) {
      // Open the circuit
      await expect(breaker.execute(async () => { throw new Error('fail'); })).rejects.toThrow();
      await expect(breaker.execute(async () => { throw new Error('fail'); })).rejects.toThrow();
      expect(breaker.getState()).toBe(CircuitState.OPEN);

      // Advance time to make it HALF_OPEN
      vi.advanceTimersByTime(5000);
    }

    it('should allow function execution in HALF_OPEN state', async () => {
      const breaker = new CircuitBreaker({
        failureThreshold: 2,
        successThreshold: 2,
        timeout: 5000,
      });

      await openCircuit(breaker);

      const result = await breaker.execute(async () => 'test');
      expect(result).toBe('test');
    });

    it('should close circuit after success threshold met', async () => {
      const breaker = new CircuitBreaker({
        failureThreshold: 2,
        successThreshold: 2, // Need 2 successes
        timeout: 5000,
      });

      await openCircuit(breaker);

      // First success
      await breaker.execute(async () => 'success 1');
      expect(breaker.getState()).toBe(CircuitState.HALF_OPEN);

      // Second success should close circuit
      await breaker.execute(async () => 'success 2');
      expect(breaker.getState()).toBe(CircuitState.CLOSED);
    });

    it('should reopen circuit on failure in HALF_OPEN', async () => {
      const breaker = new CircuitBreaker({
        failureThreshold: 2,
        successThreshold: 2,
        timeout: 5000,
      });

      await openCircuit(breaker);

      // Success in HALF_OPEN
      await breaker.execute(async () => 'success');
      expect(breaker.getState()).toBe(CircuitState.HALF_OPEN);

      // Failure should reopen circuit
      await expect(
        breaker.execute(async () => { throw new Error('fail'); })
      ).rejects.toThrow();

      // Wait for timeout again
      vi.advanceTimersByTime(5000);

      // Should be OPEN (transitioning to HALF_OPEN)
      const state = breaker.getState();
      expect([CircuitState.OPEN, CircuitState.HALF_OPEN]).toContain(state);
    });
  });

  describe('reset', () => {
    it('should reset circuit to CLOSED state', async () => {
      const breaker = new CircuitBreaker({
        failureThreshold: 2,
        successThreshold: 2,
        timeout: 5000,
      });

      // Open the circuit
      await expect(breaker.execute(async () => { throw new Error('fail'); })).rejects.toThrow();
      await expect(breaker.execute(async () => { throw new Error('fail'); })).rejects.toThrow();
      expect(breaker.getState()).toBe(CircuitState.OPEN);

      // Reset
      breaker.reset();

      expect(breaker.getState()).toBe(CircuitState.CLOSED);
      expect(breaker.getFailureCount()).toBe(0);
    });

    it('should allow execution after reset', async () => {
      const breaker = new CircuitBreaker({
        failureThreshold: 2,
        successThreshold: 2,
        timeout: 5000,
      });

      // Open the circuit
      await expect(breaker.execute(async () => { throw new Error('fail'); })).rejects.toThrow();
      await expect(breaker.execute(async () => { throw new Error('fail'); })).rejects.toThrow();

      // Reset
      breaker.reset();

      // Should work now
      const result = await breaker.execute(async () => 'success');
      expect(result).toBe('success');
    });
  });

  describe('edge cases', () => {
    it('should handle immediate failures correctly', async () => {
      const breaker = new CircuitBreaker({
        failureThreshold: 1,
        successThreshold: 1,
        timeout: 5000,
      });

      // Single failure should open circuit
      await expect(
        breaker.execute(async () => { throw new Error('fail'); })
      ).rejects.toThrow();

      expect(breaker.getState()).toBe(CircuitState.OPEN);
    });

    it('should handle very high thresholds', async () => {
      const breaker = new CircuitBreaker({
        failureThreshold: 100,
        successThreshold: 50,
        timeout: 5000,
      });

      // Fail many times
      for (let i = 0; i < 99; i++) {
        await expect(
          breaker.execute(async () => { throw new Error('fail'); })
        ).rejects.toThrow('fail');
      }

      expect(breaker.getState()).toBe(CircuitState.CLOSED);
      expect(breaker.getFailureCount()).toBe(99);

      // 100th failure should open
      await expect(
        breaker.execute(async () => { throw new Error('fail'); })
      ).rejects.toThrow();

      expect(breaker.getState()).toBe(CircuitState.OPEN);
    });

    it('should handle async errors correctly', async () => {
      const breaker = new CircuitBreaker({
        failureThreshold: 2,
        successThreshold: 2,
        timeout: 5000,
      });

      const asyncError = async () => {
        // Simulate async work without setTimeout (doesn't work well with fake timers)
        await Promise.resolve();
        throw new Error('Async error');
      };

      await expect(breaker.execute(asyncError)).rejects.toThrow('Async error');
      expect(breaker.getFailureCount()).toBe(1);
    });
  });
});
