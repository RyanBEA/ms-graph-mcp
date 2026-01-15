/**
 * Secure Microsoft Graph API client.
 * Includes rate limiting, circuit breaker, retry logic, and input validation.
 */

import { logger } from '../security/logger.js';
import { GraphAPIError, AuthenticationError } from '../security/errors.js';
import { RateLimiter, RateLimiterConfig } from './rate-limiter.js';
import { CircuitBreaker, CircuitBreakerConfig, CircuitState } from './circuit-breaker.js';
import { TokenRefresher } from '../auth/token-refresher.js';

export interface GraphClientConfig {
  rateLimiter?: RateLimiterConfig;
  circuitBreaker?: CircuitBreakerConfig;
  maxRetries?: number;
  baseRetryDelayMs?: number;
}

export interface GraphResponse<T> {
  value?: T[];
  '@odata.nextLink'?: string;
  '@odata.count'?: number;
}

/**
 * Secure Microsoft Graph API client with comprehensive error handling.
 */
export class SecureGraphClient {
  private readonly baseUrl = 'https://graph.microsoft.com/v1.0/';  // Trailing slash is critical!
  private readonly rateLimiter: RateLimiter;
  private readonly circuitBreaker: CircuitBreaker;
  private readonly tokenRefresher: TokenRefresher;
  private readonly maxRetries: number;
  private readonly baseRetryDelayMs: number;

  constructor(
    tokenRefresher: TokenRefresher,
    config?: GraphClientConfig
  ) {
    this.tokenRefresher = tokenRefresher;
    this.maxRetries = config?.maxRetries ?? 3;
    this.baseRetryDelayMs = config?.baseRetryDelayMs ?? 1000;

    // Initialize rate limiter (default: 60 req/min)
    this.rateLimiter = new RateLimiter(
      config?.rateLimiter ?? { maxRequestsPerMinute: 60 }
    );

    // Initialize circuit breaker
    this.circuitBreaker = new CircuitBreaker(
      config?.circuitBreaker ?? {
        failureThreshold: 5,
        successThreshold: 2,
        timeout: 60000, // 1 minute
      }
    );

    logger.info('SecureGraphClient initialized', {
      baseUrl: this.baseUrl,
      maxRetries: this.maxRetries,
      baseRetryDelayMs: this.baseRetryDelayMs,
    });
  }

  /**
   * Make a GET request to Microsoft Graph API.
   * Includes automatic token refresh, rate limiting, retries, and circuit breaker.
   *
   * @param endpoint - API endpoint (e.g., '/me/todo/lists')
   * @param options - Request options
   * @returns Response data
   */
  async get<T = any>(
    endpoint: string,
    options?: {
      queryParams?: Record<string, string>;
      headers?: Record<string, string>;
    }
  ): Promise<GraphResponse<T>> {
    // Check rate limit
    await this.rateLimiter.checkLimit();

    // Build URL with query params
    const url = this.buildUrl(endpoint, options?.queryParams);

    // Execute with circuit breaker and retry logic
    return this.circuitBreaker.execute(async () => {
      return this.executeWithRetry<GraphResponse<T>>(async () => {
        // Get fresh access token
        const accessToken = await this.tokenRefresher.getValidAccessToken();

        // Log the request (sanitized - endpoint only, no full URL)
        logger.debug('Making Graph API request', {
          endpoint: endpoint,
          hasQueryParams: !!options?.queryParams
        });

        // Make request
        const response = await fetch(url, {
          method: 'GET',
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json',
            ...options?.headers,
          },
        });

        return this.handleResponse<GraphResponse<T>>(response, url, endpoint);
      });
    });
  }

  /**
   * Make a POST request to Microsoft Graph API.
   * Used for creating resources.
   */
  async post<T = any>(
    endpoint: string,
    body: Record<string, any>,
    options?: { headers?: Record<string, string> }
  ): Promise<T> {
    await this.rateLimiter.checkLimit();
    const url = this.buildUrl(endpoint);

    return this.circuitBreaker.execute(async () => {
      return this.executeWithRetry<T>(async () => {
        const accessToken = await this.tokenRefresher.getValidAccessToken();

        logger.debug('Making Graph API POST request', {
          endpoint: endpoint,
        });

        const response = await fetch(url, {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json',
            ...options?.headers,
          },
          body: JSON.stringify(body),
        });

        return this.handleResponse<T>(response, url, endpoint);
      });
    });
  }

  /**
   * Make a PATCH request to Microsoft Graph API.
   * Used for updating resources.
   */
  async patch<T = any>(
    endpoint: string,
    body: Record<string, any>,
    options?: { headers?: Record<string, string> }
  ): Promise<T> {
    await this.rateLimiter.checkLimit();
    const url = this.buildUrl(endpoint);

    return this.circuitBreaker.execute(async () => {
      return this.executeWithRetry<T>(async () => {
        const accessToken = await this.tokenRefresher.getValidAccessToken();

        logger.debug('Making Graph API PATCH request', {
          endpoint: endpoint,
        });

        const response = await fetch(url, {
          method: 'PATCH',
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json',
            ...options?.headers,
          },
          body: JSON.stringify(body),
        });

        // PATCH may return 204 No Content
        if (response.status === 204) {
          return {} as T;
        }

        return this.handleResponse<T>(response, url, endpoint);
      });
    });
  }

  /**
   * Make a DELETE request to Microsoft Graph API.
   * Used for deleting resources.
   */
  async delete(
    endpoint: string,
    options?: { headers?: Record<string, string> }
  ): Promise<void> {
    await this.rateLimiter.checkLimit();
    const url = this.buildUrl(endpoint);

    return this.circuitBreaker.execute(async () => {
      return this.executeWithRetry<void>(async () => {
        const accessToken = await this.tokenRefresher.getValidAccessToken();

        logger.debug('Making Graph API DELETE request', {
          endpoint: endpoint,
        });

        const response = await fetch(url, {
          method: 'DELETE',
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            ...options?.headers,
          },
        });

        // DELETE typically returns 204 No Content
        if (response.status === 204 || response.ok) {
          return;
        }

        await this.handleResponse(response, url, endpoint);
      });
    });
  }

  /**
   * Execute a request with exponential backoff retry logic.
   */
  private async executeWithRetry<T>(
    fn: () => Promise<T>,
    attempt: number = 0
  ): Promise<T> {
    try {
      return await fn();
    } catch (error) {
      // Don't retry non-transient errors
      if (!this.isTransientError(error)) {
        throw error;
      }

      // Max retries exceeded
      if (attempt >= this.maxRetries) {
        logger.error('Max retries exceeded', {
          attempt,
          maxRetries: this.maxRetries,
        });
        throw error;
      }

      // Calculate exponential backoff delay
      const delayMs = this.baseRetryDelayMs * Math.pow(2, attempt);
      const jitter = Math.random() * 0.3 * delayMs; // Add 0-30% jitter
      const totalDelay = Math.floor(delayMs + jitter);

      logger.info('Retrying request after delay', {
        attempt: attempt + 1,
        maxRetries: this.maxRetries,
        delayMs: totalDelay,
      });

      // Wait before retrying
      await this.sleep(totalDelay);

      // Retry
      return this.executeWithRetry(fn, attempt + 1);
    }
  }

  /**
   * Handle HTTP response and parse JSON.
   */
  private async handleResponse<T>(response: Response, _url?: string, _endpoint?: string): Promise<T> {
    // Success
    if (response.ok) {
      const data: unknown = await response.json();
      logger.debug('Graph API request successful', {
        status: response.status,
        hasValue: typeof data === 'object' && data !== null && 'value' in data,
      });
      return data as T;
    }

    // Parse error response
    let errorData: any;
    try {
      errorData = await response.json();
    } catch {
      errorData = { message: response.statusText };
    }

    // Log error (sanitized - no URLs, IDs, or error details that could leak sensitive info)
    logger.error('Graph API request failed', {
      status: response.status,
      errorCode: errorData?.error?.code,
    });

    // Handle specific error codes
    switch (response.status) {
      case 401:
        throw new AuthenticationError('Authentication token invalid or expired');

      case 429:
        // Rate limit from Graph API (shouldn't happen with our rate limiter)
        const retryAfter = response.headers.get('Retry-After');
        logger.warn('Graph API rate limit hit', {
          retryAfter,
        });
        throw new GraphAPIError('Rate limit exceeded. Please try again later.');

      case 503:
      case 504:
        // Service unavailable - transient error
        throw new GraphAPIError('Microsoft Graph service temporarily unavailable');

      default:
        // Generic error
        throw new GraphAPIError(
          `Graph API request failed with status ${response.status}`
        );
    }
  }

  /**
   * Determine if an error is transient and should be retried.
   */
  private isTransientError(error: any): boolean {
    if (error instanceof GraphAPIError) {
      // Retry on service unavailable errors
      return error.message.includes('temporarily unavailable');
    }

    if (error instanceof AuthenticationError) {
      // Don't retry auth errors - token refresh should handle this
      return false;
    }

    // Retry on network errors
    if (error.name === 'FetchError' || error.name === 'NetworkError') {
      return true;
    }

    return false;
  }

  /**
   * Build full URL with query parameters.
   * Handles endpoints with or without leading slashes.
   */
  private buildUrl(
    endpoint: string,
    queryParams?: Record<string, string>
  ): string {
    // Remove leading slash if present to avoid URL constructor treating it as absolute path
    // This prevents '/me/todo/lists' from dropping the '/v1.0' base path
    const cleanEndpoint = endpoint.startsWith('/') ? endpoint.slice(1) : endpoint;
    const url = new URL(cleanEndpoint, this.baseUrl);

    logger.debug('Building Graph API URL', {
      originalEndpoint: endpoint,
      cleanEndpoint,
      baseUrl: this.baseUrl,
      finalUrl: url.toString()
    });

    if (queryParams) {
      Object.entries(queryParams).forEach(([key, value]) => {
        url.searchParams.append(key, value);
      });
    }

    return url.toString();
  }

  /**
   * Sleep for specified milliseconds.
   */
  private sleep(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  /**
   * Get circuit breaker state (for monitoring).
   */
  getCircuitState(): CircuitState {
    return this.circuitBreaker.getState();
  }

  /**
   * Get available rate limit tokens (for monitoring).
   */
  getAvailableTokens(): number {
    return this.rateLimiter.getAvailableTokens();
  }

  /**
   * Reset rate limiter and circuit breaker (for testing).
   */
  reset(): void {
    this.rateLimiter.reset();
    this.circuitBreaker.reset();
    logger.debug('Graph client reset');
  }
}
