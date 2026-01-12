/**
 * Base error class for all application errors.
 * Ensures errors never leak sensitive information.
 */
export abstract class AppError extends Error {
  abstract readonly statusCode: number;
  abstract readonly isOperational: boolean;

  constructor(message: string) {
    super(message);
    this.name = this.constructor.name;
    Error.captureStackTrace(this, this.constructor);
  }

  /**
   * Get sanitized error for client response
   */
  toJSON() {
    return {
      name: this.name,
      message: this.message,
      // Never include stack traces in client responses
    };
  }
}

export class AuthenticationError extends AppError {
  readonly statusCode = 401;
  readonly isOperational = true;

  constructor(message: string = 'Authentication required') {
    super(message);
  }
}

export class AuthorizationError extends AppError {
  readonly statusCode = 403;
  readonly isOperational = true;

  constructor(message: string = 'Access denied') {
    super(message);
  }
}

export class ValidationError extends AppError {
  readonly statusCode = 400;
  readonly isOperational = true;

  constructor(message: string = 'Invalid input') {
    super(message);
  }
}

export class RateLimitError extends AppError {
  readonly statusCode = 429;
  readonly isOperational = true;
  readonly retryAfterMs: number;

  constructor(message: string = 'Rate limit exceeded', retryAfterMs: number = 0) {
    super(message);
    this.retryAfterMs = retryAfterMs;
  }

  toJSON() {
    return {
      ...super.toJSON(),
      retryAfterSeconds: Math.ceil(this.retryAfterMs / 1000),
    };
  }
}

export class GraphAPIError extends AppError {
  readonly statusCode = 502;
  readonly isOperational = false;

  constructor(message: string = 'External service error') {
    super(message);
  }
}

export class TokenStorageError extends AppError {
  readonly statusCode = 500;
  readonly isOperational = false;

  constructor(message: string = 'Token storage error') {
    super(message);
  }
}

export class ConfigurationError extends AppError {
  readonly statusCode = 500;
  readonly isOperational = false;

  constructor(message: string = 'Configuration error') {
    super(message);
  }
}
