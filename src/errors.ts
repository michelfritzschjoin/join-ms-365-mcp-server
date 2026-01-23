/**
 * Custom error types for Microsoft Graph API errors
 */

export class GraphApiError extends Error {
  public readonly statusCode: number;
  public readonly retryable: boolean;
  public readonly retryAfter?: number;
  public readonly originalError?: unknown;

  constructor(
    message: string,
    statusCode: number,
    retryable = false,
    retryAfter?: number,
    originalError?: unknown
  ) {
    super(message);
    this.name = 'GraphApiError';
    this.statusCode = statusCode;
    this.retryable = retryable;
    this.retryAfter = retryAfter;
    this.originalError = originalError;

    // Maintains proper stack trace for where our error was thrown (only available on V8)
    if (Error.captureStackTrace) {
      Error.captureStackTrace(this, GraphApiError);
    }
  }
}

export class RetryableError extends GraphApiError {
  constructor(message: string, statusCode: number, retryAfter?: number, originalError?: unknown) {
    super(message, statusCode, true, retryAfter, originalError);
    this.name = 'RetryableError';
  }
}

export class RateLimitError extends RetryableError {
  constructor(retryAfter?: number, originalError?: unknown) {
    super(
      `Microsoft Graph API rate limit exceeded. Retry after ${retryAfter || 'unknown'} seconds.`,
      429,
      retryAfter,
      originalError
    );
    this.name = 'RateLimitError';
  }
}

export class ServiceUnavailableError extends RetryableError {
  constructor(retryAfter?: number, originalError?: unknown) {
    super('Microsoft Graph API service temporarily unavailable.', 503, retryAfter, originalError);
    this.name = 'ServiceUnavailableError';
  }
}

export class AuthenticationError extends GraphApiError {
  constructor(message: string, originalError?: unknown) {
    super(message, 401, false, undefined, originalError);
    this.name = 'AuthenticationError';
  }
}

export class AuthorizationError extends GraphApiError {
  constructor(message: string, originalError?: unknown) {
    super(message, 403, false, undefined, originalError);
    this.name = 'AuthorizationError';
  }
}

export class NotFoundError extends GraphApiError {
  constructor(message: string, originalError?: unknown) {
    super(message, 404, false, undefined, originalError);
    this.name = 'NotFoundError';
  }
}

/**
 * Check if an error is retryable
 */
export function isRetryableError(error: unknown): boolean {
  if (error instanceof GraphApiError) {
    return error.retryable;
  }
  return false;
}

/**
 * Extract retry-after value from error
 */
export function getRetryAfter(error: unknown): number | undefined {
  if (error instanceof RetryableError) {
    return error.retryAfter;
  }
  return undefined;
}
