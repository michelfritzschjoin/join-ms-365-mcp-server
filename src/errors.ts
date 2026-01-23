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

/**
 * Parsed Graph API error information
 */
export interface ParsedGraphApiError {
  errorCode?: string;
  message: string;
  statusCode: number;
  missingScopes?: string[];
  invalidParameters?: string[];
  deprecatedApi?: boolean;
  resourceNotFound?: boolean;
  rawError?: unknown;
}

/**
 * Parse Microsoft Graph API error response to extract structured information
 * Note: For Response objects, the response body should be read first before calling this function
 */
export function parseGraphApiError(error: unknown, errorText?: string): ParsedGraphApiError {
  const defaultResult: ParsedGraphApiError = {
    message: error instanceof Error ? error.message : String(error),
    statusCode: error instanceof GraphApiError ? error.statusCode : 500,
    rawError: error,
  };

  // If errorText is provided (from Response.text()), try to parse it
  if (errorText) {
    try {
      const errorBody = JSON.parse(errorText);
      return parseGraphApiErrorFromObject(errorBody, defaultResult);
    } catch {
      // Not JSON, use text as message
      return parseGraphApiErrorFromMessage(errorText, defaultResult);
    }
  }

  // If it's a GraphApiError, try to extract more info from originalError
  if (error instanceof GraphApiError && error.originalError) {
    const originalError = error.originalError;

    // Try to parse error object (Response objects should be handled separately)
    if (
      typeof originalError === 'object' &&
      originalError !== null &&
      !(originalError instanceof Response)
    ) {
      return parseGraphApiErrorFromObject(originalError, defaultResult);
    }
  }

  // Try to extract from error message
  const message = error instanceof Error ? error.message : String(error);
  return parseGraphApiErrorFromMessage(message, defaultResult);
}

/**
 * Parse error from object (typically from Graph API error response)
 */
function parseGraphApiErrorFromObject(
  errorObj: unknown,
  defaultResult: ParsedGraphApiError
): ParsedGraphApiError {
  if (typeof errorObj !== 'object' || errorObj === null) {
    return defaultResult;
  }

  const obj = errorObj as Record<string, unknown>;
  const result: ParsedGraphApiError = { ...defaultResult };

  // Extract error code
  if (typeof obj.error === 'object' && obj.error !== null) {
    const error = obj.error as Record<string, unknown>;

    // Error code (e.g., "Request_ResourceNotFound")
    if (typeof error.code === 'string') {
      result.errorCode = error.code;
    }

    // Error message
    if (typeof error.message === 'string') {
      result.message = error.message;
    }

    // Inner error (nested error details)
    if (error.innerError) {
      const innerError = error.innerError as Record<string, unknown>;
      if (typeof innerError['request-id'] === 'string') {
        // Can be used for correlation
      }
    }
  } else if (typeof obj.error === 'string') {
    result.errorCode = obj.error;
  }

  // Check for common error patterns
  const message = result.message.toLowerCase();

  // Resource not found
  if (
    result.errorCode?.includes('ResourceNotFound') ||
    result.errorCode?.includes('NotFound') ||
    message.includes('not found') ||
    message.includes('does not exist')
  ) {
    result.resourceNotFound = true;
  }

  // Missing scopes/permissions
  if (
    message.includes('insufficient privileges') ||
    message.includes('permission') ||
    message.includes('scope') ||
    message.includes('authorization') ||
    result.statusCode === 403
  ) {
    // Try to extract missing scopes from message
    const scopeMatches = message.match(/([A-Za-z]+\.[A-Za-z]+(?:\.[A-Za-z]+)?)/g);
    if (scopeMatches) {
      result.missingScopes = [...new Set(scopeMatches)];
    }
  }

  // Invalid parameters
  if (
    message.includes('invalid') ||
    message.includes('parameter') ||
    message.includes('property') ||
    result.statusCode === 400
  ) {
    // Try to extract parameter names from message
    const paramMatches = message.match(/(?:parameter|property)\s+['"]?(\w+)['"]?/gi);
    if (paramMatches) {
      result.invalidParameters = paramMatches.map((m) =>
        m.replace(/^(?:parameter|property)\s+['"]?/i, '').replace(/['"]?$/i, '')
      );
    }
  }

  // Deprecated API
  if (
    message.includes('deprecated') ||
    message.includes('obsolete') ||
    message.includes('no longer supported')
  ) {
    result.deprecatedApi = true;
  }

  return result;
}

/**
 * Parse error from message string (fallback)
 */
function parseGraphApiErrorFromMessage(
  message: string,
  defaultResult: ParsedGraphApiError
): ParsedGraphApiError {
  const result: ParsedGraphApiError = { ...defaultResult, message };

  const lowerMessage = message.toLowerCase();

  // Check for common patterns in message
  if (lowerMessage.includes('not found') || lowerMessage.includes('does not exist')) {
    result.resourceNotFound = true;
  }

  if (
    lowerMessage.includes('insufficient privileges') ||
    lowerMessage.includes('permission') ||
    lowerMessage.includes('scope') ||
    lowerMessage.includes('authorization')
  ) {
    const scopeMatches = message.match(/([A-Za-z]+\.[A-Za-z]+(?:\.[A-Za-z]+)?)/g);
    if (scopeMatches) {
      result.missingScopes = [...new Set(scopeMatches)];
    }
  }

  if (lowerMessage.includes('deprecated') || lowerMessage.includes('obsolete')) {
    result.deprecatedApi = true;
  }

  return result;
}
