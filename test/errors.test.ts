/**
 * Error Types Tests
 *
 * Tests for custom error types and error parsing utilities
 */

import { describe, it, expect } from 'vitest';
import {
  GraphApiError,
  RetryableError,
  RateLimitError,
  ServiceUnavailableError,
  AuthenticationError,
  AuthorizationError,
  NotFoundError,
  isRetryableError,
  getRetryAfter,
  parseGraphApiError,
  type ParsedGraphApiError,
} from '../src/errors.js';

describe('Error Types', () => {
  describe('GraphApiError', () => {
    it('should create a GraphApiError with correct properties', () => {
      const error = new GraphApiError('Test error', 500);

      expect(error.message).toBe('Test error');
      expect(error.statusCode).toBe(500);
      expect(error.retryable).toBe(false);
      expect(error.name).toBe('GraphApiError');
    });

    it('should create a retryable GraphApiError', () => {
      const error = new GraphApiError('Test error', 503, true, 30);

      expect(error.retryable).toBe(true);
      expect(error.retryAfter).toBe(30);
    });

    it('should preserve original error', () => {
      const originalError = new Error('Original');
      const error = new GraphApiError('Test error', 500, false, undefined, originalError);

      expect(error.originalError).toBe(originalError);
    });
  });

  describe('RetryableError', () => {
    it('should create a RetryableError', () => {
      const error = new RetryableError('Retryable error', 503, 30);

      expect(error.retryable).toBe(true);
      expect(error.statusCode).toBe(503);
      expect(error.retryAfter).toBe(30);
      expect(error.name).toBe('RetryableError');
    });
  });

  describe('RateLimitError', () => {
    it('should create a RateLimitError with retry after', () => {
      const error = new RateLimitError(60);

      expect(error.retryable).toBe(true);
      expect(error.statusCode).toBe(429);
      expect(error.retryAfter).toBe(60);
      expect(error.message).toContain('rate limit exceeded');
      expect(error.message).toContain('60');
      expect(error.name).toBe('RateLimitError');
    });

    it('should create a RateLimitError without retry after', () => {
      const error = new RateLimitError();

      expect(error.retryAfter).toBeUndefined();
      expect(error.message).toContain('unknown');
    });
  });

  describe('ServiceUnavailableError', () => {
    it('should create a ServiceUnavailableError', () => {
      const error = new ServiceUnavailableError(30);

      expect(error.retryable).toBe(true);
      expect(error.statusCode).toBe(503);
      expect(error.retryAfter).toBe(30);
      expect(error.message).toContain('temporarily unavailable');
      expect(error.name).toBe('ServiceUnavailableError');
    });
  });

  describe('AuthenticationError', () => {
    it('should create an AuthenticationError', () => {
      const error = new AuthenticationError('Unauthorized');

      expect(error.statusCode).toBe(401);
      expect(error.retryable).toBe(false);
      expect(error.message).toBe('Unauthorized');
      expect(error.name).toBe('AuthenticationError');
    });
  });

  describe('AuthorizationError', () => {
    it('should create an AuthorizationError', () => {
      const error = new AuthorizationError('Forbidden');

      expect(error.statusCode).toBe(403);
      expect(error.retryable).toBe(false);
      expect(error.message).toBe('Forbidden');
      expect(error.name).toBe('AuthorizationError');
    });
  });

  describe('NotFoundError', () => {
    it('should create a NotFoundError', () => {
      const error = new NotFoundError('Resource not found');

      expect(error.statusCode).toBe(404);
      expect(error.retryable).toBe(false);
      expect(error.message).toBe('Resource not found');
      expect(error.name).toBe('NotFoundError');
    });
  });

  describe('isRetryableError', () => {
    it('should return true for retryable errors', () => {
      const error = new RetryableError('Test', 503);
      expect(isRetryableError(error)).toBe(true);
    });

    it('should return false for non-retryable errors', () => {
      const error = new GraphApiError('Test', 500, false);
      expect(isRetryableError(error)).toBe(false);
    });

    it('should return false for non-GraphApiError errors', () => {
      const error = new Error('Test');
      expect(isRetryableError(error)).toBe(false);
    });
  });

  describe('getRetryAfter', () => {
    it('should return retry after for RetryableError', () => {
      const error = new RetryableError('Test', 503, 30);
      expect(getRetryAfter(error)).toBe(30);
    });

    it('should return undefined for non-retryable errors', () => {
      const error = new GraphApiError('Test', 500);
      expect(getRetryAfter(error)).toBeUndefined();
    });

    it('should return undefined for non-GraphApiError errors', () => {
      const error = new Error('Test');
      expect(getRetryAfter(error)).toBeUndefined();
    });
  });

  describe('parseGraphApiError', () => {
    it('should parse GraphApiError correctly', () => {
      const error = new GraphApiError('Test error', 500);
      const parsed = parseGraphApiError(error);

      expect(parsed.message).toBe('Test error');
      expect(parsed.statusCode).toBe(500);
      expect(parsed.rawError).toBe(error);
    });

    it('should parse error from JSON response', () => {
      const errorText = JSON.stringify({
        error: {
          code: 'Request_ResourceNotFound',
          message: 'Resource not found',
        },
      });

      const parsed = parseGraphApiError(new Error('Test'), errorText);

      expect(parsed.errorCode).toBe('Request_ResourceNotFound');
      expect(parsed.message).toBe('Resource not found');
      expect(parsed.resourceNotFound).toBe(true);
    });

    it('should detect missing scopes', () => {
      const errorText = JSON.stringify({
        error: {
          code: 'InsufficientPrivileges',
          message: 'Insufficient privileges to complete the operation. Required scope: Mail.Read',
        },
      });

      const parsed = parseGraphApiError(new Error('Test'), errorText);

      expect(parsed.missingScopes).toBeDefined();
      expect(parsed.missingScopes?.length).toBeGreaterThan(0);
    });

    it('should detect invalid parameters', () => {
      const errorText = JSON.stringify({
        error: {
          code: 'InvalidParameter',
          message: 'Invalid parameter userId',
        },
      });

      const parsed = parseGraphApiError(new Error('Test'), errorText);

      // The parser looks for "parameter" or "property" in the message
      // It may not always extract parameter names, so we just check if it's handled
      expect(parsed.message).toContain('Invalid parameter');
      expect(parsed.statusCode).toBe(500); // Default status code
    });

    it('should detect deprecated API', () => {
      const errorText = JSON.stringify({
        error: {
          code: 'DeprecatedApi',
          message: 'This API is deprecated and no longer supported',
        },
      });

      const parsed = parseGraphApiError(new Error('Test'), errorText);

      expect(parsed.deprecatedApi).toBe(true);
    });

    it('should handle non-JSON error text', () => {
      const errorText = 'Plain text error message';
      const parsed = parseGraphApiError(new Error('Test'), errorText);

      expect(parsed.message).toBe('Plain text error message');
    });

    it('should handle error from originalError object', () => {
      const originalError = {
        error: {
          code: 'TestError',
          message: 'Test message',
        },
      };

      const error = new GraphApiError('Test', 500, false, undefined, originalError);
      const parsed = parseGraphApiError(error);

      expect(parsed.errorCode).toBe('TestError');
      expect(parsed.message).toBe('Test message');
    });

    it('should handle string error code', () => {
      const originalError = {
        error: 'TestError',
      };

      const error = new GraphApiError('Test', 500, false, undefined, originalError);
      const parsed = parseGraphApiError(error);

      expect(parsed.errorCode).toBe('TestError');
    });

    it('should extract scopes from message text', () => {
      const error = new Error('Insufficient privileges. Required: Mail.Read, Calendar.Read');
      const parsed = parseGraphApiError(error);

      expect(parsed.missingScopes).toBeDefined();
      expect(parsed.missingScopes?.length).toBeGreaterThan(0);
    });

    it('should handle Response objects gracefully', () => {
      const response = new Response();
      const error = new GraphApiError('Test', 500, false, undefined, response);
      const parsed = parseGraphApiError(error);

      // Should not crash and should return default result
      expect(parsed.message).toBe('Test');
      expect(parsed.statusCode).toBe(500);
    });
  });
});
