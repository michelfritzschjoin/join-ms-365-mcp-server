/**
 * CORS Middleware Tests
 *
 * Tests for CORS middleware functionality
 */

import { describe, it, expect, beforeEach, vi, afterEach } from 'vitest';
import { Request, Response, NextFunction } from 'express';
import { corsMiddleware } from '../../src/middleware/cors.js';

describe('CORS Middleware', () => {
  let mockRequest: Partial<Request>;
  let mockResponse: Partial<Response>;
  let mockNext: NextFunction;
  let originalEnv: typeof process.env;

  beforeEach(() => {
    originalEnv = { ...process.env };
    process.env = { ...originalEnv };

    mockRequest = {
      method: 'GET',
      path: '/api/test',
      headers: {
        origin: 'https://example.com',
      },
    };

    mockResponse = {
      setHeader: vi.fn(),
      status: vi.fn().mockReturnThis(),
      end: vi.fn(),
    };

    mockNext = vi.fn();
  });

  afterEach(() => {
    process.env = originalEnv;
    vi.clearAllMocks();
  });

  describe('corsMiddleware', () => {
    it('should not set CORS headers when no origins configured', () => {
      delete process.env.MS365_MCP_CORS_ORIGINS;
      delete process.env.MS365_MCP_CORS_ORIGIN;

      corsMiddleware(mockRequest as Request, mockResponse as Response, mockNext);

      expect(mockResponse.setHeader).not.toHaveBeenCalledWith(
        'Access-Control-Allow-Origin',
        expect.any(String)
      );
      expect(mockNext).toHaveBeenCalled();
    });

    it('should allow configured origin', () => {
      process.env.MS365_MCP_CORS_ORIGINS = 'https://example.com,https://test.com';

      corsMiddleware(mockRequest as Request, mockResponse as Response, mockNext);

      expect(mockResponse.setHeader).toHaveBeenCalledWith(
        'Access-Control-Allow-Origin',
        'https://example.com'
      );
      expect(mockNext).toHaveBeenCalled();
    });

    it('should reject non-configured origin', () => {
      process.env.MS365_MCP_CORS_ORIGINS = 'https://allowed.com';
      mockRequest.headers = { origin: 'https://blocked.com' };

      corsMiddleware(mockRequest as Request, mockResponse as Response, mockNext);

      expect(mockResponse.setHeader).not.toHaveBeenCalledWith(
        'Access-Control-Allow-Origin',
        'https://blocked.com'
      );
      expect(mockNext).toHaveBeenCalled();
    });

    it('should support legacy single origin', () => {
      process.env.MS365_MCP_CORS_ORIGIN = 'https://example.com';

      corsMiddleware(mockRequest as Request, mockResponse as Response, mockNext);

      expect(mockResponse.setHeader).toHaveBeenCalledWith(
        'Access-Control-Allow-Origin',
        'https://example.com'
      );
    });

    it('should support legacy wildcard origin', () => {
      process.env.MS365_MCP_CORS_ORIGIN = '*';

      corsMiddleware(mockRequest as Request, mockResponse as Response, mockNext);

      expect(mockResponse.setHeader).toHaveBeenCalledWith('Access-Control-Allow-Origin', '*');
    });

    it('should set CORS method headers', () => {
      process.env.MS365_MCP_CORS_ORIGINS = 'https://example.com';

      corsMiddleware(mockRequest as Request, mockResponse as Response, mockNext);

      expect(mockResponse.setHeader).toHaveBeenCalledWith(
        'Access-Control-Allow-Methods',
        expect.any(String)
      );
    });

    it('should set CORS header headers', () => {
      process.env.MS365_MCP_CORS_ORIGINS = 'https://example.com';

      corsMiddleware(mockRequest as Request, mockResponse as Response, mockNext);

      expect(mockResponse.setHeader).toHaveBeenCalledWith(
        'Access-Control-Allow-Headers',
        expect.any(String)
      );
    });

    it('should set exposed headers', () => {
      process.env.MS365_MCP_CORS_ORIGINS = 'https://example.com';

      corsMiddleware(mockRequest as Request, mockResponse as Response, mockNext);

      expect(mockResponse.setHeader).toHaveBeenCalledWith(
        'Access-Control-Expose-Headers',
        expect.any(String)
      );
    });

    it('should set max age header', () => {
      process.env.MS365_MCP_CORS_ORIGINS = 'https://example.com';

      corsMiddleware(mockRequest as Request, mockResponse as Response, mockNext);

      expect(mockResponse.setHeader).toHaveBeenCalledWith(
        'Access-Control-Max-Age',
        expect.any(String)
      );
    });

    it('should handle OPTIONS preflight request', () => {
      process.env.MS365_MCP_CORS_ORIGINS = 'https://example.com';
      mockRequest.method = 'OPTIONS';

      corsMiddleware(mockRequest as Request, mockResponse as Response, mockNext);

      expect(mockResponse.status).toHaveBeenCalledWith(200);
      expect(mockResponse.end).toHaveBeenCalled();
      expect(mockNext).not.toHaveBeenCalled();
    });

    it('should handle request without origin header', () => {
      process.env.MS365_MCP_CORS_ORIGINS = 'https://example.com';
      delete mockRequest.headers?.origin;

      corsMiddleware(mockRequest as Request, mockResponse as Response, mockNext);

      expect(mockNext).toHaveBeenCalled();
    });

    it('should use custom CORS methods from env', () => {
      process.env.MS365_MCP_CORS_ORIGINS = 'https://example.com';
      process.env.MS365_MCP_CORS_METHODS = 'GET, POST';

      corsMiddleware(mockRequest as Request, mockResponse as Response, mockNext);

      expect(mockResponse.setHeader).toHaveBeenCalledWith(
        'Access-Control-Allow-Methods',
        'GET, POST'
      );
    });

    it('should use custom CORS headers from env', () => {
      process.env.MS365_MCP_CORS_ORIGINS = 'https://example.com';
      process.env.MS365_MCP_CORS_HEADERS = 'Custom-Header';

      corsMiddleware(mockRequest as Request, mockResponse as Response, mockNext);

      expect(mockResponse.setHeader).toHaveBeenCalledWith(
        'Access-Control-Allow-Headers',
        'Custom-Header'
      );
    });
  });
});
