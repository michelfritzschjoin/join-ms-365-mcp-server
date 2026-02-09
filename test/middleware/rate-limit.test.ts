/**
 * Rate Limit Middleware Tests
 *
 * Tests for rate limiting functionality
 */

import { describe, it, expect, beforeEach, vi } from 'vitest';
import { Request, Response, NextFunction } from 'express';
import { createRateLimitMiddleware } from '../../src/middleware/rate-limit.js';

describe('Rate Limit Middleware', () => {
  let mockRequest: Partial<Request>;
  let mockResponse: Partial<Response>;
  let mockNext: NextFunction;
  let middleware: ReturnType<typeof createRateLimitMiddleware>;

  beforeEach(() => {
    mockRequest = {
      ip: '127.0.0.1',
      socket: {
        remoteAddress: '127.0.0.1',
      } as unknown,
    };

    mockResponse = {
      setHeader: vi.fn(),
      status: vi.fn().mockReturnThis(),
      json: vi.fn().mockReturnThis(),
    };

    mockNext = vi.fn();
  });

  describe('createRateLimitMiddleware', () => {
    it('should allow requests within limit', () => {
      middleware = createRateLimitMiddleware(60000, 5); // 5 requests per minute

      for (let i = 0; i < 5; i++) {
        middleware(mockRequest as Request, mockResponse as Response, mockNext);
      }

      expect(mockNext).toHaveBeenCalledTimes(5);
      expect(mockResponse.status).not.toHaveBeenCalledWith(429);
    });

    it('should block requests exceeding limit', () => {
      middleware = createRateLimitMiddleware(60000, 2); // 2 requests per minute

      // Make 2 allowed requests
      middleware(mockRequest as Request, mockResponse as Response, mockNext);
      middleware(mockRequest as Request, mockResponse as Response, mockNext);

      // Third request should be blocked
      middleware(mockRequest as Request, mockResponse as Response, mockNext);

      expect(mockNext).toHaveBeenCalledTimes(2);
      expect(mockResponse.status).toHaveBeenCalledWith(429);
      expect(mockResponse.json).toHaveBeenCalledWith(
        expect.objectContaining({
          error: 'Too Many Requests',
        })
      );
    });

    it('should set rate limit headers', () => {
      middleware = createRateLimitMiddleware(60000, 10);

      middleware(mockRequest as Request, mockResponse as Response, mockNext);

      expect(mockResponse.setHeader).toHaveBeenCalledWith('X-RateLimit-Limit', '10');
      expect(mockResponse.setHeader).toHaveBeenCalledWith(
        'X-RateLimit-Remaining',
        expect.any(String)
      );
      expect(mockResponse.setHeader).toHaveBeenCalledWith('X-RateLimit-Reset', expect.any(String));
    });

    it('should use IP address as client identifier', () => {
      middleware = createRateLimitMiddleware(60000, 1);

      const request1 = { ...mockRequest, ip: '192.168.1.1' } as Request;
      const request2 = { ...mockRequest, ip: '192.168.1.2' } as Request;

      middleware(request1, mockResponse as Response, mockNext);
      middleware(request2, mockResponse as Response, mockNext);

      // Both should be allowed (different IPs)
      expect(mockNext).toHaveBeenCalledTimes(2);
    });

    it('should use socket remote address as fallback', () => {
      middleware = createRateLimitMiddleware(60000, 1);

      const request = {
        socket: { remoteAddress: '192.168.1.1' },
      } as unknown as Request;

      middleware(request, mockResponse as Response, mockNext);

      expect(mockNext).toHaveBeenCalled();
    });

    it('should reset limit after window expires', async () => {
      middleware = createRateLimitMiddleware(100, 1); // 1 request per 100ms

      // First request
      middleware(mockRequest as Request, mockResponse as Response, mockNext);
      expect(mockNext).toHaveBeenCalledTimes(1);

      // Second request immediately (should be blocked)
      middleware(mockRequest as Request, mockResponse as Response, mockNext);
      expect(mockResponse.status).toHaveBeenCalledWith(429);

      // Wait for window to expire
      await new Promise((resolve) => setTimeout(resolve, 150));

      // Reset mocks
      vi.clearAllMocks();

      // Request after window should be allowed
      middleware(mockRequest as Request, mockResponse as Response, mockNext);
      expect(mockNext).toHaveBeenCalled();
    });
  });
});
