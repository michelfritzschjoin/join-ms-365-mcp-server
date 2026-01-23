/**
 * Rate Limiting Middleware for HTTP Endpoints
 */

import { Request, Response, NextFunction } from 'express';
import logger from '../logger.js';

interface RateLimitStore {
  [key: string]: {
    count: number;
    resetTime: number;
  };
}

class MemoryRateLimitStore {
  private store: RateLimitStore = {};
  private readonly windowMs: number;
  private readonly maxRequests: number;

  constructor(windowMs: number, maxRequests: number) {
    this.windowMs = windowMs;
    this.maxRequests = maxRequests;

    // Clean up expired entries every minute
    setInterval(() => {
      this.cleanup();
    }, 60000);
  }

  /**
   * Check if request is allowed
   */
  check(key: string): { allowed: boolean; remaining: number; resetTime: number } {
    const now = Date.now();
    const entry = this.store[key];

    if (!entry || now > entry.resetTime) {
      // Create new entry
      this.store[key] = {
        count: 1,
        resetTime: now + this.windowMs,
      };
      return {
        allowed: true,
        remaining: this.maxRequests - 1,
        resetTime: now + this.windowMs,
      };
    }

    if (entry.count >= this.maxRequests) {
      return {
        allowed: false,
        remaining: 0,
        resetTime: entry.resetTime,
      };
    }

    entry.count++;
    return {
      allowed: true,
      remaining: this.maxRequests - entry.count,
      resetTime: entry.resetTime,
    };
  }

  /**
   * Clean up expired entries
   */
  private cleanup(): void {
    const now = Date.now();
    for (const key in this.store) {
      if (this.store[key].resetTime < now) {
        delete this.store[key];
      }
    }
  }
}

/**
 * Create rate limit middleware
 */
export function createRateLimitMiddleware(
  windowMs: number = 15 * 60 * 1000, // 15 minutes
  maxRequests: number = 100
) {
  const store = new MemoryRateLimitStore(windowMs, maxRequests);

  return (req: Request, res: Response, next: NextFunction): void => {
    // Get client identifier (IP address or user ID)
    const clientId = req.ip || req.socket.remoteAddress || 'unknown';

    const result = store.check(clientId);

    // Set rate limit headers
    res.setHeader('X-RateLimit-Limit', maxRequests.toString());
    res.setHeader('X-RateLimit-Remaining', result.remaining.toString());
    res.setHeader('X-RateLimit-Reset', new Date(result.resetTime).toISOString());

    if (!result.allowed) {
      logger.warn(`Rate limit exceeded for ${clientId}`);
      res.status(429).json({
        error: 'Too Many Requests',
        message: 'Rate limit exceeded. Please try again later.',
        retryAfter: Math.ceil((result.resetTime - Date.now()) / 1000),
      });
      return;
    }

    next();
  };
}

/**
 * Default rate limit middleware (100 requests per 15 minutes)
 */
export const rateLimitMiddleware = createRateLimitMiddleware(
  parseInt(process.env.MS365_MCP_RATE_LIMIT_WINDOW_MS || '900000', 10), // 15 minutes
  parseInt(process.env.MS365_MCP_RATE_LIMIT_MAX_REQUESTS || '100', 10)
);

export default rateLimitMiddleware;
