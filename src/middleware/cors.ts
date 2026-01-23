/**
 * CORS Middleware with improved configuration
 */

import { Request, Response, NextFunction } from 'express';
import logger from '../logger.js';

/**
 * CORS middleware with configurable origins
 */
export function corsMiddleware(req: Request, res: Response, next: NextFunction): void {
  // Get allowed origins from environment (comma-separated)
  const allowedOrigins = process.env.MS365_MCP_CORS_ORIGINS
    ? process.env.MS365_MCP_CORS_ORIGINS.split(',').map((o) => o.trim())
    : [];

  // Default: no CORS (more secure) unless explicitly configured
  // Legacy: MS365_MCP_CORS_ORIGIN for single origin
  const legacyOrigin = process.env.MS365_MCP_CORS_ORIGIN;

  const origin = req.headers.origin;

  // Determine allowed origin
  let allowedOrigin: string | undefined;

  if (allowedOrigins.length > 0) {
    // Check if origin is in allowed list
    if (origin && allowedOrigins.includes(origin)) {
      allowedOrigin = origin;
    }
  } else if (legacyOrigin) {
    // Legacy support: single origin or '*'
    if (legacyOrigin === '*' || (origin && legacyOrigin === origin)) {
      allowedOrigin = legacyOrigin === '*' ? '*' : origin;
    }
  }

  // Set CORS headers
  if (allowedOrigin) {
    res.setHeader('Access-Control-Allow-Origin', allowedOrigin);
    res.setHeader('Access-Control-Allow-Credentials', 'true');
  }

  // Set allowed methods
  const allowedMethods = process.env.MS365_MCP_CORS_METHODS || 'GET, POST, PUT, DELETE, OPTIONS';
  res.setHeader('Access-Control-Allow-Methods', allowedMethods);

  // Set allowed headers
  const allowedHeaders =
    process.env.MS365_MCP_CORS_HEADERS ||
    'Origin, X-Requested-With, Content-Type, Accept, Authorization, mcp-protocol-version';
  res.setHeader('Access-Control-Allow-Headers', allowedHeaders);

  // Set exposed headers
  const exposedHeaders = process.env.MS365_MCP_CORS_EXPOSED_HEADERS || 'X-RateLimit-Limit, X-RateLimit-Remaining, X-RateLimit-Reset';
  res.setHeader('Access-Control-Expose-Headers', exposedHeaders);

  // Set max age for preflight
  const maxAge = process.env.MS365_MCP_CORS_MAX_AGE || '86400'; // 24 hours
  res.setHeader('Access-Control-Max-Age', maxAge);

  // Handle preflight requests
  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  // Log CORS violations
  if (origin && !allowedOrigin && allowedOrigins.length > 0) {
    logger.warn(`CORS violation: Origin ${origin} not in allowed list`);
  }

  // #region agent log
  if (req.path === '/mcp' || req.url?.includes('/mcp')) {
    fetch('http://127.0.0.1:7245/ingest/76c7865f-57f2-4bf0-8001-38b29d141bbc', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        location: 'middleware/cors.ts:68',
        message: 'CORS check for MCP request',
        data: {
          origin,
          allowedOrigin,
          allowedOrigins,
          hasLegacyOrigin: !!legacyOrigin,
          method: req.method,
          path: req.path,
          corsAllowed: !!allowedOrigin
        },
        timestamp: Date.now(),
        sessionId: 'debug-session',
        runId: 'run1',
        hypothesisId: 'F'
      })
    }).catch(() => {});
  }
  // #endregion

  next();
}

export default corsMiddleware;

