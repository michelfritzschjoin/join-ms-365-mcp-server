/**
 * Security Headers Middleware
 */

import { Request, Response, NextFunction } from 'express';
import logger from '../logger.js';

/**
 * Security headers middleware
 */
export function securityHeadersMiddleware(req: Request, res: Response, next: NextFunction): void {
  // X-Content-Type-Options: Prevent MIME type sniffing
  res.setHeader('X-Content-Type-Options', 'nosniff');

  // X-Frame-Options: Prevent clickjacking
  const frameOptions = process.env.MS365_MCP_X_FRAME_OPTIONS || 'DENY';
  res.setHeader('X-Frame-Options', frameOptions);

  // X-XSS-Protection: Enable XSS filter (legacy, but still useful)
  res.setHeader('X-XSS-Protection', '1; mode=block');

  // Referrer-Policy: Control referrer information
  const referrerPolicy = process.env.MS365_MCP_REFERRER_POLICY || 'strict-origin-when-cross-origin';
  res.setHeader('Referrer-Policy', referrerPolicy);

  // Permissions-Policy: Control browser features
  const permissionsPolicy =
    process.env.MS365_MCP_PERMISSIONS_POLICY || 'geolocation=(), microphone=(), camera=()';
  res.setHeader('Permissions-Policy', permissionsPolicy);

  // Content-Security-Policy: Control resource loading
  const csp =
    process.env.MS365_MCP_CSP ||
    "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline';";
  res.setHeader('Content-Security-Policy', csp);

  // Strict-Transport-Security: Force HTTPS (only if HTTPS is used)
  if (req.secure || req.headers['x-forwarded-proto'] === 'https') {
    const maxAge = process.env.MS365_MCP_HSTS_MAX_AGE || '31536000'; // 1 year
    res.setHeader('Strict-Transport-Security', `max-age=${maxAge}; includeSubDomains`);
  }

  next();
}

export default securityHeadersMiddleware;
