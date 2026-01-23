import { Request, Response, NextFunction } from 'express';
import logger from '../logger.js';

/**
 * Middleware to log all HTTP requests for debugging purposes.
 * Logs method, URL, headers, query parameters, and body (if present).
 */
export function requestLoggerMiddleware(req: Request, res: Response, next: NextFunction): void {
  try {
    const startTime = Date.now();
    const requestId = Math.random().toString(36).substring(7);
    let responseLogged = false; // Track if response has been logged to prevent duplicates

    // Log request details
    const requestInfo: Record<string, unknown> = {
      requestId,
      method: req.method,
      url: req.url,
      path: req.path,
      query: req.query,
      headers: sanitizeHeaders(req.headers),
      ip: req.ip || req.socket.remoteAddress || 'unknown',
      userAgent: req.get('user-agent') || 'unknown',
    };

    // Log body if present (but limit size to avoid logging huge payloads)
    if (req.body && Object.keys(req.body).length > 0) {
      try {
        const bodyStr = JSON.stringify(req.body);
        if (bodyStr.length < 10000) {
          // Only log if body is less than 10KB
          requestInfo.body = sanitizeBody(req.body);
        } else {
          requestInfo.bodySize = bodyStr.length;
          requestInfo.bodyPreview = bodyStr.substring(0, 500) + '... (truncated)';
        }
      } catch (error) {
        requestInfo.bodyError = 'Failed to serialize body';
      }
    }

    logger.info(`[${requestId}] Incoming ${req.method} ${req.path}`, requestInfo);

    // Log response when finished
    const originalSend = res.send;
    const originalJson = res.json;
    const originalEnd = res.end;

    res.send = function (body: unknown) {
      try {
        if (!responseLogged) {
          logResponse(requestId, req, res, startTime, body);
          responseLogged = true;
        }
      } catch (error) {
        logger.error(`[${requestId}] Error logging response:`, error);
      }
      return originalSend.call(this, body);
    };

    res.json = function (body: unknown) {
      try {
        if (!responseLogged) {
          logResponse(requestId, req, res, startTime, body);
          responseLogged = true;
        }
      } catch (error) {
        logger.error(`[${requestId}] Error logging response:`, error);
      }
      return originalJson.call(this, body);
    };

    res.end = function (chunk?: unknown, encoding?: unknown) {
      try {
        if (chunk && !res.headersSent && !responseLogged) {
          logResponse(requestId, req, res, startTime, chunk);
          responseLogged = true;
        }
      } catch (error) {
        logger.error(`[${requestId}] Error logging response:`, error);
      }
      return originalEnd.call(this, chunk, encoding);
    };

    // Log errors
    res.on('error', (error: Error) => {
      logger.error(`[${requestId}] Response error:`, {
        requestId,
        method: req.method,
        path: req.path,
        error: error.message,
        stack: error.stack,
      });
    });

    next();
  } catch (error) {
    // If logging fails, don't break the request
    logger.error('Request logger middleware error:', error);
    next();
  }
}

/**
 * Helper function to log response details
 */
function logResponse(
  requestId: string,
  req: Request,
  res: Response,
  startTime: number,
  body: unknown
): void {
  const duration = Date.now() - startTime;
  const responseInfo: Record<string, unknown> = {
    requestId,
    method: req.method,
    path: req.path,
    statusCode: res.statusCode,
    duration: `${duration}ms`,
    headers: sanitizeHeaders(res.getHeaders()),
  };

  // Log response body if it's small enough
  if (body) {
    // Handle Buffer objects (common in Express responses)
    if (Buffer.isBuffer(body)) {
      try {
        const bodyStr = body.toString('utf8');
        if (bodyStr.length < 5000) {
          try {
            const parsed = JSON.parse(bodyStr);
            responseInfo.body = sanitizeResponseBody(parsed);
          } catch {
            // Not JSON, log first 500 chars
            responseInfo.bodyPreview = bodyStr.substring(0, 500) + '... (truncated)';
          }
        } else {
          responseInfo.bodySize = bodyStr.length;
          responseInfo.bodyPreview = bodyStr.substring(0, 200) + '... (truncated)';
        }
      } catch {
        responseInfo.bodyType = 'Buffer';
        responseInfo.bodySize = body.length;
      }
    } else if (typeof body === 'string' && body.length < 5000) {
      try {
        const parsed = JSON.parse(body);
        responseInfo.body = sanitizeResponseBody(parsed);
      } catch {
        // Not JSON, log first 500 chars
        responseInfo.bodyPreview = body.substring(0, 500);
      }
    } else if (typeof body === 'string') {
      responseInfo.bodySize = body.length;
      responseInfo.bodyPreview = body.substring(0, 200) + '... (truncated)';
    } else if (typeof body === 'object' && body !== null) {
      // Check if it's a Buffer-like object (has numeric indices)
      if (Array.isArray(body) || (body.constructor?.name === 'Object' && Object.keys(body).some(key => /^\d+$/.test(key)))) {
        // Might be a serialized Buffer, try to reconstruct
        try {
          const bodyStr = JSON.stringify(body);
          if (bodyStr.length < 5000) {
            responseInfo.body = sanitizeResponseBody(body);
          } else {
            responseInfo.bodySize = bodyStr.length;
            responseInfo.bodyPreview = bodyStr.substring(0, 200) + '... (truncated)';
          }
        } catch {
          responseInfo.bodyType = typeof body;
        }
      } else {
        // Regular object
        try {
          const bodyStr = JSON.stringify(body);
          if (bodyStr.length < 5000) {
            responseInfo.body = sanitizeResponseBody(body);
          } else {
            responseInfo.bodySize = bodyStr.length;
            responseInfo.bodyPreview = bodyStr.substring(0, 200) + '... (truncated)';
          }
        } catch {
          responseInfo.bodyType = typeof body;
        }
      }
    }
  }

  logger.info(`[${requestId}] Response ${res.statusCode} ${req.method} ${req.path}`, responseInfo);
}

/**
 * Sanitize headers to remove sensitive information
 */
function sanitizeHeaders(headers: Record<string, unknown>): Record<string, unknown> {
  const sanitized: Record<string, unknown> = {};
  const sensitiveHeaders = ['authorization', 'cookie', 'x-api-key', 'x-auth-token'];

  for (const [key, value] of Object.entries(headers)) {
    const lowerKey = key.toLowerCase();
    if (sensitiveHeaders.some((h) => lowerKey.includes(h))) {
      sanitized[key] = '[REDACTED]';
    } else {
      sanitized[key] = value;
    }
  }

  return sanitized;
}

/**
 * Sanitize request body to remove sensitive information
 */
function sanitizeBody(body: unknown): unknown {
  if (typeof body !== 'object' || body === null) {
    return body;
  }

  if (Array.isArray(body)) {
    return body.map(sanitizeBody);
  }

  const sanitized: Record<string, unknown> = {};
  const sensitiveFields = [
    'password',
    'client_secret',
    'access_token',
    'refresh_token',
    'code',
    'code_verifier',
    'authorization',
  ];

  for (const [key, value] of Object.entries(body)) {
    const lowerKey = key.toLowerCase();
    if (sensitiveFields.some((field) => lowerKey.includes(field))) {
      sanitized[key] = '[REDACTED]';
    } else if (typeof value === 'object' && value !== null) {
      sanitized[key] = sanitizeBody(value);
    } else {
      sanitized[key] = value;
    }
  }

  return sanitized;
}

/**
 * Sanitize response body to remove sensitive information
 */
function sanitizeResponseBody(body: unknown): unknown {
  if (typeof body !== 'object' || body === null) {
    return body;
  }

  if (Array.isArray(body)) {
    return body.map(sanitizeResponseBody);
  }

  const sanitized: Record<string, unknown> = {};
  const sensitiveFields = [
    'access_token',
    'refresh_token',
    'id_token',
    'token',
    'password',
    'client_secret',
  ];

  for (const [key, value] of Object.entries(body)) {
    const lowerKey = key.toLowerCase();
    if (sensitiveFields.some((field) => lowerKey.includes(field))) {
      sanitized[key] = '[REDACTED]';
    } else if (typeof value === 'object' && value !== null) {
      sanitized[key] = sanitizeResponseBody(value);
    } else {
      sanitized[key] = value;
    }
  }

  return sanitized;
}
