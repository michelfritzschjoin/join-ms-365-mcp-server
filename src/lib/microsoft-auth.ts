import { Request, Response, NextFunction } from 'express';
import logger from '../logger.js';
import { getCloudEndpoints, type CloudType } from '../cloud-config.js';

/**
 * Configuration for Microsoft Bearer Token Auth Middleware
 */
export interface MicrosoftAuthMiddlewareConfig {
  /** Whether to require OAuth authentication (default: true in production) */
  requireAuth?: boolean;
  /** Base URL for OAuth discovery endpoints */
  baseUrl?: string;
}

/**
 * Creates Microsoft Bearer Token Auth Middleware that validates access tokens
 * The token is passed in the Authorization header as a Bearer token
 *
 * IMPORTANT: By default, authentication is REQUIRED. The middleware will return
 * a 401 Unauthorized response with WWW-Authenticate header pointing to the
 * OAuth discovery endpoint, as required by RFC 9728 (Protected Resource Metadata).
 */
export function createMicrosoftBearerTokenAuthMiddleware(
  config: MicrosoftAuthMiddlewareConfig = {}
) {
  const requireAuth = config.requireAuth ?? true;
  const baseUrl = config.baseUrl ?? '';

  return (
    req: Request & { microsoftAuth?: { accessToken: string; refreshToken: string } },
    res: Response,
    next: NextFunction
  ): void => {
    const authHeader = req.headers.authorization;

    if (authHeader && authHeader.startsWith('Bearer ')) {
      const accessToken = authHeader.substring(7);

      // For Microsoft Graph, we don't validate the token here - we'll let the API calls fail if it's invalid
      // and handle token refresh in the GraphClient

      // Extract refresh token from a custom header (if provided)
      const refreshToken = (req.headers['x-microsoft-refresh-token'] as string) || '';

      // Store tokens in request for later use
      req.microsoftAuth = {
        accessToken,
        refreshToken,
      };

      next();
    } else if (!requireAuth) {
      // Authentication not required - allow request to continue (ONLY for explicit testing mode)
      logger.debug('Request without Bearer token - authentication not required in current mode');
      next();
    } else {
      // Authentication REQUIRED but no token provided
      // Return 401 with WWW-Authenticate header per RFC 9728
      logger.warn('Authentication required but no Bearer token provided');

      // Determine the resource-metadata URL for OAuth discovery
      const protocol = req.secure ? 'https' : 'http';
      const host = req.get('host') || 'localhost';
      const discoveryUrl = baseUrl || `${protocol}://${host}`;

      // Set WWW-Authenticate header as per RFC 9728 (OAuth 2.0 Protected Resource Metadata)
      res.set(
        'WWW-Authenticate',
        `Bearer resource_metadata="${discoveryUrl}/.well-known/oauth-protected-resource"`
      );

      res.status(401).json({
        jsonrpc: '2.0',
        error: {
          code: -32001,
          message: 'Authentication required',
          data: {
            type: 'oauth_required',
            description:
              'You must authenticate with Microsoft 365 before accessing this resource. ' +
              'Please complete the OAuth flow to obtain an access token.',
            oauth_discovery: `${discoveryUrl}/.well-known/oauth-protected-resource`,
            authorization_endpoint: `${discoveryUrl}/authorize`,
          },
        },
        id: null,
      });
    }
  };
}

/**
 * Legacy middleware for backward compatibility
 * @deprecated Use createMicrosoftBearerTokenAuthMiddleware instead
 *
 * SECURITY WARNING: This middleware does NOT require authentication and should
 * only be used for local testing/development. In production, use
 * createMicrosoftBearerTokenAuthMiddleware with requireAuth: true.
 */
export const microsoftBearerTokenAuthMiddleware = (
  req: Request & { microsoftAuth?: { accessToken: string; refreshToken: string } },
  res: Response,
  next: NextFunction
): void => {
  const authHeader = req.headers.authorization;

  if (authHeader && authHeader.startsWith('Bearer ')) {
    const accessToken = authHeader.substring(7);

    // For Microsoft Graph, we don't validate the token here - we'll let the API calls fail if it's invalid
    // and handle token refresh in the GraphClient

    // Extract refresh token from a custom header (if provided)
    const refreshToken = (req.headers['x-microsoft-refresh-token'] as string) || '';

    // Store tokens in request for later use
    req.microsoftAuth = {
      accessToken,
      refreshToken,
    };
  } else {
    // No token provided - this is OK for inspector/testing mode
    // The request will continue, but API calls may fail if they require authentication
    logger.debug(
      'Request without Bearer token - continuing without authentication (OK for inspector/testing)'
    );
  }

  next();
};

/**
 * Exchange authorization code for access token
 */
export async function exchangeCodeForToken(
  code: string,
  redirectUri: string,
  clientId: string,
  clientSecret: string | undefined,
  tenantId: string = 'common',
  codeVerifier?: string,
  cloudType: CloudType = 'global'
): Promise<{
  access_token: string;
  token_type: string;
  scope: string;
  expires_in: number;
  refresh_token: string;
}> {
  const cloudEndpoints = getCloudEndpoints(cloudType);
  const params = new URLSearchParams({
    grant_type: 'authorization_code',
    code,
    redirect_uri: redirectUri,
    client_id: clientId,
  });

  // Add client_secret for confidential clients
  if (clientSecret) {
    params.append('client_secret', clientSecret);
  }

  // Add code_verifier for PKCE flow
  if (codeVerifier) {
    params.append('code_verifier', codeVerifier);
  }

  const response = await fetch(`${cloudEndpoints.authority}/${tenantId}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: params,
  });

  if (!response.ok) {
    const errorText = await response.text();
    let errorMessage = `Token exchange failed: ${response.status}`;

    try {
      // Try to parse error as JSON for better error messages
      const errorJson = JSON.parse(errorText);
      if (errorJson.error) {
        errorMessage = `Token exchange failed: ${errorJson.error} - ${errorJson.error_description || errorText}`;
      } else {
        errorMessage = `Token exchange failed: ${response.status} - ${errorText}`;
      }
    } catch {
      // Not JSON, use text as-is
      errorMessage = `Token exchange failed: ${response.status} - ${errorText}`;
    }

    logger.error(`Failed to exchange code for token:`, {
      status: response.status,
      statusText: response.statusText,
      error: errorText,
      redirectUri,
      hasCodeVerifier: !!codeVerifier,
      hasClientSecret: !!clientSecret,
    });

    throw new Error(errorMessage);
  }

  return response.json();
}

/**
 * Refresh an access token
 */
export async function refreshAccessToken(
  refreshToken: string,
  clientId: string,
  clientSecret: string | undefined,
  tenantId: string = 'common',
  cloudType: CloudType = 'global'
): Promise<{
  access_token: string;
  token_type: string;
  scope: string;
  expires_in: number;
  refresh_token?: string;
}> {
  const cloudEndpoints = getCloudEndpoints(cloudType);
  const params = new URLSearchParams({
    grant_type: 'refresh_token',
    refresh_token: refreshToken,
    client_id: clientId,
  });

  if (clientSecret) {
    params.append('client_secret', clientSecret);
  }

  const response = await fetch(`${cloudEndpoints.authority}/${tenantId}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: params,
  });

  if (!response.ok) {
    const error = await response.text();
    logger.error(`Failed to refresh token: ${error}`);
    throw new Error(`Failed to refresh token: ${error}`);
  }

  return response.json();
}
