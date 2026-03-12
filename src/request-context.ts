import { AsyncLocalStorage } from 'node:async_hooks';
import { createHash, randomUUID } from 'node:crypto';
import type { UserProfile, ProfessionProfile } from './user-profile.js';

/**
 * Request context that is passed through async operations
 * Contains authentication tokens, a unique session ID, and chat session information.
 * Each MCP request gets its own sessionId so that concurrent requests from different
 * users never share or mix content.
 */
export interface RequestContext {
  /** Unique session ID for this request (one per request, prevents cross-request mixing). Set by HTTP handlers; optional for tests/stdio. */
  sessionId?: string;
  /** Microsoft Graph API access token */
  accessToken: string;
  /** Microsoft Graph API refresh token */
  refreshToken?: string;
  /** OpenWebUI Chat ID for per-chat memory */
  chatId?: string;
  /** User ID extracted from token for user-scoped memory */
  userId?: string;
  /** Token hash for logging (never log actual token) */
  tokenHash?: string;
  /** User profile with job title, department, and profession profile */
  userProfile?: UserProfile;
}

/**
 * Generate a unique session ID for each MCP request.
 * Used to isolate concurrent requests so that two employees never see mixed content.
 */
export function generateSessionId(): string {
  return randomUUID();
}

/**
 * AsyncLocalStorage for request context propagation
 */
export const requestContext = new AsyncLocalStorage<RequestContext>();

/**
 * SECURITY: Create a safe hash of the token for logging purposes
 * Never log the actual token, only its hash for correlation.
 *
 * @param token - The access token to hash
 * @returns A truncated SHA-256 hash (first 16 chars)
 */
export function createTokenHash(token: string): string {
  if (!token) return 'no-token';
  return createHash('sha256').update(token).digest('hex').substring(0, 16);
}

/**
 * SECURITY: Validate that we have a proper authentication context
 * @returns True if we have valid authentication tokens
 */
export function hasValidAuth(): boolean {
  const ctx = requestContext.getStore();
  return !!(ctx?.accessToken && ctx.accessToken.length > 0);
}

/**
 * SECURITY: Check if we have user identification
 * @returns True if we have a user ID from the token
 */
export function hasUserIdentity(): boolean {
  const ctx = requestContext.getStore();
  return !!(ctx?.userId && ctx.userId.length > 0);
}

/**
 * Get the current request tokens from context
 * @returns RequestContext or undefined if not in request context
 */
export function getRequestTokens(): RequestContext | undefined {
  return requestContext.getStore();
}

/**
 * Get the current request session ID from context.
 * Every MCP request has a unique sessionId; use for logging and transport isolation.
 * @returns Session ID or undefined if not in request context
 */
export function getSessionId(): string | undefined {
  return requestContext.getStore()?.sessionId;
}

/**
 * Get the current chat ID from context
 * @returns Chat ID or undefined if not available
 */
export function getChatId(): string | undefined {
  return requestContext.getStore()?.chatId;
}

/**
 * Get the current user ID from context
 * @returns User ID or undefined if not available
 */
export function getUserId(): string | undefined {
  return requestContext.getStore()?.userId;
}

/**
 * SECURITY: Get user ID with requirement check
 * Throws error if user ID is not available (for protected operations)
 *
 * @returns User ID
 * @throws Error if user ID is not available
 */
export function requireUserId(): string {
  const userId = getUserId();
  if (!userId) {
    throw new Error(
      'SECURITY: User identification required but not available. ' +
        'This operation requires authentication with a valid Microsoft Graph token.'
    );
  }
  return userId;
}

/**
 * SECURITY: Get user context for logging (anonymized)
 * Returns an object suitable for logging without exposing sensitive data.
 */
export function getSecureLogContext(): {
  sessionId?: string;
  hasAuth: boolean;
  hasUserId: boolean;
  userIdPrefix?: string;
  chatIdPrefix?: string;
  tokenHash?: string;
  professionProfile?: string;
} {
  const ctx = requestContext.getStore();
  return {
    sessionId: ctx?.sessionId,
    hasAuth: hasValidAuth(),
    hasUserId: hasUserIdentity(),
    userIdPrefix: ctx?.userId?.substring(0, 8),
    chatIdPrefix: ctx?.chatId?.substring(0, 8),
    tokenHash: ctx?.tokenHash,
    professionProfile: ctx?.userProfile?.professionProfile?.id,
  };
}

/**
 * Get the current user profile from context
 * @returns UserProfile or undefined if not available
 */
export function getUserProfile(): UserProfile | undefined {
  return requestContext.getStore()?.userProfile;
}

/**
 * Get the current profession profile from context
 * @returns ProfessionProfile or undefined if not available
 */
export function getProfessionProfile(): ProfessionProfile | undefined {
  return requestContext.getStore()?.userProfile?.professionProfile;
}

/**
 * Check if we have a profession profile in context
 * @returns True if profession profile is available
 */
export function hasProfessionProfile(): boolean {
  return !!getProfessionProfile();
}
