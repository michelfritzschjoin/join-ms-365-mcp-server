import { AsyncLocalStorage } from 'node:async_hooks';

/**
 * Request context that is passed through async operations
 * Contains authentication tokens and chat session information
 */
export interface RequestContext {
  /** Microsoft Graph API access token */
  accessToken: string;
  /** Microsoft Graph API refresh token */
  refreshToken?: string;
  /** OpenWebUI Chat ID for per-chat memory */
  chatId?: string;
  /** User ID extracted from token for user-scoped memory */
  userId?: string;
}

/**
 * AsyncLocalStorage for request context propagation
 */
export const requestContext = new AsyncLocalStorage<RequestContext>();

/**
 * Get the current request tokens from context
 * @returns RequestContext or undefined if not in request context
 */
export function getRequestTokens(): RequestContext | undefined {
  return requestContext.getStore();
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
