/**
 * Query Logger - Middleware to log all MCP tool calls
 *
 * Integrates with QueryStore to persist all user queries
 * for auditing and analytics purposes.
 */

import { getQueryStore } from './query-store.js';
import { getUserId, getChatId } from './request-context.js';
import logger from './logger.js';

export interface QueryLogContext {
  toolName: string;
  parameters: Record<string, unknown>;
  ipAddress?: string;
  userAgent?: string;
}

export interface QueryLogResult {
  success: boolean;
  responseSummary?: string;
  errorMessage?: string;
  durationMs: number;
}

/**
 * Log a query to the store
 */
export function logQuery(context: QueryLogContext, result: QueryLogResult): void {
  try {
    const queryStore = getQueryStore();
    const userId = getUserId();
    const chatId = getChatId();

    queryStore.storeQuery({
      userIdHash: queryStore.hashUserId(userId || 'anonymous'),
      chatId,
      toolName: context.toolName,
      parameters: context.parameters,
      responseSummary: result.responseSummary,
      success: result.success,
      errorMessage: result.errorMessage,
      durationMs: result.durationMs,
      ipAnonymized: context.ipAddress ? queryStore.anonymizeIp(context.ipAddress) : undefined,
      userAgent: context.userAgent,
    });
  } catch (error) {
    // Don't let logging errors break the application
    logger.error('Failed to log query:', error);
  }
}

/**
 * Wrap a tool handler to automatically log queries
 */
export function withQueryLogging<T>(
  toolName: string,
  handler: (params: Record<string, unknown>) => Promise<T>,
  getSummary?: (result: T) => string
): (params: Record<string, unknown>) => Promise<T> {
  return async (params: Record<string, unknown>): Promise<T> => {
    const startTime = Date.now();

    try {
      const result = await handler(params);
      const durationMs = Date.now() - startTime;

      logQuery(
        { toolName, parameters: params },
        {
          success: true,
          responseSummary: getSummary ? getSummary(result) : undefined,
          durationMs,
        }
      );

      return result;
    } catch (error) {
      const durationMs = Date.now() - startTime;

      logQuery(
        { toolName, parameters: params },
        {
          success: false,
          errorMessage: error instanceof Error ? error.message : 'Unknown error',
          durationMs,
        }
      );

      throw error;
    }
  };
}

export default logQuery;
