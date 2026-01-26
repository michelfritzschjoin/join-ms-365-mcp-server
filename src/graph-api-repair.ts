/**
 * Graph API Self-Repair Manager
 * Coordinates repair strategies to automatically fix Graph API request errors
 */

import logger from './logger.js';
import type { ParsedGraphApiError } from './errors.js';
import { parseGraphApiError } from './errors.js';
import type { RepairRequest, RepairResult, RepairStrategy } from './repair-strategies.js';
import { getRepairStrategies } from './repair-strategies.js';
import { getRepairConfig, type RepairConfig } from './repair-config.js';

/**
 * Repair history entry
 */
export interface RepairHistoryEntry {
  timestamp: number;
  endpoint: string;
  error: ParsedGraphApiError;
  strategy: string;
  success: boolean;
  message: string;
  metadata?: Record<string, unknown>;
}

/**
 * Graph API Repair Manager
 */
export class GraphApiRepairManager {
  private config: RepairConfig;
  private strategies: RepairStrategy[];
  private history: RepairHistoryEntry[] = [];
  private readonly maxHistorySize = 1000;

  constructor(config?: RepairConfig) {
    this.config = config || getRepairConfig();
    this.strategies = getRepairStrategies();
  }

  /**
   * Check if self-repair is enabled
   */
  isEnabled(): boolean {
    return this.config.enabled;
  }

  /**
   * Analyze error and determine applicable repair strategies
   */
  analyzeError(
    error: unknown,
    errorText?: string
  ): {
    parsedError: ParsedGraphApiError;
    applicableStrategies: RepairStrategy[];
  } {
    const parsedError = parseGraphApiError(error, errorText);
    const applicableStrategies = this.strategies.filter((strategy) => {
      // Check if strategy is enabled in config
      const strategyEnabled = this.isStrategyEnabled(strategy.name);
      if (!strategyEnabled) {
        return false;
      }
      return strategy.canRepair(parsedError);
    });

    logger.debug(`Error analysis: ${applicableStrategies.length} applicable strategies`, {
      errorCode: parsedError.errorCode,
      statusCode: parsedError.statusCode,
      strategies: applicableStrategies.map((s) => s.name),
    });

    return {
      parsedError,
      applicableStrategies,
    };
  }

  /**
   * Check if a strategy is enabled
   */
  private isStrategyEnabled(strategyName: string): boolean {
    switch (strategyName) {
      case 'endpoint':
        return this.config.endpointRepair;
      case 'parameter':
        return this.config.parameterRepair;
      case 'scope':
        return this.config.scopeRepair;
      case 'version':
        return this.config.versionRepair;
      case 'ratelimit':
        return this.config.rateLimitRepair;
      default:
        return this.config.enabledStrategies.includes(strategyName);
    }
  }

  /**
   * Attempt to repair a request using available strategies
   */
  async attemptRepair(
    request: RepairRequest,
    strategies?: RepairStrategy[]
  ): Promise<RepairResult | null> {
    if (!this.config.enabled) {
      return null;
    }

    const strategiesToTry =
      strategies || this.strategies.filter((s) => this.isStrategyEnabled(s.name));

    let currentRequest = { ...request };
    let finalResult: RepairResult | null = null;

    // Try strategies in order
    for (const strategy of strategiesToTry) {
      if (!strategy.canRepair(request.originalError)) {
        continue;
      }

      try {
        const result = await strategy.repair(currentRequest);

        if (result.success && result.repairedRequest) {
          // Update current request for next strategy in chain
          currentRequest = result.repairedRequest;
          finalResult = result;

          logger.info(`Repair strategy ${strategy.name} succeeded`, {
            endpoint: currentRequest.endpoint,
          });
        }
      } catch (error) {
        logger.error(`Repair strategy ${strategy.name} threw error: ${error}`);
      }
    }

    if (finalResult) {
      this.recordRepair(request, 'chained', finalResult);
      return finalResult;
    }

    return null;
  }

  /**
   * Record repair attempt in history
   */
  private recordRepair(request: RepairRequest, strategy: string, result: RepairResult): void {
    const entry: RepairHistoryEntry = {
      timestamp: Date.now(),
      endpoint: request.endpoint,
      error: request.originalError,
      strategy,
      success: result.success,
      message: result.message,
      metadata: result.metadata,
    };

    this.history.push(entry);

    // Limit history size
    if (this.history.length > this.maxHistorySize) {
      this.history = this.history.slice(-this.maxHistorySize);
    }
  }

  /**
   * Get repair history
   */
  getHistory(limit?: number): RepairHistoryEntry[] {
    if (limit) {
      return this.history.slice(-limit);
    }
    return [...this.history];
  }

  /**
   * Get repair statistics
   */
  getStatistics(): {
    totalRepairs: number;
    successfulRepairs: number;
    failedRepairs: number;
    successRate: number;
    byStrategy: Record<string, { total: number; successful: number }>;
  } {
    const total = this.history.length;
    const successful = this.history.filter((e) => e.success).length;
    const failed = total - successful;
    const successRate = total > 0 ? successful / total : 0;

    const byStrategy: Record<string, { total: number; successful: number }> = {};
    for (const entry of this.history) {
      if (!byStrategy[entry.strategy]) {
        byStrategy[entry.strategy] = { total: 0, successful: 0 };
      }
      byStrategy[entry.strategy].total++;
      if (entry.success) {
        byStrategy[entry.strategy].successful++;
      }
    }

    return {
      totalRepairs: total,
      successfulRepairs: successful,
      failedRepairs: failed,
      successRate,
      byStrategy,
    };
  }

  /**
   * Clear repair history
   */
  clearHistory(): void {
    this.history = [];
  }

  /**
   * Should retry after repair attempt
   */
  shouldRetry(attempt: number, maxAttempts: number): boolean {
    return attempt < maxAttempts;
  }

  /**
   * Create repair request from error and original request
   */
  createRepairRequest(
    endpoint: string,
    options: Record<string, unknown>,
    error: unknown,
    errorText?: string,
    params?: Record<string, unknown>
  ): RepairRequest {
    const parsedError = parseGraphApiError(error, errorText);

    return {
      endpoint,
      method: (options.method as string) || 'GET',
      options: {
        headers: options.headers as Record<string, string> | undefined,
        body: options.body as string | undefined,
        ...options,
      },
      params,
      originalError: parsedError,
    };
  }
}
