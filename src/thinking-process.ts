/**
 * Thinking Process Manager
 *
 * Tracks and formats reasoning steps during tool execution.
 * Provides transparency into how the MCP server processes requests.
 *
 * Features:
 * - Tracks thinking steps with timestamps and categories
 * - Supports multiple detail levels (minimal, normal, verbose)
 * - Formats output for OpenWebUI frontend display
 * - Thread-safe via AsyncLocalStorage integration
 */

import logger from './logger.js';

/**
 * Type of thinking step
 */
export type ThinkingStepType = 'reasoning' | 'decision' | 'action' | 'result' | 'error' | 'info';

/**
 * Category of thinking step
 */
export type ThinkingCategory =
  | 'intent'
  | 'parameters'
  | 'validation'
  | 'api-call'
  | 'processing'
  | 'formatting'
  | 'error-handling'
  | 'optimization'
  | 'search'
  | 'aggregation';

/**
 * Thinking level configuration
 */
export type ThinkingLevel = 'minimal' | 'normal' | 'verbose';

/**
 * Individual thinking step
 */
export interface ThinkingStep {
  /** Unix timestamp when step occurred */
  timestamp: number;
  /** Type of step */
  type: ThinkingStepType;
  /** Category of step */
  category: ThinkingCategory;
  /** Human-readable message */
  message: string;
  /** Additional details (only included in verbose mode) */
  details?: Record<string, unknown>;
  /** Duration in ms (for action steps) */
  duration?: number;
  /** Icon for frontend display */
  icon?: string;
}

/**
 * Complete thinking process result
 */
export interface ThinkingResult {
  /** All thinking steps */
  steps: ThinkingStep[];
  /** Summary of the thinking process */
  summary: string;
  /** Total duration of thinking process */
  totalDuration: number;
  /** Tool name being executed */
  toolName: string;
  /** Thinking level used */
  level: ThinkingLevel;
}

/**
 * Check if thinking process is enabled
 */
export function isThinkingEnabled(): boolean {
  const enabled = process.env.MS365_MCP_THINKING_ENABLED;
  // Default to true if not set
  return enabled !== 'false' && enabled !== '0';
}

/**
 * Get the thinking level from environment
 */
export function getThinkingLevel(): ThinkingLevel {
  const level = process.env.MS365_MCP_THINKING_LEVEL?.toLowerCase();
  if (level === 'minimal' || level === 'normal' || level === 'verbose') {
    return level;
  }
  return 'normal'; // Default
}

/**
 * Icon mapping for step types
 */
const STEP_ICONS: Record<ThinkingStepType, string> = {
  reasoning: '🧠',
  decision: '🎯',
  action: '⚡',
  result: '✅',
  error: '❌',
  info: 'ℹ️',
};

/**
 * Icon mapping for categories
 */
const CATEGORY_ICONS: Record<ThinkingCategory, string> = {
  intent: '🎯',
  parameters: '📋',
  validation: '✓',
  'api-call': '🌐',
  processing: '⚙️',
  formatting: '📝',
  'error-handling': '🔧',
  optimization: '⚡',
  search: '🔍',
  aggregation: '📊',
};

/**
 * ThinkingProcessManager - Tracks reasoning steps during tool execution
 */
export class ThinkingProcessManager {
  private steps: ThinkingStep[] = [];
  private startTime: number;
  private toolName: string;
  private level: ThinkingLevel;
  private enabled: boolean;
  private lastActionStart: number | null = null;

  constructor(toolName: string) {
    this.toolName = toolName;
    this.startTime = Date.now();
    this.level = getThinkingLevel();
    this.enabled = isThinkingEnabled();
  }

  /**
   * Check if thinking is enabled
   */
  isEnabled(): boolean {
    return this.enabled;
  }

  /**
   * Add a reasoning step (explaining why)
   */
  addReasoning(
    category: ThinkingCategory,
    message: string,
    details?: Record<string, unknown>
  ): void {
    this.addStep('reasoning', category, message, details);
  }

  /**
   * Add a decision step (what was decided)
   */
  addDecision(
    category: ThinkingCategory,
    message: string,
    details?: Record<string, unknown>
  ): void {
    this.addStep('decision', category, message, details);
  }

  /**
   * Start tracking an action (records start time)
   */
  startAction(category: ThinkingCategory, message: string): void {
    this.lastActionStart = Date.now();
    this.addStep('action', category, message);
  }

  /**
   * Complete an action (includes duration)
   */
  completeAction(
    category: ThinkingCategory,
    message: string,
    details?: Record<string, unknown>
  ): void {
    const duration = this.lastActionStart ? Date.now() - this.lastActionStart : undefined;
    this.lastActionStart = null;
    this.addStep('result', category, message, details, duration);
  }

  /**
   * Add an info step
   */
  addInfo(category: ThinkingCategory, message: string, details?: Record<string, unknown>): void {
    this.addStep('info', category, message, details);
  }

  /**
   * Add an error step
   */
  addError(category: ThinkingCategory, message: string, details?: Record<string, unknown>): void {
    this.addStep('error', category, message, details);
  }

  /**
   * Add a generic step
   */
  addStep(
    type: ThinkingStepType,
    category: ThinkingCategory,
    message: string,
    details?: Record<string, unknown>,
    duration?: number
  ): void {
    if (!this.enabled) {
      return;
    }

    // Filter based on level
    if (this.level === 'minimal' && type !== 'decision' && type !== 'result' && type !== 'error') {
      return;
    }

    const step: ThinkingStep = {
      timestamp: Date.now(),
      type,
      category,
      message,
      icon: STEP_ICONS[type] || CATEGORY_ICONS[category],
    };

    // Include details only in verbose mode
    if (this.level === 'verbose' && details) {
      step.details = this.sanitizeDetails(details);
    }

    if (duration !== undefined) {
      step.duration = duration;
    }

    this.steps.push(step);

    // Log for debugging
    logger.debug(`[Thinking] ${step.icon} ${message}`, {
      type,
      category,
      duration,
    });
  }

  /**
   * Sanitize details to remove sensitive information
   */
  private sanitizeDetails(details: Record<string, unknown>): Record<string, unknown> {
    const sensitiveKeys = [
      'password',
      'token',
      'secret',
      'key',
      'authorization',
      'bearer',
      'credential',
      'apikey',
      'api_key',
      'access_token',
      'refresh_token',
      'client_secret',
    ];

    const sanitized: Record<string, unknown> = {};

    for (const [key, value] of Object.entries(details)) {
      const lowerKey = key.toLowerCase();
      if (sensitiveKeys.some((sensitive) => lowerKey.includes(sensitive))) {
        sanitized[key] = '[REDACTED]';
      } else if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
        sanitized[key] = this.sanitizeDetails(value as Record<string, unknown>);
      } else {
        sanitized[key] = value;
      }
    }

    return sanitized;
  }

  /**
   * Generate a summary of the thinking process
   */
  private generateSummary(): string {
    const actionCount = this.steps.filter((s) => s.type === 'action').length;
    const decisionCount = this.steps.filter((s) => s.type === 'decision').length;
    const errorCount = this.steps.filter((s) => s.type === 'error').length;
    const totalDuration = Date.now() - this.startTime;

    if (errorCount > 0) {
      return `Tool "${this.toolName}" encountered ${errorCount} error(s) during execution (${totalDuration}ms)`;
    }

    return `Tool "${this.toolName}" executed ${actionCount} action(s) with ${decisionCount} decision(s) in ${totalDuration}ms`;
  }

  /**
   * Finish the thinking process and return the result
   */
  finish(): ThinkingResult {
    const totalDuration = Date.now() - this.startTime;

    return {
      steps: this.steps,
      summary: this.generateSummary(),
      totalDuration,
      toolName: this.toolName,
      level: this.level,
    };
  }

  /**
   * Format thinking result for inclusion in tool response
   * Returns a structured format that OpenWebUI can display
   */
  formatForResponse(): Record<string, unknown> {
    if (!this.enabled || this.steps.length === 0) {
      return {};
    }

    const result = this.finish();

    // Create markdown-formatted thinking process
    const markdownSteps = this.steps.map((step) => {
      const icon = step.icon || '';
      const duration = step.duration ? ` (${step.duration}ms)` : '';
      return `${icon} **${step.category}**: ${step.message}${duration}`;
    });

    return {
      thinking: {
        enabled: true,
        level: this.level,
        summary: result.summary,
        totalDuration: result.totalDuration,
        stepCount: this.steps.length,
        // Structured steps for programmatic access
        steps: this.steps.map((step) => ({
          type: step.type,
          category: step.category,
          message: step.message,
          duration: step.duration,
          icon: step.icon,
          ...(step.details ? { details: step.details } : {}),
        })),
        // Markdown for display in OpenWebUI
        markdown: markdownSteps.join('\n'),
      },
    };
  }

  /**
   * Get steps formatted as processing steps array (for backward compatibility)
   */
  getProcessingSteps(): string[] {
    return this.steps.map((step) => {
      const icon = step.icon || '';
      const duration = step.duration ? ` (${step.duration}ms)` : '';
      return `${icon} ${step.message}${duration}`;
    });
  }
}

/**
 * Create a new ThinkingProcessManager for a tool
 */
export function createThinkingProcess(toolName: string): ThinkingProcessManager {
  return new ThinkingProcessManager(toolName);
}

/**
 * Add thinking steps to a response string
 * Used by Super-Tools for simple thinking process output
 *
 * @param result - The actual result data (JSON string or plain text)
 * @param thinkingSteps - Array of thinking step messages
 * @returns Formatted response with thinking process prepended
 */
export function addThinkingToResponse(result: string, thinkingSteps: string[]): string {
  if (!isThinkingEnabled() || thinkingSteps.length === 0) {
    return result;
  }

  const level = getThinkingLevel();

  // Format thinking steps based on level
  let thinkingOutput = '';

  if (level === 'minimal') {
    // Just show the last step
    thinkingOutput = `💭 ${thinkingSteps[thinkingSteps.length - 1]}`;
  } else if (level === 'normal') {
    // Show all steps with icons
    thinkingOutput = thinkingSteps.map((step) => `💭 ${step}`).join('\n');
  } else {
    // Verbose: show with timestamps and full detail
    const timestamp = new Date().toISOString();
    thinkingOutput =
      `🧠 Thinking Process (${timestamp}):\n` + thinkingSteps.map((step, i) => `  ${i + 1}. ${step}`).join('\n');
  }

  // Combine thinking with result
  return `${thinkingOutput}\n\n---\n\n${result}`;
}

export default ThinkingProcessManager;
