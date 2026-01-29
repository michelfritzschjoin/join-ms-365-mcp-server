/**
 * Token Controller
 *
 * Manages token budgets for LLM-friendly output formatting.
 * Optimizes response size while maintaining information quality.
 */

import type { ResponseData } from '../i18n/response-templates.js';

/**
 * Token budget configuration
 */
export interface TokenBudget {
  /** Total token budget */
  total: number;
  /** Budget for summary */
  summary: number;
  /** Budget for facts */
  facts: number;
  /** Budget for timeline */
  timeline: number;
  /** Budget for recommendations */
  recommendations: number;
  /** Reserved for metadata/formatting */
  overhead: number;
}

/**
 * Token controller configuration
 */
export interface TokenControllerConfig {
  /** Maximum tokens for response */
  maxTokens?: number;
  /** Chars per token estimate */
  charsPerToken?: number;
  /** Enable aggressive compression */
  aggressiveCompression?: boolean;
}

/**
 * Compact result with token info
 */
export interface CompactResult {
  data: ResponseData;
  originalTokens: number;
  compactTokens: number;
  compressionRatio: number;
}

/**
 * Default token budgets
 */
const DEFAULT_BUDGETS: Record<'minimal' | 'normal' | 'expanded', TokenBudget> = {
  minimal: {
    total: 500,
    summary: 100,
    facts: 200,
    timeline: 100,
    recommendations: 50,
    overhead: 50,
  },
  normal: {
    total: 1500,
    summary: 200,
    facts: 600,
    timeline: 400,
    recommendations: 200,
    overhead: 100,
  },
  expanded: {
    total: 3000,
    summary: 400,
    facts: 1200,
    timeline: 800,
    recommendations: 400,
    overhead: 200,
  },
};

/**
 * TokenController - Manages token budgets
 */
export class TokenController {
  private config: Required<TokenControllerConfig>;
  private budget: TokenBudget;

  constructor(config: TokenControllerConfig = {}) {
    this.config = {
      maxTokens: config.maxTokens ?? 1500,
      charsPerToken: config.charsPerToken ?? 4,
      aggressiveCompression: config.aggressiveCompression ?? false,
    };

    // Select budget based on max tokens
    if (this.config.maxTokens <= 500) {
      this.budget = { ...DEFAULT_BUDGETS.minimal };
    } else if (this.config.maxTokens <= 1500) {
      this.budget = { ...DEFAULT_BUDGETS.normal };
    } else {
      this.budget = { ...DEFAULT_BUDGETS.expanded };
    }

    // Scale budget to match max tokens
    this.scaleBudget(this.config.maxTokens);
  }

  /**
   * Scale budget to match target
   */
  private scaleBudget(target: number): void {
    const currentTotal = this.budget.total;
    const scale = target / currentTotal;

    this.budget = {
      total: target,
      summary: Math.floor(this.budget.summary * scale),
      facts: Math.floor(this.budget.facts * scale),
      timeline: Math.floor(this.budget.timeline * scale),
      recommendations: Math.floor(this.budget.recommendations * scale),
      overhead: Math.floor(this.budget.overhead * scale),
    };
  }

  /**
   * Optimize response data to fit token budget
   */
  optimize(data: ResponseData): ResponseData {
    const original = this.estimateTokens(JSON.stringify(data));

    // If already within budget, return as-is
    if (original <= this.config.maxTokens && !this.config.aggressiveCompression) {
      return data;
    }

    const optimized: ResponseData = {
      ...data,
      summary: this.truncateToTokens(data.summary, this.budget.summary),
      facts: this.optimizeFacts(data.facts, this.budget.facts),
      timeline: data.timeline
        ? this.optimizeTimeline(data.timeline, this.budget.timeline)
        : undefined,
      recommendations: data.recommendations
        ? this.optimizeRecommendations(data.recommendations, this.budget.recommendations)
        : undefined,
    };

    return optimized;
  }

  /**
   * Get compact result with metrics
   */
  compactWithMetrics(data: ResponseData): CompactResult {
    const originalStr = JSON.stringify(data);
    const originalTokens = this.estimateTokens(originalStr);

    const optimized = this.optimize(data);
    const compactStr = JSON.stringify(optimized);
    const compactTokens = this.estimateTokens(compactStr);

    return {
      data: optimized,
      originalTokens,
      compactTokens,
      compressionRatio: originalTokens > 0 ? compactTokens / originalTokens : 1,
    };
  }

  /**
   * Optimize facts array
   */
  private optimizeFacts(facts: string[], tokenBudget: number): string[] {
    if (facts.length === 0) return facts;

    const result: string[] = [];
    let usedTokens = 0;
    const tokensPerFact = Math.floor(tokenBudget / Math.min(facts.length, 5));

    for (const fact of facts) {
      const factTokens = this.estimateTokens(fact);

      if (usedTokens + factTokens > tokenBudget) {
        // Try to truncate and add
        const remaining = tokenBudget - usedTokens;
        if (remaining > 20) {
          result.push(this.truncateToTokens(fact, remaining));
        }
        break;
      }

      result.push(
        fact.length > tokensPerFact * this.config.charsPerToken
          ? this.truncateToTokens(fact, tokensPerFact)
          : fact
      );
      usedTokens += factTokens;
    }

    return result;
  }

  /**
   * Optimize timeline
   */
  private optimizeTimeline(
    timeline: NonNullable<ResponseData['timeline']>,
    tokenBudget: number
  ): ResponseData['timeline'] {
    if (!timeline || timeline.length === 0) return timeline;

    const result: NonNullable<ResponseData['timeline']> = [];
    let usedTokens = 0;

    // Limit to most recent entries that fit
    for (const entry of timeline) {
      const entryTokens = this.estimateTokens(entry.title) + 10; // ~10 tokens for date/type

      if (usedTokens + entryTokens > tokenBudget) {
        break;
      }

      // Truncate title if needed
      const maxTitleTokens = Math.floor(tokenBudget / Math.min(timeline.length, 5)) - 10;
      result.push({
        ...entry,
        title: this.truncateToTokens(entry.title, maxTitleTokens),
      });

      usedTokens += entryTokens;
    }

    return result;
  }

  /**
   * Optimize recommendations
   */
  private optimizeRecommendations(recommendations: string[], tokenBudget: number): string[] {
    if (recommendations.length === 0) return recommendations;

    // Limit to top 3 recommendations
    const limited = recommendations.slice(0, 3);
    const tokensPerRec = Math.floor(tokenBudget / limited.length);

    return limited.map((rec) => this.truncateToTokens(rec, tokensPerRec));
  }

  /**
   * Truncate text to fit token budget
   */
  truncateToTokens(text: string, tokenBudget: number): string {
    const maxChars = tokenBudget * this.config.charsPerToken;

    if (text.length <= maxChars) {
      return text;
    }

    // Truncate with ellipsis
    const truncated = text.substring(0, maxChars - 3);
    // Try to cut at word boundary
    const lastSpace = truncated.lastIndexOf(' ');
    if (lastSpace > maxChars * 0.7) {
      return truncated.substring(0, lastSpace) + '...';
    }
    return truncated + '...';
  }

  /**
   * Estimate token count for text
   */
  estimateTokens(text: string): number {
    if (!text) return 0;
    return Math.ceil(text.length / this.config.charsPerToken);
  }

  /**
   * Check if text fits in budget
   */
  fitsInBudget(text: string, budget?: number): boolean {
    const tokens = this.estimateTokens(text);
    return tokens <= (budget ?? this.config.maxTokens);
  }

  /**
   * Get remaining budget
   */
  getRemainingBudget(usedTokens: number): number {
    return Math.max(0, this.config.maxTokens - usedTokens);
  }

  /**
   * Get current budget allocation
   */
  getBudget(): TokenBudget {
    return { ...this.budget };
  }

  /**
   * Set custom budget
   */
  setBudget(budget: Partial<TokenBudget>): void {
    this.budget = {
      ...this.budget,
      ...budget,
      total: budget.total ?? this.budget.total,
    };
  }

  /**
   * Get config
   */
  getConfig(): Required<TokenControllerConfig> {
    return { ...this.config };
  }

  /**
   * Create compact summary for very limited budgets
   */
  createMinimalSummary(data: ResponseData): string {
    const parts: string[] = [];

    // One-line summary
    const shortSummary = this.truncateToTokens(data.summary, 50);
    parts.push(`📋 ${shortSummary}`);

    // Compact confidence
    const bar = '█'.repeat(Math.round(data.confidence * 5));
    parts.push(`[${bar}] ${Math.round(data.confidence * 100)}%`);

    // Top fact only
    if (data.facts.length > 0) {
      parts.push(`• ${this.truncateToTokens(data.facts[0], 30)}`);
    }

    return parts.join('\n');
  }

  /**
   * Format data with specific token target
   */
  formatToTarget(data: ResponseData, targetTokens: number): string {
    const controller = new TokenController({ maxTokens: targetTokens });
    const optimized = controller.optimize(data);

    // Build minimal output
    const lines: string[] = [];
    lines.push(optimized.summary);

    if (optimized.facts.length > 0) {
      lines.push('');
      optimized.facts.forEach((f) => lines.push(`• ${f}`));
    }

    return lines.join('\n');
  }
}

export default TokenController;
