/**
 * Learning Analytics for tracking and analyzing learning system performance
 */

import logger from './logger.js';
import KnowledgeBase from './knowledge-base.js';
import type { KnowledgeBaseData } from './knowledge-base.js';

export interface PerformanceMetrics {
  totalQueries: number;
  successfulQueries: number;
  failedQueries: number;
  successRate: number;
  averageResultsPerQuery: number;
  averageConfidence: number;
  queryImprovementRate?: number;
  topPatterns: Array<{
    pattern: string;
    successCount: number;
    confidence: number;
    lastUsed: Date;
  }>;
  toolUsageStats: Array<{
    tool: string;
    successRate: number;
    averageResults: number;
    totalUses: number;
  }>;
  confidenceDistribution: {
    high: number; // >= 0.7
    medium: number; // 0.4 - 0.7
    low: number; // < 0.4
  };
}

export class LearningAnalytics {
  private knowledgeBase: KnowledgeBase;

  constructor(knowledgeBase: KnowledgeBase) {
    this.knowledgeBase = knowledgeBase;
  }

  /**
   * Calculate performance metrics
   */
  calculatePerformanceMetrics(): PerformanceMetrics {
    const data = this.knowledgeBase.getAllData();

    const totalQueries = Object.keys(data.successfulQueries).length;
    const successfulQueries = Object.values(data.queryPatterns).reduce(
      (sum, p) => sum + p.successCount,
      0
    );
    const failedQueries = totalQueries - successfulQueries;
    const successRate = totalQueries > 0 ? successfulQueries / totalQueries : 0;

    // Calculate average results per query
    const averageResultsPerQuery =
      totalQueries > 0
        ? Object.values(data.successfulQueries).reduce((sum, q) => sum + q.results, 0) /
          totalQueries
        : 0;

    // Calculate average confidence
    const confidenceScores = Object.values(data.confidenceScores);
    const averageConfidence =
      confidenceScores.length > 0
        ? confidenceScores.reduce((sum, score) => sum + score, 0) / confidenceScores.length
        : 0.5; // Default to 0.5 if no confidence scores

    // Get top patterns
    const topPatterns = Object.entries(data.queryPatterns)
      .map(([key, pattern]) => ({
        pattern: key,
        successCount: pattern.successCount,
        confidence: this.knowledgeBase.getConfidenceScore(key),
        lastUsed: pattern.lastUsed,
      }))
      .sort((a, b) => {
        // Sort by confidence first, then by success count
        if (Math.abs(a.confidence - b.confidence) > 0.1) {
          return b.confidence - a.confidence;
        }
        return b.successCount - a.successCount;
      })
      .slice(0, 10);

    // Get tool usage stats
    const toolUsageStats = Object.values(data.toolUsagePatterns)
      .map((pattern) => {
        const totalUses = pattern.successCount + pattern.failureCount;
        return {
          tool: pattern.toolName,
          successRate: totalUses > 0 ? pattern.successCount / totalUses : 0,
          averageResults: pattern.averageResults || 0,
          totalUses,
        };
      })
      .sort((a, b) => b.successRate - a.successRate);

    // Calculate confidence distribution
    const confidenceDistribution = {
      high: 0,
      medium: 0,
      low: 0,
    };

    for (const score of confidenceScores) {
      if (score >= 0.7) {
        confidenceDistribution.high++;
      } else if (score >= 0.4) {
        confidenceDistribution.medium++;
      } else {
        confidenceDistribution.low++;
      }
    }

    // Calculate query improvement rate (simplified - compares recent vs older patterns)
    const queryImprovementRate = this.calculateQueryImprovementRate(data);

    return {
      totalQueries,
      successfulQueries,
      failedQueries,
      successRate,
      averageResultsPerQuery,
      averageConfidence,
      queryImprovementRate,
      topPatterns,
      toolUsageStats,
      confidenceDistribution,
    };
  }

  /**
   * Calculate query improvement rate
   * Compares success rates of recent queries vs older queries
   */
  private calculateQueryImprovementRate(data: KnowledgeBaseData): number | undefined {
    const now = new Date();
    const thirtyDaysAgo = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);

    const recentPatterns: Array<{ successCount: number; totalUses: number }> = [];
    const olderPatterns: Array<{ successCount: number; totalUses: number }> = [];

    for (const pattern of Object.values(data.queryPatterns)) {
      const patternData = {
        successCount: pattern.successCount,
        totalUses: pattern.successCount, // Simplified - assume all uses are tracked
      };

      if (pattern.lastUsed >= thirtyDaysAgo) {
        recentPatterns.push(patternData);
      } else {
        olderPatterns.push(patternData);
      }
    }

    if (recentPatterns.length === 0 || olderPatterns.length === 0) {
      return undefined;
    }

    const recentSuccessRate =
      recentPatterns.reduce((sum, p) => sum + p.successCount / Math.max(1, p.totalUses), 0) /
      recentPatterns.length;

    const olderSuccessRate =
      olderPatterns.reduce((sum, p) => sum + p.successCount / Math.max(1, p.totalUses), 0) /
      olderPatterns.length;

    if (olderSuccessRate === 0) {
      return recentSuccessRate > 0 ? 100 : 0;
    }

    return ((recentSuccessRate - olderSuccessRate) / olderSuccessRate) * 100;
  }

  /**
   * Get confidence score for a pattern
   * Calculates based on success rate, usage count, recency, and user feedback
   */
  calculateConfidenceScore(
    patternKey: string,
    successCount: number,
    totalUses: number,
    lastUsed: Date,
    userFeedbackCount?: number
  ): number {
    // Base confidence from success rate
    const successRate = totalUses > 0 ? successCount / totalUses : 0.5;
    let confidence = successRate;

    // Boost confidence based on usage count (more uses = more reliable)
    const usageBoost = Math.min(0.2, Math.log10(totalUses + 1) * 0.05);
    confidence += usageBoost;

    // Recency boost (recent patterns are more relevant)
    const now = new Date();
    const daysSinceLastUse = (now.getTime() - lastUsed.getTime()) / (24 * 60 * 60 * 1000);
    const recencyBoost = Math.max(0, 0.1 * (1 - daysSinceLastUse / 90)); // Decay over 90 days
    confidence += recencyBoost;

    // User feedback boost (explicit feedback is more valuable)
    if (userFeedbackCount && userFeedbackCount > 0) {
      const feedbackBoost = Math.min(0.15, userFeedbackCount * 0.05);
      confidence += feedbackBoost;
    }

    // Clamp between 0 and 1
    return Math.max(0, Math.min(1, confidence));
  }

  /**
   * Update confidence scores for all patterns
   */
  updateAllConfidenceScores(): void {
    const data = this.knowledgeBase.getAllData();
    const now = new Date();

    // Update confidence for query patterns
    for (const [key, pattern] of Object.entries(data.queryPatterns)) {
      const totalUses = pattern.successCount; // Simplified
      const userFeedback = data.userFeedback[key]?.length || 0;

      const confidence = this.calculateConfidenceScore(
        key,
        pattern.successCount,
        totalUses,
        pattern.lastUsed,
        userFeedback
      );

      this.knowledgeBase.calculateConfidence(key, confidence);
    }

    // Update confidence for entity mappings
    for (const [key, mapping] of Object.entries(data.entityMappings)) {
      const totalUses = mapping.successCount; // Simplified
      const userFeedback = data.userFeedback[key]?.length || 0;

      const confidence = this.calculateConfidenceScore(
        key,
        mapping.successCount,
        totalUses,
        mapping.lastUsed,
        userFeedback
      );

      this.knowledgeBase.calculateConfidence(key, confidence);
    }

    logger.debug('Updated confidence scores for all patterns');
  }

  /**
   * Export analytics data as JSON
   */
  exportAnalytics(): string {
    const metrics = this.calculatePerformanceMetrics();
    return JSON.stringify(metrics, null, 2);
  }

  /**
   * Get performance trends over time
   */
  getPerformanceTrends(days: number = 30): {
    date: string;
    successRate: number;
    averageResults: number;
    averageConfidence: number;
  }[] {
    // Simplified implementation - in a real system, you'd track metrics over time
    const data = this.knowledgeBase.getAllData();
    const trends: Array<{
      date: string;
      successRate: number;
      averageResults: number;
      averageConfidence: number;
    }> = [];

    const now = new Date();
    for (let i = days; i >= 0; i--) {
      const date = new Date(now.getTime() - i * 24 * 60 * 60 * 1000);

      // Filter patterns by date (simplified - use lastUsed)
      const patternsForDate = Object.values(data.queryPatterns).filter(
        (p) => p.lastUsed <= date && p.lastUsed >= new Date(date.getTime() - 24 * 60 * 60 * 1000)
      );

      if (patternsForDate.length > 0) {
        const successRate =
          patternsForDate.reduce((sum, p) => sum + p.successCount, 0) / patternsForDate.length;
        const averageResults = 0; // Would need to track this over time
        const averageConfidence = 0.5; // Would need to track this over time

        trends.push({
          date: date.toISOString().split('T')[0],
          successRate,
          averageResults,
          averageConfidence,
        });
      }
    }

    return trends;
  }
}

export default LearningAnalytics;
