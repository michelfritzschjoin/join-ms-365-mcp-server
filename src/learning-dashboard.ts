/**
 * Learning Dashboard API for insights and statistics
 */

import logger from './logger.js';
import LearningSystem from './learning-system.js';
import LearningAnalytics, { type PerformanceMetrics } from './learning-analytics.js';
import PatternClusterer from './pattern-clusterer.js';
import KnowledgeBase from './knowledge-base.js';
import NLPEnhancer from './nlp-enhancer.js';

export interface LearningInsights {
  stats: PerformanceMetrics;
  topPatterns: Array<{
    pattern: string;
    successCount: number;
    confidence: number;
    lastUsed: Date;
  }>;
  confidenceDistribution: {
    high: number;
    medium: number;
    low: number;
  };
  toolUsageStats: Array<{
    tool: string;
    successRate: number;
    averageResults: number;
    totalUses: number;
  }>;
  performanceTrends: Array<{
    date: string;
    successRate: number;
    averageResults: number;
    averageConfidence: number;
  }>;
  clusters?: Array<{
    id: string;
    patternCount: number;
    averageConfidence: number;
    representativePattern: string;
  }>;
}

export class LearningDashboard {
  private learningSystem: LearningSystem;
  private analytics: LearningAnalytics;
  private patternClusterer: PatternClusterer;

  constructor(
    learningSystem: LearningSystem,
    knowledgeBase: KnowledgeBase,
    nlpEnhancer?: NLPEnhancer
  ) {
    this.learningSystem = learningSystem;
    this.analytics = new LearningAnalytics(knowledgeBase);
    this.patternClusterer = new PatternClusterer(knowledgeBase, nlpEnhancer);
  }

  /**
   * Get learning statistics
   */
  getLearningStats(): PerformanceMetrics {
    return this.analytics.calculatePerformanceMetrics();
  }

  /**
   * Get top patterns
   */
  getTopPatterns(limit: number = 10): Array<{
    pattern: string;
    successCount: number;
    confidence: number;
    lastUsed: Date;
  }> {
    const metrics = this.analytics.calculatePerformanceMetrics();
    return metrics.topPatterns.slice(0, limit);
  }

  /**
   * Get confidence distribution
   */
  getConfidenceDistribution(): {
    high: number;
    medium: number;
    low: number;
  } {
    const metrics = this.analytics.calculatePerformanceMetrics();
    return metrics.confidenceDistribution;
  }

  /**
   * Get tool usage statistics
   */
  getToolUsageStats(limit: number = 10): Array<{
    tool: string;
    successRate: number;
    averageResults: number;
    totalUses: number;
  }> {
    const metrics = this.analytics.calculatePerformanceMetrics();
    return metrics.toolUsageStats.slice(0, limit);
  }

  /**
   * Get performance trends over time
   */
  getPerformanceTrends(days: number = 30): Array<{
    date: string;
    successRate: number;
    averageResults: number;
    averageConfidence: number;
  }> {
    return this.analytics.getPerformanceTrends(days);
  }

  /**
   * Get all learning insights
   */
  getLearningInsights(): LearningInsights {
    const stats = this.getLearningStats();
    const clusters = this.patternClusterer.getAllClusters();

    return {
      stats,
      topPatterns: stats.topPatterns,
      confidenceDistribution: stats.confidenceDistribution,
      toolUsageStats: stats.toolUsageStats,
      performanceTrends: this.getPerformanceTrends(30),
      clusters: clusters.map((c) => ({
        id: c.id,
        patternCount: c.patterns.length,
        averageConfidence: c.averageConfidence,
        representativePattern: c.representativePattern,
      })),
    };
  }
}

export default LearningDashboard;
