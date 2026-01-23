/**
 * Pattern Clusterer for grouping similar query patterns
 */

import logger from './logger.js';
import KnowledgeBase from './knowledge-base.js';
import NLPEnhancer from './nlp-enhancer.js';

export interface PatternCluster {
  id: string;
  patterns: string[];
  representativePattern: string;
  averageConfidence: number;
  totalSuccessCount: number;
  lastUsed: Date;
}

export class PatternClusterer {
  private knowledgeBase: KnowledgeBase;
  private nlpEnhancer: NLPEnhancer;
  private readonly enabled: boolean;
  private readonly similarityThreshold: number;

  constructor(knowledgeBase: KnowledgeBase, nlpEnhancer?: NLPEnhancer) {
    this.knowledgeBase = knowledgeBase;
    this.nlpEnhancer = nlpEnhancer || new NLPEnhancer();
    this.enabled =
      process.env.MS365_MCP_LEARNING_CLUSTER_ENABLED === 'true' ||
      process.env.MS365_MCP_LEARNING_CLUSTER_ENABLED !== 'false';
    this.similarityThreshold = parseFloat(
      process.env.MS365_MCP_LEARNING_CLUSTER_THRESHOLD || '0.7'
    );
  }

  /**
   * Cluster patterns using hierarchical clustering
   */
  clusterPatterns(): Record<string, string[]> {
    if (!this.enabled) {
      return {};
    }

    const data = this.knowledgeBase.getAllData();
    const patterns = Object.keys(data.queryPatterns);

    if (patterns.length === 0) {
      return {};
    }

    // Use knowledge base clustering method
    this.knowledgeBase.clusterPatterns(this.similarityThreshold);
    return this.knowledgeBase.getAllData().patternClusters;
  }

  /**
   * Calculate similarity between two patterns
   * Uses multiple similarity metrics
   */
  calculateSimilarity(pattern1: string, pattern2: string): number {
    // Use NLP enhancer for better similarity calculation
    if (this.nlpEnhancer) {
      return this.nlpEnhancer.calculateSimilarity(pattern1, pattern2);
    }

    // Fallback to simple Jaccard similarity
    return this.jaccardSimilarity(pattern1, pattern2);
  }

  /**
   * Jaccard similarity between two patterns
   */
  private jaccardSimilarity(pattern1: string, pattern2: string): number {
    const words1 = new Set(pattern1.toLowerCase().split(/\s+/));
    const words2 = new Set(pattern2.toLowerCase().split(/\s+/));

    const intersection = new Set([...words1].filter((x) => words2.has(x)));
    const union = new Set([...words1, ...words2]);

    return union.size > 0 ? intersection.size / union.size : 0;
  }

  /**
   * Get cluster information
   */
  getClusterInfo(clusterId: string): PatternCluster | null {
    const data = this.knowledgeBase.getAllData();
    const clusterPatterns = data.patternClusters[clusterId];

    if (!clusterPatterns || clusterPatterns.length === 0) {
      return null;
    }

    // Calculate cluster statistics
    let totalSuccessCount = 0;
    let totalConfidence = 0;
    let latestUsed = new Date(0);
    let representativePattern = clusterPatterns[0];

    for (const patternKey of clusterPatterns) {
      const pattern = data.queryPatterns[patternKey];
      if (pattern) {
        totalSuccessCount += pattern.successCount;
        totalConfidence += this.knowledgeBase.getConfidenceScore(patternKey);
        if (pattern.lastUsed > latestUsed) {
          latestUsed = pattern.lastUsed;
          representativePattern = patternKey;
        }
      }
    }

    const averageConfidence =
      clusterPatterns.length > 0 ? totalConfidence / clusterPatterns.length : 0;

    return {
      id: clusterId,
      patterns: clusterPatterns,
      representativePattern,
      averageConfidence,
      totalSuccessCount,
      lastUsed: latestUsed,
    };
  }

  /**
   * Get all clusters with their information
   */
  getAllClusters(): PatternCluster[] {
    const clusters = this.clusterPatterns();
    const clusterInfos: PatternCluster[] = [];

    for (const clusterId of Object.keys(clusters)) {
      const info = this.getClusterInfo(clusterId);
      if (info) {
        clusterInfos.push(info);
      }
    }

    // Sort by average confidence and success count
    clusterInfos.sort((a, b) => {
      if (Math.abs(a.averageConfidence - b.averageConfidence) > 0.1) {
        return b.averageConfidence - a.averageConfidence;
      }
      return b.totalSuccessCount - a.totalSuccessCount;
    });

    return clusterInfos;
  }

  /**
   * Find cluster for a pattern
   */
  findClusterForPattern(pattern: string): PatternCluster | null {
    const clusters = this.clusterPatterns();

    for (const [clusterId, clusterPatterns] of Object.entries(clusters)) {
      for (const clusterPattern of clusterPatterns) {
        const similarity = this.calculateSimilarity(pattern, clusterPattern);
        if (similarity >= this.similarityThreshold) {
          return this.getClusterInfo(clusterId);
        }
      }
    }

    return null;
  }

  /**
   * Get recommended patterns from same cluster
   */
  getRecommendedPatternsFromCluster(pattern: string, limit: number = 5): string[] {
    const cluster = this.findClusterForPattern(pattern);
    if (!cluster) {
      return [];
    }

    const data = this.knowledgeBase.getAllData();
    const recommendations: Array<{ pattern: string; confidence: number }> = [];

    for (const patternKey of cluster.patterns) {
      if (patternKey !== pattern) {
        const confidence = this.knowledgeBase.getConfidenceScore(patternKey);
        recommendations.push({ pattern: patternKey, confidence });
      }
    }

    return recommendations
      .sort((a, b) => b.confidence - a.confidence)
      .slice(0, limit)
      .map((r) => r.pattern);
  }

  /**
   * Update clusters (should be called periodically)
   */
  updateClusters(): void {
    if (!this.enabled) {
      return;
    }

    logger.debug('Updating pattern clusters...');
    const clusters = this.clusterPatterns();
    logger.debug(`Created ${Object.keys(clusters).length} pattern clusters`);
  }
}

export default PatternClusterer;
