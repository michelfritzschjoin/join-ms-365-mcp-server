/**
 * UQAS Pro - Core Module
 *
 * Universal Question Answering System for Microsoft 365:
 * - Multi-layer adaptive analysis (L1-L5)
 * - Entity relationship graph
 * - Token budget management
 * - Smart caching
 */

export { AdaptiveLayerController, type LayerResult, type AnalysisDepth } from './adaptive-layer.js';
export {
  EntityGraphBuilder,
  type EntityGraph,
  type EntityNode,
  type EntityEdge,
} from './entity-graph.js';
export { TokenController, type TokenBudget, type CompactResult } from './token-controller.js';
export { CacheManager, type CacheEntry, type CacheConfig } from './cache-manager.js';
