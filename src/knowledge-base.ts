/**
 * Knowledge Base for persisting learned synonyms, query patterns, and entity mappings
 */

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import logger from './logger.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

export interface SuccessfulQuery {
  query: string;
  results: number;
  sources: string[];
  timestamp: Date;
  context?: string;
}

export interface QueryPattern {
  pattern: string;
  entityTypes: string[];
  successCount: number;
  lastUsed: Date;
  context?: string;
}

export interface EntityMapping {
  keyword: string;
  entityTypes: string[];
  successCount: number;
  lastUsed: Date;
}

export interface DataLocation {
  dataType: string;
  sources: string[];
  successCount: number;
  lastUsed: Date;
}

export interface KnowledgeBaseData {
  successfulQueries: Record<string, SuccessfulQuery>;
  learnedSynonyms: Record<string, string[]>;
  queryPatterns: Record<string, QueryPattern>;
  entityMappings: Record<string, EntityMapping>;
  dataLocations: Record<string, DataLocation>;
  version: number;
  lastUpdated: string;
}

export class KnowledgeBase {
  private data: KnowledgeBaseData;
  private readonly filePath: string;
  private readonly maxEntries = 10000; // Limit knowledge base size

  constructor(filePath?: string) {
    this.filePath =
      filePath ||
      process.env.MS365_MCP_KNOWLEDGE_BASE_PATH ||
      path.join(__dirname, '..', 'knowledge-base.json');
    this.data = this.load();
  }

  /**
   * Load knowledge base from file
   */
  private load(): KnowledgeBaseData {
    try {
      if (fs.existsSync(this.filePath)) {
        const content = fs.readFileSync(this.filePath, 'utf8');
        const loaded = JSON.parse(content) as KnowledgeBaseData;
        // Convert date strings back to Date objects
        Object.values(loaded.successfulQueries || {}).forEach((q) => {
          q.timestamp = new Date(q.timestamp);
        });
        Object.values(loaded.queryPatterns || {}).forEach((p) => {
          p.lastUsed = new Date(p.lastUsed);
        });
        Object.values(loaded.entityMappings || {}).forEach((m) => {
          m.lastUsed = new Date(m.lastUsed);
        });
        Object.values(loaded.dataLocations || {}).forEach((l) => {
          l.lastUsed = new Date(l.lastUsed);
        });
        return loaded;
      }
    } catch (error) {
      logger.warn(`Failed to load knowledge base from ${this.filePath}: ${error}`);
    }

    // Return empty knowledge base
    return {
      successfulQueries: {},
      learnedSynonyms: {},
      queryPatterns: {},
      entityMappings: {},
      dataLocations: {},
      version: 1,
      lastUpdated: new Date().toISOString(),
    };
  }

  /**
   * Save knowledge base to file
   */
  async save(): Promise<void> {
    try {
      // Ensure directory exists
      const dir = path.dirname(this.filePath);
      if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
      }

      // Limit entries to prevent file from growing too large
      this.trimEntries();

      this.data.lastUpdated = new Date().toISOString();
      fs.writeFileSync(this.filePath, JSON.stringify(this.data, null, 2), 'utf8');
      logger.debug(`Knowledge base saved to ${this.filePath}`);
    } catch (error) {
      logger.error(`Failed to save knowledge base to ${this.filePath}: ${error}`);
    }
  }

  /**
   * Trim entries to prevent knowledge base from growing too large
   */
  private trimEntries(): void {
    // Keep only most recent and most successful entries
    const trimMap = <T extends { lastUsed?: Date; successCount?: number }>(
      map: Record<string, T>,
      maxSize: number
    ): Record<string, T> => {
      const entries = Object.entries(map);
      if (entries.length <= maxSize) {
        return map;
      }

      // Sort by lastUsed (most recent first) and successCount (highest first)
      entries.sort((a, b) => {
        const aDate = a[1].lastUsed?.getTime() || 0;
        const bDate = b[1].lastUsed?.getTime() || 0;
        if (bDate !== aDate) {
          return bDate - aDate;
        }
        return (b[1].successCount || 0) - (a[1].successCount || 0);
      });

      return Object.fromEntries(entries.slice(0, maxSize));
    };

    const maxPerType = Math.floor(this.maxEntries / 5);
    this.data.successfulQueries = trimMap(this.data.successfulQueries, maxPerType);
    this.data.queryPatterns = trimMap(this.data.queryPatterns, maxPerType);
    this.data.entityMappings = trimMap(this.data.entityMappings, maxPerType);
    this.data.dataLocations = trimMap(this.data.dataLocations, maxPerType);
  }

  /**
   * Record a successful query
   */
  recordSuccessfulQuery(
    query: string,
    results: number,
    sources: string[],
    context?: string
  ): void {
    const key = this.normalizeKey(query);
    this.data.successfulQueries[key] = {
      query,
      results,
      sources,
      timestamp: new Date(),
      context,
    };
  }

  /**
   * Get successful queries matching a pattern
   */
  getSuccessfulQueries(pattern: string, limit = 10): SuccessfulQuery[] {
    const patternLower = pattern.toLowerCase();
    return Object.values(this.data.successfulQueries)
      .filter((q) => q.query.toLowerCase().includes(patternLower))
      .sort((a, b) => b.timestamp.getTime() - a.timestamp.getTime())
      .slice(0, limit);
  }

  /**
   * Record a learned synonym
   */
  recordSynonym(original: string, synonym: string): void {
    const key = this.normalizeKey(original);
    if (!this.data.learnedSynonyms[key]) {
      this.data.learnedSynonyms[key] = [];
    }
    if (!this.data.learnedSynonyms[key].includes(synonym)) {
      this.data.learnedSynonyms[key].push(synonym);
    }
  }

  /**
   * Get learned synonyms for a word
   */
  getLearnedSynonyms(word: string): string[] {
    const key = this.normalizeKey(word);
    return this.data.learnedSynonyms[key] || [];
  }

  /**
   * Record a query pattern
   */
  recordQueryPattern(
    pattern: string,
    entityTypes: string[],
    success: boolean,
    context?: string
  ): void {
    const key = this.normalizeKey(pattern);
    if (!this.data.queryPatterns[key]) {
      this.data.queryPatterns[key] = {
        pattern,
        entityTypes: [],
        successCount: 0,
        lastUsed: new Date(),
        context,
      };
    }

    const existing = this.data.queryPatterns[key];
    if (success) {
      existing.successCount++;
    }
    existing.lastUsed = new Date();

    // Update entity types (merge with existing)
    for (const et of entityTypes) {
      if (!existing.entityTypes.includes(et)) {
        existing.entityTypes.push(et);
      }
    }
  }

  /**
   * Get query pattern for a query type
   */
  getQueryPattern(pattern: string): QueryPattern | undefined {
    const key = this.normalizeKey(pattern);
    return this.data.queryPatterns[key];
  }

  /**
   * Record entity mapping
   */
  recordEntityMapping(keyword: string, entityTypes: string[], success: boolean): void {
    const key = this.normalizeKey(keyword);
    if (!this.data.entityMappings[key]) {
      this.data.entityMappings[key] = {
        keyword,
        entityTypes: [],
        successCount: 0,
        lastUsed: new Date(),
      };
    }

    const existing = this.data.entityMappings[key];
    if (success) {
      existing.successCount++;
    }
    existing.lastUsed = new Date();

    // Update entity types (merge with existing)
    for (const et of entityTypes) {
      if (!existing.entityTypes.includes(et)) {
        existing.entityTypes.push(et);
      }
    }
  }

  /**
   * Get entity mapping for a keyword
   */
  getEntityMapping(keyword: string): string[] {
    const key = this.normalizeKey(keyword);
    const mapping = this.data.entityMappings[key];
    return mapping ? mapping.entityTypes : [];
  }

  /**
   * Record data location (where certain data types are typically found)
   */
  recordDataLocation(dataType: string, sources: string[], success: boolean): void {
    const key = this.normalizeKey(dataType);
    if (!this.data.dataLocations[key]) {
      this.data.dataLocations[key] = {
        dataType,
        sources: [],
        successCount: 0,
        lastUsed: new Date(),
      };
    }

    const existing = this.data.dataLocations[key];
    if (success) {
      existing.successCount++;
    }
    existing.lastUsed = new Date();

    // Update sources (merge with existing)
    for (const source of sources) {
      if (!existing.sources.includes(source)) {
        existing.sources.push(source);
      }
    }
  }

  /**
   * Get data location for a data type
   */
  getDataLocation(dataType: string): string[] {
    const key = this.normalizeKey(dataType);
    const location = this.data.dataLocations[key];
    return location ? location.sources : [];
  }

  /**
   * Suggest query variants based on learned patterns
   */
  suggestVariants(query: string, context?: string): string[] {
    const variants = new Set<string>([query]);

    // Check successful queries
    const successful = this.getSuccessfulQueries(query, 5);
    for (const sq of successful) {
      if (!context || !sq.context || sq.context === context) {
        variants.add(sq.query);
      }
    }

    // Check learned synonyms
    const words = query.toLowerCase().split(/\s+/);
    for (const word of words) {
      const synonyms = this.getLearnedSynonyms(word);
      for (const synonym of synonyms) {
        const variant = query.toLowerCase().replace(word, synonym);
        variants.add(variant);
      }
    }

    return Array.from(variants).slice(0, 10);
  }

  /**
   * Normalize key for storage
   */
  private normalizeKey(key: string): string {
    return key.toLowerCase().trim();
  }

  /**
   * Get all data (for testing/debugging)
   */
  getAllData(): KnowledgeBaseData {
    return this.data;
  }

  /**
   * Clear all data
   */
  clear(): void {
    this.data = {
      successfulQueries: {},
      learnedSynonyms: {},
      queryPatterns: {},
      entityMappings: {},
      dataLocations: {},
      version: 1,
      lastUpdated: new Date().toISOString(),
    };
  }
}

export default KnowledgeBase;

