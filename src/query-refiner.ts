/**
 * Query Refinement Engine for generating query variants when no results are found
 */

import SynonymExpander from './synonym-expander.js';
import logger from './logger.js';

export interface SearchContext {
  entityTypes?: string[];
  sources?: string[];
  timeRange?: string;
  maxVariants?: number;
}

export class QueryRefiner {
  private synonymExpander: SynonymExpander;
  private readonly maxVariants: number;

  constructor(synonymExpander: SynonymExpander) {
    this.synonymExpander = synonymExpander;
    this.maxVariants = parseInt(process.env.MS365_MCP_MAX_QUERY_VARIANTS || '5', 10);
  }

  /**
   * Refine query when no results are found
   */
  async refineQuery(
    originalQuery: string,
    noResults: boolean,
    context?: SearchContext
  ): Promise<string[]> {
    if (!noResults) {
      return [originalQuery];
    }

    logger.info(`Refining query "${originalQuery}" - no results found`);

    const variants = new Set<string>([originalQuery]);

    // 1. Try synonym expansion
    const synonymVariants = this.synonymExpander.expandQuery(
      originalQuery,
      context?.entityTypes?.join(',')
    );
    for (const variant of synonymVariants) {
      variants.add(variant);
    }

    // 2. Try query decomposition (split into parts)
    const decomposed = this.decomposeQuery(originalQuery);
    for (const variant of decomposed) {
      variants.add(variant);
    }

    // 3. Try broader terms (remove modifiers)
    const broaderTerms = this.getBroaderTerms(originalQuery);
    for (const variant of broaderTerms) {
      variants.add(variant);
    }

    // 4. Try different word orders
    const reordered = this.reorderWords(originalQuery);
    for (const variant of reordered) {
      variants.add(variant);
    }

    // 5. Try removing stop words
    const withoutStopWords = this.removeStopWords(originalQuery);
    for (const variant of withoutStopWords) {
      variants.add(variant);
    }

    // 6. Try adding/removing common suffixes
    const withSuffixes = this.addCommonSuffixes(originalQuery);
    for (const variant of withSuffixes) {
      variants.add(variant);
    }

    // Limit and return
    const result = Array.from(variants).slice(0, this.maxVariants);
    logger.debug(`Generated ${result.length} query variants`);
    return result;
  }

  /**
   * Decompose query into parts
   */
  private decomposeQuery(query: string): string[] {
    const words = query.split(/\s+/);
    if (words.length <= 1) {
      return [];
    }

    const variants: string[] = [];

    // Try removing first word
    if (words.length > 1) {
      variants.push(words.slice(1).join(' '));
    }

    // Try removing last word
    if (words.length > 1) {
      variants.push(words.slice(0, -1).join(' '));
    }

    // Try keeping only first 2-3 words
    if (words.length > 3) {
      variants.push(words.slice(0, 2).join(' '));
      variants.push(words.slice(0, 3).join(' '));
    }

    return variants;
  }

  /**
   * Get broader terms (remove specific modifiers)
   */
  private getBroaderTerms(query: string): string[] {
    const variants: string[] = [];
    const queryLower = query.toLowerCase();

    // Remove common modifiers
    const modifiers = [
      /\b(q1|q2|q3|q4|quarter)\s+/gi,
      /\b(2024|2023|2025|2026)\s*/gi,
      /\b(january|february|march|april|may|june|july|august|september|october|november|december)\s+/gi,
      /\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\s+/gi,
      /\b(projekt|project|initiative)\s+/gi,
      /\b(meeting|besprechung|termin)\s+/gi,
    ];

    let broader = queryLower;
    for (const modifier of modifiers) {
      broader = broader.replace(modifier, ' ').trim();
    }

    if (broader !== queryLower && broader.length > 0) {
      variants.push(broader);
    }

    // Try keeping only the main noun (last word)
    const words = query.split(/\s+/);
    if (words.length > 1) {
      variants.push(words[words.length - 1]);
    }

    return variants;
  }

  /**
   * Reorder words in query
   */
  private reorderWords(query: string): string[] {
    const words = query.split(/\s+/);
    if (words.length <= 1) {
      return [];
    }

    const variants: string[] = [];

    // Reverse order
    if (words.length >= 2 && words.length <= 4) {
      variants.push(words.slice().reverse().join(' '));
    }

    // Move last word to front
    if (words.length >= 2) {
      const reordered = [words[words.length - 1], ...words.slice(0, -1)];
      variants.push(reordered.join(' '));
    }

    return variants;
  }

  /**
   * Remove stop words
   */
  private removeStopWords(query: string): string[] {
    const stopWords = [
      'the',
      'a',
      'an',
      'and',
      'or',
      'but',
      'in',
      'on',
      'at',
      'to',
      'for',
      'of',
      'with',
      'by',
      'der',
      'die',
      'das',
      'und',
      'oder',
      'mit',
      'von',
      'für',
      'zu',
    ];

    const words = query.toLowerCase().split(/\s+/);
    const filtered = words.filter((w) => !stopWords.includes(w));

    if (filtered.length < words.length && filtered.length > 0) {
      return [filtered.join(' ')];
    }

    return [];
  }

  /**
   * Add common suffixes/variations
   */
  private addCommonSuffixes(query: string): string[] {
    const variants: string[] = [];
    const queryLower = query.toLowerCase();

    // Try adding year
    const currentYear = new Date().getFullYear();
    if (!queryLower.includes(currentYear.toString())) {
      variants.push(`${query} ${currentYear}`);
      variants.push(`${query} ${currentYear - 1}`);
    }

    // Try adding common project suffixes
    if (!queryLower.includes('projekt') && !queryLower.includes('project')) {
      variants.push(`${query} projekt`);
      variants.push(`${query} project`);
    }

    return variants;
  }

  /**
   * Limit variants to maximum number
   */
  private limitVariants(variants: string[], max: number): string[] {
    return variants.slice(0, max);
  }
}

export default QueryRefiner;
