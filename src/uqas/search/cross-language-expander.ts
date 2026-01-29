/**
 * Cross-Language Search Expander
 *
 * Expands search queries to include both German and English variants
 * for maximum coverage when searching Microsoft 365 content.
 */

import type { SupportedLanguage } from '../i18n/index.js';
import { BilingualThesaurus } from '../i18n/bilingual-thesaurus.js';

/**
 * Expanded search query set
 */
export interface SearchQuerySet {
  /** Original query */
  primary: string;
  /** Variants in source language */
  sourceLangVariants: string[];
  /** Variants in other language */
  crossLangVariants: string[];
  /** Combined OR query for MS Search API */
  combined: string;
  /** Keywords extracted */
  keywords: string[];
  /** Detected language */
  language: SupportedLanguage;
}

/**
 * Single expanded query
 */
export interface ExpandedQuery {
  query: string;
  language: SupportedLanguage;
  isOriginal: boolean;
  confidence: number;
}

/**
 * Stop words to exclude from expansion
 */
const STOP_WORDS = {
  de: new Set([
    'der',
    'die',
    'das',
    'den',
    'dem',
    'des',
    'ein',
    'eine',
    'einer',
    'einem',
    'und',
    'oder',
    'aber',
    'in',
    'im',
    'auf',
    'an',
    'am',
    'zu',
    'zum',
    'zur',
    'für',
    'von',
    'vom',
    'mit',
    'bei',
    'nach',
    'vor',
    'über',
    'unter',
    'aus',
    'ist',
    'sind',
    'war',
    'hat',
    'haben',
    'wird',
    'werden',
    'kann',
    'können',
    'ich',
    'du',
    'er',
    'sie',
    'es',
    'wir',
    'ihr',
    'mich',
    'dir',
    'mir',
    'was',
    'wer',
    'wie',
    'wo',
    'wann',
    'warum',
    'welche',
    'welcher',
  ]),
  en: new Set([
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
    'from',
    'up',
    'about',
    'into',
    'over',
    'after',
    'is',
    'are',
    'was',
    'were',
    'has',
    'have',
    'had',
    'will',
    'would',
    'can',
    'i',
    'you',
    'he',
    'she',
    'it',
    'we',
    'they',
    'me',
    'him',
    'her',
    'us',
    'what',
    'who',
    'how',
    'where',
    'when',
    'why',
    'which',
  ]),
};

/**
 * CrossLanguageSearchExpander - Expands queries for bilingual search
 */
export class CrossLanguageSearchExpander {
  private thesaurus: BilingualThesaurus;

  constructor(thesaurus?: BilingualThesaurus) {
    this.thesaurus = thesaurus || new BilingualThesaurus();
  }

  /**
   * Expand query to include both language variants
   */
  expand(query: string, sourceLang: SupportedLanguage): SearchQuerySet {
    const keywords = this.extractKeywords(query, sourceLang);

    // Generate source language variants
    const sourceLangVariants = this.generateVariants(query, keywords, sourceLang);

    // Generate cross-language variants
    const crossLangVariants = this.generateCrossLanguageVariants(query, keywords, sourceLang);

    // Build combined OR query (limited to prevent explosion)
    const combined = this.buildCombinedQuery([
      query,
      ...sourceLangVariants.slice(0, 2),
      ...crossLangVariants.slice(0, 2),
    ]);

    return {
      primary: query,
      sourceLangVariants,
      crossLangVariants,
      combined,
      keywords,
      language: sourceLang,
    };
  }

  /**
   * Get all expanded queries with metadata
   */
  expandWithMetadata(query: string, sourceLang: SupportedLanguage): ExpandedQuery[] {
    const result: ExpandedQuery[] = [];
    const querySet = this.expand(query, sourceLang);

    // Original query
    result.push({
      query: querySet.primary,
      language: sourceLang,
      isOriginal: true,
      confidence: 1.0,
    });

    // Source language variants
    for (const variant of querySet.sourceLangVariants) {
      result.push({
        query: variant,
        language: sourceLang,
        isOriginal: false,
        confidence: 0.8,
      });
    }

    // Cross-language variants
    const targetLang = sourceLang === 'de' ? 'en' : 'de';
    for (const variant of querySet.crossLangVariants) {
      result.push({
        query: variant,
        language: targetLang,
        isOriginal: false,
        confidence: 0.7,
      });
    }

    return result;
  }

  /**
   * Extract searchable keywords from query
   */
  extractKeywords(query: string, lang: SupportedLanguage): string[] {
    const words = query
      .toLowerCase()
      .replace(/[^\w\säöüÄÖÜß-]/g, ' ')
      .split(/\s+/)
      .filter((w) => w.length > 2);

    // Filter out stop words
    const stopWords = STOP_WORDS[lang];
    return words.filter((w) => !stopWords.has(w));
  }

  /**
   * Generate variants in the same language
   */
  private generateVariants(query: string, keywords: string[], lang: SupportedLanguage): string[] {
    const variants = new Set<string>();

    for (const keyword of keywords) {
      // Get synonyms from thesaurus
      const synonyms = this.thesaurus.getVariantsInLanguage(keyword, lang);

      for (const synonym of synonyms) {
        // Create variant by replacing keyword with synonym
        const variant = query.toLowerCase().replace(keyword, synonym);
        if (variant !== query.toLowerCase()) {
          variants.add(variant);
        }
      }
    }

    return Array.from(variants).slice(0, 5); // Limit variants
  }

  /**
   * Generate cross-language variants
   */
  private generateCrossLanguageVariants(
    query: string,
    keywords: string[],
    sourceLang: SupportedLanguage
  ): string[] {
    const variants = new Set<string>();
    const targetLang = sourceLang === 'de' ? 'en' : 'de';

    for (const keyword of keywords) {
      // Get translations from thesaurus
      const translations = this.thesaurus.getCrossLanguageVariants(keyword, sourceLang);

      for (const translation of translations) {
        // Create variant by replacing keyword with translation
        const variant = query.toLowerCase().replace(keyword, translation);
        if (variant !== query.toLowerCase()) {
          variants.add(variant);
        }
      }
    }

    // Also create a fully translated variant for common patterns
    const fullTranslation = this.translateKeyTerms(keywords, targetLang);
    if (fullTranslation.length > 0) {
      variants.add(fullTranslation.join(' '));
    }

    return Array.from(variants).slice(0, 5); // Limit variants
  }

  /**
   * Translate key terms to target language
   */
  private translateKeyTerms(keywords: string[], targetLang: SupportedLanguage): string[] {
    const translated: string[] = [];

    for (const keyword of keywords) {
      const translations = this.thesaurus.getVariantsInLanguage(keyword, targetLang);
      if (translations.length > 0) {
        translated.push(translations[0]); // Use first translation
      } else {
        translated.push(keyword); // Keep original if no translation
      }
    }

    return translated;
  }

  /**
   * Build combined OR query for MS Search API
   */
  private buildCombinedQuery(queries: string[]): string {
    // Deduplicate and clean
    const unique = [...new Set(queries.map((q) => q.trim().toLowerCase()))];

    // Limit to prevent query explosion
    const limited = unique.slice(0, 5);

    // Build OR query
    if (limited.length === 1) {
      return limited[0];
    }

    // Use parentheses for complex OR queries
    return limited.map((q) => (q.includes(' ') ? `"${q}"` : q)).join(' OR ');
  }

  /**
   * Build KQL (Keyword Query Language) query for SharePoint search
   */
  buildKQLQuery(querySet: SearchQuerySet): string {
    const terms: string[] = [];

    // Add primary query
    terms.push(querySet.primary);

    // Add top variants
    for (const variant of querySet.sourceLangVariants.slice(0, 2)) {
      terms.push(variant);
    }

    // Combine with OR
    return terms.map((t) => (t.includes(' ') ? `"${t}"` : t)).join(' OR ');
  }

  /**
   * Build OData $search parameter
   */
  buildODataSearch(querySet: SearchQuerySet): string {
    // OData search is simpler - just use combined
    return querySet.combined;
  }

  /**
   * Check if query contains a specific keyword
   */
  containsKeyword(query: string, keyword: string): boolean {
    const lower = query.toLowerCase();
    const keywordLower = keyword.toLowerCase();

    // Direct match
    if (lower.includes(keywordLower)) {
      return true;
    }

    // Check synonyms
    const allVariants = this.thesaurus.getAllVariants(keyword);
    return allVariants.some((v) => lower.includes(v.toLowerCase()));
  }

  /**
   * Get the thesaurus for direct access
   */
  getThesaurus(): BilingualThesaurus {
    return this.thesaurus;
  }
}

export default CrossLanguageSearchExpander;
