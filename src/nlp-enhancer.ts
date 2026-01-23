/**
 * NLP Enhancer for improved query processing with stemming, entity recognition, and intent classification
 */

import logger from './logger.js';

export interface IntentClassification {
  intent: 'search' | 'find' | 'list' | 'get' | 'create' | 'update' | 'delete' | 'unknown';
  confidence: number;
  entities: string[];
}

export interface NormalizedQuery {
  original: string;
  normalized: string;
  stemmed: string[];
  entities: string[];
  intent: IntentClassification;
}

/**
 * Simple stemmer for English and German words
 * Based on Porter Stemmer algorithm (simplified)
 */
class SimpleStemmer {
  private readonly suffixes = [
    // English
    'ing',
    'ed',
    'er',
    'est',
    'ly',
    'tion',
    'sion',
    'ness',
    'ment',
    'able',
    'ible',
    // German
    'ung',
    'en',
    'er',
    'est',
    'lich',
    'keit',
    'heit',
    'schaft',
    'bar',
  ];

  private readonly stopWords = new Set([
    // English
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
    'is',
    'are',
    'was',
    'were',
    'be',
    'been',
    'being',
    'have',
    'has',
    'had',
    'do',
    'does',
    'did',
    'will',
    'would',
    'should',
    'could',
    'may',
    'might',
    'must',
    'can',
    // German
    'der',
    'die',
    'das',
    'und',
    'oder',
    'aber',
    'in',
    'auf',
    'an',
    'zu',
    'für',
    'von',
    'mit',
    'durch',
    'ist',
    'sind',
    'war',
    'waren',
    'sein',
    'gewesen',
    'haben',
    'hat',
    'hatte',
    'tun',
    'tut',
    'tat',
    'wird',
    'würde',
    'sollte',
    'könnte',
    'kann',
    'muss',
  ]);

  /**
   * Stem a word (simplified Porter-like algorithm)
   */
  stem(word: string): string {
    const lower = word.toLowerCase();

    // Remove common suffixes
    for (const suffix of this.suffixes) {
      if (lower.endsWith(suffix) && lower.length > suffix.length + 2) {
        return lower.slice(0, -suffix.length);
      }
    }

    // Remove plural 's' or 'es'
    if (lower.endsWith('es') && lower.length > 4) {
      return lower.slice(0, -2);
    }
    if (lower.endsWith('s') && lower.length > 3) {
      return lower.slice(0, -1);
    }

    return lower;
  }

  /**
   * Check if word is a stop word
   */
  isStopWord(word: string): boolean {
    return this.stopWords.has(word.toLowerCase());
  }
}

/**
 * Entity Recognizer for common business entities
 */
class EntityRecognizer {
  private readonly entityPatterns: Record<string, RegExp[]> = {
    email: [/\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g],
    date: [
      /\b\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\b/g,
      /\b(january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{1,2},?\s+\d{4}\b/gi,
      /\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\s+\d{1,2},?\s+\d{4}\b/gi,
      /\b(januar|februar|märz|april|mai|juni|juli|august|september|oktober|november|dezember)\s+\d{1,2},?\s+\d{4}\b/gi,
    ],
    person: [
      /\b(mr|mrs|ms|dr|prof|herr|frau|doktor|professor)\.?\s+[A-Z][a-z]+/gi,
      /\b[A-Z][a-z]+\s+[A-Z][a-z]+/g, // First Last name pattern
    ],
    project: [
      /\b(projekt|project|initiative|programm|program)\s+[A-Z][a-z]+/gi,
      /\b[A-Z]{2,}\s*(projekt|project)/gi, // Acronym + project
    ],
    company: [
      /\b(inc|llc|ltd|gmbh|ag|corp|corporation|company|unternehmen)\b/gi,
      /\b[A-Z][a-z]+\s+(inc|llc|ltd|gmbh|ag|corp)/gi,
    ],
    number: [/\b\d+\b/g],
    url: [/\bhttps?:\/\/[^\s]+/gi, /\bwww\.[^\s]+/gi],
  };

  /**
   * Extract entities from text
   */
  extractEntities(text: string): string[] {
    const entities: string[] = [];
    const lower = text.toLowerCase();

    for (const [entityType, patterns] of Object.entries(this.entityPatterns)) {
      for (const pattern of patterns) {
        const matches = text.match(pattern);
        if (matches) {
          entities.push(...matches.map((m) => `${entityType}:${m}`));
        }
      }
    }

    // Extract potential keywords (capitalized words, acronyms)
    const words = text.split(/\s+/);
    for (const word of words) {
      if (word.length > 2) {
        // Acronym pattern (all caps, 2+ chars)
        if (/^[A-Z]{2,}$/.test(word)) {
          entities.push(`acronym:${word}`);
        }
        // Capitalized word (likely proper noun)
        else if (/^[A-Z][a-z]+$/.test(word)) {
          entities.push(`keyword:${word}`);
        }
      }
    }

    return [...new Set(entities)]; // Remove duplicates
  }
}

/**
 * Intent Classifier for query intent detection
 */
class IntentClassifier {
  private readonly intentPatterns: Record<string, RegExp[]> = {
    search: [
      /\b(search|find|look|suchen|finden|suche)\b/gi,
      /\b(what|where|who|when|wie|wo|wer|wann)\b/gi,
    ],
    find: [/\b(find|locate|get|fetch|holen|abrufen)\b/gi],
    list: [/\b(list|show|display|all|alle|zeigen|anzeigen)\b/gi],
    get: [/\b(get|fetch|retrieve|abrufen|holen)\b/gi],
    create: [/\b(create|new|add|make|erstellen|neu|hinzufügen)\b/gi],
    update: [/\b(update|edit|modify|change|ändern|bearbeiten)\b/gi],
    delete: [/\b(delete|remove|drop|remove|löschen|entfernen)\b/gi],
  };

  /**
   * Classify query intent
   */
  classifyIntent(query: string): IntentClassification {
    const lower = query.toLowerCase();
    let bestIntent:
      | 'search'
      | 'find'
      | 'list'
      | 'get'
      | 'create'
      | 'update'
      | 'delete'
      | 'unknown' = 'unknown';
    let maxMatches = 0;
    const entities: string[] = [];

    for (const [intent, patterns] of Object.entries(this.intentPatterns)) {
      let matches = 0;
      for (const pattern of patterns) {
        if (pattern.test(lower)) {
          matches++;
        }
      }
      if (matches > maxMatches) {
        maxMatches = matches;
        bestIntent = intent as IntentClassification['intent'];
      }
    }

    // Extract entities from query
    const entityRecognizer = new EntityRecognizer();
    entities.push(...entityRecognizer.extractEntities(query));

    // Default to 'search' if no clear intent
    if (bestIntent === 'unknown' && query.length > 0) {
      bestIntent = 'search';
    }

    const confidence = maxMatches > 0 ? Math.min(1.0, maxMatches * 0.3) : 0.5;

    return {
      intent: bestIntent,
      confidence,
      entities,
    };
  }
}

/**
 * NLP Enhancer for query processing
 */
export class NLPEnhancer {
  private stemmer: SimpleStemmer;
  private entityRecognizer: EntityRecognizer;
  private intentClassifier: IntentClassifier;
  private readonly enabled: boolean;

  constructor() {
    this.stemmer = new SimpleStemmer();
    this.entityRecognizer = new EntityRecognizer();
    this.intentClassifier = new IntentClassifier();
    this.enabled =
      process.env.MS365_MCP_LEARNING_NLP_ENABLED === 'true' ||
      process.env.MS365_MCP_LEARNING_NLP_ENABLED !== 'false';
  }

  /**
   * Normalize a query
   */
  normalizeQuery(query: string): NormalizedQuery {
    if (!this.enabled) {
      return {
        original: query,
        normalized: query.toLowerCase().trim(),
        stemmed: query.toLowerCase().split(/\s+/),
        entities: [],
        intent: {
          intent: 'search',
          confidence: 0.5,
          entities: [],
        },
      };
    }

    // Normalize: lowercase, trim, remove extra spaces
    const normalized = query.toLowerCase().trim().replace(/\s+/g, ' ');

    // Extract entities
    const entities = this.entityRecognizer.extractEntities(query);

    // Stem words
    const words = normalized.split(/\s+/);
    const stemmed = words
      .filter((w) => !this.stemmer.isStopWord(w))
      .map((w) => this.stemmer.stem(w));

    // Classify intent
    const intent = this.intentClassifier.classifyIntent(query);

    return {
      original: query,
      normalized,
      stemmed,
      entities,
      intent,
    };
  }

  /**
   * Extract entities from text
   */
  extractEntities(text: string): string[] {
    if (!this.enabled) {
      return [];
    }
    return this.entityRecognizer.extractEntities(text);
  }

  /**
   * Classify intent
   */
  classifyIntent(query: string): IntentClassification {
    if (!this.enabled) {
      return {
        intent: 'search',
        confidence: 0.5,
        entities: [],
      };
    }
    return this.intentClassifier.classifyIntent(query);
  }

  /**
   * Stem a word
   */
  stem(word: string): string {
    if (!this.enabled) {
      return word.toLowerCase();
    }
    return this.stemmer.stem(word);
  }

  /**
   * Check if word is a stop word
   */
  isStopWord(word: string): boolean {
    if (!this.enabled) {
      return false;
    }
    return this.stemmer.isStopWord(word);
  }

  /**
   * Calculate similarity between two queries using stemming
   */
  calculateSimilarity(query1: string, query2: string): number {
    if (!this.enabled) {
      // Simple word overlap
      const words1 = new Set(query1.toLowerCase().split(/\s+/));
      const words2 = new Set(query2.toLowerCase().split(/\s+/));
      const intersection = new Set([...words1].filter((x) => words2.has(x)));
      const union = new Set([...words1, ...words2]);
      return union.size > 0 ? intersection.size / union.size : 0;
    }

    const norm1 = this.normalizeQuery(query1);
    const norm2 = this.normalizeQuery(query2);

    const stemmed1 = new Set(norm1.stemmed);
    const stemmed2 = new Set(norm2.stemmed);

    const intersection = new Set([...stemmed1].filter((x) => stemmed2.has(x)));
    const union = new Set([...stemmed1, ...stemmed2]);

    return union.size > 0 ? intersection.size / union.size : 0;
  }
}

export default NLPEnhancer;
