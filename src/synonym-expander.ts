/**
 * Synonym Expander for query expansion with business terminology and learned synonyms
 */

import logger from './logger.js';
import NLPEnhancer from './nlp-enhancer.js';

interface LearnedSynonym {
  original: string;
  synonym: string;
  successCount: number;
  failureCount: number;
  lastUsed: Date;
  context?: string;
}

/**
 * Business terminology synonyms
 */
const BUSINESS_SYNONYMS: Record<string, string[]> = {
  projekt: ['project', 'initiative', 'programm', 'vorhaben'],
  meeting: ['besprechung', 'termin', 'call', 'conference'],
  kunde: ['customer', 'client', 'mandant'],
  vertrag: ['contract', 'agreement', 'vereinbarung'],
  team: ['gruppe', 'abteilung', 'department'],
  dokument: ['document', 'file', 'datei', 'unterlage'],
  entscheidung: ['decision', 'beschluss', 'entscheid'],
  budget: ['haushalt', 'etat', 'finanzen'],
  präsentation: ['presentation', 'vortrag', 'deck'],
  notiz: ['note', 'notiz', 'memo', 'annotation'],
  aufgabe: ['task', 'aufgabe', 'todo', 'assignment'],
  kalender: ['calendar', 'kalender', 'terminplan'],
  email: ['mail', 'email', 'nachricht', 'message'],
  datei: ['file', 'datei', 'dokument', 'document'],
  ordner: ['folder', 'ordner', 'verzeichnis'],
  seite: ['page', 'seite', 'site'],
  kanal: ['channel', 'kanal'],
  chat: ['chat', 'unterhaltung', 'conversation'],
};

/**
 * Common abbreviations and their expansions
 */
const ABBREVIATIONS: Record<string, string[]> = {
  q1: ['first quarter', 'q1', 'quarter 1'],
  q2: ['second quarter', 'q2', 'quarter 2'],
  q3: ['third quarter', 'q3', 'quarter 3'],
  q4: ['fourth quarter', 'q4', 'quarter 4'],
  hr: ['human resources', 'hr', 'personal'],
  it: ['information technology', 'it'],
  cfo: ['chief financial officer', 'cfo'],
  ceo: ['chief executive officer', 'ceo'],
  cto: ['chief technology officer', 'cto'],
};

export class SynonymExpander {
  private learnedSynonyms: Map<string, LearnedSynonym[]>;
  private readonly minSuccessRatio = 0.6; // Synonym must be successful 60% of the time
  private nlpEnhancer: NLPEnhancer | null = null;

  constructor() {
    this.learnedSynonyms = new Map();
    // Initialize NLP enhancer if enabled
    try {
      this.nlpEnhancer = new NLPEnhancer();
    } catch (error) {
      logger.warn(`Failed to initialize NLP enhancer: ${error}`);
    }
  }

  /**
   * Expand query with synonyms
   */
  expandQuery(query: string, context?: string): string[] {
    const variants = new Set<string>([query.toLowerCase()]);
    const queryLower = query.toLowerCase();
    const words = queryLower.split(/\s+/);

    // Use NLP enhancer for better normalization if available
    if (this.nlpEnhancer) {
      const normalized = this.nlpEnhancer.normalizeQuery(query);
      // Add stemmed variants
      if (normalized.stemmed.length > 0) {
        variants.add(normalized.stemmed.join(' '));
      }
    }

    // 1. Check learned synonyms
    for (const word of words) {
      const learned = this.learnedSynonyms.get(word);
      if (learned) {
        const relevantSynonyms = learned
          .filter((s) => !context || !s.context || s.context === context)
          .filter((s) => {
            const total = s.successCount + s.failureCount;
            return total > 0 && s.successCount / total >= this.minSuccessRatio;
          })
          .map((s) => s.synonym);

        for (const synonym of relevantSynonyms) {
          const variant = queryLower.replace(word, synonym);
          variants.add(variant);
        }
      }
    }

    // 2. Check business terminology
    for (const [key, synonyms] of Object.entries(BUSINESS_SYNONYMS)) {
      if (queryLower.includes(key)) {
        for (const synonym of synonyms) {
          const variant = queryLower.replace(key, synonym);
          variants.add(variant);
          // Also try with original word replaced
          variants.add(queryLower.replace(new RegExp(key, 'gi'), synonym));
        }
      }
    }

    // 3. Expand abbreviations
    for (const [abbr, expansions] of Object.entries(ABBREVIATIONS)) {
      if (queryLower.includes(abbr)) {
        for (const expansion of expansions) {
          const variant = queryLower.replace(abbr, expansion);
          variants.add(variant);
        }
      }
    }

    // 4. Try word order variations (for 2-3 word queries)
    if (words.length >= 2 && words.length <= 3) {
      const reversed = words.slice().reverse().join(' ');
      variants.add(reversed);
    }

    // 5. Remove duplicates and return as array
    return Array.from(variants).slice(0, 10); // Limit to 10 variants
  }

  /**
   * Learn a synonym from search results
   */
  learnSynonym(original: string, synonym: string, success: boolean, context?: string): void {
    const key = original.toLowerCase();
    if (!this.learnedSynonyms.has(key)) {
      this.learnedSynonyms.set(key, []);
    }

    const synonyms = this.learnedSynonyms.get(key)!;
    let learned = synonyms.find((s) => s.synonym === synonym && s.context === context);

    if (!learned) {
      learned = {
        original,
        synonym,
        successCount: 0,
        failureCount: 0,
        lastUsed: new Date(),
        context,
      };
      synonyms.push(learned);
    }

    if (success) {
      learned.successCount++;
    } else {
      learned.failureCount++;
    }
    learned.lastUsed = new Date();

    logger.debug(`Learned synonym: "${original}" ≈ "${synonym}" (success: ${success})`);
  }

  /**
   * Get learned synonyms for a word
   */
  getLearnedSynonyms(word: string, context?: string): string[] {
    const key = word.toLowerCase();
    const learned = this.learnedSynonyms.get(key);
    if (!learned) {
      return [];
    }

    return learned
      .filter((s) => !context || !s.context || s.context === context)
      .filter((s) => {
        const total = s.successCount + s.failureCount;
        return total > 0 && s.successCount / total >= this.minSuccessRatio;
      })
      .map((s) => s.synonym);
  }

  /**
   * Load learned synonyms (for persistence)
   */
  loadLearnedSynonyms(data: LearnedSynonym[]): void {
    for (const item of data) {
      const key = item.original.toLowerCase();
      if (!this.learnedSynonyms.has(key)) {
        this.learnedSynonyms.set(key, []);
      }
      this.learnedSynonyms.get(key)!.push(item);
    }
  }

  /**
   * Get all learned synonyms for persistence
   */
  getLearnedSynonymsData(): LearnedSynonym[] {
    const result: LearnedSynonym[] = [];
    for (const synonyms of this.learnedSynonyms.values()) {
      result.push(...synonyms);
    }
    return result;
  }
}

export default SynonymExpander;
