/**
 * Language Detector for DE/EN
 *
 * Detects whether user input is German or English with confidence scoring.
 * Handles mixed-language input (code-switching) common in business contexts.
 */

import type { SupportedLanguage } from './index.js';

/**
 * Result of language detection
 */
export interface LanguageDetectionResult {
  /** Detected primary language */
  lang: SupportedLanguage;
  /** Confidence score (0-1) */
  confidence: number;
  /** Whether input contains both languages */
  hasCodeSwitch: boolean;
  /** Breakdown of language indicators found */
  indicators: {
    de: number;
    en: number;
  };
}

/**
 * Language indicator patterns
 */
const LANG_INDICATORS = {
  de: {
    articles: ['der', 'die', 'das', 'den', 'dem', 'des', 'ein', 'eine', 'einer', 'einem', 'einen'],
    prepositions: [
      'mit',
      'bei',
      'für',
      'nach',
      'von',
      'zu',
      'aus',
      'über',
      'unter',
      'zwischen',
      'durch',
      'gegen',
      'ohne',
      'um',
    ],
    verbs: [
      'ist',
      'sind',
      'war',
      'waren',
      'hat',
      'haben',
      'hatte',
      'hatten',
      'wird',
      'werden',
      'wurde',
      'wurden',
      'kann',
      'können',
      'muss',
      'müssen',
      'soll',
      'sollen',
      'möchte',
      'möchten',
    ],
    questions: [
      'wann',
      'was',
      'wer',
      'wo',
      'wie',
      'warum',
      'weshalb',
      'wieso',
      'welche',
      'welcher',
      'welches',
      'welchen',
      'welchem',
      'wieviel',
      'wieviele',
    ],
    connectors: [
      'und',
      'oder',
      'aber',
      'denn',
      'weil',
      'dass',
      'damit',
      'obwohl',
      'wenn',
      'falls',
      'als',
      'bevor',
      'nachdem',
      'während',
      'sobald',
    ],
    pronouns: [
      'ich',
      'du',
      'er',
      'sie',
      'es',
      'wir',
      'ihr',
      'mich',
      'dich',
      'sich',
      'uns',
      'euch',
      'mir',
      'dir',
      'ihm',
      'ihr',
      'ihnen',
      'mein',
      'dein',
      'sein',
      'unser',
      'euer',
    ],
  },
  en: {
    articles: ['the', 'a', 'an'],
    prepositions: [
      'with',
      'at',
      'for',
      'after',
      'from',
      'to',
      'about',
      'into',
      'over',
      'between',
      'through',
      'against',
      'without',
      'around',
    ],
    verbs: [
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
      'could',
      'must',
      'should',
      'shall',
      'may',
      'might',
    ],
    questions: ['when', 'what', 'who', 'where', 'how', 'why', 'which', 'whose', 'whom'],
    connectors: [
      'and',
      'or',
      'but',
      'because',
      'that',
      'so',
      'although',
      'if',
      'when',
      'while',
      'before',
      'after',
      'since',
      'unless',
    ],
    pronouns: [
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
      'them',
      'my',
      'your',
      'his',
      'our',
      'their',
    ],
  },
} as const;

/**
 * Regex for German umlauts and ß
 */
const GERMAN_CHARS_REGEX = /[äöüÄÖÜß]/;

/**
 * Common English-only words (rarely used in German)
 */
const ENGLISH_ONLY = new Set([
  'the',
  'of',
  'and',
  'to',
  'in',
  'is',
  'it',
  'you',
  'that',
  'he',
  'she',
  'for',
  'on',
  'are',
  'as',
  'with',
  'his',
  'they',
  'at',
  'be',
  'this',
  'from',
  'or',
  'have',
  'by',
  'not',
  'but',
  'what',
  'all',
  'were',
  'when',
  'we',
  'there',
  'can',
  'an',
  'your',
  'which',
  'their',
  'if',
  'do',
  'will',
  'each',
  'about',
  'how',
  'up',
  'out',
  'them',
  'then',
  'these',
  'so',
  'some',
  'would',
  'into',
  'has',
  'more',
  'her',
  'two',
  'like',
  'him',
  'see',
  'time',
  'could',
  'no',
  'make',
  'than',
  'been',
  'its',
  'now',
  'way',
  'may',
  'down',
  'did',
  'get',
  'come',
  'made',
  'find',
  'work',
  'here',
  'must',
  'before',
  'through',
  'back',
  'much',
  'where',
  'those',
  'after',
  'around',
  'should',
]);

/**
 * Common German-only words (rarely used in English)
 */
const GERMAN_ONLY = new Set([
  'der',
  'die',
  'das',
  'und',
  'ist',
  'von',
  'nicht',
  'mit',
  'ein',
  'eine',
  'als',
  'auch',
  'es',
  'ich',
  'auf',
  'für',
  'sie',
  'sich',
  'den',
  'des',
  'dem',
  'werden',
  'bei',
  'hat',
  'aus',
  'er',
  'haben',
  'oder',
  'aber',
  'nach',
  'noch',
  'kann',
  'vor',
  'wenn',
  'nur',
  'diese',
  'über',
  'so',
  'wie',
  'dieser',
  'diesem',
  'um',
  'durch',
  'zum',
  'zur',
  'bis',
  'schon',
  'weil',
  'dann',
  'muss',
  'zwischen',
  'unter',
  'gegen',
  'immer',
  'wieder',
  'gibt',
  'wurde',
  'waren',
  'hatte',
  'alle',
  'alles',
  'andere',
  'anderen',
  'heute',
  'morgen',
  'gestern',
  'jetzt',
  'hier',
  'dort',
  'ganz',
  'sehr',
  'viel',
  'mehr',
  'macht',
  'machen',
  'gehen',
  'kommen',
  'sollen',
  'müssen',
  'können',
  'wollen',
  'dürfen',
  'möchten',
]);

/**
 * LanguageDetector - Detects DE/EN with confidence scoring
 */
export class LanguageDetector {
  /**
   * Detect the language of input text
   */
  detect(text: string): LanguageDetectionResult {
    if (!text || text.trim().length === 0) {
      return {
        lang: 'en',
        confidence: 0,
        hasCodeSwitch: false,
        indicators: { de: 0, en: 0 },
      };
    }

    const words = this.tokenize(text);
    let deScore = 0;
    let enScore = 0;

    // Check for German special characters (strong indicator)
    if (GERMAN_CHARS_REGEX.test(text)) {
      deScore += 3;
    }

    // Check each word against language indicators
    for (const word of words) {
      const lower = word.toLowerCase();

      // Check German-only words
      if (GERMAN_ONLY.has(lower)) {
        deScore += 2;
        continue;
      }

      // Check English-only words
      if (ENGLISH_ONLY.has(lower)) {
        enScore += 2;
        continue;
      }

      // Check indicator categories
      for (const category of Object.keys(LANG_INDICATORS.de) as Array<
        keyof typeof LANG_INDICATORS.de
      >) {
        if (LANG_INDICATORS.de[category].includes(lower)) {
          deScore += 1;
        }
        if (LANG_INDICATORS.en[category].includes(lower)) {
          enScore += 1;
        }
      }
    }

    // Calculate confidence
    const total = deScore + enScore;
    const hasCodeSwitch =
      deScore > 0 && enScore > 0 && Math.min(deScore, enScore) / Math.max(deScore, enScore) > 0.3;

    let confidence: number;
    if (total === 0) {
      confidence = 0.5; // No clear indicators
    } else {
      confidence = Math.max(deScore, enScore) / total;
    }

    return {
      lang: deScore >= enScore ? 'de' : 'en',
      confidence: Math.min(1, confidence),
      hasCodeSwitch,
      indicators: { de: deScore, en: enScore },
    };
  }

  /**
   * Detect language with mixed-language handling
   */
  detectMixed(text: string): {
    primary: SupportedLanguage;
    hasCodeSwitch: boolean;
    segments?: Array<{ text: string; lang: SupportedLanguage }>;
  } {
    const result = this.detect(text);

    if (!result.hasCodeSwitch) {
      return {
        primary: result.lang,
        hasCodeSwitch: false,
      };
    }

    // Attempt to segment the text by language
    const segments: Array<{ text: string; lang: SupportedLanguage }> = [];
    const sentences = text.split(/[.!?]+/).filter((s) => s.trim().length > 0);

    for (const sentence of sentences) {
      const sentenceResult = this.detect(sentence);
      segments.push({
        text: sentence.trim(),
        lang: sentenceResult.lang,
      });
    }

    return {
      primary: result.lang,
      hasCodeSwitch: true,
      segments,
    };
  }

  /**
   * Quick check if text is likely German
   */
  isGerman(text: string): boolean {
    const result = this.detect(text);
    return result.lang === 'de' && result.confidence > 0.6;
  }

  /**
   * Quick check if text is likely English
   */
  isEnglish(text: string): boolean {
    const result = this.detect(text);
    return result.lang === 'en' && result.confidence > 0.6;
  }

  /**
   * Tokenize text into words
   */
  private tokenize(text: string): string[] {
    return text
      .toLowerCase()
      .replace(/[^\w\säöüÄÖÜß-]/g, ' ')
      .split(/\s+/)
      .filter((w) => w.length > 1);
  }
}

export default LanguageDetector;
