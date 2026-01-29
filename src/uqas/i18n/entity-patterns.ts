/**
 * Bilingual Entity Patterns (DE/EN)
 *
 * Extracts entities from text: persons, dates, times, files, projects, etc.
 * Handles German and English patterns including umlauts and date formats.
 */

import type { SupportedLanguage } from './index.js';

/**
 * Entity types
 */
export type EntityType =
  | 'person'
  | 'date'
  | 'time'
  | 'datetime'
  | 'duration'
  | 'file'
  | 'email'
  | 'project'
  | 'organization'
  | 'location'
  | 'money'
  | 'number'
  | 'unknown';

/**
 * Extracted entity
 */
export interface ExtractedEntity {
  type: EntityType;
  value: string;
  normalized?: string;
  confidence: number;
  position: { start: number; end: number };
  language: SupportedLanguage;
}

/**
 * Temporal expression
 */
export interface TemporalExpression {
  type: 'relative' | 'absolute' | 'range';
  original: string;
  normalized: string;
  date?: Date;
  startDate?: Date;
  endDate?: Date;
  relativeDays?: number;
}

/**
 * Person name patterns
 */
export const PERSON_PATTERNS = {
  de: [
    // Titles + Name
    /(?:Herr|Frau|Hr\.|Fr\.|Dr\.|Prof\.|Ing\.)\s+([A-ZÄÖÜ][a-zäöüß]+(?:\s+[A-ZÄÖÜ][a-zäöüß]+)?)/gi,
    // "mit/von/bei Person"
    /(?:mit|von|bei|an|für)\s+([A-ZÄÖÜ][a-zäöüß]+(?:\s+[A-ZÄÖÜ][a-zäöüß]+)?)/gi,
    // Full name pattern (First Last)
    /\b([A-ZÄÖÜ][a-zäöüß]+\s+[A-ZÄÖÜ][a-zäöüß]+)\b/g,
  ],
  en: [
    // Titles + Name
    /(?:Mr\.?|Mrs\.?|Ms\.?|Miss|Dr\.?|Prof\.?)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)/gi,
    // "with/from/by Person"
    /(?:with|from|by|to|for)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)/gi,
    // Full name pattern (First Last)
    /\b([A-Z][a-z]+\s+[A-Z][a-z]+)\b/g,
  ],
};

/**
 * Temporal patterns
 */
export const TEMPORAL_PATTERNS = {
  de: {
    relative: {
      // Today/Tomorrow/Yesterday
      heute: { days: 0, type: 'day' as const },
      morgen: { days: 1, type: 'day' as const },
      übermorgen: { days: 2, type: 'day' as const },
      gestern: { days: -1, type: 'day' as const },
      vorgestern: { days: -2, type: 'day' as const },
      // Week
      'diese woche': { days: 0, type: 'week' as const },
      'nächste woche': { days: 7, type: 'week' as const },
      'letzte woche': { days: -7, type: 'week' as const },
      'kommende woche': { days: 7, type: 'week' as const },
      'vergangene woche': { days: -7, type: 'week' as const },
      // Month
      'diesen monat': { days: 0, type: 'month' as const },
      'nächsten monat': { days: 30, type: 'month' as const },
      'letzten monat': { days: -30, type: 'month' as const },
      // Year
      'dieses jahr': { days: 0, type: 'year' as const },
      'nächstes jahr': { days: 365, type: 'year' as const },
      'letztes jahr': { days: -365, type: 'year' as const },
    },
    absolutePatterns: [
      // DD.MM.YYYY or DD.MM.YY
      /\b(\d{1,2})\.(\d{1,2})\.(\d{2,4})\b/g,
      // DD. Month YYYY
      /\b(\d{1,2})\.\s*(Januar|Februar|März|April|Mai|Juni|Juli|August|September|Oktober|November|Dezember)\s*(\d{4})?\b/gi,
    ],
    timePatterns: [
      // HH:MM or H:MM
      /\b(\d{1,2}):(\d{2})\s*(?:Uhr)?\b/g,
      // "um X Uhr"
      /um\s+(\d{1,2})\s*Uhr/gi,
    ],
  },
  en: {
    relative: {
      // Today/Tomorrow/Yesterday
      today: { days: 0, type: 'day' as const },
      tomorrow: { days: 1, type: 'day' as const },
      yesterday: { days: -1, type: 'day' as const },
      // Week
      'this week': { days: 0, type: 'week' as const },
      'next week': { days: 7, type: 'week' as const },
      'last week': { days: -7, type: 'week' as const },
      // Month
      'this month': { days: 0, type: 'month' as const },
      'next month': { days: 30, type: 'month' as const },
      'last month': { days: -30, type: 'month' as const },
      // Year
      'this year': { days: 0, type: 'year' as const },
      'next year': { days: 365, type: 'year' as const },
      'last year': { days: -365, type: 'year' as const },
    },
    absolutePatterns: [
      // MM/DD/YYYY or M/D/YY
      /\b(\d{1,2})\/(\d{1,2})\/(\d{2,4})\b/g,
      // Month DD, YYYY
      /\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2})(?:st|nd|rd|th)?,?\s*(\d{4})?\b/gi,
    ],
    timePatterns: [
      // HH:MM AM/PM
      /\b(\d{1,2}):(\d{2})\s*(AM|PM)?\b/gi,
      // "at X o'clock"
      /at\s+(\d{1,2})\s*(?:o'clock)?/gi,
    ],
  },
};

/**
 * File patterns
 */
const FILE_PATTERNS = [
  // File with extension
  /\b([\w\-_.]+\.(pdf|docx?|xlsx?|pptx?|txt|csv|jpg|jpeg|png|gif|zip|mp4|mp3))\b/gi,
  // Quoted filename
  /"([^"]+\.(pdf|docx?|xlsx?|pptx?|txt|csv))"/gi,
  /'([^']+\.(pdf|docx?|xlsx?|pptx?|txt|csv))'/gi,
];

/**
 * Email patterns
 */
const EMAIL_PATTERN = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g;

/**
 * Project patterns
 */
const PROJECT_PATTERNS = {
  de: [
    /Projekt\s+([A-ZÄÖÜ][a-zA-ZäöüÄÖÜß0-9_-]+)/gi,
    /\b([A-Z]{2,}[_-]?\d+)\b/g, // Acronym-based project codes like ABC-123
  ],
  en: [/Project\s+([A-Z][a-zA-Z0-9_-]+)/gi, /\b([A-Z]{2,}[_-]?\d+)\b/g],
};

/**
 * Money patterns
 */
const MONEY_PATTERNS = {
  de: [/\b(\d+(?:[.,]\d{2})?)\s*(?:€|EUR|Euro)\b/gi, /\b(?:€|EUR)\s*(\d+(?:[.,]\d{2})?)\b/gi],
  en: [
    /\b\$(\d+(?:,\d{3})*(?:\.\d{2})?)\b/g,
    /\b(\d+(?:,\d{3})*(?:\.\d{2})?)\s*(?:USD|dollars?)\b/gi,
    /\b(?:€|EUR)\s*(\d+(?:[.,]\d{2})?)\b/gi,
  ],
};

/**
 * German month names to numbers
 */
const GERMAN_MONTHS: Record<string, number> = {
  januar: 1,
  februar: 2,
  märz: 3,
  april: 4,
  mai: 5,
  juni: 6,
  juli: 7,
  august: 8,
  september: 9,
  oktober: 10,
  november: 11,
  dezember: 12,
};

/**
 * English month names to numbers
 */
const ENGLISH_MONTHS: Record<string, number> = {
  january: 1,
  february: 2,
  march: 3,
  april: 4,
  may: 5,
  june: 6,
  july: 7,
  august: 8,
  september: 9,
  october: 10,
  november: 11,
  december: 12,
};

/**
 * BilingualEntityRecognizer - Extracts entities from DE/EN text
 */
export class BilingualEntityRecognizer {
  /**
   * Extract all entities from text
   */
  extractAll(text: string, lang: SupportedLanguage = 'de'): ExtractedEntity[] {
    const entities: ExtractedEntity[] = [];

    // Extract persons
    entities.push(...this.extractPersons(text, lang));

    // Extract temporal expressions
    entities.push(...this.extractTemporalEntities(text, lang));

    // Extract files
    entities.push(...this.extractFiles(text));

    // Extract emails
    entities.push(...this.extractEmails(text));

    // Extract projects
    entities.push(...this.extractProjects(text, lang));

    // Extract money
    entities.push(...this.extractMoney(text, lang));

    // Sort by position
    entities.sort((a, b) => a.position.start - b.position.start);

    // Remove duplicates (overlapping entities)
    return this.deduplicateEntities(entities);
  }

  /**
   * Extract person names
   */
  extractPersons(text: string, lang: SupportedLanguage): ExtractedEntity[] {
    const entities: ExtractedEntity[] = [];
    const patterns = PERSON_PATTERNS[lang];

    for (const pattern of patterns) {
      // Reset regex
      pattern.lastIndex = 0;
      let match;
      while ((match = pattern.exec(text)) !== null) {
        const value = match[1] || match[0];
        // Skip common non-person words
        if (this.isCommonWord(value)) continue;

        entities.push({
          type: 'person',
          value: value.trim(),
          confidence: 0.75,
          position: { start: match.index, end: match.index + match[0].length },
          language: lang,
        });
      }
    }

    return entities;
  }

  /**
   * Extract temporal expressions
   */
  extractTemporalEntities(text: string, lang: SupportedLanguage): ExtractedEntity[] {
    const entities: ExtractedEntity[] = [];
    const patterns = TEMPORAL_PATTERNS[lang];

    // Check relative expressions
    const lowerText = text.toLowerCase();
    for (const [expression, info] of Object.entries(patterns.relative)) {
      const idx = lowerText.indexOf(expression);
      if (idx !== -1) {
        const temporal = this.resolveRelativeTemporal(expression, info, lang);
        entities.push({
          type: 'date',
          value: expression,
          normalized: temporal.normalized,
          confidence: 0.9,
          position: { start: idx, end: idx + expression.length },
          language: lang,
        });
      }
    }

    // Check absolute date patterns
    for (const pattern of patterns.absolutePatterns) {
      pattern.lastIndex = 0;
      let match;
      while ((match = pattern.exec(text)) !== null) {
        entities.push({
          type: 'date',
          value: match[0],
          normalized: this.normalizeDate(match, lang),
          confidence: 0.85,
          position: { start: match.index, end: match.index + match[0].length },
          language: lang,
        });
      }
    }

    // Check time patterns
    for (const pattern of patterns.timePatterns) {
      pattern.lastIndex = 0;
      let match;
      while ((match = pattern.exec(text)) !== null) {
        entities.push({
          type: 'time',
          value: match[0],
          normalized: this.normalizeTime(match),
          confidence: 0.85,
          position: { start: match.index, end: match.index + match[0].length },
          language: lang,
        });
      }
    }

    return entities;
  }

  /**
   * Extract temporal expression (combined method for backward compatibility)
   */
  extractTemporal(text: string, lang: SupportedLanguage): TemporalExpression | null {
    const patterns = TEMPORAL_PATTERNS[lang];
    const lowerText = text.toLowerCase();

    // Check relative expressions first
    for (const [expression, info] of Object.entries(patterns.relative)) {
      if (lowerText.includes(expression)) {
        return this.resolveRelativeTemporal(expression, info, lang);
      }
    }

    // Check absolute patterns
    for (const pattern of patterns.absolutePatterns) {
      pattern.lastIndex = 0;
      const match = pattern.exec(text);
      if (match) {
        const normalized = this.normalizeDate(match, lang);
        return {
          type: 'absolute',
          original: match[0],
          normalized,
          date: new Date(normalized),
        };
      }
    }

    return null;
  }

  /**
   * Extract file references
   */
  extractFiles(text: string): ExtractedEntity[] {
    const entities: ExtractedEntity[] = [];

    for (const pattern of FILE_PATTERNS) {
      pattern.lastIndex = 0;
      let match;
      while ((match = pattern.exec(text)) !== null) {
        entities.push({
          type: 'file',
          value: match[1] || match[0],
          confidence: 0.9,
          position: { start: match.index, end: match.index + match[0].length },
          language: 'en', // File names are typically language-neutral
        });
      }
    }

    return entities;
  }

  /**
   * Extract email addresses
   */
  extractEmails(text: string): ExtractedEntity[] {
    const entities: ExtractedEntity[] = [];
    EMAIL_PATTERN.lastIndex = 0;
    let match;

    while ((match = EMAIL_PATTERN.exec(text)) !== null) {
      entities.push({
        type: 'email',
        value: match[0],
        confidence: 0.95,
        position: { start: match.index, end: match.index + match[0].length },
        language: 'en',
      });
    }

    return entities;
  }

  /**
   * Extract project references
   */
  extractProjects(text: string, lang: SupportedLanguage): ExtractedEntity[] {
    const entities: ExtractedEntity[] = [];
    const patterns = PROJECT_PATTERNS[lang];

    for (const pattern of patterns) {
      pattern.lastIndex = 0;
      let match;
      while ((match = pattern.exec(text)) !== null) {
        entities.push({
          type: 'project',
          value: match[1] || match[0],
          confidence: 0.8,
          position: { start: match.index, end: match.index + match[0].length },
          language: lang,
        });
      }
    }

    return entities;
  }

  /**
   * Extract money amounts
   */
  extractMoney(text: string, lang: SupportedLanguage): ExtractedEntity[] {
    const entities: ExtractedEntity[] = [];
    const patterns = MONEY_PATTERNS[lang];

    for (const pattern of patterns) {
      pattern.lastIndex = 0;
      let match;
      while ((match = pattern.exec(text)) !== null) {
        entities.push({
          type: 'money',
          value: match[0],
          normalized: match[1]?.replace(',', '.'),
          confidence: 0.85,
          position: { start: match.index, end: match.index + match[0].length },
          language: lang,
        });
      }
    }

    return entities;
  }

  /**
   * Resolve relative temporal expression
   */
  private resolveRelativeTemporal(
    expression: string,
    info: { days: number; type: 'day' | 'week' | 'month' | 'year' },
    _lang: SupportedLanguage
  ): TemporalExpression {
    const now = new Date();
    const targetDate = new Date(now);
    targetDate.setDate(targetDate.getDate() + info.days);

    return {
      type: 'relative',
      original: expression,
      normalized: targetDate.toISOString().split('T')[0],
      date: targetDate,
      relativeDays: info.days,
    };
  }

  /**
   * Normalize date from regex match
   */
  private normalizeDate(match: RegExpExecArray, lang: SupportedLanguage): string {
    if (lang === 'de') {
      // DD.MM.YYYY format
      if (match[1] && match[2] && match[3]) {
        const day = match[1].padStart(2, '0');
        const month = match[2].padStart(2, '0');
        let year = match[3];
        if (year.length === 2) {
          year = (parseInt(year) > 50 ? '19' : '20') + year;
        }
        return `${year}-${month}-${day}`;
      }
      // DD. Month YYYY
      if (match[1] && match[2]) {
        const day = match[1].padStart(2, '0');
        const monthNum = GERMAN_MONTHS[match[2].toLowerCase()];
        const year = match[3] || new Date().getFullYear().toString();
        if (monthNum) {
          return `${year}-${String(monthNum).padStart(2, '0')}-${day}`;
        }
      }
    } else {
      // MM/DD/YYYY format
      if (match[1] && match[2] && match[3]) {
        const month = match[1].padStart(2, '0');
        const day = match[2].padStart(2, '0');
        let year = match[3];
        if (year.length === 2) {
          year = (parseInt(year) > 50 ? '19' : '20') + year;
        }
        return `${year}-${month}-${day}`;
      }
      // Month DD, YYYY
      if (match[1] && match[2]) {
        const monthNum = ENGLISH_MONTHS[match[1].toLowerCase()];
        const day = match[2].padStart(2, '0');
        const year = match[3] || new Date().getFullYear().toString();
        if (monthNum) {
          return `${year}-${String(monthNum).padStart(2, '0')}-${day}`;
        }
      }
    }
    return match[0];
  }

  /**
   * Normalize time from regex match
   */
  private normalizeTime(match: RegExpExecArray): string {
    let hours = parseInt(match[1]);
    const minutes = match[2] || '00';
    const ampm = match[3]?.toUpperCase();

    if (ampm === 'PM' && hours < 12) hours += 12;
    if (ampm === 'AM' && hours === 12) hours = 0;

    return `${String(hours).padStart(2, '0')}:${minutes}`;
  }

  /**
   * Check if a string is a common word (not a person name)
   */
  private isCommonWord(value: string): boolean {
    const commonWords = new Set([
      'mit',
      'von',
      'bei',
      'an',
      'für',
      'zu',
      'aus',
      'nach',
      'with',
      'from',
      'by',
      'to',
      'for',
      'at',
      'about',
      'der',
      'die',
      'das',
      'the',
      'a',
      'an',
      'und',
      'oder',
      'and',
      'or',
      'Meeting',
      'Email',
      'Termin',
      'Projekt',
      'Project',
      'Task',
    ]);
    return commonWords.has(value) || value.length < 2;
  }

  /**
   * Remove duplicate/overlapping entities
   */
  private deduplicateEntities(entities: ExtractedEntity[]): ExtractedEntity[] {
    const result: ExtractedEntity[] = [];

    for (const entity of entities) {
      // Check if this entity overlaps with any existing entity
      const overlaps = result.some(
        (e) =>
          (entity.position.start >= e.position.start && entity.position.start < e.position.end) ||
          (entity.position.end > e.position.start && entity.position.end <= e.position.end)
      );

      if (!overlaps) {
        result.push(entity);
      } else {
        // Keep the one with higher confidence
        const overlapIdx = result.findIndex(
          (e) =>
            (entity.position.start >= e.position.start && entity.position.start < e.position.end) ||
            (entity.position.end > e.position.start && entity.position.end <= e.position.end)
        );
        if (overlapIdx !== -1 && entity.confidence > result[overlapIdx].confidence) {
          result[overlapIdx] = entity;
        }
      }
    }

    return result;
  }
}

export default BilingualEntityRecognizer;
