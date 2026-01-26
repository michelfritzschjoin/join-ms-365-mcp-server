/**
 * NLP Enhancer for improved query processing with stemming, entity recognition, and intent classification
 */

import logger from './logger.js';

export interface IntentClassification {
  intent: 'search' | 'find' | 'list' | 'get' | 'create' | 'update' | 'delete' | 'unknown';
  confidence: number;
  entities: string[];
}

/**
 * Decomposed query result for NLP-based query splitting
 */
export interface DecomposedQuery {
  original: string;
  entity: string | null;
  entities: ExtractedEntity[];
  temporal: TemporalExpression | null;
  intent: QueryIntent;
  action: string | null;
  subQueries: string[];
  semanticVariants: string[];
  compoundParts: string[];
  ms365Context: MS365Context | null;
  urgency: UrgencyLevel;
  confidence: number;
  /** Structured Markdown summary of the query decomposition */
  markdown: string;
}

export interface ExtractedEntity {
  value: string;
  type: EntityType;
  confidence: number;
  position: { start: number; end: number };
}

export type EntityType =
  | 'person'
  | 'organization'
  | 'product'
  | 'food'
  | 'location'
  | 'date'
  | 'time'
  | 'email'
  | 'file'
  | 'project'
  | 'event'
  | 'task'
  | 'unknown';

export interface TemporalExpression {
  type: 'past' | 'future' | 'present' | 'range' | 'specific';
  expression: string;
  normalized: string | null;
  relativeDays?: number;
}

export interface QueryIntent {
  type:
    | 'when'
    | 'what'
    | 'who'
    | 'where'
    | 'how'
    | 'why'
    | 'list'
    | 'find'
    | 'count'
    | 'compare'
    | 'existence'
    | 'frequency'
    | 'last_occurrence'
    | 'unknown';
  confidence: number;
}

export interface MS365Context {
  service: 'calendar' | 'mail' | 'files' | 'teams' | 'tasks' | 'contacts' | 'notes' | 'search';
  suggestedTools: string[];
  searchScopes: string[];
}

export type UrgencyLevel = 'high' | 'medium' | 'low' | 'none';

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
 * Query Decomposer for splitting complex queries into sub-queries
 * Enhanced with compound word splitting, semantic variants, and MS365 context
 */
class QueryDecomposer {
  private readonly temporalPatterns: {
    pattern: RegExp;
    type: TemporalExpression['type'];
    normalized: string;
    relativeDays?: number;
  }[] = [
    // German past expressions with relative days
    { pattern: /\bzuletzt\b/gi, type: 'past', normalized: 'last', relativeDays: -7 },
    { pattern: /\bletzte[rns]?\b/gi, type: 'past', normalized: 'last', relativeDays: -7 },
    { pattern: /\bgestern\b/gi, type: 'past', normalized: 'yesterday', relativeDays: -1 },
    {
      pattern: /\bvorgestern\b/gi,
      type: 'past',
      normalized: 'day before yesterday',
      relativeDays: -2,
    },
    { pattern: /\bletzte woche\b/gi, type: 'past', normalized: 'last week', relativeDays: -7 },
    { pattern: /\bletzten monat\b/gi, type: 'past', normalized: 'last month', relativeDays: -30 },
    { pattern: /\bletztes jahr\b/gi, type: 'past', normalized: 'last year', relativeDays: -365 },
    {
      pattern: /\bvor (\d+) (tag|tagen|woche|wochen|monat|monaten|jahr|jahren)\b/gi,
      type: 'past',
      normalized: 'past',
    },
    { pattern: /\bkürzlich\b/gi, type: 'past', normalized: 'recently', relativeDays: -14 },
    { pattern: /\bneulich\b/gi, type: 'past', normalized: 'recently', relativeDays: -14 },
    { pattern: /\bfrüher\b/gi, type: 'past', normalized: 'earlier', relativeDays: -30 },
    // German future expressions
    { pattern: /\bmorgen\b/gi, type: 'future', normalized: 'tomorrow', relativeDays: 1 },
    {
      pattern: /\bübermorgen\b/gi,
      type: 'future',
      normalized: 'day after tomorrow',
      relativeDays: 2,
    },
    { pattern: /\bnächste[rns]?\b/gi, type: 'future', normalized: 'next', relativeDays: 7 },
    { pattern: /\bnächste woche\b/gi, type: 'future', normalized: 'next week', relativeDays: 7 },
    { pattern: /\bnächsten monat\b/gi, type: 'future', normalized: 'next month', relativeDays: 30 },
    { pattern: /\bnächstes jahr\b/gi, type: 'future', normalized: 'next year', relativeDays: 365 },
    {
      pattern: /\bin (\d+) (tag|tagen|woche|wochen|monat|monaten|jahr|jahren)\b/gi,
      type: 'future',
      normalized: 'future',
    },
    { pattern: /\bbald\b/gi, type: 'future', normalized: 'soon', relativeDays: 7 },
    { pattern: /\bdemnächst\b/gi, type: 'future', normalized: 'soon', relativeDays: 14 },
    // German present expressions
    { pattern: /\bheute\b/gi, type: 'present', normalized: 'today', relativeDays: 0 },
    { pattern: /\bjetzt\b/gi, type: 'present', normalized: 'now', relativeDays: 0 },
    { pattern: /\baktuell\b/gi, type: 'present', normalized: 'current', relativeDays: 0 },
    { pattern: /\bdiese woche\b/gi, type: 'present', normalized: 'this week', relativeDays: 0 },
    { pattern: /\bdiesen monat\b/gi, type: 'present', normalized: 'this month', relativeDays: 0 },
    { pattern: /\bdieses jahr\b/gi, type: 'present', normalized: 'this year', relativeDays: 0 },
    // English past expressions
    { pattern: /\blast\b/gi, type: 'past', normalized: 'last', relativeDays: -7 },
    { pattern: /\byesterday\b/gi, type: 'past', normalized: 'yesterday', relativeDays: -1 },
    { pattern: /\brecently\b/gi, type: 'past', normalized: 'recently', relativeDays: -14 },
    { pattern: /\bago\b/gi, type: 'past', normalized: 'ago' },
    { pattern: /\bprevious\b/gi, type: 'past', normalized: 'previous', relativeDays: -7 },
    { pattern: /\bpast\b/gi, type: 'past', normalized: 'past' },
    // English future expressions
    { pattern: /\btomorrow\b/gi, type: 'future', normalized: 'tomorrow', relativeDays: 1 },
    { pattern: /\bnext\b/gi, type: 'future', normalized: 'next', relativeDays: 7 },
    { pattern: /\bupcoming\b/gi, type: 'future', normalized: 'upcoming', relativeDays: 14 },
    { pattern: /\bsoon\b/gi, type: 'future', normalized: 'soon', relativeDays: 7 },
    // English present expressions
    { pattern: /\btoday\b/gi, type: 'present', normalized: 'today', relativeDays: 0 },
    { pattern: /\bnow\b/gi, type: 'present', normalized: 'now', relativeDays: 0 },
    { pattern: /\bcurrent\b/gi, type: 'present', normalized: 'current', relativeDays: 0 },
    { pattern: /\bthis week\b/gi, type: 'present', normalized: 'this week', relativeDays: 0 },
    { pattern: /\bthis month\b/gi, type: 'present', normalized: 'this month', relativeDays: 0 },
    { pattern: /\bthis year\b/gi, type: 'present', normalized: 'this year', relativeDays: 0 },
    // Specific date patterns
    { pattern: /\b\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\b/g, type: 'specific', normalized: 'date' },
    {
      pattern:
        /\b(januar|februar|märz|april|mai|juni|juli|august|september|oktober|november|dezember)\s+\d{1,4}\b/gi,
      type: 'specific',
      normalized: 'date',
    },
    {
      pattern:
        /\b(january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{1,4}\b/gi,
      type: 'specific',
      normalized: 'date',
    },
  ];

  private readonly questionWords: { pattern: RegExp; type: QueryIntent['type'] }[] = [
    // German question words - specific patterns first
    { pattern: /\bwann.*zuletzt\b/gi, type: 'last_occurrence' },
    { pattern: /\bwann.*letzte\b/gi, type: 'last_occurrence' },
    { pattern: /\bgab es\b/gi, type: 'existence' },
    { pattern: /\bwie oft\b/gi, type: 'frequency' },
    { pattern: /\bwie häufig\b/gi, type: 'frequency' },
    { pattern: /^wann\b/gi, type: 'when' },
    { pattern: /^was\b/gi, type: 'what' },
    { pattern: /^wer\b/gi, type: 'who' },
    { pattern: /^wo\b/gi, type: 'where' },
    { pattern: /^wie\b/gi, type: 'how' },
    { pattern: /^warum\b/gi, type: 'why' },
    { pattern: /^welche[rns]?\b/gi, type: 'what' },
    { pattern: /^wieviel[e]?\b/gi, type: 'count' },
    { pattern: /\bwann\b/gi, type: 'when' },
    // English question words - specific patterns first
    { pattern: /\bwhen.*last\b/gi, type: 'last_occurrence' },
    { pattern: /\bhow often\b/gi, type: 'frequency' },
    { pattern: /\bhow frequently\b/gi, type: 'frequency' },
    { pattern: /\bwas there\b/gi, type: 'existence' },
    { pattern: /\bdid.*exist\b/gi, type: 'existence' },
    { pattern: /^when\b/gi, type: 'when' },
    { pattern: /^what\b/gi, type: 'what' },
    { pattern: /^who\b/gi, type: 'who' },
    { pattern: /^where\b/gi, type: 'where' },
    { pattern: /^how\b/gi, type: 'how' },
    { pattern: /^why\b/gi, type: 'why' },
    { pattern: /^which\b/gi, type: 'what' },
    { pattern: /^how many\b/gi, type: 'count' },
    { pattern: /^how much\b/gi, type: 'count' },
    // Action patterns
    { pattern: /\bfind(e|en)?\b/gi, type: 'find' },
    { pattern: /\bsuche[n]?\b/gi, type: 'find' },
    { pattern: /\bzeig(e|en)?\b/gi, type: 'list' },
    { pattern: /\blist(e|en)?\b/gi, type: 'list' },
    { pattern: /\balle\b/gi, type: 'list' },
    { pattern: /\ball\b/gi, type: 'list' },
    { pattern: /\bvergleich(e|en)?\b/gi, type: 'compare' },
    { pattern: /\bcompare\b/gi, type: 'compare' },
  ];

  // German compound word patterns for splitting
  private readonly germanCompoundSuffixes = [
    'kraut',
    'pflanze',
    'blatt',
    'blume',
    'baum',
    'strauch',
    'gras',
    'moos',
    'pilz', // Plants
    'essen',
    'gericht',
    'speise',
    'suppe',
    'salat',
    'brot',
    'kuchen',
    'torte', // Food
    'bericht',
    'protokoll',
    'dokument',
    'datei',
    'plan',
    'liste',
    'tabelle', // Documents
    'meeting',
    'termin',
    'besprechung',
    'konferenz',
    'sitzung',
    'call', // Meetings
    'projekt',
    'aufgabe',
    'task',
    'arbeit',
    'vorgang', // Tasks
    'mail',
    'nachricht',
    'brief',
    'mitteilung',
    'info', // Messages
  ];

  // Semantic synonyms for common terms
  private readonly semanticSynonyms: Record<string, string[]> = {
    // Food-related
    essen: ['speise', 'gericht', 'mahlzeit', 'menü', 'kantine', 'mittagessen'],
    mittagessen: ['lunch', 'essen', 'mahlzeit', 'kantine'],
    speiseplan: ['menü', 'essensplan', 'kantinenplan', 'wochenplan'],
    // Meeting-related
    meeting: ['besprechung', 'termin', 'konferenz', 'call', 'sitzung'],
    termin: ['meeting', 'besprechung', 'verabredung', 'appointment'],
    // Document-related
    dokument: ['datei', 'file', 'unterlage', 'dokumente'],
    bericht: ['report', 'protokoll', 'zusammenfassung'],
    // Task-related
    aufgabe: ['task', 'todo', 'arbeit', 'vorgang'],
    // Person-related
    kollege: ['mitarbeiter', 'teammitglied', 'person'],
    chef: ['vorgesetzter', 'manager', 'leiter'],
  };

  // MS365 service detection patterns
  private readonly ms365Patterns: {
    pattern: RegExp;
    service: MS365Context['service'];
    tools: string[];
    scopes: string[];
  }[] = [
    {
      pattern: /\b(mail|e-?mail|nachricht|inbox|posteingang|send|senden|empfangen)\b/gi,
      service: 'mail',
      tools: ['list-mail-messages', 'search-mail', 'smart-query'],
      scopes: ['message'],
    },
    {
      pattern: /\b(kalender|calendar|termin|meeting|besprechung|event|einladung)\b/gi,
      service: 'calendar',
      tools: ['list-calendar-events', 'get-calendar-view', 'smart-query'],
      scopes: ['event'],
    },
    {
      pattern: /\b(datei|file|dokument|ordner|folder|onedrive|sharepoint|ablage)\b/gi,
      service: 'files',
      tools: ['list-drive-items', 'search-drive-items', 'smart-query'],
      scopes: ['driveItem'],
    },
    {
      pattern: /\b(teams|channel|chat|nachricht|team|gruppe)\b/gi,
      service: 'teams',
      tools: ['list-teams', 'list-team-channels', 'smart-query'],
      scopes: ['chatMessage', 'channel'],
    },
    {
      pattern: /\b(aufgabe|task|todo|planner|to-?do|erledigen)\b/gi,
      service: 'tasks',
      tools: ['list-planner-tasks', 'list-todo-lists', 'smart-query'],
      scopes: ['task'],
    },
    {
      pattern: /\b(kontakt|contact|person|adresse|telefon|people)\b/gi,
      service: 'contacts',
      tools: ['list-contacts', 'list-users', 'smart-query'],
      scopes: ['person', 'contact'],
    },
    {
      pattern: /\b(notiz|note|onenote|notizbuch|notebook)\b/gi,
      service: 'notes',
      tools: ['list-onenote-notebooks', 'smart-query'],
      scopes: ['note'],
    },
  ];

  // Urgency patterns
  private readonly urgencyPatterns: { pattern: RegExp; level: UrgencyLevel }[] = [
    {
      pattern: /\b(dringend|urgent|asap|sofort|immediately|wichtig|critical|kritisch)\b/gi,
      level: 'high',
    },
    { pattern: /\b(bald|soon|zeitnah|priorität|priority)\b/gi, level: 'medium' },
    { pattern: /\b(irgendwann|sometime|gelegentlich|wenn zeit)\b/gi, level: 'low' },
  ];

  private readonly stopWordsDE = new Set([
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
    'einen',
    'eines',
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
    'durch',
    'über',
    'unter',
    'vor',
    'nach',
    'bei',
    'aus',
    'ist',
    'sind',
    'war',
    'waren',
    'sein',
    'gewesen',
    'haben',
    'hat',
    'hatte',
    'hatten',
    'wird',
    'werden',
    'wurde',
    'wurden',
    'es',
    'gab',
    'gibt',
    'geben',
    'noch',
    'schon',
    'auch',
    'nur',
    'doch',
    'mal',
    'wann',
    'was',
    'wer',
    'wo',
    'wie',
    'warum',
    'welche',
    'welcher',
    'welches',
    'welchen',
    'welchem',
  ]);

  private readonly stopWordsEN = new Set([
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
    'can',
    'may',
    'might',
    'must',
    'shall',
    'it',
    'this',
    'that',
    'these',
    'those',
    'there',
    'when',
    'what',
    'who',
    'where',
    'how',
    'why',
    'which',
  ]);

  /**
   * Decompose a natural language query into structured parts
   */
  decompose(query: string): DecomposedQuery {
    const queryTrimmed = query.trim();

    // 1. Detect intent from question words
    const intent = this.detectIntent(queryTrimmed);

    // 2. Extract temporal expressions
    const temporal = this.extractTemporal(queryTrimmed);

    // 3. Extract main entity (what the query is about)
    const entity = this.extractEntity(queryTrimmed, temporal);

    // 4. Extract all entities with types
    const entities = this.extractAllEntities(queryTrimmed);

    // 5. Extract action if any
    const action = this.extractAction(queryTrimmed);

    // 6. Split German compound words
    const compoundParts = entity ? this.splitCompoundWord(entity) : [];

    // 7. Generate semantic variants
    const semanticVariants = this.generateSemanticVariants(entity, compoundParts);

    // 8. Detect MS365 context
    const ms365Context = this.detectMS365Context(queryTrimmed);

    // 9. Detect urgency
    const urgency = this.detectUrgency(queryTrimmed);

    // 10. Generate sub-queries (enhanced)
    const subQueries = this.generateSubQueries(
      entity,
      entities,
      temporal,
      intent,
      compoundParts,
      semanticVariants,
      ms365Context,
      queryTrimmed
    );

    // 11. Calculate confidence
    const confidence = this.calculateConfidence(entity, temporal, intent, entities.length);

    // 12. Generate markdown summary
    const markdown = this.generateMarkdownSummary({
      original: queryTrimmed,
      entity,
      entities,
      temporal,
      intent,
      action,
      subQueries,
      semanticVariants,
      compoundParts,
      ms365Context,
      urgency,
      confidence,
    });

    return {
      original: queryTrimmed,
      entity,
      entities,
      temporal,
      intent,
      action,
      subQueries,
      semanticVariants,
      compoundParts,
      ms365Context,
      urgency,
      confidence,
      markdown,
    };
  }

  /**
   * Generate a structured markdown summary of the query decomposition
   */
  private generateMarkdownSummary(data: Omit<DecomposedQuery, 'markdown'>): string {
    const lines: string[] = [];

    lines.push('## 🔍 Query Analysis');
    lines.push('');
    lines.push(`**Original Query:** "${data.original}"`);
    lines.push('');

    // Entity section
    if (data.entity) {
      lines.push('### 🎯 Extracted Entity');
      lines.push(`- **Main Entity:** \`${data.entity}\``);
      if (data.compoundParts.length > 0) {
        lines.push(`- **Compound Parts:** ${data.compoundParts.map((p) => `\`${p}\``).join(', ')}`);
      }
      lines.push('');
    }

    // All entities
    if (data.entities.length > 0) {
      lines.push('### 📋 Detected Entities');
      for (const entity of data.entities) {
        lines.push(
          `- **${entity.type}:** \`${entity.value}\` (${Math.round(entity.confidence * 100)}%)`
        );
      }
      lines.push('');
    }

    // Intent
    lines.push('### 🎭 Intent');
    lines.push(`- **Type:** ${data.intent.type}`);
    lines.push(`- **Confidence:** ${Math.round(data.intent.confidence * 100)}%`);
    lines.push('');

    // Temporal
    if (data.temporal) {
      lines.push('### ⏰ Temporal Context');
      lines.push(`- **Expression:** "${data.temporal.expression}"`);
      lines.push(`- **Type:** ${data.temporal.type}`);
      if (data.temporal.relativeDays !== undefined) {
        lines.push(`- **Relative Days:** ${data.temporal.relativeDays}`);
      }
      lines.push('');
    }

    // MS365 Context
    if (data.ms365Context) {
      lines.push('### 📧 MS365 Context');
      lines.push(`- **Service:** ${data.ms365Context.service}`);
      lines.push(
        `- **Suggested Tools:** ${data.ms365Context.suggestedTools.map((t) => `\`${t}\``).join(', ')}`
      );
      lines.push(`- **Search Scopes:** ${data.ms365Context.searchScopes.join(', ')}`);
      lines.push('');
    }

    // Urgency
    if (data.urgency !== 'none') {
      lines.push('### ⚡ Urgency');
      lines.push(`- **Level:** ${data.urgency}`);
      lines.push('');
    }

    // Sub-queries
    if (data.subQueries.length > 0) {
      lines.push('### 🔎 Generated Sub-Queries');
      lines.push('| # | Query |');
      lines.push('|---|-------|');
      data.subQueries.slice(0, 10).forEach((q, i) => {
        lines.push(`| ${i + 1} | \`${q}\` |`);
      });
      if (data.subQueries.length > 10) {
        lines.push(`| ... | *+${data.subQueries.length - 10} more* |`);
      }
      lines.push('');
    }

    // Semantic variants
    if (data.semanticVariants.length > 0) {
      lines.push('### 💡 Semantic Variants');
      lines.push(
        data.semanticVariants
          .slice(0, 5)
          .map((v) => `\`${v}\``)
          .join(', ')
      );
      lines.push('');
    }

    // Overall confidence
    lines.push('---');
    lines.push(`**Overall Confidence:** ${Math.round(data.confidence * 100)}%`);

    return lines.join('\n');
  }

  private detectIntent(query: string): QueryIntent {
    for (const { pattern, type } of this.questionWords) {
      // Reset lastIndex for global patterns
      pattern.lastIndex = 0;
      if (pattern.test(query)) {
        return { type, confidence: 0.85 };
      }
    }
    return { type: 'unknown', confidence: 0.3 };
  }

  private extractTemporal(query: string): TemporalExpression | null {
    for (const { pattern, type, normalized, relativeDays } of this.temporalPatterns) {
      pattern.lastIndex = 0;
      const match = query.match(pattern);
      if (match) {
        return {
          type,
          expression: match[0],
          normalized,
          relativeDays,
        };
      }
    }
    return null;
  }

  private extractEntity(query: string, temporal: TemporalExpression | null): string | null {
    let cleanedQuery = query;

    // Remove question words
    cleanedQuery = cleanedQuery.replace(
      /^(wann|was|wer|wo|wie|warum|welche[rns]?|when|what|who|where|how|why|which)\s+/gi,
      ''
    );

    // Remove temporal expression if found
    if (temporal) {
      cleanedQuery = cleanedQuery.replace(
        new RegExp(this.escapeRegExp(temporal.expression), 'gi'),
        ''
      );
    }

    // Remove common verbs and auxiliary words
    cleanedQuery = cleanedQuery.replace(
      /\b(gab|gibt|geben|ist|sind|war|waren|haben|hat|hatte|es|there|is|are|was|were|have|has|had)\b/gi,
      ''
    );

    // Remove punctuation
    cleanedQuery = cleanedQuery.replace(/[?!.,;:]/g, '');

    // Split into words and filter stop words
    const words = cleanedQuery.split(/\s+/).filter((word) => {
      const lower = word.toLowerCase();
      return word.length > 1 && !this.stopWordsDE.has(lower) && !this.stopWordsEN.has(lower);
    });

    if (words.length === 0) {
      return null;
    }

    // Return remaining words as entity
    return words.join(' ').trim() || null;
  }

  private extractAllEntities(query: string): ExtractedEntity[] {
    const entities: ExtractedEntity[] = [];
    const queryLower = query.toLowerCase();

    // Email pattern
    const emailPattern = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g;
    let match;
    while ((match = emailPattern.exec(query)) !== null) {
      entities.push({
        value: match[0],
        type: 'email',
        confidence: 0.95,
        position: { start: match.index, end: match.index + match[0].length },
      });
    }

    // Person patterns (names with titles)
    const personPattern =
      /\b(Herr|Frau|Mr\.?|Mrs\.?|Ms\.?|Dr\.?)\s+([A-ZÄÖÜ][a-zäöüß]+(?:\s+[A-ZÄÖÜ][a-zäöüß]+)?)\b/g;
    while ((match = personPattern.exec(query)) !== null) {
      entities.push({
        value: match[0],
        type: 'person',
        confidence: 0.9,
        position: { start: match.index, end: match.index + match[0].length },
      });
    }

    // File patterns
    const filePattern = /\b[\w-]+\.(pdf|docx?|xlsx?|pptx?|txt|csv|jpg|png|gif)\b/gi;
    while ((match = filePattern.exec(query)) !== null) {
      entities.push({
        value: match[0],
        type: 'file',
        confidence: 0.95,
        position: { start: match.index, end: match.index + match[0].length },
      });
    }

    // Project patterns
    const projectPattern = /\b(Projekt|Project)\s+([A-ZÄÖÜ][a-zA-ZäöüÄÖÜß0-9-]+)\b/gi;
    while ((match = projectPattern.exec(query)) !== null) {
      entities.push({
        value: match[0],
        type: 'project',
        confidence: 0.85,
        position: { start: match.index, end: match.index + match[0].length },
      });
    }

    // Food patterns (German)
    const foodPatterns = [
      /\b(Schicht|Blumen|Rot|Weiß|Grün|Spitz|Rosen|Chinakohl)\s*kohl\b/gi,
      /\b\w+(kraut|suppe|salat|gericht|essen|kuchen|torte|brot)\b/gi,
      /\b(Nudeln|Reis|Kartoffeln?|Fleisch|Fisch|Gemüse|Obst)\b/gi,
    ];
    for (const pattern of foodPatterns) {
      while ((match = pattern.exec(query)) !== null) {
        entities.push({
          value: match[0],
          type: 'food',
          confidence: 0.8,
          position: { start: match.index, end: match.index + match[0].length },
        });
      }
    }

    return entities;
  }

  private splitCompoundWord(word: string): string[] {
    const parts: string[] = [];
    const wordLower = word.toLowerCase();

    // Check for known German compound suffixes
    for (const suffix of this.germanCompoundSuffixes) {
      if (wordLower.endsWith(suffix) && wordLower.length > suffix.length + 2) {
        const prefix = word.slice(0, word.length - suffix.length);
        if (prefix.length >= 3) {
          parts.push(prefix);
          parts.push(suffix);
          // Also add capitalized versions
          parts.push(prefix.charAt(0).toUpperCase() + prefix.slice(1).toLowerCase());
          parts.push(suffix.charAt(0).toUpperCase() + suffix.slice(1).toLowerCase());
        }
        break;
      }
    }

    // If word is camelCase or PascalCase, split it
    const camelParts = word.split(/(?=[A-Z])/).filter((p) => p.length > 0);
    if (camelParts.length > 1) {
      parts.push(...camelParts);
    }

    // Try splitting on common connecting characters
    if (word.includes('-')) {
      parts.push(...word.split('-').filter((p) => p.length > 1));
    }
    if (word.includes('_')) {
      parts.push(...word.split('_').filter((p) => p.length > 1));
    }

    return [...new Set(parts)];
  }

  private generateSemanticVariants(entity: string | null, compoundParts: string[]): string[] {
    const variants: string[] = [];

    if (!entity) return variants;

    const entityLower = entity.toLowerCase();

    // Check direct synonyms
    for (const [term, synonyms] of Object.entries(this.semanticSynonyms)) {
      if (entityLower.includes(term)) {
        for (const synonym of synonyms) {
          variants.push(entityLower.replace(term, synonym));
        }
      }
      // Also check if any synonym matches
      for (const synonym of synonyms) {
        if (entityLower.includes(synonym)) {
          variants.push(entityLower.replace(synonym, term));
        }
      }
    }

    // Add compound parts as variants
    for (const part of compoundParts) {
      if (part.length >= 3) {
        variants.push(part);
        // Check if compound parts have synonyms
        const partLower = part.toLowerCase();
        if (this.semanticSynonyms[partLower]) {
          variants.push(...this.semanticSynonyms[partLower]);
        }
      }
    }

    // Add common spelling variations
    if (entityLower.includes('ä')) variants.push(entityLower.replace(/ä/g, 'ae'));
    if (entityLower.includes('ö')) variants.push(entityLower.replace(/ö/g, 'oe'));
    if (entityLower.includes('ü')) variants.push(entityLower.replace(/ü/g, 'ue'));
    if (entityLower.includes('ß')) variants.push(entityLower.replace(/ß/g, 'ss'));

    return [...new Set(variants)].filter((v) => v !== entityLower && v.length > 2);
  }

  private detectMS365Context(query: string): MS365Context | null {
    for (const { pattern, service, tools, scopes } of this.ms365Patterns) {
      pattern.lastIndex = 0;
      if (pattern.test(query)) {
        return {
          service,
          suggestedTools: tools,
          searchScopes: scopes,
        };
      }
    }

    // Default to search if no specific context detected
    return {
      service: 'search',
      suggestedTools: ['smart-query', 'search-content'],
      searchScopes: ['driveItem', 'message', 'event'],
    };
  }

  private detectUrgency(query: string): UrgencyLevel {
    for (const { pattern, level } of this.urgencyPatterns) {
      pattern.lastIndex = 0;
      if (pattern.test(query)) {
        return level;
      }
    }
    return 'none';
  }

  private extractAction(query: string): string | null {
    const actionPatterns = [
      { pattern: /\b(erstellen|create|anlegen|machen|neu)\b/gi, action: 'create' },
      { pattern: /\b(löschen|delete|entfernen|remove)\b/gi, action: 'delete' },
      { pattern: /\b(ändern|update|bearbeiten|edit|modify|aktualisieren)\b/gi, action: 'update' },
      { pattern: /\b(senden|send|schicken|verschicken)\b/gi, action: 'send' },
      { pattern: /\b(suchen|search|finden|find|durchsuchen)\b/gi, action: 'search' },
      { pattern: /\b(anzeigen|show|display|zeigen|öffnen|open)\b/gi, action: 'view' },
      { pattern: /\b(herunterladen|download|exportieren|export)\b/gi, action: 'download' },
      { pattern: /\b(teilen|share|freigeben)\b/gi, action: 'share' },
    ];

    for (const { pattern, action } of actionPatterns) {
      pattern.lastIndex = 0;
      if (pattern.test(query)) {
        return action;
      }
    }
    return null;
  }

  private generateSubQueries(
    entity: string | null,
    entities: ExtractedEntity[],
    temporal: TemporalExpression | null,
    intent: QueryIntent,
    compoundParts: string[],
    semanticVariants: string[],
    ms365Context: MS365Context | null,
    original: string
  ): string[] {
    const subQueries: string[] = [];

    // 1. Entity-only query (most important for search)
    if (entity) {
      subQueries.push(entity);
    }

    // 2. Entity with temporal context
    if (entity && temporal) {
      subQueries.push(`${entity} ${temporal.expression}`);
      if (temporal.normalized && temporal.normalized !== temporal.expression.toLowerCase()) {
        subQueries.push(`${entity} ${temporal.normalized}`);
      }
    }

    // 3. Compound parts as separate queries
    for (const part of compoundParts) {
      if (part.length >= 3) {
        subQueries.push(part);
        if (temporal) {
          subQueries.push(`${part} ${temporal.expression}`);
        }
      }
    }

    // 4. Semantic variants
    for (const variant of semanticVariants.slice(0, 5)) {
      subQueries.push(variant);
      if (temporal) {
        subQueries.push(`${variant} ${temporal.expression}`);
      }
    }

    // 5. Intent-specific queries
    if (entity) {
      switch (intent.type) {
        case 'when':
        case 'last_occurrence':
          subQueries.push(`${entity} Datum`);
          subQueries.push(`${entity} date`);
          subQueries.push(`${entity} wann`);
          subQueries.push(`${entity} letzte`);
          break;
        case 'who':
          subQueries.push(`${entity} Person`);
          subQueries.push(`${entity} Teilnehmer`);
          subQueries.push(`${entity} Organisator`);
          break;
        case 'where':
          subQueries.push(`${entity} Ort`);
          subQueries.push(`${entity} location`);
          subQueries.push(`${entity} Raum`);
          break;
        case 'count':
        case 'frequency':
          subQueries.push(`Anzahl ${entity}`);
          subQueries.push(`${entity} count`);
          subQueries.push(`${entity} häufigkeit`);
          break;
        case 'existence':
          subQueries.push(`${entity}`);
          subQueries.push(`${entity} vorhanden`);
          break;
        default:
          break;
      }
    }

    // 6. MS365 context-specific queries
    if (entity && ms365Context) {
      switch (ms365Context.service) {
        case 'calendar':
          subQueries.push(`${entity} termin`);
          subQueries.push(`${entity} meeting`);
          subQueries.push(`${entity} event`);
          break;
        case 'mail':
          subQueries.push(`${entity} mail`);
          subQueries.push(`${entity} email`);
          subQueries.push(`${entity} nachricht`);
          break;
        case 'files':
          subQueries.push(`${entity} datei`);
          subQueries.push(`${entity} dokument`);
          subQueries.push(`${entity} file`);
          break;
        case 'tasks':
          subQueries.push(`${entity} aufgabe`);
          subQueries.push(`${entity} task`);
          subQueries.push(`${entity} todo`);
          break;
        default:
          break;
      }
    }

    // 7. Add extracted entity values
    for (const extractedEntity of entities) {
      if (extractedEntity.value !== entity && extractedEntity.value.length >= 3) {
        subQueries.push(extractedEntity.value);
      }
    }

    // 8. Add variations without stop words
    if (entity) {
      const entityWords = entity.split(/\s+/);
      if (entityWords.length > 1) {
        for (const word of entityWords) {
          if (
            word.length > 3 &&
            !this.stopWordsDE.has(word.toLowerCase()) &&
            !this.stopWordsEN.has(word.toLowerCase())
          ) {
            subQueries.push(word);
          }
        }
      }
    }

    // 9. Fuzzy variants (common typo corrections)
    if (entity && entity.length >= 4) {
      // Add wildcard-style variants for partial matching
      subQueries.push(`${entity.slice(0, Math.ceil(entity.length * 0.7))}*`);
    }

    // Remove duplicates, empty strings, and limit
    return [...new Set(subQueries)].filter((q) => q.trim().length > 0).slice(0, 20); // Limit to top 20 queries
  }

  private calculateConfidence(
    entity: string | null,
    temporal: TemporalExpression | null,
    intent: QueryIntent,
    entityCount: number
  ): number {
    let confidence = 0.2; // Base confidence

    if (entity) {
      confidence += 0.3;
    }
    if (temporal) {
      confidence += 0.15;
    }
    if (intent.type !== 'unknown') {
      confidence += intent.confidence * 0.25;
    }
    if (entityCount > 0) {
      confidence += Math.min(0.1, entityCount * 0.03);
    }

    return Math.min(1.0, confidence);
  }

  private escapeRegExp(string: string): string {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
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
  private queryDecomposer: QueryDecomposer;
  private readonly enabled: boolean;

  constructor() {
    this.stemmer = new SimpleStemmer();
    this.entityRecognizer = new EntityRecognizer();
    this.intentClassifier = new IntentClassifier();
    this.queryDecomposer = new QueryDecomposer();
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
   * Decompose a complex query into structured parts and sub-queries
   *
   * Example:
   * Input: "Wann gab es Schichtkraut zuletzt?"
   * Output: {
   *   original: "Wann gab es Schichtkraut zuletzt?",
   *   entity: "Schichtkraut",
   *   entities: [{ value: "Schichtkraut", type: "food", confidence: 0.8, position: {...} }],
   *   temporal: { type: "past", expression: "zuletzt", normalized: "last", relativeDays: -7 },
   *   intent: { type: "last_occurrence", confidence: 0.85 },
   *   action: null,
   *   subQueries: ["Schichtkraut", "Schichtkraut zuletzt", "Schicht", "kraut", "Schichtkraut Datum"],
   *   semanticVariants: ["schicht", "kraut"],
   *   compoundParts: ["Schicht", "kraut"],
   *   ms365Context: { service: "search", suggestedTools: ["smart-query"], searchScopes: ["driveItem", "message", "event"] },
   *   urgency: "none",
   *   confidence: 0.85
   * }
   */
  decomposeQuery(query: string): DecomposedQuery {
    if (!this.enabled) {
      const fallbackResult: Omit<DecomposedQuery, 'markdown'> = {
        original: query,
        entity: query.trim(),
        entities: [],
        temporal: null,
        intent: { type: 'unknown', confidence: 0.3 },
        action: null,
        subQueries: [query.trim()],
        semanticVariants: [],
        compoundParts: [],
        ms365Context: null,
        urgency: 'none',
        confidence: 0.3,
      };
      return {
        ...fallbackResult,
        markdown: `## 🔍 Query Analysis\n\n**Original Query:** "${query}"\n\n**Note:** NLP processing is disabled.\n\n**Sub-Query:** \`${query.trim()}\``,
      };
    }
    return this.queryDecomposer.decompose(query);
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
