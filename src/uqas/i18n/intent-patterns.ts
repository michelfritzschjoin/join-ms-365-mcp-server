/**
 * Bilingual Intent Patterns (DE/EN)
 *
 * Recognizes user intent from questions in both German and English.
 * Maps to M365 data sources and optimal query strategies.
 */

import type { SupportedLanguage } from './index.js';

/**
 * Intent types for M365 queries
 */
export type IntentType =
  | 'when' // Time-based questions
  | 'what' // Content/description questions
  | 'who' // Person-related questions
  | 'where' // Location questions
  | 'how_many' // Count/quantity questions
  | 'list' // List/enumeration requests
  | 'find' // Search/locate requests
  | 'summary' // Summary/overview requests
  | 'compare' // Comparison requests
  | 'action' // Action requests (create, update, delete)
  | 'status' // Status inquiries
  | 'unknown'; // Unrecognized intent

/**
 * Recognized intent with metadata
 */
export interface BilingualIntent {
  type: IntentType;
  confidence: number;
  language: SupportedLanguage;
  matchedPattern: string;
  suggestedSources: string[];
  suggestedActions: string[];
}

/**
 * Intent pattern definition
 */
interface IntentPatternDef {
  patterns: RegExp[];
  sources: string[];
  actions: string[];
}

/**
 * Bilingual intent patterns
 */
export const INTENT_PATTERNS: Record<IntentType, { de: IntentPatternDef; en: IntentPatternDef }> = {
  when: {
    de: {
      patterns: [
        /^wann\b/i,
        /\bwann\b.*\?/i,
        /zu welchem zeitpunkt/i,
        /um wieviel uhr/i,
        /an welchem tag/i,
        /wann.*zuletzt/i,
        /wann.*nächst/i,
        /seit wann/i,
        /bis wann/i,
      ],
      sources: ['calendar', 'email'],
      actions: ['get-calendar-view', 'list-calendar-events'],
    },
    en: {
      patterns: [
        /^when\b/i,
        /\bwhen\b.*\?/i,
        /at what time/i,
        /on what day/i,
        /when.*last/i,
        /when.*next/i,
        /since when/i,
        /until when/i,
      ],
      sources: ['calendar', 'email'],
      actions: ['get-calendar-view', 'list-calendar-events'],
    },
  },

  what: {
    de: {
      patterns: [
        /^was\b/i,
        /\bwas\b.*\?/i,
        /welche[rns]?\b/i,
        /was für ein/i,
        /worum.*geht/i,
        /was.*inhalt/i,
        /was.*thema/i,
      ],
      sources: ['email', 'files', 'calendar'],
      actions: ['search', 'get-mail-message', 'get-drive-item'],
    },
    en: {
      patterns: [
        /^what\b/i,
        /\bwhat\b.*\?/i,
        /which\b/i,
        /what kind of/i,
        /what.*about/i,
        /what.*content/i,
        /what.*topic/i,
      ],
      sources: ['email', 'files', 'calendar'],
      actions: ['search', 'get-mail-message', 'get-drive-item'],
    },
  },

  who: {
    de: {
      patterns: [
        /^wer\b/i,
        /\bwer\b.*\?/i,
        /mit wem/i,
        /von wem/i,
        /an wen/i,
        /für wen/i,
        /wessen\b/i,
        /welche person/i,
        /welcher kollege/i,
      ],
      sources: ['people', 'email', 'calendar', 'teams'],
      actions: ['list-users', 'search-people', 'list-attendees'],
    },
    en: {
      patterns: [
        /^who\b/i,
        /\bwho\b.*\?/i,
        /with whom/i,
        /from whom/i,
        /to whom/i,
        /for whom/i,
        /whose\b/i,
        /which person/i,
        /which colleague/i,
      ],
      sources: ['people', 'email', 'calendar', 'teams'],
      actions: ['list-users', 'search-people', 'list-attendees'],
    },
  },

  where: {
    de: {
      patterns: [
        /^wo\b/i,
        /\bwo\b.*\?/i,
        /wohin\b/i,
        /an welchem ort/i,
        /in welchem ordner/i,
        /wo.*gespeichert/i,
        /wo.*abgelegt/i,
      ],
      sources: ['files', 'calendar', 'sharepoint'],
      actions: ['search-files', 'list-drive-items', 'search-sites'],
    },
    en: {
      patterns: [
        /^where\b/i,
        /\bwhere\b.*\?/i,
        /in which folder/i,
        /where.*saved/i,
        /where.*stored/i,
        /where.*located/i,
      ],
      sources: ['files', 'calendar', 'sharepoint'],
      actions: ['search-files', 'list-drive-items', 'search-sites'],
    },
  },

  how_many: {
    de: {
      patterns: [
        /wie ?viele?\b/i,
        /wieviel\b/i,
        /anzahl\b/i,
        /wie oft\b/i,
        /häufigkeit\b/i,
        /zähle?\b/i,
      ],
      sources: ['calendar', 'email', 'files', 'tasks'],
      actions: ['count', 'list-with-count'],
    },
    en: {
      patterns: [
        /how many\b/i,
        /how much\b/i,
        /count\b/i,
        /number of\b/i,
        /how often\b/i,
        /frequency\b/i,
      ],
      sources: ['calendar', 'email', 'files', 'tasks'],
      actions: ['count', 'list-with-count'],
    },
  },

  list: {
    de: {
      patterns: [
        /^zeig(e|en)?\b/i,
        /^liste?\b/i,
        /alle\b.*\b(e-?mails?|termine?|dateien|aufgaben|kontakte)/i,
        /übersicht\b/i,
        /auflistung\b/i,
        /gib mir.*liste/i,
      ],
      sources: ['email', 'calendar', 'files', 'tasks', 'contacts'],
      actions: ['list-mail-messages', 'list-calendar-events', 'list-drive-items', 'list-tasks'],
    },
    en: {
      patterns: [
        /^show\b/i,
        /^list\b/i,
        /all\b.*\b(emails?|events?|files?|tasks?|contacts?)/i,
        /overview\b/i,
        /give me.*list/i,
        /display\b/i,
      ],
      sources: ['email', 'calendar', 'files', 'tasks', 'contacts'],
      actions: ['list-mail-messages', 'list-calendar-events', 'list-drive-items', 'list-tasks'],
    },
  },

  find: {
    de: {
      patterns: [
        /^find(e|en)?\b/i,
        /^such(e|en)?\b/i,
        /wo ist\b/i,
        /wo sind\b/i,
        /lokalisiere?\b/i,
        /hast du.*gefunden/i,
      ],
      sources: ['files', 'email', 'calendar', 'sharepoint'],
      actions: ['search', 'search-files', 'search-mail'],
    },
    en: {
      patterns: [
        /^find\b/i,
        /^search\b/i,
        /^look for\b/i,
        /locate\b/i,
        /where is\b/i,
        /where are\b/i,
      ],
      sources: ['files', 'email', 'calendar', 'sharepoint'],
      actions: ['search', 'search-files', 'search-mail'],
    },
  },

  summary: {
    de: {
      patterns: [
        /zusammenfassung\b/i,
        /überblick\b/i,
        /was.*passiert/i,
        /was.*neu/i,
        /status.*update/i,
        /kurzfassung\b/i,
        /gib mir einen überblick/i,
        /fass.*zusammen/i,
      ],
      sources: ['email', 'calendar', 'tasks', 'teams'],
      actions: ['my-day', 'my-week', 'summary'],
    },
    en: {
      patterns: [
        /summary\b/i,
        /overview\b/i,
        /what.*happened/i,
        /what.*new/i,
        /status.*update/i,
        /recap\b/i,
        /give me an overview/i,
        /summarize\b/i,
      ],
      sources: ['email', 'calendar', 'tasks', 'teams'],
      actions: ['my-day', 'my-week', 'summary'],
    },
  },

  compare: {
    de: {
      patterns: [
        /vergleich(e|en)?\b/i,
        /unterschied\b/i,
        /anders als\b/i,
        /im vergleich zu\b/i,
        /gegenüberstellung\b/i,
      ],
      sources: ['files', 'calendar', 'email'],
      actions: ['compare', 'diff'],
    },
    en: {
      patterns: [
        /compare\b/i,
        /difference\b/i,
        /different from\b/i,
        /compared to\b/i,
        /versus\b/i,
        /vs\.?\b/i,
      ],
      sources: ['files', 'calendar', 'email'],
      actions: ['compare', 'diff'],
    },
  },

  action: {
    de: {
      patterns: [
        /erstell(e|en)?\b/i,
        /anleg(e|en)?\b/i,
        /send(e|en)?\b/i,
        /schick(e|en)?\b/i,
        /lösch(e|en)?\b/i,
        /entfern(e|en)?\b/i,
        /aktualisier(e|en)?\b/i,
        /änder(e|n)?\b/i,
        /verschib(e|en)?\b/i,
        /kopier(e|en)?\b/i,
      ],
      sources: ['email', 'calendar', 'files', 'tasks'],
      actions: ['create', 'send', 'delete', 'update', 'move', 'copy'],
    },
    en: {
      patterns: [
        /create\b/i,
        /make\b/i,
        /send\b/i,
        /delete\b/i,
        /remove\b/i,
        /update\b/i,
        /change\b/i,
        /move\b/i,
        /copy\b/i,
      ],
      sources: ['email', 'calendar', 'files', 'tasks'],
      actions: ['create', 'send', 'delete', 'update', 'move', 'copy'],
    },
  },

  status: {
    de: {
      patterns: [/status\b/i, /stand\b/i, /wie.*läuft/i, /fortschritt\b/i, /was macht\b/i],
      sources: ['tasks', 'email', 'planner'],
      actions: ['get-task-status', 'list-planner-tasks'],
    },
    en: {
      patterns: [/status\b/i, /state\b/i, /how.*going/i, /progress\b/i, /what.*doing/i],
      sources: ['tasks', 'email', 'planner'],
      actions: ['get-task-status', 'list-planner-tasks'],
    },
  },

  unknown: {
    de: {
      patterns: [],
      sources: ['search'],
      actions: ['smart-query'],
    },
    en: {
      patterns: [],
      sources: ['search'],
      actions: ['smart-query'],
    },
  },
};

/**
 * BilingualIntentRecognizer - Recognizes intent from DE/EN questions
 */
export class BilingualIntentRecognizer {
  /**
   * Recognize intent from question text
   */
  recognize(text: string): BilingualIntent {
    const normalized = text.trim();

    // Try each intent type
    for (const [intentType, patterns] of Object.entries(INTENT_PATTERNS)) {
      if (intentType === 'unknown') continue;

      // Check German patterns
      for (const pattern of patterns.de.patterns) {
        if (pattern.test(normalized)) {
          return {
            type: intentType as IntentType,
            confidence: 0.85,
            language: 'de',
            matchedPattern: pattern.source,
            suggestedSources: patterns.de.sources,
            suggestedActions: patterns.de.actions,
          };
        }
      }

      // Check English patterns
      for (const pattern of patterns.en.patterns) {
        if (pattern.test(normalized)) {
          return {
            type: intentType as IntentType,
            confidence: 0.85,
            language: 'en',
            matchedPattern: pattern.source,
            suggestedSources: patterns.en.sources,
            suggestedActions: patterns.en.actions,
          };
        }
      }
    }

    // No pattern matched - return unknown
    return {
      type: 'unknown',
      confidence: 0.3,
      language: 'en', // Default
      matchedPattern: '',
      suggestedSources: ['search'],
      suggestedActions: ['smart-query'],
    };
  }

  /**
   * Get all intents that could match (for ambiguous queries)
   */
  recognizeAll(text: string): BilingualIntent[] {
    const matches: BilingualIntent[] = [];
    const normalized = text.trim();

    for (const [intentType, patterns] of Object.entries(INTENT_PATTERNS)) {
      if (intentType === 'unknown') continue;

      // Check German patterns
      for (const pattern of patterns.de.patterns) {
        if (pattern.test(normalized)) {
          matches.push({
            type: intentType as IntentType,
            confidence: 0.85,
            language: 'de',
            matchedPattern: pattern.source,
            suggestedSources: patterns.de.sources,
            suggestedActions: patterns.de.actions,
          });
          break; // Only one match per intent type
        }
      }

      // Check English patterns
      for (const pattern of patterns.en.patterns) {
        if (pattern.test(normalized)) {
          // Avoid duplicate if already matched German
          const exists = matches.some((m) => m.type === intentType);
          if (!exists) {
            matches.push({
              type: intentType as IntentType,
              confidence: 0.85,
              language: 'en',
              matchedPattern: pattern.source,
              suggestedSources: patterns.en.sources,
              suggestedActions: patterns.en.actions,
            });
          }
          break;
        }
      }
    }

    // Sort by confidence
    return matches.sort((a, b) => b.confidence - a.confidence);
  }

  /**
   * Get suggested data sources for an intent
   */
  getSuggestedSources(intentType: IntentType, lang: SupportedLanguage): string[] {
    const pattern = INTENT_PATTERNS[intentType];
    return pattern?.[lang]?.sources ?? ['search'];
  }

  /**
   * Get suggested actions for an intent
   */
  getSuggestedActions(intentType: IntentType, lang: SupportedLanguage): string[] {
    const pattern = INTENT_PATTERNS[intentType];
    return pattern?.[lang]?.actions ?? ['smart-query'];
  }
}

export default BilingualIntentRecognizer;
