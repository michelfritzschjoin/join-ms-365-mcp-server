/**
 * Bilingual Thesaurus (DE/EN)
 *
 * Cross-language synonym mappings optimized for Microsoft 365 search.
 * Enables finding content regardless of whether it was created in German or English.
 */

import type { SupportedLanguage } from './index.js';

/**
 * Thesaurus entry with DE and EN variants
 */
export interface ThesaurusEntry {
  de: string[];
  en: string[];
}

/**
 * Comprehensive DE<->EN thesaurus for M365 content
 */
export const BILINGUAL_THESAURUS: Record<string, ThesaurusEntry> = {
  // ============================================
  // CALENDAR & MEETINGS
  // ============================================
  meeting: {
    de: ['meeting', 'besprechung', 'termin', 'sitzung', 'konferenz', 'treffen', 'abstimmung'],
    en: ['meeting', 'appointment', 'session', 'conference', 'call', 'sync', 'standup', 'huddle'],
  },
  besprechung: {
    de: ['besprechung', 'meeting', 'termin', 'sitzung', 'abstimmung'],
    en: ['meeting', 'discussion', 'session', 'briefing'],
  },
  termin: {
    de: ['termin', 'meeting', 'besprechung', 'verabredung', 'ereignis'],
    en: ['appointment', 'meeting', 'event', 'date', 'schedule'],
  },
  calendar: {
    de: ['kalender', 'terminkalender', 'zeitplan'],
    en: ['calendar', 'schedule', 'agenda', 'planner'],
  },
  reminder: {
    de: ['erinnerung', 'mahnung', 'hinweis'],
    en: ['reminder', 'notification', 'alert'],
  },

  // ============================================
  // EMAIL & MESSAGES
  // ============================================
  email: {
    de: ['e-mail', 'email', 'mail', 'nachricht', 'mitteilung', 'schreiben'],
    en: ['email', 'mail', 'message', 'correspondence'],
  },
  message: {
    de: ['nachricht', 'mitteilung', 'meldung', 'botschaft'],
    en: ['message', 'note', 'communication', 'memo'],
  },
  inbox: {
    de: ['posteingang', 'eingang', 'inbox'],
    en: ['inbox', 'mailbox', 'incoming'],
  },
  draft: {
    de: ['entwurf', 'draft', 'vorlage'],
    en: ['draft', 'template', 'outline'],
  },
  attachment: {
    de: ['anhang', 'anlage', 'beilage', 'attachment'],
    en: ['attachment', 'file', 'enclosure'],
  },
  reply: {
    de: ['antwort', 'erwiderung', 'rückmeldung'],
    en: ['reply', 'response', 'answer'],
  },
  forward: {
    de: ['weiterleitung', 'weiterleiten'],
    en: ['forward', 'forwarding'],
  },

  // ============================================
  // FILES & DOCUMENTS
  // ============================================
  document: {
    de: ['dokument', 'datei', 'unterlage', 'akte', 'schriftstück'],
    en: ['document', 'file', 'record', 'paper'],
  },
  file: {
    de: ['datei', 'dokument', 'file', 'anlage'],
    en: ['file', 'document', 'attachment'],
  },
  folder: {
    de: ['ordner', 'verzeichnis', 'mappe', 'ablage'],
    en: ['folder', 'directory', 'binder'],
  },
  report: {
    de: ['bericht', 'report', 'auswertung', 'analyse', 'zusammenfassung'],
    en: ['report', 'analysis', 'summary', 'review'],
  },
  presentation: {
    de: ['präsentation', 'vortrag', 'vorstellung', 'folie', 'powerpoint'],
    en: ['presentation', 'slides', 'deck', 'powerpoint', 'ppt'],
  },
  spreadsheet: {
    de: ['tabelle', 'excel', 'kalkulation', 'arbeitsblatt'],
    en: ['spreadsheet', 'excel', 'workbook', 'sheet'],
  },
  contract: {
    de: ['vertrag', 'kontrakt', 'vereinbarung', 'abkommen'],
    en: ['contract', 'agreement', 'deal'],
  },
  proposal: {
    de: ['angebot', 'vorschlag', 'offerte', 'proposal'],
    en: ['proposal', 'offer', 'quote', 'bid'],
  },
  invoice: {
    de: ['rechnung', 'faktura', 'invoice'],
    en: ['invoice', 'bill', 'receipt'],
  },

  // ============================================
  // PEOPLE & TEAMS
  // ============================================
  colleague: {
    de: ['kollege', 'kollegin', 'mitarbeiter', 'mitarbeiterin', 'teammitglied'],
    en: ['colleague', 'coworker', 'team member', 'peer'],
  },
  manager: {
    de: ['manager', 'vorgesetzter', 'chef', 'chefin', 'leiter', 'leiterin', 'führungskraft'],
    en: ['manager', 'boss', 'supervisor', 'lead', 'head'],
  },
  team: {
    de: ['team', 'gruppe', 'abteilung', 'bereich', 'mannschaft'],
    en: ['team', 'group', 'department', 'unit', 'squad'],
  },
  customer: {
    de: ['kunde', 'kundin', 'auftraggeber', 'klient', 'mandant'],
    en: ['customer', 'client', 'buyer', 'account'],
  },
  contact: {
    de: ['kontakt', 'ansprechpartner', 'ansprechpartnerin'],
    en: ['contact', 'person', 'point of contact', 'poc'],
  },

  // ============================================
  // TASKS & PROJECTS
  // ============================================
  task: {
    de: ['aufgabe', 'task', 'todo', 'arbeit', 'tätigkeit', 'aktion'],
    en: ['task', 'todo', 'assignment', 'work item', 'action item'],
  },
  project: {
    de: ['projekt', 'vorhaben', 'initiative'],
    en: ['project', 'initiative', 'program'],
  },
  deadline: {
    de: ['deadline', 'frist', 'termin', 'fälligkeitsdatum', 'stichtag'],
    en: ['deadline', 'due date', 'target date'],
  },
  priority: {
    de: ['priorität', 'wichtigkeit', 'dringlichkeit'],
    en: ['priority', 'importance', 'urgency'],
  },
  status: {
    de: ['status', 'stand', 'zustand', 'fortschritt'],
    en: ['status', 'state', 'progress', 'update'],
  },

  // ============================================
  // TIME EXPRESSIONS
  // ============================================
  today: {
    de: ['heute', 'heutigen', 'aktuell'],
    en: ['today', 'current', 'this day'],
  },
  tomorrow: {
    de: ['morgen', 'morgigen'],
    en: ['tomorrow', 'next day'],
  },
  yesterday: {
    de: ['gestern', 'gestrigen'],
    en: ['yesterday', 'previous day'],
  },
  week: {
    de: ['woche', 'wöchentlich', 'kalenderwoche'],
    en: ['week', 'weekly'],
  },
  month: {
    de: ['monat', 'monatlich'],
    en: ['month', 'monthly'],
  },
  year: {
    de: ['jahr', 'jährlich', 'jahres'],
    en: ['year', 'yearly', 'annual'],
  },
  next: {
    de: ['nächste', 'nächsten', 'nächster', 'nächstes', 'kommende', 'kommenden'],
    en: ['next', 'upcoming', 'coming', 'following'],
  },
  last: {
    de: ['letzte', 'letzten', 'letzter', 'letztes', 'vergangene', 'vergangenen', 'vorige'],
    en: ['last', 'previous', 'past', 'prior'],
  },
  this: {
    de: ['diese', 'dieser', 'dieses', 'diesen', 'diesem', 'aktuelle'],
    en: ['this', 'current'],
  },
  recent: {
    de: ['kürzlich', 'neulich', 'vor kurzem', 'aktuelle'],
    en: ['recent', 'recently', 'lately', 'latest'],
  },

  // ============================================
  // ACTIONS
  // ============================================
  send: {
    de: ['senden', 'schicken', 'verschicken', 'abschicken'],
    en: ['send', 'submit', 'dispatch'],
  },
  receive: {
    de: ['empfangen', 'erhalten', 'bekommen'],
    en: ['receive', 'get', 'obtain'],
  },
  create: {
    de: ['erstellen', 'anlegen', 'erzeugen', 'neu'],
    en: ['create', 'make', 'new', 'add'],
  },
  update: {
    de: ['aktualisieren', 'ändern', 'bearbeiten', 'updaten'],
    en: ['update', 'edit', 'modify', 'change'],
  },
  delete: {
    de: ['löschen', 'entfernen', 'delete'],
    en: ['delete', 'remove', 'discard'],
  },
  share: {
    de: ['teilen', 'freigeben', 'sharen'],
    en: ['share', 'distribute'],
  },
  download: {
    de: ['herunterladen', 'download', 'runterladen'],
    en: ['download', 'get', 'fetch'],
  },
  upload: {
    de: ['hochladen', 'upload', 'raufladen'],
    en: ['upload', 'submit'],
  },

  // ============================================
  // QUESTIONS & SEARCH
  // ============================================
  find: {
    de: ['finden', 'suchen', 'lokalisieren'],
    en: ['find', 'search', 'locate', 'look for'],
  },
  show: {
    de: ['zeigen', 'anzeigen', 'darstellen', 'auflisten'],
    en: ['show', 'display', 'list', 'present'],
  },
  all: {
    de: ['alle', 'alles', 'sämtliche', 'gesamt'],
    en: ['all', 'every', 'each', 'entire'],
  },
  summary: {
    de: ['zusammenfassung', 'überblick', 'übersicht', 'kurzfassung'],
    en: ['summary', 'overview', 'synopsis', 'recap'],
  },

  // ============================================
  // COMMON BUSINESS TERMS
  // ============================================
  budget: {
    de: ['budget', 'etat', 'haushalt', 'finanzplan'],
    en: ['budget', 'funds', 'allocation'],
  },
  review: {
    de: ['überprüfung', 'review', 'prüfung', 'bewertung'],
    en: ['review', 'evaluation', 'assessment', 'check'],
  },
  approval: {
    de: ['genehmigung', 'freigabe', 'zustimmung', 'bewilligung'],
    en: ['approval', 'authorization', 'sign-off'],
  },
  request: {
    de: ['anfrage', 'anforderung', 'bitte', 'ersuchen'],
    en: ['request', 'inquiry', 'ask'],
  },
  feedback: {
    de: ['feedback', 'rückmeldung', 'kommentar', 'stellungnahme'],
    en: ['feedback', 'comment', 'input', 'response'],
  },
};

/**
 * BilingualThesaurus - Cross-language synonym lookup
 */
export class BilingualThesaurus {
  private thesaurus: Record<string, ThesaurusEntry>;
  private reverseIndex: Map<string, string[]>;

  constructor(customEntries?: Record<string, ThesaurusEntry>) {
    this.thesaurus = { ...BILINGUAL_THESAURUS, ...customEntries };
    this.reverseIndex = this.buildReverseIndex();
  }

  /**
   * Build reverse index for fast lookups
   */
  private buildReverseIndex(): Map<string, string[]> {
    const index = new Map<string, string[]>();

    for (const [key, entry] of Object.entries(this.thesaurus)) {
      // Index all DE variants
      for (const word of entry.de) {
        const lower = word.toLowerCase();
        if (!index.has(lower)) {
          index.set(lower, []);
        }
        index.get(lower)!.push(key);
      }
      // Index all EN variants
      for (const word of entry.en) {
        const lower = word.toLowerCase();
        if (!index.has(lower)) {
          index.set(lower, []);
        }
        index.get(lower)!.push(key);
      }
    }

    return index;
  }

  /**
   * Get all variants for a word (both DE and EN)
   */
  getAllVariants(word: string): string[] {
    const lower = word.toLowerCase();
    const keys = this.reverseIndex.get(lower);
    if (!keys) return [];

    const variants = new Set<string>();
    for (const key of keys) {
      const entry = this.thesaurus[key];
      if (entry) {
        entry.de.forEach((v) => variants.add(v));
        entry.en.forEach((v) => variants.add(v));
      }
    }

    // Remove the original word
    variants.delete(lower);
    return Array.from(variants);
  }

  /**
   * Get variants in a specific language
   */
  getVariantsInLanguage(word: string, targetLang: SupportedLanguage): string[] {
    const lower = word.toLowerCase();
    const keys = this.reverseIndex.get(lower);
    if (!keys) return [];

    const variants = new Set<string>();
    for (const key of keys) {
      const entry = this.thesaurus[key];
      if (entry) {
        entry[targetLang].forEach((v) => variants.add(v));
      }
    }

    variants.delete(lower);
    return Array.from(variants);
  }

  /**
   * Get cross-language variants (translate to other language)
   */
  getCrossLanguageVariants(word: string, sourceLang: SupportedLanguage): string[] {
    const targetLang = sourceLang === 'de' ? 'en' : 'de';
    return this.getVariantsInLanguage(word, targetLang);
  }

  /**
   * Check if a word has known synonyms
   */
  hasEntry(word: string): boolean {
    return this.reverseIndex.has(word.toLowerCase());
  }

  /**
   * Get entry by key
   */
  getEntry(key: string): ThesaurusEntry | undefined {
    return this.thesaurus[key.toLowerCase()];
  }

  /**
   * Add a custom entry
   */
  addEntry(key: string, entry: ThesaurusEntry): void {
    this.thesaurus[key.toLowerCase()] = entry;
    // Rebuild reverse index
    this.reverseIndex = this.buildReverseIndex();
  }

  /**
   * Get all entry keys
   */
  getKeys(): string[] {
    return Object.keys(this.thesaurus);
  }

  /**
   * Get total number of entries
   */
  get size(): number {
    return Object.keys(this.thesaurus).length;
  }
}

export default BilingualThesaurus;
