/**
 * Bilingual Response Templates (DE/EN)
 *
 * Token-optimized response formatting for UQAS answers.
 * Uses compact formats while maintaining readability in both languages.
 */

import type { SupportedLanguage } from './index.js';

/**
 * Source information
 */
export interface SourceInfo {
  type: string;
  name: string;
  webUrl?: string;
  downloadUrl?: string;
  relevance: number;
}

/**
 * Document link information
 */
export interface DocumentLink {
  fileName: string;
  webUrl: string;
  downloadUrl?: string;
  type: string;
}

/**
 * Response data structure
 */
export interface ResponseData {
  summary: string;
  confidence: number;
  sourceCount: number;
  depth: number;
  facts: string[];
  timeline?: TimelineEntry[];
  recommendations?: string[];
  entities?: Map<string, unknown>;
  tokenBudget?: number;
  sources?: SourceInfo[];
  documentLinks?: DocumentLink[];
  importantDocuments?: SourceInfo[];
}

/**
 * Timeline entry
 */
export interface TimelineEntry {
  date: string;
  time?: string;
  title: string;
  type: 'email' | 'meeting' | 'file' | 'task' | 'other';
}

/**
 * Response template strings for both languages
 */
export interface ResponseTemplates {
  answer_prefix: string;
  confidence: string;
  sources: string;
  depth: string;
  key_facts: string;
  timeline: string;
  recommendations: string;
  details: string;
  no_results: string;
  partial_results: string;
  items: string;
  // Time formatting
  time_format: string;
  time_ago: string;
  units: {
    minutes: string;
    hours: string;
    days: string;
    weeks: string;
    months: string;
  };
  // Days of week (short)
  days: {
    mon: string;
    tue: string;
    wed: string;
    thu: string;
    fri: string;
    sat: string;
    sun: string;
  };
  // Entity type labels
  entityLabels: {
    person: string;
    email: string;
    meeting: string;
    file: string;
    task: string;
  };
  // Source and document sections
  important_documents: string;
  document_links: string;
  all_sources: string;
}

/**
 * Bilingual response templates
 */
export const RESPONSE_TEMPLATES: Record<SupportedLanguage, ResponseTemplates> = {
  de: {
    answer_prefix: '## 📋 Antwort\n\n',
    confidence: 'Konfidenz',
    sources: 'Quellen',
    depth: 'Tiefe',
    key_facts: '### 🔑 Wichtige Fakten',
    timeline: '### ⏱️ Zeitverlauf',
    recommendations: '### 💡 Empfehlungen',
    details: 'Details anzeigen',
    no_results: 'Keine Ergebnisse gefunden für',
    partial_results: 'Teilweise Ergebnisse gefunden',
    items: 'Einträge',
    time_format: 'DD.MM.YYYY HH:mm',
    time_ago: 'vor {n} {unit}',
    units: {
      minutes: 'Minuten',
      hours: 'Stunden',
      days: 'Tagen',
      weeks: 'Wochen',
      months: 'Monaten',
    },
    days: {
      mon: 'Mo',
      tue: 'Di',
      wed: 'Mi',
      thu: 'Do',
      fri: 'Fr',
      sat: 'Sa',
      sun: 'So',
    },
    entityLabels: {
      person: 'Person',
      email: 'E-Mail',
      meeting: 'Termin',
      file: 'Datei',
      task: 'Aufgabe',
    },
    important_documents: '### 📄 Wichtige Dokumente',
    document_links: '### 🔗 Dokumenten-Links',
    all_sources: '### 📚 Alle Quellen',
  },
  en: {
    answer_prefix: '## 📋 Answer\n\n',
    confidence: 'Confidence',
    sources: 'Sources',
    depth: 'Depth',
    key_facts: '### 🔑 Key Facts',
    timeline: '### ⏱️ Timeline',
    recommendations: '### 💡 Recommendations',
    details: 'Show details',
    no_results: 'No results found for',
    partial_results: 'Partial results found',
    items: 'items',
    time_format: 'MM/DD/YYYY h:mm A',
    time_ago: '{n} {unit} ago',
    units: {
      minutes: 'minutes',
      hours: 'hours',
      days: 'days',
      weeks: 'weeks',
      months: 'months',
    },
    days: {
      mon: 'Mon',
      tue: 'Tue',
      wed: 'Wed',
      thu: 'Thu',
      fri: 'Fri',
      sat: 'Sat',
      sun: 'Sun',
    },
    entityLabels: {
      person: 'Person',
      email: 'Email',
      meeting: 'Meeting',
      file: 'File',
      task: 'Task',
    },
    important_documents: '### 📄 Important Documents',
    document_links: '### 🔗 Document Links',
    all_sources: '### 📚 All Sources',
  },
};

/**
 * Universal abbreviations (work in both languages)
 */
export const UNIVERSAL_ABBREV: Record<string, string> = {
  // Time
  mtg: 'meeting',
  msg: 'message',
  doc: 'document',
  att: 'attachment',
  min: 'minute',
  hr: 'hour',
  // Icons for compact display
  '📧': 'email',
  '📅': 'calendar',
  '📁': 'file',
  '✅': 'task',
  '👤': 'person',
  '💬': 'chat',
};

/**
 * Compact labels for minimum token usage
 */
export const COMPACT_LABELS: Record<SupportedLanguage, Record<string, string>> = {
  de: {
    from: 'Von',
    to: 'An',
    subject: 'Betr',
    date: 'Datum',
    loc: 'Ort',
    dur: 'Dauer',
    att: 'Anh',
  },
  en: {
    from: 'From',
    to: 'To',
    subject: 'Subj',
    date: 'Date',
    loc: 'Loc',
    dur: 'Dur',
    att: 'Att',
  },
};

/**
 * BilingualResponseBuilder - Builds token-optimized responses
 */
export class BilingualResponseBuilder {
  private lang: SupportedLanguage;
  private templates: ResponseTemplates;
  private compact: Record<string, string>;

  constructor(lang: SupportedLanguage) {
    this.lang = lang;
    this.templates = RESPONSE_TEMPLATES[lang];
    this.compact = COMPACT_LABELS[lang];
  }

  /**
   * Build full response from data
   */
  buildResponse(data: ResponseData): string {
    const parts: string[] = [];

    // Header with answer
    parts.push(this.templates.answer_prefix + data.summary);

    // Confidence bar
    parts.push(this.buildConfidenceLine(data));
    parts.push('\n---\n');

    // Key facts
    if (data.facts.length > 0) {
      parts.push(this.templates.key_facts);
      parts.push(data.facts.map((f) => `• ${f}`).join('\n'));
      parts.push('');
    }

    // Timeline (compact)
    if (data.timeline && data.timeline.length > 0) {
      parts.push(this.templates.timeline);
      parts.push(this.formatTimeline(data.timeline));
      parts.push('');
    }

    // Important documents
    if (data.importantDocuments && data.importantDocuments.length > 0) {
      parts.push(this.templates.important_documents);
      for (const doc of data.importantDocuments.slice(0, 10)) {
        const link = doc.webUrl ? `[${doc.name}](${doc.webUrl})` : doc.name;
        const download = doc.downloadUrl ? ` [📥 Download](${doc.downloadUrl})` : '';
        parts.push(`• ${link}${download} (${doc.type})`);
      }
      parts.push('');
    }

    // Document links
    if (data.documentLinks && data.documentLinks.length > 0) {
      parts.push(this.templates.document_links);
      for (const link of data.documentLinks.slice(0, 20)) {
        const download = link.downloadUrl ? ` [📥 Download](${link.downloadUrl})` : '';
        parts.push(`• [${link.fileName}](${link.webUrl})${download}`);
      }
      if (data.documentLinks.length > 20) {
        const remaining = this.lang === 'de' ? 'weitere Links' : 'more links';
        parts.push(
          `  ... ${this.lang === 'de' ? 'und' : 'and'} ${data.documentLinks.length - 20} ${remaining}`
        );
      }
      parts.push('');
    }

    // All sources
    if (data.sources && data.sources.length > 0) {
      parts.push(this.templates.all_sources);
      for (const source of data.sources.slice(0, 15)) {
        if (source.webUrl) {
          parts.push(`• [${source.name}](${source.webUrl}) (${source.type})`);
        } else {
          parts.push(`• ${source.name} (${source.type})`);
        }
      }
      if (data.sources.length > 15) {
        const remaining = this.lang === 'de' ? 'weitere Quellen' : 'more sources';
        parts.push(
          `  ... ${this.lang === 'de' ? 'und' : 'and'} ${data.sources.length - 15} ${remaining}`
        );
      }
      parts.push('');
    }

    // Recommendations
    if (data.recommendations && data.recommendations.length > 0) {
      parts.push(this.templates.recommendations);
      parts.push(data.recommendations.map((r, i) => `${i + 1}. ${r}`).join('\n'));
    }

    return parts.join('\n');
  }

  /**
   * Build compact response (for token-limited situations)
   */
  buildCompactResponse(data: ResponseData): string {
    const parts: string[] = [];

    // One-line summary
    parts.push(`📋 ${data.summary}`);

    // Compact metrics
    const confidenceBar =
      '█'.repeat(Math.round(data.confidence * 5)) + '░'.repeat(5 - Math.round(data.confidence * 5));
    parts.push(
      `[${confidenceBar}] ${Math.round(data.confidence * 100)}% | ${data.sourceCount} ${this.templates.sources} | L${data.depth}`
    );

    // Top 3 facts only
    if (data.facts.length > 0) {
      parts.push('');
      data.facts.slice(0, 3).forEach((f) => parts.push(`• ${f}`));
      if (data.facts.length > 3) {
        parts.push(`  (+${data.facts.length - 3} ${this.templates.items})`);
      }
    }

    return parts.join('\n');
  }

  /**
   * Build confidence line
   */
  private buildConfidenceLine(data: ResponseData): string {
    const confidenceBar =
      '█'.repeat(Math.round(data.confidence * 10)) +
      '░'.repeat(10 - Math.round(data.confidence * 10));
    return (
      `**${this.templates.confidence}**: ${confidenceBar} ${Math.round(data.confidence * 100)}% | ` +
      `**${this.templates.sources}**: ${data.sourceCount} | ` +
      `**${this.templates.depth}**: L${data.depth}`
    );
  }

  /**
   * Format timeline entries
   */
  formatTimeline(entries: TimelineEntry[]): string {
    return entries
      .map((entry) => {
        const icon = this.getTypeIcon(entry.type);
        const dateStr = this.formatDateCompact(entry.date);
        const time = entry.time ? ` ${entry.time}` : '';
        return `\`${dateStr}\`${time} ${icon} ${entry.title}`;
      })
      .join('\n');
  }

  /**
   * Get icon for entry type
   */
  private getTypeIcon(type: TimelineEntry['type']): string {
    const icons: Record<TimelineEntry['type'], string> = {
      email: '📧',
      meeting: '📅',
      file: '📁',
      task: '✅',
      other: '📌',
    };
    return icons[type];
  }

  /**
   * Format date in compact form
   */
  formatDateCompact(dateStr: string): string {
    const date = new Date(dateStr);
    const now = new Date();
    const diffDays = Math.floor((date.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));

    // Today/Yesterday/Tomorrow
    if (diffDays === 0) return this.lang === 'de' ? 'Heute' : 'Today';
    if (diffDays === -1) return this.lang === 'de' ? 'Gestern' : 'Yesterday';
    if (diffDays === 1) return this.lang === 'de' ? 'Morgen' : 'Tomorrow';

    // Day of week + date
    const dayNames = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'] as const;
    const dayAbbr = this.templates.days[dayNames[date.getDay()]];

    if (this.lang === 'de') {
      return `${dayAbbr} ${date.getDate().toString().padStart(2, '0')}.${(date.getMonth() + 1).toString().padStart(2, '0')}`;
    } else {
      return `${dayAbbr} ${(date.getMonth() + 1).toString().padStart(2, '0')}/${date.getDate().toString().padStart(2, '0')}`;
    }
  }

  /**
   * Format time ago
   */
  formatTimeAgo(date: Date): string {
    const now = new Date();
    const diffMs = now.getTime() - date.getTime();
    const diffMinutes = Math.floor(diffMs / (1000 * 60));
    const diffHours = Math.floor(diffMs / (1000 * 60 * 60));
    const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
    const diffWeeks = Math.floor(diffDays / 7);

    let n: number;
    let unit: keyof ResponseTemplates['units'];

    if (diffMinutes < 60) {
      n = diffMinutes;
      unit = 'minutes';
    } else if (diffHours < 24) {
      n = diffHours;
      unit = 'hours';
    } else if (diffDays < 7) {
      n = diffDays;
      unit = 'days';
    } else {
      n = diffWeeks;
      unit = 'weeks';
    }

    return this.templates.time_ago
      .replace('{n}', n.toString())
      .replace('{unit}', this.templates.units[unit]);
  }

  /**
   * Build no results message
   */
  buildNoResults(query: string): string {
    return `❌ ${this.templates.no_results} "${query}"`;
  }

  /**
   * Build partial results message
   */
  buildPartialResults(found: number, total: number): string {
    return `⚠️ ${this.templates.partial_results}: ${found}/${total} ${this.templates.items}`;
  }

  /**
   * Get current language
   */
  getLanguage(): SupportedLanguage {
    return this.lang;
  }

  /**
   * Get compact label
   */
  getCompactLabel(key: string): string {
    return this.compact[key] || key;
  }
}

export default BilingualResponseBuilder;
