/**
 * Universal Business Content Extractor
 *
 * Extracts structured data from business documents and communications.
 * Supports multiple business domains: Project Management, Customer Management,
 * Meeting Management, Document Management, Sales & Business Development, HR & Personal.
 */

import logger from '../logger.js';

// ============================================================================
// TYPE DEFINITIONS
// ============================================================================

/**
 * Document types that can be detected and extracted
 */
export type DocumentType =
  | 'roadmap'
  | 'project_plan'
  | 'status_update'
  | 'customer_info'
  | 'contract'
  | 'proposal'
  | 'meeting_agenda'
  | 'meeting_notes'
  | 'action_items'
  | 'invoice'
  | 'report'
  | 'executive_summary'
  | 'offer'
  | 'budget'
  | 'forecast'
  | 'onboarding'
  | 'review'
  | 'application'
  | 'unknown';

/**
 * Extractor configuration for registry
 */
export interface ExtractorConfig {
  type: DocumentType;
  extractor: new () => BaseExtractor;
  patterns: RegExp[];
  keywords: string[];
  priority: number;
}

/**
 * Options for extraction
 */
export interface ExtractorOptions {
  strictMode?: boolean; // Only extract with high confidence
  includeMetadata?: boolean;
  includeEntities?: boolean;
  includeSummary?: boolean;
  language?: 'de' | 'en' | 'auto';
  customPatterns?: RegExp[];
}

/**
 * Base extraction result
 */
export interface ExtractedContent {
  type: DocumentType;
  confidence: number;
  content: unknown;
  metadata?: Record<string, unknown>;
}

// ============================================================================
// BUSINESS CONTENT TYPE DEFINITIONS
// ============================================================================

/**
 * Person entity
 */
export interface Person {
  name: string;
  email?: string;
  role?: string;
}

/**
 * Organization entity
 */
export interface Organization {
  name: string;
  type?: string;
}

/**
 * Amount with currency
 */
export interface Amount {
  value: number;
  currency: string;
  context?: string;
}

/**
 * Resource entity
 */
export interface Resource {
  name: string;
  type: 'person' | 'team' | 'tool' | 'other';
  id?: string;
}

/**
 * Action Item
 */
export interface ActionItem {
  text: string;
  responsible?: Person;
  deadline?: Date;
  priority?: 'high' | 'medium' | 'low';
  status?: 'pending' | 'in-progress' | 'completed';
}

/**
 * Decision
 */
export interface Decision {
  text: string;
  context?: string;
  decisionMaker?: Person;
  date?: Date;
  rationale?: string;
}

/**
 * Milestone
 */
export interface Milestone {
  name: string;
  date?: Date;
  status?: 'planned' | 'in-progress' | 'completed' | 'blocked';
  responsible?: Person;
  description?: string;
}

/**
 * Timeline Entry
 */
export interface TimelineEntry {
  date: Date;
  event: string;
  type: 'milestone' | 'deadline' | 'meeting' | 'delivery' | 'other';
  description?: string;
}

/**
 * Table structure
 */
export interface Table {
  headers: string[];
  rows: string[][];
  caption?: string;
}

/**
 * List structure
 */
export interface List {
  items: string[];
  type: 'ordered' | 'unordered';
  title?: string;
}

/**
 * Project Content
 */
export interface ProjectContent {
  projectCode?: string;
  projectName?: string;
  milestones?: Milestone[];
  timeline?: TimelineEntry[];
  tasks?: Array<{
    id?: string;
    name: string;
    status?: string;
    responsible?: Person;
    deadline?: Date;
  }>;
  resources?: Resource[];
  dependencies?: Array<{
    from: string;
    to: string;
    type: string;
  }>;
  progress?: number;
  blockers?: string[];
  risks?: string[];
}

/**
 * Customer Content
 */
export interface CustomerContent {
  customerName?: string;
  customerId?: string;
  contactPerson?: Person;
  company?: Organization;
  contracts?: Array<{
    number?: string;
    startDate?: Date;
    endDate?: Date;
    conditions?: string;
    cancellationPeriod?: string;
  }>;
  offers?: Array<{
    number?: string;
    price?: Amount;
    validity?: Date;
    discount?: number;
  }>;
  touchpoints?: Array<{
    date: Date;
    type: string;
    description: string;
  }>;
}

/**
 * Meeting Content
 */
export interface MeetingContent {
  agenda?: Array<{
    point: string;
    time?: string;
    presenter?: Person;
  }>;
  actionItems?: ActionItem[];
  decisions?: Decision[];
  participants?: Person[];
  notes?: string[];
  keyPoints?: string[];
  followUps?: string[];
}

/**
 * Document Content
 */
export interface DocumentContent {
  documentId?: string;
  documentType?: 'contract' | 'invoice' | 'report' | 'proposal' | 'other';
  parties?: Organization[];
  invoiceNumber?: string;
  amount?: Amount;
  dueDate?: Date;
  reportSummary?: string;
  kpis?: Array<{
    name: string;
    value: string | number;
  }>;
  recommendations?: string[];
}

/**
 * Sales Content
 */
export interface SalesContent {
  offerNumber?: string;
  offerPrice?: Amount;
  discount?: number;
  validity?: Date;
  budgets?: Array<{
    id?: string;
    amount: Amount;
    costCenter?: string;
    limit?: Amount;
  }>;
  forecasts?: Array<{
    id?: string;
    revenue: Amount;
    probability?: number;
    stage?: string;
  }>;
  deals?: Array<{
    name: string;
    stage?: string;
    value?: Amount;
    probability?: number;
  }>;
}

/**
 * HR Content
 */
export interface HRContent {
  onboarding?: {
    checklist?: string[];
    dates?: Date[];
    responsible?: Person[];
    documents?: string[];
  };
  reviews?: Array<{
    id?: string;
    type: 'performance' | 'goal' | 'feedback';
    goals?: string[];
    feedback?: string;
    rating?: number;
  }>;
  applications?: Array<{
    candidateId?: string;
    candidateName?: string;
    status?: string;
    interviewDate?: Date;
    feedback?: string;
  }>;
  trainings?: Array<{
    id?: string;
    name: string;
    certificate?: string;
    deadline?: Date;
    participants?: Person[];
  }>;
}

/**
 * Complete Business Content Extraction Result
 */
export interface BusinessContentExtraction {
  detectedType: DocumentType;
  confidence: number;
  extracted: {
    project?: ProjectContent;
    customer?: CustomerContent;
    meeting?: MeetingContent;
    document?: DocumentContent;
    sales?: SalesContent;
    hr?: HRContent;
  };
  metadata: {
    priorities?: string[];
    statuses?: string[];
    deadlines?: Date[];
    tags?: string[];
  };
  entities: {
    // Basis-Entitäten
    people?: Person[];
    organizations?: Organization[];
    dates?: Date[];
    amounts?: Amount[];
    // Projekt-Entitäten
    projectCodes?: string[];
    taskIds?: string[];
    resources?: Resource[];
    // Customer-Entitäten
    customerNames?: string[];
    contractNumbers?: string[];
    offerNumbers?: string[];
    customerIds?: string[];
    // Meeting-Entitäten
    participants?: Person[];
    actionItemIds?: string[];
    decisionIds?: string[];
    agendaPoints?: string[];
    // Dokument-Entitäten
    invoiceNumbers?: string[];
    documentIds?: string[];
    reportNumbers?: string[];
    // Sales-Entitäten
    dealNames?: string[];
    budgetIds?: string[];
    forecastIds?: string[];
    // HR-Entitäten
    candidateIds?: string[];
    reviewIds?: string[];
    trainingIds?: string[];
    onboardingIds?: string[];
  };
  summary?: {
    actionItems?: ActionItem[];
    decisions?: Decision[];
    keyPoints?: string[];
  };
}

/**
 * Base class for all extractors
 */
export abstract class BaseExtractor {
  /**
   * Extract content from text
   */
  abstract extract(content: string, options?: ExtractorOptions): ExtractedContent;

  /**
   * Get patterns used by this extractor
   */
  abstract getPatterns(): RegExp[];

  /**
   * Get keywords used by this extractor
   */
  abstract getKeywords(): string[];

  /**
   * Calculate confidence score for extraction
   */
  protected calculateConfidence(content: string, matches: number, totalPatterns: number): number {
    if (totalPatterns === 0) return 0;
    const matchRatio = matches / totalPatterns;
    const keywordMatches = this.countKeywordMatches(content);
    const keywordScore = Math.min(keywordMatches / 3, 1); // Max 3 keywords = full score
    return matchRatio * 0.6 + keywordScore * 0.4;
  }

  /**
   * Count keyword matches in content
   */
  private countKeywordMatches(content: string): number {
    const keywords = this.getKeywords();
    const contentLower = content.toLowerCase();
    return keywords.filter((keyword) => contentLower.includes(keyword.toLowerCase())).length;
  }

  /**
   * Sanitize HTML content to plain text
   */
  protected sanitizeHtml(html: string): string {
    if (!html || typeof html !== 'string') {
      return '';
    }

    // Remove script, style, and other dangerous tags
    let result = html;
    const dangerousTags = ['script', 'style', 'iframe', 'object', 'form', 'textarea', 'select'];
    for (const tag of dangerousTags) {
      const regex = new RegExp(`<${tag}[^>]*>.*?</${tag}>`, 'gis');
      result = result.replace(regex, '');
    }

    // Remove all HTML tags
    result = result.replace(/<[^>]*>/g, ' ');

    // Decode HTML entities
    const entityMap: Record<string, string> = {
      '&nbsp;': ' ',
      '&amp;': '&',
      '&lt;': '<',
      '&gt;': '>',
      '&quot;': '"',
      '&#39;': "'",
      '&apos;': "'",
    };

    for (const [entity, replacement] of Object.entries(entityMap)) {
      result = result.replace(new RegExp(entity, 'gi'), replacement);
    }

    // Clean up whitespace
    result = result.replace(/\s+/g, ' ').trim();

    return result;
  }
}

// ============================================================================
// EXTRACTOR REGISTRY
// ============================================================================

/**
 * Registry for managing all extractors
 */
export class ExtractorRegistry {
  private extractors: Map<DocumentType, ExtractorConfig> = new Map();
  private extractorInstances: Map<DocumentType, BaseExtractor> = new Map();

  /**
   * Register an extractor
   */
  register(config: ExtractorConfig): void {
    this.extractors.set(config.type, config);
    logger.info(`Registered extractor for type: ${config.type}`);
  }

  /**
   * Detect document type from content
   */
  detectType(content: string): DocumentType {
    if (!content || typeof content !== 'string') {
      return 'unknown';
    }

    const contentLower = content.toLowerCase();
    let bestMatch: { type: DocumentType; score: number } | null = null;

    for (const [type, config] of this.extractors.entries()) {
      let score = 0;

      // Check keyword matches
      const keywordMatches = config.keywords.filter((keyword) =>
        contentLower.includes(keyword.toLowerCase())
      ).length;
      score += (keywordMatches / config.keywords.length) * 50;

      // Check pattern matches
      const patternMatches = config.patterns.filter((pattern) => pattern.test(content)).length;
      score += (patternMatches / config.patterns.length) * 50;

      // Apply priority multiplier
      score *= config.priority / 10;

      if (!bestMatch || score > bestMatch.score) {
        bestMatch = { type, score };
      }
    }

    if (bestMatch && bestMatch.score > 20) {
      return bestMatch.type;
    }

    return 'unknown';
  }

  /**
   * Get extractor instance for a document type
   */
  getExtractor(type: DocumentType): BaseExtractor | null {
    if (this.extractorInstances.has(type)) {
      return this.extractorInstances.get(type)!;
    }

    const config = this.extractors.get(type);
    if (!config) {
      return null;
    }

    const instance = new config.extractor();
    this.extractorInstances.set(type, instance);
    return instance;
  }

  /**
   * Get all registered types
   */
  getRegisteredTypes(): DocumentType[] {
    return Array.from(this.extractors.keys());
  }
}

// ============================================================================
// PATTERN-BASED EXTRACTOR ENGINE
// ============================================================================

/**
 * Pattern-based extractor engine that automatically detects and extracts content
 */
export class PatternBasedExtractor {
  private registry: ExtractorRegistry;

  constructor(registry: ExtractorRegistry) {
    this.registry = registry;
  }

  /**
   * Detect document type from content
   */
  detectDocumentType(content: string): DocumentType {
    return this.registry.detectType(content);
  }

  /**
   * Extract content based on detected or specified type
   */
  extract(content: string, type?: DocumentType, options?: ExtractorOptions): ExtractedContent {
    const detectedType = type || this.detectDocumentType(content);

    if (detectedType === 'unknown') {
      return {
        type: 'unknown',
        confidence: 0,
        content: null,
      };
    }

    const extractor = this.registry.getExtractor(detectedType);
    if (!extractor) {
      logger.warn(`No extractor found for type: ${detectedType}`);
      return {
        type: detectedType,
        confidence: 0,
        content: null,
      };
    }

    try {
      const result = extractor.extract(content, options);
      return result;
    } catch (error) {
      logger.error(`Error extracting content: ${error}`);
      return {
        type: detectedType,
        confidence: 0,
        content: null,
      };
    }
  }
}

// ============================================================================
// BUSINESS EXTRACTORS
// ============================================================================

/**
 * Project Management Extractor
 */
export class ProjectExtractor extends BaseExtractor {
  private tableParser = new TableListParser();
  private dateParser = new DateParser();
  private metadataExtractor = new MetadataExtractor();
  private projectEntityExtractor = new ProjectEntityExtractor();

  getPatterns(): RegExp[] {
    return [
      /\b(roadmap|meilenstein|timeline|milestone)\b/gi,
      /\b(projekt|project|task|aufgabe)\b/gi,
      /\b(status|fortschritt|progress|blocked)\b/gi,
      /\b([A-Z]{1,5}[-]?\d{2,6})\b/g, // Project codes
    ];
  }

  getKeywords(): string[] {
    return [
      'roadmap',
      'meilenstein',
      'timeline',
      'milestone',
      'projekt',
      'project',
      'task',
      'aufgabe',
      'status',
      'fortschritt',
      'progress',
    ];
  }

  extract(content: string, options?: ExtractorOptions): ExtractedContent {
    const text = this.sanitizeHtml(content);
    const contentLower = text.toLowerCase();

    const projectContent: ProjectContent = {};
    let matches = 0;
    const patterns = this.getPatterns();

    // Extract project codes
    const projectEntities = this.projectEntityExtractor.extract(text);
    if (projectEntities.projectCodes.length > 0) {
      projectContent.projectCode = projectEntities.projectCodes[0];
      matches++;
    }

    // Extract milestones from tables or structured text
    if (contentLower.includes('milestone') || contentLower.includes('meilenstein')) {
      const tables = this.tableParser.parseHtmlTable(content);
      const milestones: Milestone[] = [];

      for (const table of tables) {
        if (
          table.headers.some(
            (h) => h.toLowerCase().includes('milestone') || h.toLowerCase().includes('meilenstein')
          )
        ) {
          for (const row of table.rows) {
            if (row.length >= 2) {
              const milestone: Milestone = {
                name: row[0] || '',
              };

              // Try to extract date
              if (row[1]) {
                const date = this.dateParser.parse(row[1]);
                if (date) milestone.date = date;
              }

              // Try to extract status
              if (row.length >= 3 && row[2]) {
                const statuses = this.metadataExtractor.extractStatuses(row[2]);
                if (statuses.length > 0) {
                  milestone.status = statuses[0] as
                    | 'planned'
                    | 'in-progress'
                    | 'completed'
                    | 'blocked';
                }
              }

              // Try to extract responsible
              if (row.length >= 4 && row[3]) {
                const persons = this.projectEntityExtractor['extractPersons'](row[3]);
                if (persons.length > 0) {
                  milestone.responsible = persons[0];
                }
              }

              milestones.push(milestone);
            }
          }
        }
      }

      if (milestones.length > 0) {
        projectContent.milestones = milestones;
        matches++;
      }
    }

    // Extract tasks
    if (contentLower.includes('task') || contentLower.includes('aufgabe')) {
      const taskIds = projectEntities.taskIds;
      if (taskIds.length > 0) {
        projectContent.tasks = taskIds.map((id) => ({
          id,
          name: `Task ${id}`,
        }));
        matches++;
      }
    }

    // Extract progress percentage
    const progressMatch = text.match(/(?:progress|fortschritt)[:\s]+(\d+)%/i);
    if (progressMatch) {
      projectContent.progress = parseInt(progressMatch[1], 10);
      matches++;
    }

    // Extract blockers
    if (contentLower.includes('blocked') || contentLower.includes('blockiert')) {
      const blockerPattern = /(?:blocked|blockiert|blocker)[:\s]+(.+?)(?:\.|$)/gi;
      const blockers: string[] = [];
      let match;
      while ((match = blockerPattern.exec(text)) !== null) {
        blockers.push(match[1].trim());
      }
      if (blockers.length > 0) {
        projectContent.blockers = blockers;
        matches++;
      }
    }

    const confidence = this.calculateConfidence(text, matches, patterns.length);

    return {
      type: contentLower.includes('roadmap')
        ? 'roadmap'
        : contentLower.includes('status')
          ? 'status_update'
          : 'project_plan',
      confidence,
      content: projectContent,
    };
  }
}

/**
 * Customer Management Extractor
 */
export class CustomerExtractor extends BaseExtractor {
  private dateParser = new DateParser();
  private amountParser = new AmountParser();
  private customerEntityExtractor = new CustomerEntityExtractor();

  getPatterns(): RegExp[] {
    return [
      /\b(kunde|customer|client|mandant)\b/gi,
      /\b(vertrag|contract|agreement|vereinbarung)\b/gi,
      /\b(angebot|offer|quote|anfrage)\b/gi,
      /\b(CONTRACT|VERTRAG|VERT)[-:]?\d+\b/gi,
    ];
  }

  getKeywords(): string[] {
    return [
      'kunde',
      'customer',
      'client',
      'mandant',
      'vertrag',
      'contract',
      'agreement',
      'angebot',
      'offer',
      'quote',
    ];
  }

  extract(content: string, options?: ExtractorOptions): ExtractedContent {
    const text = this.sanitizeHtml(content);
    const contentLower = text.toLowerCase();

    const customerContent: CustomerContent = {};
    let matches = 0;
    const patterns = this.getPatterns();

    // Extract customer entities
    const entities = this.customerEntityExtractor.extract(text);
    if (entities.customerNames.length > 0) {
      customerContent.customerName = entities.customerNames[0];
      matches++;
    }
    if (entities.customerIds.length > 0) {
      customerContent.customerId = entities.customerIds[0];
      matches++;
    }

    // Extract contracts
    if (contentLower.includes('vertrag') || contentLower.includes('contract')) {
      const contracts: CustomerContent['contracts'] = [];
      if (entities.contractNumbers.length > 0) {
        for (const contractNum of entities.contractNumbers) {
          contracts.push({ number: contractNum });
        }
        customerContent.contracts = contracts;
        matches++;
      }
    }

    // Extract offers
    if (contentLower.includes('angebot') || contentLower.includes('offer')) {
      const offers: CustomerContent['offers'] = [];
      if (entities.offerNumbers.length > 0) {
        for (const offerNum of entities.offerNumbers) {
          const offer: { number?: string; price?: Amount; validity?: Date; discount?: number } = {
            number: offerNum,
          };
          const amounts = this.amountParser.parseAll(text);
          if (amounts.length > 0) {
            offer.price = amounts[0];
          }
          offers.push(offer);
        }
        customerContent.offers = offers;
        matches++;
      }
    }

    const confidence = this.calculateConfidence(text, matches, patterns.length);

    return {
      type:
        contentLower.includes('vertrag') || contentLower.includes('contract')
          ? 'contract'
          : 'customer_info',
      confidence,
      content: customerContent,
    };
  }
}

/**
 * Meeting Management Extractor
 */
export class MeetingExtractor extends BaseExtractor {
  private summaryGenerator = new SummaryGenerator();
  private meetingEntityExtractor = new MeetingEntityExtractor();
  private dateParser = new DateParser();

  getPatterns(): RegExp[] {
    return [
      /\b(agenda|tagesordnung|top|points)\b/gi,
      /\b(action\s+item|aufgabe|todo|follow-up)\b/gi,
      /\b(decision|entscheidung|beschluss|agreed)\b/gi,
      /\b(meeting|besprechung|termin)\b/gi,
    ];
  }

  getKeywords(): string[] {
    return [
      'agenda',
      'tagesordnung',
      'top',
      'action item',
      'aufgabe',
      'todo',
      'decision',
      'entscheidung',
      'meeting',
      'besprechung',
    ];
  }

  extract(content: string, options?: ExtractorOptions): ExtractedContent {
    const text = this.sanitizeHtml(content);
    const contentLower = text.toLowerCase();

    const meetingContent: MeetingContent = {};
    let matches = 0;
    const patterns = this.getPatterns();

    // Extract agenda
    if (contentLower.includes('agenda') || contentLower.includes('tagesordnung')) {
      const agenda: MeetingContent['agenda'] = [];
      const agendaPattern = /(?:top|agenda|punkt)[\s:]+(\d+)[\s:]+(.+?)(?:\n|$)/gi;
      let match;
      while ((match = agendaPattern.exec(text)) !== null) {
        agenda.push({
          point: match[2].trim(),
        });
      }
      if (agenda.length > 0) {
        meetingContent.agenda = agenda;
        matches++;
      }
    }

    // Extract action items
    const actionItems = this.summaryGenerator.generateActionItems(text);
    if (actionItems.length > 0) {
      meetingContent.actionItems = actionItems;
      matches++;
    }

    // Extract decisions
    const decisions = this.summaryGenerator.generateDecisions(text);
    if (decisions.length > 0) {
      meetingContent.decisions = decisions;
      matches++;
    }

    // Extract participants
    const entities = this.meetingEntityExtractor.extract(text);
    if (entities.participants.length > 0) {
      meetingContent.participants = entities.participants;
      matches++;
    }

    // Extract key points
    const keyPoints = this.summaryGenerator.generateKeyPoints(text);
    if (keyPoints.length > 0) {
      meetingContent.keyPoints = keyPoints;
      matches++;
    }

    const confidence = this.calculateConfidence(text, matches, patterns.length);

    return {
      type: contentLower.includes('agenda') ? 'meeting_agenda' : 'meeting_notes',
      confidence,
      content: meetingContent,
    };
  }
}

/**
 * Document Management Extractor
 */
export class DocumentExtractor extends BaseExtractor {
  private amountParser = new AmountParser();
  private dateParser = new DateParser();
  private documentEntityExtractor = new DocumentEntityExtractor();

  getPatterns(): RegExp[] {
    return [
      /\b(rechnung|invoice|bill|zahlung)\b/gi,
      /\b(bericht|report|summary|zusammenfassung)\b/gi,
      /\b(vertrag|contract|agreement)\b/gi,
      /\b(INV|INVOICE|RECHNUNG)[-:]?\d+\b/gi,
    ];
  }

  getKeywords(): string[] {
    return ['rechnung', 'invoice', 'bill', 'bericht', 'report', 'summary', 'vertrag', 'contract'];
  }

  extract(content: string, options?: ExtractorOptions): ExtractedContent {
    const text = this.sanitizeHtml(content);
    const contentLower = text.toLowerCase();

    const documentContent: DocumentContent = {};
    let matches = 0;
    const patterns = this.getPatterns();

    // Extract invoice information
    if (contentLower.includes('rechnung') || contentLower.includes('invoice')) {
      const entities = this.documentEntityExtractor.extract(text);
      if (entities.invoiceNumbers.length > 0) {
        documentContent.invoiceNumber = entities.invoiceNumbers[0];
        matches++;
      }

      const amounts = this.amountParser.parseAll(text);
      if (amounts.length > 0) {
        documentContent.amount = amounts[0];
        matches++;
      }

      // Extract due date
      const dueDatePattern = /(?:due\s+date|zahlungsfrist|fällig)[:\s]+(.+?)(?:\n|$)/gi;
      let match;
      while ((match = dueDatePattern.exec(text)) !== null) {
        const date = this.dateParser.parse(match[1]);
        if (date) {
          documentContent.dueDate = date;
          matches++;
          break;
        }
      }

      documentContent.documentType = 'invoice';
    }

    // Extract report information
    if (contentLower.includes('bericht') || contentLower.includes('report')) {
      const entities = this.documentEntityExtractor.extract(text);
      if (entities.reportNumbers.length > 0) {
        documentContent.documentId = entities.reportNumbers[0];
        matches++;
      }

      // Extract summary
      const summaryPattern =
        /(?:summary|zusammenfassung|executive\s+summary)[:\s]+(.+?)(?:\n\n|$)/gis;
      const summaryMatch = summaryPattern.exec(text);
      if (summaryMatch) {
        documentContent.reportSummary = summaryMatch[1].trim();
        matches++;
      }

      documentContent.documentType = 'report';
    }

    const confidence = this.calculateConfidence(text, matches, patterns.length);

    return {
      type:
        contentLower.includes('rechnung') || contentLower.includes('invoice')
          ? 'invoice'
          : contentLower.includes('bericht') || contentLower.includes('report')
            ? 'report'
            : 'unknown',
      confidence,
      content: documentContent,
    };
  }
}

/**
 * Sales & Business Development Extractor
 */
export class SalesExtractor extends BaseExtractor {
  private amountParser = new AmountParser();
  private dateParser = new DateParser();
  private salesEntityExtractor = new SalesEntityExtractor();

  getPatterns(): RegExp[] {
    return [
      /\b(angebot|offer|quote|preis)\b/gi,
      /\b(budget|kosten|cost|ausgabe)\b/gi,
      /\b(forecast|prognose|pipeline|umsatz)\b/gi,
      /\b(deal|opportunity|geschäft)\b/gi,
    ];
  }

  getKeywords(): string[] {
    return [
      'angebot',
      'offer',
      'quote',
      'budget',
      'kosten',
      'forecast',
      'prognose',
      'pipeline',
      'deal',
      'opportunity',
    ];
  }

  extract(content: string, options?: ExtractorOptions): ExtractedContent {
    const text = this.sanitizeHtml(content);
    const contentLower = text.toLowerCase();

    const salesContent: SalesContent = {};
    let matches = 0;
    const patterns = this.getPatterns();

    // Extract offers
    if (contentLower.includes('angebot') || contentLower.includes('offer')) {
      const entities = this.salesEntityExtractor.extract(text);
      const amounts = this.amountParser.parseAll(text);
      if (amounts.length > 0) {
        salesContent.offerPrice = amounts[0];
        matches++;
      }
    }

    // Extract budgets
    if (contentLower.includes('budget') || contentLower.includes('kosten')) {
      const entities = this.salesEntityExtractor.extract(text);
      if (entities.budgetIds.length > 0) {
        const budgets: SalesContent['budgets'] = [];
        const amounts = this.amountParser.parseAll(text);
        for (const budgetId of entities.budgetIds) {
          budgets.push({
            id: budgetId,
            amount: amounts[0] || { value: 0, currency: 'EUR' },
          });
        }
        salesContent.budgets = budgets;
        matches++;
      }
    }

    // Extract forecasts
    if (contentLower.includes('forecast') || contentLower.includes('prognose')) {
      const entities = this.salesEntityExtractor.extract(text);
      if (entities.forecastIds.length > 0) {
        const forecasts: SalesContent['forecasts'] = [];
        const amounts = this.amountParser.parseAll(text);
        for (const forecastId of entities.forecastIds) {
          forecasts.push({
            id: forecastId,
            revenue: amounts[0] || { value: 0, currency: 'EUR' },
          });
        }
        salesContent.forecasts = forecasts;
        matches++;
      }
    }

    // Extract deals
    if (contentLower.includes('deal') || contentLower.includes('opportunity')) {
      const entities = this.salesEntityExtractor.extract(text);
      if (entities.dealNames.length > 0) {
        const deals: SalesContent['deals'] = [];
        const amounts = this.amountParser.parseAll(text);
        for (const dealName of entities.dealNames) {
          deals.push({
            name: dealName,
            value: amounts[0] || { value: 0, currency: 'EUR' },
          });
        }
        salesContent.deals = deals;
        matches++;
      }
    }

    const confidence = this.calculateConfidence(text, matches, patterns.length);

    return {
      type: contentLower.includes('budget')
        ? 'budget'
        : contentLower.includes('forecast')
          ? 'forecast'
          : 'offer',
      confidence,
      content: salesContent,
    };
  }
}

/**
 * HR & Personal Extractor
 */
export class HRExtractor extends BaseExtractor {
  private dateParser = new DateParser();
  private hrEntityExtractor = new HREntityExtractor();

  getPatterns(): RegExp[] {
    return [
      /\b(onboarding|einarbeitung|checkliste)\b/gi,
      /\b(review|bewertung|performance|ziel)\b/gi,
      /\b(bewerbung|application|kandidat|interview)\b/gi,
      /\b(training|schulung|zertifikat)\b/gi,
    ];
  }

  getKeywords(): string[] {
    return [
      'onboarding',
      'einarbeitung',
      'checkliste',
      'review',
      'bewertung',
      'performance',
      'bewerbung',
      'application',
      'kandidat',
      'training',
      'schulung',
    ];
  }

  extract(content: string, options?: ExtractorOptions): ExtractedContent {
    const text = this.sanitizeHtml(content);
    const contentLower = text.toLowerCase();

    const hrContent: HRContent = {};
    let matches = 0;
    const patterns = this.getPatterns();

    // Extract onboarding
    if (contentLower.includes('onboarding') || contentLower.includes('einarbeitung')) {
      const entities = this.hrEntityExtractor.extract(text);
      hrContent.onboarding = {
        checklist: [],
        dates: [],
        responsible: [],
        documents: [],
      };

      if (entities.onboardingIds.length > 0) {
        matches++;
      }

      // Extract checklist items
      const checklistPattern = /[-*]\s*(.+?)(?:\n|$)/g;
      const checklist: string[] = [];
      let match;
      while ((match = checklistPattern.exec(text)) !== null) {
        checklist.push(match[1].trim());
      }
      if (checklist.length > 0) {
        hrContent.onboarding.checklist = checklist;
        matches++;
      }
    }

    // Extract reviews
    if (contentLower.includes('review') || contentLower.includes('bewertung')) {
      const entities = this.hrEntityExtractor.extract(text);
      if (entities.reviewIds.length > 0) {
        hrContent.reviews = entities.reviewIds.map((id) => ({
          id,
          type: 'performance',
        }));
        matches++;
      }
    }

    // Extract applications
    if (contentLower.includes('bewerbung') || contentLower.includes('application')) {
      const entities = this.hrEntityExtractor.extract(text);
      if (entities.candidateIds.length > 0) {
        hrContent.applications = entities.candidateIds.map((id) => ({
          candidateId: id,
          status: 'pending',
        }));
        matches++;
      }
    }

    // Extract trainings
    if (contentLower.includes('training') || contentLower.includes('schulung')) {
      const entities = this.hrEntityExtractor.extract(text);
      if (entities.trainingIds.length > 0) {
        hrContent.trainings = entities.trainingIds.map((id) => ({
          id,
          name: `Training ${id}`,
        }));
        matches++;
      }
    }

    const confidence = this.calculateConfidence(text, matches, patterns.length);

    return {
      type: contentLower.includes('onboarding')
        ? 'onboarding'
        : contentLower.includes('review')
          ? 'review'
          : contentLower.includes('bewerbung') || contentLower.includes('application')
            ? 'application'
            : 'unknown',
      confidence,
      content: hrContent,
    };
  }
}

// ============================================================================
// SINGLETON INSTANCE
// ============================================================================

let registryInstance: ExtractorRegistry | null = null;
let patternBasedExtractorInstance: PatternBasedExtractor | null = null;

/**
 * Get the global extractor registry instance
 */
export function getExtractorRegistry(): ExtractorRegistry {
  if (!registryInstance) {
    registryInstance = new ExtractorRegistry();
  }
  return registryInstance;
}

/**
 * Get the global pattern-based extractor instance
 */
export function getPatternBasedExtractor(): PatternBasedExtractor {
  if (!patternBasedExtractorInstance) {
    patternBasedExtractorInstance = new PatternBasedExtractor(getExtractorRegistry());
    // Initialize and register all extractors
    initializeExtractors();
  }
  return patternBasedExtractorInstance;
}

// ============================================================================
// CACHING
// ============================================================================

/**
 * Simple in-memory cache for extracted content
 */
class ExtractionCache {
  private cache: Map<string, { result: BusinessContentExtraction; timestamp: number }> = new Map();
  private readonly TTL_MS = 60 * 60 * 1000; // 1 hour

  /**
   * Get cached extraction result
   */
  get(key: string): BusinessContentExtraction | null {
    const entry = this.cache.get(key);
    if (!entry) return null;

    // Check if expired
    if (Date.now() - entry.timestamp > this.TTL_MS) {
      this.cache.delete(key);
      return null;
    }

    return entry.result;
  }

  /**
   * Set cached extraction result
   */
  set(key: string, result: BusinessContentExtraction): void {
    this.cache.set(key, {
      result,
      timestamp: Date.now(),
    });
  }

  /**
   * Clear cache
   */
  clear(): void {
    this.cache.clear();
  }

  /**
   * Clear expired entries
   */
  cleanup(): void {
    const now = Date.now();
    for (const [key, entry] of this.cache.entries()) {
      if (now - entry.timestamp > this.TTL_MS) {
        this.cache.delete(key);
      }
    }
  }
}

let extractionCacheInstance: ExtractionCache | null = null;

/**
 * Get the global extraction cache instance
 */
export function getExtractionCache(): ExtractionCache {
  if (!extractionCacheInstance) {
    extractionCacheInstance = new ExtractionCache();
    // Cleanup expired entries every 30 minutes
    setInterval(
      () => {
        extractionCacheInstance?.cleanup();
      },
      30 * 60 * 1000
    );
  }
  return extractionCacheInstance;
}

/**
 * Generate cache key for email extraction
 */
export function generateCacheKey(messageId: string, options?: ExtractorOptions): string {
  const optionsStr = options
    ? JSON.stringify({
        includeMetadata: options.includeMetadata,
        includeEntities: options.includeEntities,
        includeSummary: options.includeSummary,
        language: options.language,
      })
    : '';
  return `extract:${messageId}:${optionsStr}`;
}

/**
 * Initialize and register all business extractors
 */
function initializeExtractors(): void {
  const registry = getExtractorRegistry();

  // Register Project Extractors
  registry.register({
    type: 'roadmap',
    extractor: ProjectExtractor,
    patterns: [/\b(roadmap|meilenstein|timeline|milestone)\b/gi, /\b([A-Z]{1,5}[-]?\d{2,6})\b/g],
    keywords: ['roadmap', 'meilenstein', 'timeline', 'milestone'],
    priority: 10,
  });

  registry.register({
    type: 'project_plan',
    extractor: ProjectExtractor,
    patterns: [/\b(projekt|project|task|aufgabe)\b/gi, /\b([A-Z]{1,5}[-]?\d{2,6})\b/g],
    keywords: ['projekt', 'project', 'task', 'aufgabe'],
    priority: 9,
  });

  registry.register({
    type: 'status_update',
    extractor: ProjectExtractor,
    patterns: [
      /\b(status|fortschritt|progress|blocked)\b/gi,
      /\b(\d+)%\s*(?:progress|fortschritt)/gi,
    ],
    keywords: ['status', 'fortschritt', 'progress', 'blocked', 'blockiert'],
    priority: 8,
  });

  // Register Customer Extractors
  registry.register({
    type: 'customer_info',
    extractor: CustomerExtractor,
    patterns: [/\b(kunde|customer|client|mandant)\b/gi, /\b(CUST|KUNDE|CLIENT)[-:]?\d+\b/gi],
    keywords: ['kunde', 'customer', 'client', 'mandant'],
    priority: 10,
  });

  registry.register({
    type: 'contract',
    extractor: CustomerExtractor,
    patterns: [
      /\b(vertrag|contract|agreement|vereinbarung)\b/gi,
      /\b(CONTRACT|VERTRAG|VERT)[-:]?\d+\b/gi,
    ],
    keywords: ['vertrag', 'contract', 'agreement', 'vereinbarung'],
    priority: 10,
  });

  registry.register({
    type: 'proposal',
    extractor: CustomerExtractor,
    patterns: [/\b(angebot|offer|quote|anfrage)\b/gi, /\b(OFFER|ANGEBOT|QUOTE)[-:]?\d+\b/gi],
    keywords: ['angebot', 'offer', 'quote', 'anfrage'],
    priority: 9,
  });

  // Register Meeting Extractors
  registry.register({
    type: 'meeting_agenda',
    extractor: MeetingExtractor,
    patterns: [/\b(agenda|tagesordnung|top|points)\b/gi, /\b(TOP|AGENDA)[-:]?\d+\b/gi],
    keywords: ['agenda', 'tagesordnung', 'top', 'points'],
    priority: 10,
  });

  registry.register({
    type: 'meeting_notes',
    extractor: MeetingExtractor,
    patterns: [/\b(meeting|besprechung|termin)\b/gi, /\b(notes|notizen|zusammenfassung)\b/gi],
    keywords: ['meeting', 'besprechung', 'termin', 'notes', 'notizen'],
    priority: 9,
  });

  registry.register({
    type: 'action_items',
    extractor: MeetingExtractor,
    patterns: [/\b(action\s+item|aufgabe|todo|follow-up)\b/gi, /\b(AI|ACTION|AUFGABE)[-:]?\d+\b/gi],
    keywords: ['action item', 'aufgabe', 'todo', 'follow-up'],
    priority: 9,
  });

  // Register Document Extractors
  registry.register({
    type: 'invoice',
    extractor: DocumentExtractor,
    patterns: [/\b(rechnung|invoice|bill|zahlung)\b/gi, /\b(INV|INVOICE|RECHNUNG)[-:]?\d+\b/gi],
    keywords: ['rechnung', 'invoice', 'bill', 'zahlung'],
    priority: 10,
  });

  registry.register({
    type: 'report',
    extractor: DocumentExtractor,
    patterns: [/\b(bericht|report|summary|zusammenfassung)\b/gi, /\b(REPORT|BERICHT)[-:]?\d+\b/gi],
    keywords: ['bericht', 'report', 'summary', 'zusammenfassung'],
    priority: 9,
  });

  registry.register({
    type: 'executive_summary',
    extractor: DocumentExtractor,
    patterns: [
      /\b(executive\s+summary|executive\s+zusammenfassung)\b/gi,
      /\b(kpis?|key\s+performance\s+indicators)\b/gi,
    ],
    keywords: ['executive summary', 'executive zusammenfassung', 'kpi', 'kpis'],
    priority: 8,
  });

  // Register Sales Extractors
  registry.register({
    type: 'offer',
    extractor: SalesExtractor,
    patterns: [/\b(angebot|offer|quote|preis)\b/gi, /\b(OFFER|QUOTE)[-:]?\d+\b/gi],
    keywords: ['angebot', 'offer', 'quote', 'preis'],
    priority: 10,
  });

  registry.register({
    type: 'budget',
    extractor: SalesExtractor,
    patterns: [/\b(budget|kosten|cost|ausgabe)\b/gi, /\b(BUDGET|BUD)[-:]?\d+\b/gi],
    keywords: ['budget', 'kosten', 'cost', 'ausgabe'],
    priority: 9,
  });

  registry.register({
    type: 'forecast',
    extractor: SalesExtractor,
    patterns: [/\b(forecast|prognose|pipeline|umsatz)\b/gi, /\b(FORECAST|PROGNOSE)[-:]?\d+\b/gi],
    keywords: ['forecast', 'prognose', 'pipeline', 'umsatz'],
    priority: 9,
  });

  // Register HR Extractors
  registry.register({
    type: 'onboarding',
    extractor: HRExtractor,
    patterns: [
      /\b(onboarding|einarbeitung|checkliste)\b/gi,
      /\b(ONBOARD|EINARBEITUNG)[-:]?\d+\b/gi,
    ],
    keywords: ['onboarding', 'einarbeitung', 'checkliste'],
    priority: 10,
  });

  registry.register({
    type: 'review',
    extractor: HRExtractor,
    patterns: [/\b(review|bewertung|performance|ziel)\b/gi, /\b(REVIEW|BEWERTUNG)[-:]?\d+\b/gi],
    keywords: ['review', 'bewertung', 'performance', 'ziel'],
    priority: 9,
  });

  registry.register({
    type: 'application',
    extractor: HRExtractor,
    patterns: [
      /\b(bewerbung|application|kandidat|interview)\b/gi,
      /\b(CAND|BEWERBER|APPLICANT)[-:]?\d+\b/gi,
    ],
    keywords: ['bewerbung', 'application', 'kandidat', 'interview'],
    priority: 9,
  });

  logger.info('Initialized all business content extractors');
}

// ============================================================================
// COMMON EXTRACTION COMPONENTS
// ============================================================================

/**
 * Table and List Parser
 */
export class TableListParser {
  /**
   * Parse HTML table
   */
  parseHtmlTable(html: string): Table[] {
    const tables: Table[] = [];
    const tableRegex = /<table[^>]*>(.*?)<\/table>/gis;
    let match;

    while ((match = tableRegex.exec(html)) !== null) {
      const tableHtml = match[1];
      const headers: string[] = [];
      const rows: string[][] = [];

      // Extract headers
      const headerRegex = /<th[^>]*>(.*?)<\/th>/gis;
      let headerMatch;
      while ((headerMatch = headerRegex.exec(tableHtml)) !== null) {
        headers.push(this.sanitizeText(headerMatch[1]));
      }

      // Extract rows
      const rowRegex = /<tr[^>]*>(.*?)<\/tr>/gis;
      let rowMatch;
      while ((rowMatch = rowRegex.exec(tableHtml)) !== null) {
        const rowHtml = rowMatch[1];
        const cells: string[] = [];
        const cellRegex = /<td[^>]*>(.*?)<\/td>/gis;
        let cellMatch;
        while ((cellMatch = cellRegex.exec(rowHtml)) !== null) {
          cells.push(this.sanitizeText(cellMatch[1]));
        }
        if (cells.length > 0) {
          rows.push(cells);
        }
      }

      if (headers.length > 0 || rows.length > 0) {
        tables.push({ headers, rows });
      }
    }

    return tables;
  }

  /**
   * Parse Markdown table
   */
  parseMarkdownTable(markdown: string): Table[] {
    const tables: Table[] = [];
    const lines = markdown.split('\n');
    let currentTable: { headers: string[]; rows: string[][] } | null = null;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      if (line.startsWith('|') && line.endsWith('|')) {
        const cells = line
          .split('|')
          .map((cell) => cell.trim())
          .filter((cell) => cell.length > 0);

        if (cells.length > 0) {
          // Check if it's a separator line
          if (cells.every((cell) => /^:?-+:?$/.test(cell))) {
            continue; // Skip separator
          }

          if (!currentTable) {
            currentTable = { headers: cells, rows: [] };
          } else {
            currentTable.rows.push(cells);
          }
        }
      } else {
        if (currentTable && (currentTable.headers.length > 0 || currentTable.rows.length > 0)) {
          tables.push(currentTable);
        }
        currentTable = null;
      }
    }

    if (currentTable && (currentTable.headers.length > 0 || currentTable.rows.length > 0)) {
      tables.push(currentTable);
    }

    return tables;
  }

  /**
   * Parse HTML list
   */
  parseHtmlList(html: string): List[] {
    const lists: List[] = [];
    const listRegex = /<(ul|ol)[^>]*>(.*?)<\/\1>/gis;
    let match;

    while ((match = listRegex.exec(html)) !== null) {
      const listType = match[1] === 'ol' ? 'ordered' : 'unordered';
      const listHtml = match[2];
      const items: string[] = [];

      const itemRegex = /<li[^>]*>(.*?)<\/li>/gis;
      let itemMatch;
      while ((itemMatch = itemRegex.exec(listHtml)) !== null) {
        items.push(this.sanitizeText(itemMatch[1]));
      }

      if (items.length > 0) {
        lists.push({ items, type: listType });
      }
    }

    return lists;
  }

  /**
   * Parse Markdown list
   */
  parseMarkdownList(markdown: string): List[] {
    const lists: List[] = [];
    const lines = markdown.split('\n');
    let currentList: { items: string[]; type: 'ordered' | 'unordered' } | null = null;

    for (const line of lines) {
      const trimmed = line.trim();
      const orderedMatch = trimmed.match(/^\d+\.\s+(.+)$/);
      const unorderedMatch = trimmed.match(/^[-*+]\s+(.+)$/);

      if (orderedMatch) {
        if (!currentList || currentList.type !== 'ordered') {
          if (currentList && currentList.items.length > 0) {
            lists.push(currentList);
          }
          currentList = { items: [], type: 'ordered' };
        }
        currentList.items.push(orderedMatch[1].trim());
      } else if (unorderedMatch) {
        if (!currentList || currentList.type !== 'unordered') {
          if (currentList && currentList.items.length > 0) {
            lists.push(currentList);
          }
          currentList = { items: [], type: 'unordered' };
        }
        currentList.items.push(unorderedMatch[1].trim());
      } else {
        if (currentList && currentList.items.length > 0) {
          lists.push(currentList);
        }
        currentList = null;
      }
    }

    if (currentList && currentList.items.length > 0) {
      lists.push(currentList);
    }

    return lists;
  }

  private sanitizeText(text: string): string {
    return text
      .replace(/<[^>]*>/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  }
}

/**
 * Metadata Extractor
 */
export class MetadataExtractor {
  /**
   * Extract priorities from content
   */
  extractPriorities(content: string): string[] {
    const priorities: string[] = [];
    const contentLower = content.toLowerCase();

    const priorityPatterns = [
      { pattern: /\b(high|hoch|urgent|dringend|critical|kritisch)\s+priority/gi, value: 'high' },
      { pattern: /\bpriority:\s*(high|hoch|urgent|dringend)/gi, value: 'high' },
      { pattern: /\b(medium|mittel|normal)\s+priority/gi, value: 'medium' },
      { pattern: /\bpriority:\s*(medium|mittel|normal)/gi, value: 'medium' },
      { pattern: /\b(low|niedrig|low)\s+priority/gi, value: 'low' },
      { pattern: /\bpriority:\s*(low|niedrig)/gi, value: 'low' },
    ];

    for (const { pattern, value } of priorityPatterns) {
      if (pattern.test(content)) {
        priorities.push(value);
      }
    }

    return [...new Set(priorities)];
  }

  /**
   * Extract statuses from content
   */
  extractStatuses(content: string): string[] {
    const statuses: string[] = [];
    const contentLower = content.toLowerCase();

    const statusPatterns = [
      { pattern: /\b(pending|ausstehend|offen)\b/gi, value: 'pending' },
      { pattern: /\b(in\s+progress|in\s+arbeit|laufend)\b/gi, value: 'in-progress' },
      { pattern: /\b(completed|abgeschlossen|fertig|done|erledigt)\b/gi, value: 'completed' },
      { pattern: /\b(blocked|blockiert|gestoppt)\b/gi, value: 'blocked' },
      { pattern: /\b(cancelled|abgebrochen|storniert)\b/gi, value: 'cancelled' },
    ];

    for (const { pattern, value } of statusPatterns) {
      if (pattern.test(content)) {
        statuses.push(value);
      }
    }

    return [...new Set(statuses)];
  }

  /**
   * Extract deadlines from content
   */
  extractDeadlines(content: string): Date[] {
    const deadlines: Date[] = [];
    const dateParser = new DateParser();

    const deadlinePatterns = [
      /\bdeadline[:\s]+([^\n]+)/gi,
      /\bfrist[:\s]+([^\n]+)/gi,
      /\bdue\s+date[:\s]+([^\n]+)/gi,
      /\bfällig[:\s]+([^\n]+)/gi,
      /\bdeadline\s+is\s+([^\n]+)/gi,
    ];

    for (const pattern of deadlinePatterns) {
      let match;
      while ((match = pattern.exec(content)) !== null) {
        const dateStr = match[1].trim();
        const date = dateParser.parse(dateStr);
        if (date) {
          deadlines.push(date);
        }
      }
    }

    return deadlines;
  }

  /**
   * Extract tags from content
   */
  extractTags(content: string): string[] {
    const tags: string[] = [];

    // Extract hashtags
    const hashtagPattern = /#(\w+)/g;
    let match;
    while ((match = hashtagPattern.exec(content)) !== null) {
      tags.push(match[1]);
    }

    // Extract tags in brackets
    const bracketPattern = /\[([^\]]+)\]/g;
    while ((match = bracketPattern.exec(content)) !== null) {
      const tag = match[1].trim();
      if (tag.length > 0 && tag.length < 50) {
        tags.push(tag);
      }
    }

    return [...new Set(tags)];
  }
}

/**
 * Date Parser
 */
export class DateParser {
  /**
   * Parse date from text (DE/EN)
   */
  parse(text: string): Date | null {
    if (!text || typeof text !== 'string') {
      return null;
    }

    const trimmed = text.trim();

    // Try ISO format first
    const isoMatch = trimmed.match(/^\d{4}-\d{2}-\d{2}/);
    if (isoMatch) {
      const date = new Date(trimmed);
      if (!isNaN(date.getTime())) {
        return date;
      }
    }

    // Try German date format (DD.MM.YYYY)
    const deMatch = trimmed.match(/(\d{1,2})\.(\d{1,2})\.(\d{4})/);
    if (deMatch) {
      const [, day, month, year] = deMatch;
      const date = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
      if (!isNaN(date.getTime())) {
        return date;
      }
    }

    // Try US date format (MM/DD/YYYY)
    const usMatch = trimmed.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (usMatch) {
      const [, month, day, year] = usMatch;
      const date = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
      if (!isNaN(date.getTime())) {
        return date;
      }
    }

    // Try relative dates
    const today = new Date();
    const lower = trimmed.toLowerCase();

    if (lower.includes('today') || lower.includes('heute')) {
      return today;
    }
    if (lower.includes('tomorrow') || lower.includes('morgen')) {
      const tomorrow = new Date(today);
      tomorrow.setDate(tomorrow.getDate() + 1);
      return tomorrow;
    }
    if (lower.includes('yesterday') || lower.includes('gestern')) {
      const yesterday = new Date(today);
      yesterday.setDate(yesterday.getDate() - 1);
      return yesterday;
    }

    // Try parsing as Date
    const parsed = new Date(trimmed);
    if (!isNaN(parsed.getTime())) {
      return parsed;
    }

    return null;
  }
}

/**
 * Amount Parser
 */
export class AmountParser {
  /**
   * Parse amount with currency from text
   */
  parse(text: string): Amount | null {
    if (!text || typeof text !== 'string') {
      return null;
    }

    // Common currency patterns
    // Try to match numbers with thousands separators first (more specific)
    const patterns = [
      // EUR formats - with thousands separators
      { pattern: /(\d{1,3}(?:,\d{3})+\.\d{2})\s*€/i, currency: 'EUR' },
      { pattern: /(\d{1,3}(?:\.\d{3})+\d{2})\s*€/i, currency: 'EUR' },
      // EUR formats - without thousands separators
      { pattern: /(\d+(?:[.,]\d{2})?)\s*€/i, currency: 'EUR' },
      { pattern: /€\s*(\d{1,3}(?:,\d{3})+\.\d{2})/i, currency: 'EUR' },
      { pattern: /€\s*(\d{1,3}(?:\.\d{3})+\d{2})/i, currency: 'EUR' },
      { pattern: /€\s*(\d+(?:[.,]\d{2})?)/i, currency: 'EUR' },
      { pattern: /EUR\s*(\d{1,3}(?:,\d{3})+\.\d{2})/i, currency: 'EUR' },
      { pattern: /EUR\s*(\d{1,3}(?:\.\d{3})+\d{2})/i, currency: 'EUR' },
      { pattern: /EUR\s*(\d+(?:[.,]\d{2})?)/i, currency: 'EUR' },
      // USD formats
      { pattern: /(\d+(?:[.,]\d{2})?)\s*\$/, currency: 'USD' },
      { pattern: /\$\s*(\d+(?:[.,]\d{2})?)/, currency: 'USD' },
      { pattern: /USD\s*(\d+(?:[.,]\d{2})?)/i, currency: 'USD' },
      // GBP formats
      { pattern: /(\d+(?:[.,]\d{2})?)\s*£/, currency: 'GBP' },
      { pattern: /£\s*(\d+(?:[.,]\d{2})?)/, currency: 'GBP' },
      { pattern: /GBP\s*(\d+(?:[.,]\d{2})?)/i, currency: 'GBP' },
    ];

    for (const { pattern, currency } of patterns) {
      const match = text.match(pattern);
      if (match) {
        // Handle both comma and dot as decimal separator
        // If comma is used, it might be thousands separator or decimal separator
        let valueStr = match[1];
        // Handle thousands separators and decimal separators
        // Pattern "1,234.56" = US format (comma = thousands, dot = decimal)
        // Pattern "1.234,56" = EU format (dot = thousands, comma = decimal)
        if (valueStr.includes(',') && valueStr.includes('.')) {
          // Both present - determine which is decimal separator
          const lastComma = valueStr.lastIndexOf(',');
          const lastDot = valueStr.lastIndexOf('.');
          if (lastDot > lastComma) {
            // Dot is decimal separator (US: 1,234.56)
            valueStr = valueStr.replace(/,/g, '');
          } else {
            // Comma is decimal separator (EU: 1.234,56)
            valueStr = valueStr.replace(/\./g, '').replace(',', '.');
          }
        } else if (valueStr.includes(',')) {
          // Only comma - check context
          const parts = valueStr.split(',');
          if (parts.length > 2 || (parts.length === 2 && parts[1].length > 2)) {
            // Multiple commas or >2 digits after = thousands separator
            valueStr = valueStr.replace(/,/g, '');
          } else if (parts.length === 2 && parts[1].length <= 2) {
            // Single comma with 1-2 digits = decimal separator
            valueStr = parts[0] + '.' + parts[1];
          }
        } else if (valueStr.includes('.')) {
          // Only dot - check context
          const parts = valueStr.split('.');
          if (parts.length > 2 || (parts.length === 2 && parts[1].length > 2)) {
            // Multiple dots or >2 digits after = thousands separator
            valueStr = valueStr.replace(/\./g, '');
          }
          // Single dot with <=2 digits = decimal separator (already correct)
        }
        const value = parseFloat(valueStr);
        if (!isNaN(value)) {
          return {
            value,
            currency,
            context: text.substring(0, 100), // First 100 chars as context
          };
        }
      }
    }

    return null;
  }

  /**
   * Parse multiple amounts from text
   */
  parseAll(text: string): Amount[] {
    const amounts: Amount[] = [];
    const patterns = [
      { pattern: /(\d+(?:[.,]\d{2})?)\s*€/gi, currency: 'EUR' },
      { pattern: /€\s*(\d+(?:[.,]\d{2})?)/gi, currency: 'EUR' },
      { pattern: /EUR\s*(\d+(?:[.,]\d{2})?)/gi, currency: 'EUR' },
      { pattern: /(\d+(?:[.,]\d{2})?)\s*$/gi, currency: 'USD' },
      { pattern: /\$\s*(\d+(?:[.,]\d{2})?)/gi, currency: 'USD' },
      { pattern: /USD\s*(\d+(?:[.,]\d{2})?)/gi, currency: 'USD' },
    ];

    for (const { pattern, currency } of patterns) {
      let match;
      while ((match = pattern.exec(text)) !== null) {
        const valueStr = match[1].replace(',', '.');
        const value = parseFloat(valueStr);
        if (!isNaN(value)) {
          amounts.push({
            value,
            currency,
            context: text.substring(Math.max(0, match.index - 50), match.index + 50),
          });
        }
      }
    }

    return amounts;
  }
}

// ============================================================================
// ENTITY EXTRACTORS
// ============================================================================

/**
 * Base Entity Extractor with common functionality
 */
export abstract class BaseEntityExtractor {
  /**
   * Extract entities from content
   */
  abstract extract(content: string): {
    [key: string]: unknown[];
  };

  /**
   * Get patterns used by this extractor
   */
  abstract getPatterns(): RegExp[];

  /**
   * Extract entities using patterns
   */
  protected extractByPatterns(content: string, patterns: RegExp[]): string[] {
    const entities: string[] = [];

    for (const pattern of patterns) {
      let match;
      while ((match = pattern.exec(content)) !== null) {
        const entity = match[1] || match[0];
        if (entity && !entities.includes(entity)) {
          entities.push(entity.trim());
        }
      }
    }

    return entities;
  }

  /**
   * Extract email addresses
   */
  protected extractEmails(content: string): string[] {
    const emailPattern = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g;
    return this.extractByPatterns(content, [emailPattern]);
  }

  /**
   * Extract dates
   */
  protected extractDates(content: string): Date[] {
    const dates: Date[] = [];
    const dateParser = new DateParser();

    const datePatterns = [
      /\b\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\b/g,
      /\b\d{4}-\d{2}-\d{2}\b/g,
      /\b(january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{1,2},?\s+\d{4}\b/gi,
      /\b(januar|februar|märz|april|mai|juni|juli|august|september|oktober|november|dezember)\s+\d{1,2},?\s+\d{4}\b/gi,
    ];

    for (const pattern of datePatterns) {
      let match;
      while ((match = pattern.exec(content)) !== null) {
        const date = dateParser.parse(match[0]);
        if (date && !dates.some((d) => d.getTime() === date.getTime())) {
          dates.push(date);
        }
      }
    }

    return dates;
  }

  /**
   * Extract person names
   */
  protected extractPersons(content: string): Person[] {
    const persons: Person[] = [];

    // Pattern for names with titles
    const titlePattern =
      /\b(Herr|Frau|Mr\.?|Mrs\.?|Ms\.?|Dr\.?|Prof\.?)\s+([A-ZÄÖÜ][a-zäöüß]+(?:\s+[A-ZÄÖÜ][a-zäöüß]+)?)/gi;
    let match;
    while ((match = titlePattern.exec(content)) !== null) {
      persons.push({
        name: match[2].trim(),
        role: match[1].trim(),
      });
    }

    // Pattern for "First Last" names (capitalized words)
    const namePattern = /\b([A-ZÄÖÜ][a-zäöüß]+)\s+([A-ZÄÖÜ][a-zäöüß]+)\b/g;
    while ((match = namePattern.exec(content)) !== null) {
      const fullName = `${match[1]} ${match[2]}`;
      if (!persons.some((p) => p.name === fullName)) {
        persons.push({ name: fullName });
      }
    }

    return persons;
  }
}

/**
 * Project Entity Extractor
 */
export class ProjectEntityExtractor extends BaseEntityExtractor {
  getPatterns(): RegExp[] {
    return [
      /\b([A-Z]{1,5}[-]?\d{2,6})\b/g, // Project codes: P2046, PROJ-123, AB-456
      /\b(TASK|TASK-|#)\d+\b/gi, // Task IDs: TASK-789, #12345
      /\b(Team|Team-|Ressource:)\s*([A-Z][a-z]+)\b/gi, // Resources
    ];
  }

  extract(content: string): {
    projectCodes: string[];
    taskIds: string[];
    resources: Resource[];
  } {
    const projectCodes: string[] = [];
    const taskIds: string[] = [];
    const resources: Resource[] = [];

    // Extract project codes
    const projectCodePattern = /\b([A-Z]{1,5}[-]?\d{2,6})\b/g;
    let match;
    while ((match = projectCodePattern.exec(content)) !== null) {
      const code = match[1];
      if (!projectCodes.includes(code)) {
        projectCodes.push(code);
      }
    }

    // Extract task IDs
    const taskPatterns = [/\bTASK[-:]?\s*(\d+)\b/gi, /#(\d+)\b/g];
    for (const pattern of taskPatterns) {
      while ((match = pattern.exec(content)) !== null) {
        const taskId = match[1];
        if (!taskIds.includes(taskId)) {
          taskIds.push(taskId);
        }
      }
    }

    // Extract resources
    const resourcePattern = /\b(Team|Team-|Ressource:)\s*([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)\b/gi;
    while ((match = resourcePattern.exec(content)) !== null) {
      resources.push({
        name: match[2].trim(),
        type: 'team',
      });
    }

    return { projectCodes, taskIds, resources };
  }
}

/**
 * Customer Entity Extractor
 */
export class CustomerEntityExtractor extends BaseEntityExtractor {
  getPatterns(): RegExp[] {
    return [
      /\b(CONTRACT|VERTRAG|VERT)[-:]?\s*(\d+)\b/gi,
      /\b(OFFER|ANGEBOT|QUOTE|QUOT)[-:]?\s*(\d+)\b/gi,
      /\b(CUST|KUNDE|CLIENT)[-:]?\s*(\d+)\b/gi,
    ];
  }

  extract(content: string): {
    customerNames: string[];
    contractNumbers: string[];
    offerNumbers: string[];
    customerIds: string[];
  } {
    const customerNames: string[] = [];
    const contractNumbers: string[] = [];
    const offerNumbers: string[] = [];
    const customerIds: string[] = [];

    // Extract contract numbers
    const contractPattern = /\b(CONTRACT|VERTRAG|VERT)[-:]?\s*(\d+)\b/gi;
    let match;
    while ((match = contractPattern.exec(content)) !== null) {
      const number = `${match[1]}-${match[2]}`;
      if (!contractNumbers.includes(number)) {
        contractNumbers.push(number);
      }
    }

    // Extract offer numbers
    const offerPattern = /\b(OFFER|ANGEBOT|QUOTE|QUOT)[-:]?\s*(\d+)\b/gi;
    while ((match = offerPattern.exec(content)) !== null) {
      const number = `${match[1]}-${match[2]}`;
      if (!offerNumbers.includes(number)) {
        offerNumbers.push(number);
      }
    }

    // Extract customer IDs
    const customerIdPattern = /\b(CUST|KUNDE|CLIENT)[-:]?\s*(\d+)\b/gi;
    while ((match = customerIdPattern.exec(content)) !== null) {
      const id = `${match[1]}-${match[2]}`;
      if (!customerIds.includes(id)) {
        customerIds.push(id);
      }
    }

    // Extract customer names (organization names)
    const orgPattern =
      /\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)\s+(GmbH|AG|Inc|LLC|Ltd|Corp|Company)\b/gi;
    while ((match = orgPattern.exec(content)) !== null) {
      const name = match[0].trim();
      if (!customerNames.includes(name)) {
        customerNames.push(name);
      }
    }

    return { customerNames, contractNumbers, offerNumbers, customerIds };
  }
}

/**
 * Meeting Entity Extractor
 */
export class MeetingEntityExtractor extends BaseEntityExtractor {
  getPatterns(): RegExp[] {
    return [
      /\b(AI|ACTION|AUFGABE)[-:]?\s*(\d+)\b/gi,
      /\b(DEC|DECISION|ENTSCHEIDUNG)[-:]?\s*(\d+)\b/gi,
      /\b(TOP|AGENDA|PUNKT)[-:]?\s*(\d+)\b/gi,
    ];
  }

  extract(content: string): {
    participants: Person[];
    actionItemIds: string[];
    decisionIds: string[];
    agendaPoints: string[];
  } {
    const participants = this.extractPersons(content);
    const actionItemIds: string[] = [];
    const decisionIds: string[] = [];
    const agendaPoints: string[] = [];

    // Extract action item IDs
    const actionItemPattern = /\b(AI|ACTION|AUFGABE)[-:]?\s*(\d+)\b/gi;
    let match;
    while ((match = actionItemPattern.exec(content)) !== null) {
      const id = `${match[1]}-${match[2]}`;
      if (!actionItemIds.includes(id)) {
        actionItemIds.push(id);
      }
    }

    // Extract decision IDs
    const decisionPattern = /\b(DEC|DECISION|ENTSCHEIDUNG)[-:]?\s*(\d+)\b/gi;
    while ((match = decisionPattern.exec(content)) !== null) {
      const id = `${match[1]}-${match[2]}`;
      if (!decisionIds.includes(id)) {
        decisionIds.push(id);
      }
    }

    // Extract agenda points
    const agendaPattern = /\b(TOP|AGENDA|PUNKT)[-:]?\s*(\d+)\b/gi;
    while ((match = agendaPattern.exec(content)) !== null) {
      const point = `${match[1]}-${match[2]}`;
      if (!agendaPoints.includes(point)) {
        agendaPoints.push(point);
      }
    }

    return { participants, actionItemIds, decisionIds, agendaPoints };
  }
}

/**
 * Document Entity Extractor
 */
export class DocumentEntityExtractor extends BaseEntityExtractor {
  getPatterns(): RegExp[] {
    return [
      /\b(INV|INVOICE|RECHNUNG|RECH)[-:]?\s*(\d+)\b/gi,
      /\b(DOC|DOKUMENT|FILE)[-:]?\s*(\d+)\b/gi,
      /\b(REPORT|BERICHT)[-:]?\s*(\d+)\b/gi,
    ];
  }

  extract(content: string): {
    invoiceNumbers: string[];
    documentIds: string[];
    reportNumbers: string[];
  } {
    const invoiceNumbers: string[] = [];
    const documentIds: string[] = [];
    const reportNumbers: string[] = [];

    // Extract invoice numbers
    const invoicePattern = /\b(INV|INVOICE|RECHNUNG|RECH)[-:]?\s*(\d+)\b/gi;
    let match;
    while ((match = invoicePattern.exec(content)) !== null) {
      const number = `${match[1]}-${match[2]}`;
      if (!invoiceNumbers.includes(number)) {
        invoiceNumbers.push(number);
      }
    }

    // Extract document IDs
    const docPattern = /\b(DOC|DOKUMENT|FILE)[-:]?\s*(\d+)\b/gi;
    while ((match = docPattern.exec(content)) !== null) {
      const id = `${match[1]}-${match[2]}`;
      if (!documentIds.includes(id)) {
        documentIds.push(id);
      }
    }

    // Extract report numbers
    const reportPattern = /\b(REPORT|BERICHT)[-:]?\s*(\d+)\b/gi;
    while ((match = reportPattern.exec(content)) !== null) {
      const number = `${match[1]}-${match[2]}`;
      if (!reportNumbers.includes(number)) {
        reportNumbers.push(number);
      }
    }

    return { invoiceNumbers, documentIds, reportNumbers };
  }
}

/**
 * Sales Entity Extractor
 */
export class SalesEntityExtractor extends BaseEntityExtractor {
  getPatterns(): RegExp[] {
    return [
      /\b(DEAL|OPPORTUNITY|GESCHÄFT)[-:]?\s*(\w+)\b/gi,
      /\b(BUDGET|BUD)[-:]?\s*(\d+)\b/gi,
      /\b(FORECAST|PROGNOSE)[-:]?\s*(\d+)\b/gi,
    ];
  }

  extract(content: string): {
    dealNames: string[];
    budgetIds: string[];
    forecastIds: string[];
  } {
    const dealNames: string[] = [];
    const budgetIds: string[] = [];
    const forecastIds: string[] = [];

    // Extract deal names
    const dealPattern = /\b(DEAL|OPPORTUNITY|GESCHÄFT)[-:]?\s*([A-Z0-9-]+)\b/gi;
    let match;
    while ((match = dealPattern.exec(content)) !== null) {
      // If the second group already starts with the first group, don't duplicate
      const secondPart = match[2];
      let name: string;
      if (secondPart.toUpperCase().startsWith(match[1].toUpperCase())) {
        name = secondPart;
      } else {
        name = `${match[1]}-${secondPart}`;
      }
      if (!dealNames.includes(name)) {
        dealNames.push(name);
      }
    }

    // Extract budget IDs
    const budgetPattern = /\b(BUDGET|BUD)[-:]?\s*(\d+)\b/gi;
    while ((match = budgetPattern.exec(content)) !== null) {
      const id = `${match[1]}-${match[2]}`;
      if (!budgetIds.includes(id)) {
        budgetIds.push(id);
      }
    }

    // Extract forecast IDs
    const forecastPattern = /\b(FORECAST|PROGNOSE)[-:]?\s*(\d+)\b/gi;
    while ((match = forecastPattern.exec(content)) !== null) {
      const id = `${match[1]}-${match[2]}`;
      if (!forecastIds.includes(id)) {
        forecastIds.push(id);
      }
    }

    return { dealNames, budgetIds, forecastIds };
  }
}

/**
 * HR Entity Extractor
 */
export class HREntityExtractor extends BaseEntityExtractor {
  getPatterns(): RegExp[] {
    return [
      /\b(CAND|BEWERBER|APPLICANT)[-:]?\s*(\d+)\b/gi,
      /\b(REVIEW|BEWERTUNG)[-:]?\s*(\d+)\b/gi,
      /\b(TRAINING|SCHULUNG)[-:]?\s*(\d+)\b/gi,
      /\b(ONBOARD|EINARBEITUNG)[-:]?\s*(\d+)\b/gi,
    ];
  }

  extract(content: string): {
    candidateIds: string[];
    reviewIds: string[];
    trainingIds: string[];
    onboardingIds: string[];
  } {
    const candidateIds: string[] = [];
    const reviewIds: string[] = [];
    const trainingIds: string[] = [];
    const onboardingIds: string[] = [];

    // Extract candidate IDs
    const candidatePattern = /\b(CAND|BEWERBER|APPLICANT)[-:]?\s*(\d+)\b/gi;
    let match;
    while ((match = candidatePattern.exec(content)) !== null) {
      const id = `${match[1]}-${match[2]}`;
      if (!candidateIds.includes(id)) {
        candidateIds.push(id);
      }
    }

    // Extract review IDs
    const reviewPattern = /\b(REVIEW|BEWERTUNG)[-:]?\s*(\d+)\b/gi;
    while ((match = reviewPattern.exec(content)) !== null) {
      const id = `${match[1]}-${match[2]}`;
      if (!reviewIds.includes(id)) {
        reviewIds.push(id);
      }
    }

    // Extract training IDs
    const trainingPattern = /\b(TRAINING|SCHULUNG)[-:]?\s*(\d+)\b/gi;
    while ((match = trainingPattern.exec(content)) !== null) {
      const id = `${match[1]}-${match[2]}`;
      if (!trainingIds.includes(id)) {
        trainingIds.push(id);
      }
    }

    // Extract onboarding IDs
    const onboardingPattern = /\b(ONBOARD|EINARBEITUNG)[-:]?\s*(\d+)\b/gi;
    while ((match = onboardingPattern.exec(content)) !== null) {
      const id = `${match[1]}-${match[2]}`;
      if (!onboardingIds.includes(id)) {
        onboardingIds.push(id);
      }
    }

    return { candidateIds, reviewIds, trainingIds, onboardingIds };
  }
}

/**
 * Entity Extractor Registry
 */
export class EntityExtractorRegistry {
  private extractors: Map<string, BaseEntityExtractor> = new Map();

  constructor() {
    // Register all entity extractors
    this.register('project', new ProjectEntityExtractor());
    this.register('customer', new CustomerEntityExtractor());
    this.register('meeting', new MeetingEntityExtractor());
    this.register('document', new DocumentEntityExtractor());
    this.register('sales', new SalesEntityExtractor());
    this.register('hr', new HREntityExtractor());
  }

  /**
   * Register an entity extractor
   */
  register(name: string, extractor: BaseEntityExtractor): void {
    this.extractors.set(name, extractor);
  }

  /**
   * Get entity extractor by name
   */
  get(name: string): BaseEntityExtractor | null {
    return this.extractors.get(name) || null;
  }

  /**
   * Extract all entities from content using all registered extractors
   */
  extractAll(content: string): BusinessContentExtraction['entities'] {
    const entities: BusinessContentExtraction['entities'] = {
      people: [],
      organizations: [],
      dates: [],
      amounts: [],
      projectCodes: [],
      taskIds: [],
      resources: [],
      customerNames: [],
      contractNumbers: [],
      offerNumbers: [],
      customerIds: [],
      participants: [],
      actionItemIds: [],
      decisionIds: [],
      agendaPoints: [],
      invoiceNumbers: [],
      documentIds: [],
      reportNumbers: [],
      dealNames: [],
      budgetIds: [],
      forecastIds: [],
      candidateIds: [],
      reviewIds: [],
      trainingIds: [],
      onboardingIds: [],
    };

    // Extract from project extractor
    const projectExtractor = this.get('project');
    if (projectExtractor) {
      const projectEntities = projectExtractor.extract(content);
      entities.projectCodes = projectEntities.projectCodes as string[];
      entities.taskIds = projectEntities.taskIds as string[];
      entities.resources = projectEntities.resources as Resource[];
    }

    // Extract from customer extractor
    const customerExtractor = this.get('customer');
    if (customerExtractor) {
      const customerEntities = customerExtractor.extract(content);
      entities.customerNames = customerEntities.customerNames as string[];
      entities.contractNumbers = customerEntities.contractNumbers as string[];
      entities.offerNumbers = customerEntities.offerNumbers as string[];
      entities.customerIds = customerEntities.customerIds as string[];
    }

    // Extract from meeting extractor
    const meetingExtractor = this.get('meeting');
    if (meetingExtractor) {
      const meetingEntities = meetingExtractor.extract(content);
      entities.participants = meetingEntities.participants as Person[];
      entities.actionItemIds = meetingEntities.actionItemIds as string[];
      entities.decisionIds = meetingEntities.decisionIds as string[];
      entities.agendaPoints = meetingEntities.agendaPoints as string[];
    }

    // Extract from document extractor
    const documentExtractor = this.get('document');
    if (documentExtractor) {
      const documentEntities = documentExtractor.extract(content);
      entities.invoiceNumbers = documentEntities.invoiceNumbers as string[];
      entities.documentIds = documentEntities.documentIds as string[];
      entities.reportNumbers = documentEntities.reportNumbers as string[];
    }

    // Extract from sales extractor
    const salesExtractor = this.get('sales');
    if (salesExtractor) {
      const salesEntities = salesExtractor.extract(content);
      entities.dealNames = salesEntities.dealNames as string[];
      entities.budgetIds = salesEntities.budgetIds as string[];
      entities.forecastIds = salesEntities.forecastIds as string[];
    }

    // Extract from HR extractor
    const hrExtractor = this.get('hr');
    if (hrExtractor) {
      const hrEntities = hrExtractor.extract(content);
      entities.candidateIds = hrEntities.candidateIds as string[];
      entities.reviewIds = hrEntities.reviewIds as string[];
      entities.trainingIds = hrEntities.trainingIds as string[];
      entities.onboardingIds = hrEntities.onboardingIds as string[];
    }

    // Extract common entities
    const baseExtractor = new ProjectEntityExtractor(); // Use any extractor for base methods
    entities.people = baseExtractor['extractPersons'](content);
    entities.dates = baseExtractor['extractDates'](content);

    const amountParser = new AmountParser();
    entities.amounts = amountParser.parseAll(content);

    return entities;
  }
}

/**
 * Summary Generator
 */
export class SummaryGenerator {
  /**
   * Generate action items from content
   */
  generateActionItems(content: string): ActionItem[] {
    const actionItems: ActionItem[] = [];
    const dateParser = new DateParser();
    const metadataExtractor = new MetadataExtractor();

    // Action item patterns
    const patterns = [
      {
        pattern: /(?:action\s+item|aufgabe|todo|task)[:\s]+(.+?)(?:\.|$)/gi,
        extractResponsible: true,
      },
      {
        pattern: /-?\s*(.+?)\s+(?:by|von|durch)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)/gi,
        extractResponsible: true,
      },
    ];

    for (const { pattern, extractResponsible } of patterns) {
      let match;
      while ((match = pattern.exec(content)) !== null) {
        const text = match[1]?.trim();
        if (!text || text.length < 5) continue;

        const actionItem: ActionItem = { text };

        // Extract responsible person
        if (extractResponsible && match[2]) {
          actionItem.responsible = { name: match[2].trim() };
        }

        // Extract deadline
        const deadlines = metadataExtractor.extractDeadlines(match[0]);
        if (deadlines.length > 0) {
          actionItem.deadline = deadlines[0];
        }

        // Extract priority
        const priorities = metadataExtractor.extractPriorities(match[0]);
        if (priorities.length > 0) {
          actionItem.priority = priorities[0] as 'high' | 'medium' | 'low';
        }

        // Extract status
        const statuses = metadataExtractor.extractStatuses(match[0]);
        if (statuses.length > 0) {
          actionItem.status = statuses[0] as 'pending' | 'in-progress' | 'completed';
        }

        actionItems.push(actionItem);
      }
    }

    return actionItems;
  }

  /**
   * Generate decisions from content
   */
  generateDecisions(content: string): Decision[] {
    const decisions: Decision[] = [];
    const dateParser = new DateParser();

    const patterns = [
      /(?:decided|decision|beschluss|entscheidung)[:\s]+(.+?)(?:\.|$)/gi,
      /(?:agreed|vereinbart|einigung)[:\s]+(.+?)(?:\.|$)/gi,
      /(?:we\s+will|wir\s+werden)[:\s]+(.+?)(?:\.|$)/gi,
    ];

    for (const pattern of patterns) {
      let match;
      while ((match = pattern.exec(content)) !== null) {
        const text = match[1]?.trim();
        if (!text || text.length < 5) continue;

        const decision: Decision = { text };

        // Try to extract decision maker
        const personMatch = match[0].match(/(?:by|von|durch)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)/);
        if (personMatch) {
          decision.decisionMaker = { name: personMatch[1].trim() };
        }

        // Extract date
        const dates = match[0].match(/\d{1,2}[./]\d{1,2}[./]\d{2,4}/);
        if (dates) {
          const date = dateParser.parse(dates[0]);
          if (date) {
            decision.date = date;
          }
        }

        decisions.push(decision);
      }
    }

    return decisions;
  }

  /**
   * Generate key points from content
   */
  generateKeyPoints(content: string, maxPoints: number = 10): string[] {
    const keyPoints: string[] = [];
    const sentences = content.split(/[.!?]+/).filter((s) => s.trim().length > 20);

    const importantKeywords = [
      'important',
      'wichtig',
      'key',
      'wichtigste',
      'critical',
      'kritisch',
      'summary',
      'zusammenfassung',
      'conclusion',
      'schlussfolgerung',
    ];

    for (const sentence of sentences) {
      const lower = sentence.toLowerCase();
      if (importantKeywords.some((keyword) => lower.includes(keyword))) {
        const trimmed = sentence.trim();
        if (trimmed.length > 20 && trimmed.length < 200) {
          keyPoints.push(trimmed);
          if (keyPoints.length >= maxPoints) break;
        }
      }
    }

    return keyPoints;
  }
}
