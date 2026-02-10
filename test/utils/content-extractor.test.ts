/**
 * Tests for Universal Business Content Extractor
 */

import { describe, it, expect, beforeEach } from 'vitest';
import {
  ProjectExtractor,
  CustomerExtractor,
  MeetingExtractor,
  DocumentExtractor,
  SalesExtractor,
  HRExtractor,
  ProjectEntityExtractor,
  CustomerEntityExtractor,
  MeetingEntityExtractor,
  DocumentEntityExtractor,
  SalesEntityExtractor,
  HREntityExtractor,
  EntityExtractorRegistry,
  MetadataExtractor,
  SummaryGenerator,
  DateParser,
  AmountParser,
  TableListParser,
  getPatternBasedExtractor,
  getExtractorRegistry,
  getExtractionCache,
  generateCacheKey,
} from '../../src/utils/content-extractor.js';

describe('Content Extractor', () => {
  describe('ProjectExtractor', () => {
    it('should extract roadmap information', () => {
      const extractor = new ProjectExtractor();
      const content = `
        Project Roadmap Q1 2024
        
        Milestone 1: Initial Setup
        Date: 15.01.2024
        Status: Completed
        Responsible: Max Mustermann
        
        Milestone 2: Development Phase
        Date: 01.02.2024
        Status: In Progress
        Responsible: Anna Schmidt
        
        Project Code: P2046
        Progress: 45%
      `;

      const result = extractor.extract(content);
      expect(result.type).toBe('roadmap');
      expect(result.confidence).toBeGreaterThan(0);
      expect(result.content).toBeDefined();
    });

    it('should extract project codes', () => {
      const extractor = new ProjectExtractor();
      const content = 'Working on project P2046 and PROJ-123';
      const result = extractor.extract(content);
      expect(result.type).toBe('project_plan');
    });
  });

  describe('CustomerExtractor', () => {
    it('should extract customer information', () => {
      const extractor = new CustomerExtractor();
      const content = `
        Customer: Acme Corporation GmbH
        Customer ID: CUST-456
        Offer Number: OFFER-123
      `;

      const result = extractor.extract(content);
      expect(result.type).toBe('customer_info');
      expect(result.confidence).toBeGreaterThan(0);
    });

    it('should extract contract information', () => {
      const extractor = new CustomerExtractor();
      const content = 'Contract VERTRAG-456 signed on 01.01.2024';
      const result = extractor.extract(content);
      expect(result.type).toBe('contract');
    });
  });

  describe('MeetingExtractor', () => {
    it('should extract meeting agenda', () => {
      const extractor = new MeetingExtractor();
      const content = `
        Meeting Agenda
        
        TOP 1: Project Status
        TOP 2: Budget Review
        TOP 3: Next Steps
        
        Action Items:
        - Review proposal by John Doe (deadline: 15.01.2024)
        - Update documentation by Jane Smith
        
        Decisions:
        - We decided to proceed with Option A
        - Agreed on budget increase of 10%
      `;

      const result = extractor.extract(content);
      expect(result.type).toBe('meeting_agenda');
      expect(result.confidence).toBeGreaterThan(0);
    });

    it('should extract action items', () => {
      const extractor = new MeetingExtractor();
      const content = 'Action Item: Complete the report by tomorrow';
      const result = extractor.extract(content);
      expect(result.type).toBe('meeting_notes');
    });
  });

  describe('DocumentExtractor', () => {
    it('should extract invoice information', () => {
      const extractor = new DocumentExtractor();
      const content = `
        Invoice Number: INV-12345
        Amount: 1,234.56 €
        Due Date: 31.01.2024
      `;

      const result = extractor.extract(content);
      expect(result.type).toBe('invoice');
      expect(result.confidence).toBeGreaterThan(0);
    });

    it('should extract report information', () => {
      const extractor = new DocumentExtractor();
      const content = 'Report REPORT-789: Executive Summary - Q4 results show strong growth';
      const result = extractor.extract(content);
      expect(result.type).toBe('report');
    });
  });

  describe('SalesExtractor', () => {
    it('should extract offer information', () => {
      const extractor = new SalesExtractor();
      const content = `
        Offer Number: OFFER-456
        Price: 50,000.00 €
        Valid until: 31.03.2024
      `;

      const result = extractor.extract(content);
      expect(result.type).toBe('offer');
      expect(result.confidence).toBeGreaterThan(0);
    });

    it('should extract budget information', () => {
      const extractor = new SalesExtractor();
      const content = 'Budget BUDGET-789 allocated: 100,000 €';
      const result = extractor.extract(content);
      expect(result.type).toBe('budget');
    });
  });

  describe('HRExtractor', () => {
    it('should extract onboarding information', () => {
      const extractor = new HRExtractor();
      const content = `
        Onboarding Checklist for New Employee:
        - Complete HR forms
        - Setup workstation
        - Attend orientation meeting
        
        Onboarding ID: ONBOARD-123
      `;

      const result = extractor.extract(content);
      expect(result.type).toBe('onboarding');
      expect(result.confidence).toBeGreaterThan(0);
    });

    it('should extract review information', () => {
      const extractor = new HRExtractor();
      const content = 'Performance Review REVIEW-456 scheduled for next week';
      const result = extractor.extract(content);
      expect(result.type).toBe('review');
    });
  });

  describe('Entity Extractors', () => {
    it('ProjectEntityExtractor should extract project codes', () => {
      const extractor = new ProjectEntityExtractor();
      const content = 'Projects: P2046, PROJ-123, AB-456';
      const result = extractor.extract(content);
      expect(result.projectCodes.length).toBeGreaterThan(0);
    });

    it('CustomerEntityExtractor should extract contract numbers', () => {
      const extractor = new CustomerEntityExtractor();
      const content = 'Contracts: CONTRACT-123, VERTRAG-456';
      const result = extractor.extract(content);
      expect(result.contractNumbers.length).toBeGreaterThan(0);
    });

    it('MeetingEntityExtractor should extract action item IDs', () => {
      const extractor = new MeetingEntityExtractor();
      const content = 'Action Items: AI-123, ACTION-456';
      const result = extractor.extract(content);
      expect(result.actionItemIds.length).toBeGreaterThan(0);
    });
  });

  describe('Common Components', () => {
    it('DateParser should parse dates', () => {
      const parser = new DateParser();
      expect(parser.parse('15.01.2024')).toBeInstanceOf(Date);
      expect(parser.parse('2024-01-15')).toBeInstanceOf(Date);
      expect(parser.parse('tomorrow')).toBeInstanceOf(Date);
    });

    it('AmountParser should parse amounts', () => {
      const parser = new AmountParser();
      // Test with simpler format first
      const result1 = parser.parse('1234.56 €');
      expect(result1).toBeDefined();
      expect(result1?.value).toBe(1234.56);
      expect(result1?.currency).toBe('EUR');
      // Test with thousands separator
      const result2 = parser.parse('1,234.56 €');
      expect(result2).toBeDefined();
      // The parser should handle thousands separators
      expect(result2?.value).toBeGreaterThan(0);
      expect(result2?.currency).toBe('EUR');
    });

    it('MetadataExtractor should extract priorities', () => {
      const extractor = new MetadataExtractor();
      const content = 'High priority task needs immediate attention';
      const priorities = extractor.extractPriorities(content);
      expect(priorities).toContain('high');
    });

    it('SummaryGenerator should generate action items', () => {
      const generator = new SummaryGenerator();
      const content = 'Action Item: Complete the report by John Doe (deadline: tomorrow)';
      const actionItems = generator.generateActionItems(content);
      expect(actionItems.length).toBeGreaterThan(0);
    });

    it('TableListParser should parse markdown tables', () => {
      const parser = new TableListParser();
      const markdown = `
        | Name | Date | Status |
        |------|------|--------|
        | Task 1 | 01.01.2024 | Done |
        | Task 2 | 02.01.2024 | Pending |
      `;
      const tables = parser.parseMarkdownTable(markdown);
      expect(tables.length).toBe(1);
      expect(tables[0].headers.length).toBe(3);
      expect(tables[0].rows.length).toBe(2);
    });
  });

  describe('Pattern-Based Extractor', () => {
    it('should detect document type automatically', () => {
      const extractor = getPatternBasedExtractor();
      const content = 'Project Roadmap for Q1 2024 with milestones and timeline';
      const detectedType = extractor.detectDocumentType(content);
      expect(detectedType).toBe('roadmap');
    });

    it('should extract content with automatic detection', () => {
      const extractor = getPatternBasedExtractor();
      const content = 'Meeting Agenda: TOP 1: Status Update, TOP 2: Budget Review';
      const result = extractor.extract(content);
      expect(result.type).toBe('meeting_agenda');
      expect(result.confidence).toBeGreaterThan(0);
    });
  });

  describe('Caching', () => {
    it('should cache extraction results', () => {
      const cache = getExtractionCache();
      const key = generateCacheKey('test-message-id');
      const mockResult = {
        detectedType: 'roadmap' as const,
        confidence: 0.9,
        extracted: {},
        metadata: {},
        entities: {},
      };
      cache.set(key, mockResult);
      const cached = cache.get(key);
      expect(cached).toBeDefined();
      expect(cached?.detectedType).toBe('roadmap');
    });

    it('should return null for expired cache entries', () => {
      const cache = getExtractionCache();
      const key = generateCacheKey('expired-message-id');
      const mockResult = {
        detectedType: 'roadmap' as const,
        confidence: 0.9,
        extracted: {},
        metadata: {},
        entities: {},
      };
      cache.set(key, mockResult);
      // Manually expire by setting old timestamp
      const entry = (
        cache as unknown as { cache: Map<string, { result: unknown; timestamp: number }> }
      ).cache.get(key);
      if (entry) {
        entry.timestamp = Date.now() - 2 * 60 * 60 * 1000; // 2 hours ago
      }
      const cached = cache.get(key);
      expect(cached).toBeNull();
    });
  });

  describe('EntityExtractorRegistry', () => {
    it('should extract all entities from content', () => {
      const registry = new EntityExtractorRegistry();
      const content = `
        Project P2046 with customer CUST-123
        Contract CONTRACT-456
        Action Item AI-789
        Invoice INV-123
        Deal DEAL-ABC
        Candidate CAND-456
      `;
      const entities = registry.extractAll(content);
      expect(entities.projectCodes).toContain('P2046');
      expect(entities.customerIds).toContain('CUST-123');
      expect(entities.contractNumbers).toContain('CONTRACT-456');
      expect(entities.actionItemIds).toContain('AI-789');
      expect(entities.invoiceNumbers).toContain('INV-123');
      expect(entities.dealNames).toContain('DEAL-ABC');
      expect(entities.candidateIds).toContain('CAND-456');
    });
  });
});
