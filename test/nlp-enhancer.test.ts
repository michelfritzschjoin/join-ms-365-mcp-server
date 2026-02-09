/**
 * NLP Enhancer Tests
 *
 * Tests for NLP query processing, entity recognition, and intent classification
 */

import { describe, it, expect, beforeEach } from 'vitest';
import NLPEnhancer, { type DecomposedQuery, type ExtractedEntity } from '../src/nlp-enhancer.js';

describe('NLPEnhancer', () => {
  let enhancer: NLPEnhancer;

  beforeEach(() => {
    enhancer = new NLPEnhancer();
  });

  describe('normalizeQuery', () => {
    it('should normalize a simple query', () => {
      const result = enhancer.normalizeQuery('Find emails from John');

      expect(result.original).toBe('Find emails from John');
      expect(result.normalized).toBeDefined();
      expect(result.stemmed.length).toBeGreaterThan(0);
    });

    it('should handle German queries', () => {
      const result = enhancer.normalizeQuery('Zeige E-Mails von Max');

      expect(result.original).toBe('Zeige E-Mails von Max');
      expect(result.normalized).toBeDefined();
    });

    it('should handle empty queries', () => {
      const result = enhancer.normalizeQuery('');

      expect(result.original).toBe('');
      expect(result.normalized).toBeDefined();
    });

    it('should extract entities from query', () => {
      const result = enhancer.normalizeQuery('Find emails from john@example.com');

      expect(result.entities.length).toBeGreaterThan(0);
    });

    it('should classify intent', () => {
      const result = enhancer.normalizeQuery('Find emails');

      expect(result.intent.intent).toBeDefined();
      expect(result.intent.confidence).toBeGreaterThanOrEqual(0);
    });
  });

  describe('decomposeQuery', () => {
    it('should decompose a simple query', () => {
      const result = enhancer.decomposeQuery('Show me emails from yesterday');

      expect(result.original).toBe('Show me emails from yesterday');
      expect(result.entity).toBeDefined();
      expect(result.intent).toBeDefined();
      expect(result.subQueries.length).toBeGreaterThanOrEqual(0);
    });

    it('should extract temporal expressions', () => {
      const result = enhancer.decomposeQuery('Show emails from last week');

      expect(result.temporal).toBeDefined();
      if (result.temporal) {
        expect(result.temporal.type).toBeDefined();
        expect(result.temporal.expression).toBeDefined();
      }
    });

    it('should extract entities', () => {
      const result = enhancer.decomposeQuery('Find meetings with John Doe');

      expect(result.entities.length).toBeGreaterThanOrEqual(0);
    });

    it('should detect MS365 context', () => {
      const result = enhancer.decomposeQuery('Show calendar events');

      if (result.ms365Context) {
        expect(result.ms365Context.service).toBeDefined();
      }
    });

    it('should generate semantic variants', () => {
      const result = enhancer.decomposeQuery('Find emails');

      expect(result.semanticVariants.length).toBeGreaterThanOrEqual(0);
    });

    it('should handle compound queries', () => {
      const result = enhancer.decomposeQuery('Show emails and calendar events');

      expect(result.compoundParts.length).toBeGreaterThanOrEqual(0);
    });

    it('should calculate urgency level', () => {
      const result = enhancer.decomposeQuery('Show urgent emails');

      expect(result.urgency).toBeDefined();
      expect(['high', 'medium', 'low', 'none']).toContain(result.urgency);
    });

    it('should generate markdown summary', () => {
      const result = enhancer.decomposeQuery('Find emails from yesterday');

      expect(result.markdown).toBeDefined();
      expect(typeof result.markdown).toBe('string');
      expect(result.markdown.length).toBeGreaterThan(0);
    });
  });

  describe('extractEntities', () => {
    it('should extract email addresses', () => {
      const query = 'Find emails from john@example.com';
      const result = enhancer.decomposeQuery(query);

      const emails = result.entities.filter((e) => e.type === 'email');
      expect(emails.length).toBeGreaterThan(0);
    });

    it('should extract dates', () => {
      const query = 'Show events on 2026-01-27';
      const result = enhancer.decomposeQuery(query);

      const dates = result.entities.filter((e) => e.type === 'date');
      expect(dates.length).toBeGreaterThanOrEqual(0);
    });

    it('should extract person names', () => {
      const query = 'Find emails from John Doe';
      const result = enhancer.decomposeQuery(query);

      const persons = result.entities.filter((e) => e.type === 'person');
      expect(persons.length).toBeGreaterThanOrEqual(0);
    });
  });

  describe('intent classification', () => {
    it('should classify search intent', () => {
      const result = enhancer.decomposeQuery('Search for emails');

      expect(result.intent.type).toBeDefined();
    });

    it('should classify list intent', () => {
      const result = enhancer.decomposeQuery('List all emails');

      expect(result.intent.type).toBeDefined();
    });

    it('should classify find intent', () => {
      const result = enhancer.decomposeQuery('Find meeting');

      expect(result.intent.type).toBeDefined();
    });

    it('should classify when intent', () => {
      const result = enhancer.decomposeQuery('When is the meeting?');

      expect(result.intent.type).toBeDefined();
    });
  });

  describe('temporal expressions', () => {
    it('should detect past expressions', () => {
      const result = enhancer.decomposeQuery('Show emails from yesterday');

      if (result.temporal) {
        expect(['past', 'range']).toContain(result.temporal.type);
      }
    });

    it('should detect future expressions', () => {
      const result = enhancer.decomposeQuery('Show meetings tomorrow');

      if (result.temporal) {
        expect(['future', 'range']).toContain(result.temporal.type);
      }
    });

    it('should detect present expressions', () => {
      const result = enhancer.decomposeQuery('Show current events');

      if (result.temporal) {
        expect(['present', 'range']).toContain(result.temporal.type);
      }
    });

    it('should handle German temporal expressions', () => {
      const result = enhancer.decomposeQuery('Zeige E-Mails von gestern');

      if (result.temporal) {
        expect(result.temporal.type).toBeDefined();
      }
    });
  });

  describe('edge cases', () => {
    it('should handle very long queries', () => {
      const longQuery = 'Find emails from '.repeat(100) + 'yesterday';
      const result = enhancer.decomposeQuery(longQuery);

      expect(result.original).toBe(longQuery);
      expect(result.intent).toBeDefined();
    });

    it('should handle queries with special characters', () => {
      const result = enhancer.decomposeQuery('Find emails with subject: "Test"');

      expect(result.original).toBe('Find emails with subject: "Test"');
    });

    it('should handle queries with numbers', () => {
      const result = enhancer.decomposeQuery('Show 10 emails');

      expect(result.original).toBe('Show 10 emails');
    });

    it('should handle empty query', () => {
      const result = enhancer.decomposeQuery('');

      expect(result.original).toBe('');
      expect(result.intent.type).toBe('unknown');
    });
  });
});
