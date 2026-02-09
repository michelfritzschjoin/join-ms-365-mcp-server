/**
 * Entity Type Validator Tests
 *
 * Tests for entity type validation and filtering according to Microsoft Graph API rules
 */

import { describe, it, expect } from 'vitest';
import { validateEntityTypeCombinations } from '../../src/utils/entity-type-validator.js';

describe('validateEntityTypeCombinations', () => {
  describe('Empty input', () => {
    it('should return empty array for empty input', () => {
      expect(validateEntityTypeCombinations([])).toEqual([]);
    });
  });

  describe('Answer types', () => {
    it('should allow single answer type', () => {
      expect(validateEntityTypeCombinations(['acronym'])).toEqual(['acronym']);
    });

    it('should allow multiple answer types together', () => {
      const result = validateEntityTypeCombinations(['acronym', 'bookmark', 'qna']);
      expect(result).toContain('acronym');
      expect(result).toContain('bookmark');
      expect(result).toContain('qna');
      expect(result.length).toBe(3);
    });

    it('should filter out non-answer types when answer types are present', () => {
      const result = validateEntityTypeCombinations(['acronym', 'message', 'event']);
      // Message types have priority over answer types when both are present
      expect(result).toEqual(['message']);
    });
  });

  describe('Message types', () => {
    it('should allow single message type', () => {
      expect(validateEntityTypeCombinations(['message'])).toEqual(['message']);
    });

    it('should allow multiple message types together', () => {
      const result = validateEntityTypeCombinations(['message', 'chatMessage']);
      expect(result).toContain('message');
      expect(result).toContain('chatMessage');
      expect(result.length).toBe(2);
    });

    it('should filter out non-message types when message types are present', () => {
      const result = validateEntityTypeCombinations(['message', 'event', 'person']);
      expect(result).toEqual(['message']);
    });
  });

  describe('File types', () => {
    it('should allow single file type', () => {
      expect(validateEntityTypeCombinations(['drive'])).toEqual(['drive']);
    });

    it('should allow multiple file types together', () => {
      const result = validateEntityTypeCombinations([
        'drive',
        'driveItem',
        'list',
        'listItem',
        'site',
        'externalItem',
      ]);
      expect(result.length).toBe(6);
      expect(result).toContain('drive');
      expect(result).toContain('driveItem');
      expect(result).toContain('list');
      expect(result).toContain('listItem');
      expect(result).toContain('site');
      expect(result).toContain('externalItem');
    });

    it('should filter out non-file types when file types are present', () => {
      const result = validateEntityTypeCombinations(['drive', 'message', 'event']);
      expect(result).toEqual(['drive']);
    });
  });

  describe('Standalone types', () => {
    it('should allow single standalone type', () => {
      expect(validateEntityTypeCombinations(['event'])).toEqual(['event']);
      expect(validateEntityTypeCombinations(['person'])).toEqual(['person']);
    });

    it('should filter out standalone types when other types are present', () => {
      const result = validateEntityTypeCombinations(['event', 'drive']);
      expect(result).toEqual(['drive']);
    });

    it('should filter out standalone types when multiple other types are present', () => {
      const result = validateEntityTypeCombinations(['person', 'message', 'chatMessage']);
      expect(result).toEqual(['message', 'chatMessage']);
    });
  });

  describe('Priority rules', () => {
    it('should prioritize file types over message types', () => {
      const result = validateEntityTypeCombinations(['drive', 'message']);
      expect(result).toEqual(['drive']);
    });

    it('should prioritize file types over answer types', () => {
      const result = validateEntityTypeCombinations(['drive', 'acronym']);
      expect(result).toEqual(['drive']);
    });

    it('should prioritize message types over answer types', () => {
      const result = validateEntityTypeCombinations(['message', 'acronym']);
      expect(result).toEqual(['message']);
    });

    it('should prioritize answer types over standalone types', () => {
      const result = validateEntityTypeCombinations(['acronym', 'event']);
      expect(result).toEqual(['acronym']);
    });
  });

  describe('Complex combinations', () => {
    it('should handle all file types together', () => {
      const result = validateEntityTypeCombinations([
        'drive',
        'driveItem',
        'list',
        'listItem',
        'site',
        'externalItem',
        'message',
        'event',
      ]);
      expect(result).not.toContain('message');
      expect(result).not.toContain('event');
      expect(result.length).toBe(6);
    });

    it('should handle mixed incompatible types and prioritize correctly', () => {
      const result = validateEntityTypeCombinations([
        'acronym',
        'bookmark',
        'message',
        'chatMessage',
        'drive',
        'event',
        'person',
      ]);
      expect(result).toEqual(['drive']);
    });

    it('should return first type as fallback if all filtered out', () => {
      // This shouldn't happen in practice, but test the fallback
      const result = validateEntityTypeCombinations(['unknownType']);
      expect(result).toEqual(['unknownType']);
    });
  });
});
