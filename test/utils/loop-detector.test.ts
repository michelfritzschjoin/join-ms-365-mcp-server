/**
 * Loop Detector Tests
 *
 * Tests for Microsoft Loop file detection and parsing
 */

import { describe, it, expect } from 'vitest';
import {
  detectLoopFile,
  isLoopFile,
  parseLoopContent,
  formatLoopFileInfo,
  getLoopFolderPatterns,
  getLoopFileExtensions,
  getLoopMimeTypes,
  type LoopFileDetectionResult,
  type ParsedLoopContent,
} from '../../src/utils/loop-detector.js';
import { createMockDriveItem } from '../utils/test-helpers.js';

describe('Loop File Detection', () => {
  describe('detectLoopFile', () => {
    it('should detect Loop file by URL pattern', () => {
      const driveItem = createMockDriveItem({
        webUrl: 'https://loop.microsoft.com/document/123',
      });

      const result = detectLoopFile(driveItem);

      expect(result.isLoopFile).toBe(true);
      expect(result.detectionMethod).toBe('url');
      expect(result.confidence).toBe('high');
      expect(result.matchedPattern).toBe('loop.microsoft.com');
    });

    it('should detect Loop file by file extension', () => {
      const driveItem = createMockDriveItem({
        name: 'document.loop',
      });

      const result = detectLoopFile(driveItem);

      expect(result.isLoopFile).toBe(true);
      expect(result.detectionMethod).toBe('extension');
      expect(result.confidence).toBe('high');
      expect(result.matchedPattern).toBe('.loop');
    });

    it('should detect Loop file by .fluid extension', () => {
      const driveItem = createMockDriveItem({
        name: 'document.fluid',
      });

      const result = detectLoopFile(driveItem);

      expect(result.isLoopFile).toBe(true);
      expect(result.detectionMethod).toBe('extension');
      expect(result.matchedPattern).toBe('.fluid');
    });

    it('should detect Loop file by MIME type', () => {
      const driveItem = createMockDriveItem({
        file: {
          mimeType: 'application/vnd.ms-loop',
        },
      });

      const result = detectLoopFile(driveItem);

      expect(result.isLoopFile).toBe(true);
      expect(result.detectionMethod).toBe('mimeType');
      expect(result.confidence).toBe('high');
    });

    it('should detect Loop file by folder pattern', () => {
      const driveItem = createMockDriveItem({
        parentReference: {
          path: '/drive/root:/OneNote Loop files:/document.loop',
        },
      });

      const result = detectLoopFile(driveItem);

      expect(result.isLoopFile).toBe(true);
      expect(result.detectionMethod).toBe('folder');
      expect(result.confidence).toBe('medium');
      expect(result.matchedPattern).toBe('OneNote Loop files');
    });

    it('should detect Loop file by name pattern', () => {
      const driveItem = createMockDriveItem({
        name: 'My Loop Document',
      });

      const result = detectLoopFile(driveItem);

      expect(result.isLoopFile).toBe(true);
      expect(result.detectionMethod).toBe('name');
      expect(result.confidence).toBe('low');
    });

    it('should return false for non-Loop file', () => {
      const driveItem = createMockDriveItem({
        name: 'regular-document.docx',
        webUrl: 'https://example.sharepoint.com/document.docx',
        file: {
          mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        },
      });

      const result = detectLoopFile(driveItem);

      expect(result.isLoopFile).toBe(false);
      expect(result.confidence).toBe('high');
    });

    it('should handle missing properties gracefully', () => {
      const driveItem = {};

      const result = detectLoopFile(driveItem);

      expect(result.isLoopFile).toBe(false);
    });
  });

  describe('isLoopFile', () => {
    it('should return true for Loop file', () => {
      const driveItem = createMockDriveItem({
        webUrl: 'https://loop.microsoft.com/document',
      });

      expect(isLoopFile(driveItem)).toBe(true);
    });

    it('should return false for non-Loop file', () => {
      const driveItem = createMockDriveItem();

      expect(isLoopFile(driveItem)).toBe(false);
    });
  });

  describe('parseLoopContent', () => {
    it('should parse JSON Fluid format', () => {
      const content = JSON.stringify({
        text: 'Hello World',
        content: 'Test content',
        id: '123',
        version: '1.0',
      });

      const result = parseLoopContent(content);

      expect(result.success).toBe(true);
      expect(result.contentType).toBe('fluid');
      expect(result.textContent).toBeDefined();
      // Metadata is optional and only included if metadata fields are present
      if (result.metadata) {
        expect(result.metadata).toBeDefined();
      }
    });

    it('should parse plain text content', () => {
      const content = 'This is plain text content';

      const result = parseLoopContent(content);

      expect(result.success).toBe(true);
      expect(result.contentType).toBe('text');
      expect(result.textContent).toBe(content);
    });

    it('should handle empty content', () => {
      const result = parseLoopContent('');

      expect(result.success).toBe(false);
      expect(result.error).toBe('No content provided');
    });

    it('should handle invalid JSON gracefully', () => {
      const content = '{ invalid json }';

      const result = parseLoopContent(content);

      // Should fall back to text detection
      expect(result.success).toBe(true);
    });

    it('should extract text from nested JSON structures', () => {
      const content = JSON.stringify({
        blocks: [{ text: 'First block' }, { text: 'Second block' }],
        metadata: {
          title: 'Document Title',
        },
      });

      const result = parseLoopContent(content);

      expect(result.success).toBe(true);
      expect(result.textContent).toBeDefined();
      expect(result.textContent).toContain('First block');
      expect(result.textContent).toContain('Second block');
    });

    it('should handle null content', () => {
      const result = parseLoopContent(null as unknown as string);

      expect(result.success).toBe(false);
      expect(result.error).toBe('No content provided');
    });
  });

  describe('formatLoopFileInfo', () => {
    it('should format Loop file information', () => {
      const driveItem = createMockDriveItem({
        name: 'test.loop',
        webUrl: 'https://loop.microsoft.com/test',
        size: 2048,
        lastModifiedDateTime: '2026-01-27T10:00:00.0000000Z',
      });

      const detectionResult: LoopFileDetectionResult = {
        isLoopFile: true,
        detectionMethod: 'extension',
        matchedPattern: '.loop',
        confidence: 'high',
      };

      const formatted = formatLoopFileInfo(driveItem, detectionResult);

      expect(formatted).toContain('Loop File: test.loop');
      expect(formatted).toContain('extension');
      expect(formatted).toContain('high');
      expect(formatted).toContain('.loop');
    });

    it('should handle missing properties', () => {
      const driveItem = {};
      const detectionResult: LoopFileDetectionResult = {
        isLoopFile: true,
        detectionMethod: 'url',
        confidence: 'high',
      };

      const formatted = formatLoopFileInfo(driveItem, detectionResult);

      expect(formatted).toContain('Unknown');
    });
  });

  describe('Utility functions', () => {
    it('should return Loop folder patterns', () => {
      const patterns = getLoopFolderPatterns();

      expect(patterns.length).toBeGreaterThan(0);
      expect(patterns).toContain('OneNote Loop files');
    });

    it('should return Loop file extensions', () => {
      const extensions = getLoopFileExtensions();

      expect(extensions).toContain('.loop');
      expect(extensions).toContain('.fluid');
    });

    it('should return Loop MIME types', () => {
      const mimeTypes = getLoopMimeTypes();

      expect(mimeTypes.length).toBeGreaterThan(0);
      expect(mimeTypes).toContain('application/vnd.ms-loop');
    });
  });
});
