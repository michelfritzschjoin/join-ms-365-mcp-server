/**
 * Response Formatter Tests
 *
 * Tests for response formatting utilities
 */

import { describe, it, expect } from 'vitest';
import {
  convertToLocalTime,
  formatLocalDate,
  formatLocalTime,
  formatISODate,
  formatUTCTime,
  formatUTCDate,
  formatUTCDateTime,
  formatLocalDateTime,
  calculateDuration,
  isCalendarResponse,
  isMailResponse,
} from '../src/response-formatter.js';

describe('Response Formatter', () => {
  describe('convertToLocalTime', () => {
    it('should convert UTC date string to local time', () => {
      const utcString = '2026-01-27T10:00:00.0000000Z';
      const result = convertToLocalTime(utcString);

      expect(result).toBeInstanceOf(Date);
      expect(result.getTime()).toBeDefined();
    });

    it('should handle UTC timezone explicitly', () => {
      const dateString = '2026-01-27T10:00:00.0000000';
      const result = convertToLocalTime(dateString, 'UTC');

      expect(result).toBeInstanceOf(Date);
    });

    it('should handle invalid date strings', () => {
      const result = convertToLocalTime('invalid-date');

      expect(result).toBeInstanceOf(Date);
    });

    it('should handle null/undefined input', () => {
      const result1 = convertToLocalTime(null as unknown as string);
      const result2 = convertToLocalTime(undefined as unknown as string);

      expect(result1).toBeInstanceOf(Date);
      expect(result2).toBeInstanceOf(Date);
    });
  });

  describe('formatLocalDate', () => {
    it('should format date to DD.MM.YYYY', () => {
      const date = new Date('2026-01-27T10:00:00Z');
      const result = formatLocalDate(date);

      expect(result).toMatch(/^\d{2}\.\d{2}\.\d{4}$/);
    });

    it('should handle invalid dates', () => {
      const invalidDate = new Date('invalid');
      const result = formatLocalDate(invalidDate);

      // Should return a valid date format (fallback to current date)
      expect(result).toMatch(/^\d{2}\.\d{2}\.\d{4}$/);
    });
  });

  describe('formatLocalTime', () => {
    it('should format time to HH:MM', () => {
      const date = new Date('2026-01-27T10:30:00Z');
      const result = formatLocalTime(date);

      expect(result).toMatch(/^\d{2}:\d{2}$/);
    });

    it('should handle invalid dates', () => {
      const invalidDate = new Date('invalid');
      const result = formatLocalTime(invalidDate);

      expect(result).toMatch(/^\d{2}:\d{2}$/);
    });
  });

  describe('formatISODate', () => {
    it('should format date to YYYY-MM-DD', () => {
      const date = new Date('2026-01-27T10:00:00Z');
      const result = formatISODate(date);

      expect(result).toMatch(/^\d{4}-\d{2}-\d{2}$/);
      expect(result).toContain('2026-01-27');
    });
  });

  describe('formatUTCTime', () => {
    it('should format UTC time to HH:MM', () => {
      const date = new Date('2026-01-27T10:30:00Z');
      const result = formatUTCTime(date);

      expect(result).toMatch(/^\d{2}:\d{2}$/);
      expect(result).toBe('10:30');
    });
  });

  describe('formatUTCDate', () => {
    it('should format UTC date to DD.MM.YYYY', () => {
      const date = new Date('2026-01-27T10:00:00Z');
      const result = formatUTCDate(date);

      expect(result).toMatch(/^\d{2}\.\d{2}\.\d{4}$/);
    });
  });

  describe('formatUTCDateTime', () => {
    it('should format UTC datetime to ISO string', () => {
      const date = new Date('2026-01-27T10:00:00Z');
      const result = formatUTCDateTime(date);

      expect(result).toContain('2026-01-27');
      expect(result).toContain('T');
      expect(result).toContain('Z');
    });
  });

  describe('formatLocalDateTime', () => {
    it('should format local datetime', () => {
      const date = new Date('2026-01-27T10:30:00Z');
      const result = formatLocalDateTime(date);

      expect(result).toMatch(/^\d{2}\.\d{2}\.\d{4} \d{2}:\d{2}$/);
    });
  });

  describe('calculateDuration', () => {
    it('should calculate duration in minutes', () => {
      const start = new Date('2026-01-27T10:00:00Z');
      const end = new Date('2026-01-27T10:30:00Z');
      const result = calculateDuration(start, end);

      expect(result).toContain('30');
      expect(result).toContain('min');
    });

    it('should calculate duration in hours', () => {
      const start = new Date('2026-01-27T10:00:00Z');
      const end = new Date('2026-01-27T12:00:00Z');
      const result = calculateDuration(start, end);

      expect(result).toContain('2');
      expect(result).toContain('h');
    });

    it('should handle invalid start date', () => {
      const start = new Date('invalid');
      const end = new Date('2026-01-27T10:30:00Z');
      const result = calculateDuration(start, end);

      expect(result).toBe('0 min');
    });

    it('should handle invalid end date', () => {
      const start = new Date('2026-01-27T10:00:00Z');
      const end = new Date('invalid');
      const result = calculateDuration(start, end);

      expect(result).toBe('0 min');
    });

    it('should handle end before start', () => {
      const start = new Date('2026-01-27T10:30:00Z');
      const end = new Date('2026-01-27T10:00:00Z');
      const result = calculateDuration(start, end);

      expect(result).toBe('0 min');
    });
  });

  describe('isCalendarResponse', () => {
    it('should identify calendar response', () => {
      const response = {
        value: [
          {
            id: 'event-1',
            subject: 'Meeting',
            start: { dateTime: '2026-01-27T10:00:00Z' },
            end: { dateTime: '2026-01-27T11:00:00Z' },
            isAllDay: false,
          },
        ],
      };

      expect(isCalendarResponse(response)).toBe(true);
    });

    it('should reject non-calendar response', () => {
      const response = {
        value: [
          {
            id: 'message-1',
            subject: 'Email',
            receivedDateTime: '2026-01-27T10:00:00Z',
          },
        ],
      };

      expect(isCalendarResponse(response)).toBe(false);
    });
  });

  describe('isMailResponse', () => {
    it('should identify mail response', () => {
      const response = {
        value: [
          {
            id: 'message-1',
            subject: 'Email',
            receivedDateTime: '2026-01-27T10:00:00Z',
            from: { emailAddress: { address: 'test@example.com' } },
          },
        ],
      };

      expect(isMailResponse(response)).toBe(true);
    });

    it('should reject non-mail response', () => {
      const response = {
        value: [
          {
            id: 'event-1',
            subject: 'Meeting',
            start: { dateTime: '2026-01-27T10:00:00Z' },
          },
        ],
      };

      expect(isMailResponse(response)).toBe(false);
    });
  });
});
