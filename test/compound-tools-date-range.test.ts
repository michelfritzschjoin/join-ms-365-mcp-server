/**
 * Compound Tools Date Range Tests
 *
 * Tests for date range calculation functions, especially for "tomorrow" and other timeframes
 */

import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { getDateRangeFromTimeframe } from '../src/compound-tools.js';
import { formatLocalDate } from '../src/response-formatter.js';

describe('getDateRangeFromTimeframe', () => {
  // Mock current date to 2026-02-09 20:23:00
  const mockDate = new Date('2026-02-09T20:23:00.000Z');

  beforeEach(() => {
    // Use fake timers to control Date.now() and new Date()
    vi.useFakeTimers();
    vi.setSystemTime(mockDate);
  });

  afterEach(() => {
    // Restore real timers
    vi.useRealTimers();
  });

  describe('tomorrow timeframe', () => {
    it('should calculate correct date range for tomorrow', () => {
      const result = getDateRangeFromTimeframe('tomorrow');

      // Tomorrow should be 2026-02-10
      const expectedStart = new Date('2026-02-10T00:00:00.000Z');
      const expectedEnd = new Date('2026-02-10T23:59:59.999Z');

      // Check that start is at beginning of tomorrow (00:00:00)
      expect(result.start.getFullYear()).toBe(2026);
      expect(result.start.getMonth()).toBe(1); // February (0-indexed)
      expect(result.start.getDate()).toBe(10);
      expect(result.start.getHours()).toBe(0);
      expect(result.start.getMinutes()).toBe(0);
      expect(result.start.getSeconds()).toBe(0);
      expect(result.start.getMilliseconds()).toBe(0);

      // Check that end is at end of tomorrow (23:59:59.999)
      expect(result.end.getFullYear()).toBe(2026);
      expect(result.end.getMonth()).toBe(1); // February (0-indexed)
      expect(result.end.getDate()).toBe(10);
      expect(result.end.getHours()).toBe(23);
      expect(result.end.getMinutes()).toBe(59);
      expect(result.end.getSeconds()).toBe(59);
      expect(result.end.getMilliseconds()).toBe(999);
    });

    it('should format tomorrow date correctly as DD.MM.YYYY', () => {
      const result = getDateRangeFromTimeframe('tomorrow');
      const formattedStart = formatLocalDate(result.start);
      const formattedEnd = formatLocalDate(result.end);

      // Both should be 10.02.2026
      expect(formattedStart).toBe('10.02.2026');
      expect(formattedEnd).toBe('10.02.2026');
    });

    it('should handle month boundaries correctly', () => {
      // Test with date at end of month (e.g., January 31)
      const endOfMonthDate = new Date('2026-01-31T20:23:00.000Z');
      vi.setSystemTime(endOfMonthDate);

      const result = getDateRangeFromTimeframe('tomorrow');

      // Tomorrow should be February 1, 2026
      expect(result.start.getFullYear()).toBe(2026);
      expect(result.start.getMonth()).toBe(1); // February (0-indexed)
      expect(result.start.getDate()).toBe(1);
    });

    it('should handle year boundaries correctly', () => {
      // Test with date at end of year (e.g., December 31)
      const endOfYearDate = new Date('2025-12-31T20:23:00.000Z');
      vi.setSystemTime(endOfYearDate);

      const result = getDateRangeFromTimeframe('tomorrow');

      // Tomorrow should be January 1, 2026
      expect(result.start.getFullYear()).toBe(2026);
      expect(result.start.getMonth()).toBe(0); // January (0-indexed)
      expect(result.start.getDate()).toBe(1);
    });
  });

  describe('today timeframe', () => {
    it('should calculate correct date range for today', () => {
      const result = getDateRangeFromTimeframe('today');

      // Today should be 2026-02-09
      expect(result.start.getFullYear()).toBe(2026);
      expect(result.start.getMonth()).toBe(1); // February (0-indexed)
      expect(result.start.getDate()).toBe(9);
      expect(result.start.getHours()).toBe(0);
      expect(result.start.getMinutes()).toBe(0);
      expect(result.start.getSeconds()).toBe(0);

      expect(result.end.getFullYear()).toBe(2026);
      expect(result.end.getMonth()).toBe(1); // February (0-indexed)
      expect(result.end.getDate()).toBe(9);
      expect(result.end.getHours()).toBe(23);
      expect(result.end.getMinutes()).toBe(59);
      expect(result.end.getSeconds()).toBe(59);
      expect(result.end.getMilliseconds()).toBe(999);
    });
  });

  describe('yesterday timeframe', () => {
    it('should calculate correct date range for yesterday', () => {
      const result = getDateRangeFromTimeframe('yesterday');

      // Yesterday should be 2026-02-08
      expect(result.start.getFullYear()).toBe(2026);
      expect(result.start.getMonth()).toBe(1); // February (0-indexed)
      expect(result.start.getDate()).toBe(8);
      expect(result.start.getHours()).toBe(0);
      expect(result.start.getMinutes()).toBe(0);
      expect(result.start.getSeconds()).toBe(0);

      expect(result.end.getFullYear()).toBe(2026);
      expect(result.end.getMonth()).toBe(1); // February (0-indexed)
      expect(result.end.getDate()).toBe(8);
      expect(result.end.getHours()).toBe(23);
      expect(result.end.getMinutes()).toBe(59);
      expect(result.end.getSeconds()).toBe(59);
      expect(result.end.getMilliseconds()).toBe(999);
    });
  });

  describe('date range boundaries', () => {
    it('should ensure start is always before or equal to end', () => {
      const timeframes = ['today', 'tomorrow', 'yesterday', 'thisWeek', 'lastWeek', 'nextWeek'];

      for (const timeframe of timeframes) {
        const result = getDateRangeFromTimeframe(timeframe);
        expect(result.start.getTime()).toBeLessThanOrEqual(result.end.getTime());
      }
    });

    it('should ensure start is at beginning of day (00:00:00)', () => {
      const result = getDateRangeFromTimeframe('tomorrow');
      expect(result.start.getHours()).toBe(0);
      expect(result.start.getMinutes()).toBe(0);
      expect(result.start.getSeconds()).toBe(0);
      expect(result.start.getMilliseconds()).toBe(0);
    });

    it('should ensure end is at end of day (23:59:59.999)', () => {
      const result = getDateRangeFromTimeframe('tomorrow');
      expect(result.end.getHours()).toBe(23);
      expect(result.end.getMinutes()).toBe(59);
      expect(result.end.getSeconds()).toBe(59);
      expect(result.end.getMilliseconds()).toBe(999);
    });
  });
});
