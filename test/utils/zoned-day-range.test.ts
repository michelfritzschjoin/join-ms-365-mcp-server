import { describe, it, expect } from 'vitest';
import { getUtcRangeForCalendarDayInTimeZone } from '../../src/utils/zoned-day-range.js';

describe('getUtcRangeForCalendarDayInTimeZone', () => {
  it('uses UTC civil day when timeZone is UTC', () => {
    const ref = new Date('2026-03-27T15:30:00.000Z');
    const { start, end } = getUtcRangeForCalendarDayInTimeZone(ref, 'UTC');
    expect(start.toISOString()).toBe('2026-03-27T00:00:00.000Z');
    expect(end.toISOString()).toBe('2026-03-27T23:59:59.999Z');
  });

  it('covers the Berlin calendar day for a midday UTC instant in March', () => {
    const ref = new Date('2026-03-27T12:00:00.000Z');
    const { start, end } = getUtcRangeForCalendarDayInTimeZone(ref, 'Europe/Berlin');
    const fmt = (d: Date) =>
      new Intl.DateTimeFormat('en-CA', {
        timeZone: 'Europe/Berlin',
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
      }).format(d);
    expect(fmt(start)).toBe('2026-03-27');
    expect(fmt(end)).toBe('2026-03-27');
    expect(end.getTime()).toBeGreaterThan(start.getTime());
  });

  it('falls back to UTC when timeZone is invalid', () => {
    const ref = new Date('2026-03-27T15:30:00.000Z');
    const { start, end } = getUtcRangeForCalendarDayInTimeZone(ref, 'Not/A/Zone');
    expect(start.toISOString()).toBe('2026-03-27T00:00:00.000Z');
    expect(end.toISOString()).toBe('2026-03-27T23:59:59.999Z');
  });
});
