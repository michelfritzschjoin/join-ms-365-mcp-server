import { describe, expect, it } from 'vitest';
import { calendarResponseToText, mailResponseToText } from '../src/response-formatter.js';

describe('response timezone consistency', () => {
  it('calendar text uses local-time wording instead of UTC-only wording', () => {
    const text = calendarResponseToText({
      summary: { totalEvents: 0, dateRange: '27.03.2026', timezone: 'Europe/Berlin' },
      events: [],
      groupedByDate: {},
    });

    expect(text).toContain('lokale Serverzeit');
    expect(text).not.toContain('Alle Zeiten sind UTC');
  });

  it('mail text uses local-time wording instead of UTC-only wording', () => {
    const text = mailResponseToText({
      summary: {
        totalMessages: 0,
        unreadCount: 0,
        dateRange: '27.03.2026',
        timezone: 'Europe/Berlin',
      },
      messages: [],
      groupedByDate: {},
    });

    expect(text).toContain('lokale Serverzeit');
    expect(text).not.toContain('Alle Zeiten sind UTC');
  });
});
