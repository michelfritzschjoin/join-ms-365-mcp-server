/**
 * UTC instants for the start and end of a calendar day in an IANA timezone
 * (e.g. Europe/Berlin), relative to `reference`.
 * Used for Microsoft Graph calendarView ranges so "today" matches the user's local day.
 */

function isValidIanaTimeZone(timeZone: string): boolean {
  try {
    Intl.DateTimeFormat(undefined, { timeZone });
    return true;
  } catch {
    return false;
  }
}

function formatCalendarDayInZone(date: Date, timeZone: string): string {
  return new Intl.DateTimeFormat('en-CA', {
    timeZone,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).format(date);
}

function addCalendarDaysYmd(ymd: string, days: number): string {
  const [y, m, d] = ymd.split('-').map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d + days));
  const yy = dt.getUTCFullYear();
  const mm = String(dt.getUTCMonth() + 1).padStart(2, '0');
  const dd = String(dt.getUTCDate()).padStart(2, '0');
  return `${yy}-${mm}-${dd}`;
}

/**
 * First UTC instant where the calendar date in `timeZone` equals `targetKey` (YYYY-MM-DD).
 */
function findMinInstantForCalendarDay(targetKey: string, timeZone: string, center: Date): Date {
  let lo = center.getTime() - 7 * 24 * 3600000;
  let hi = center.getTime() + 24 * 3600000;
  let guard = 0;
  while (formatCalendarDayInZone(new Date(lo), timeZone) >= targetKey && guard++ < 500) {
    lo -= 6 * 3600000;
  }
  guard = 0;
  while (formatCalendarDayInZone(new Date(hi), timeZone) < targetKey && guard++ < 500) {
    hi += 6 * 3600000;
  }
  let left = lo;
  let right = hi;
  while (right - left > 1) {
    const mid = Math.floor((left + right) / 2);
    const key = formatCalendarDayInZone(new Date(mid), timeZone);
    if (key < targetKey) {
      left = mid;
    } else {
      right = mid;
    }
  }
  return new Date(right);
}

/**
 * Returns [start, end] UTC Dates for the calendar day that contains `reference` in `timeZone`.
 * End is the last millisecond before the next calendar day in that zone.
 * For "heute" / today, pass the MCP server's current instant (`new Date()` in the server process).
 */
export function getUtcRangeForCalendarDayInTimeZone(
  reference: Date,
  timeZone: string
): { start: Date; end: Date } {
  const raw = timeZone.trim() || 'UTC';
  const tz = isValidIanaTimeZone(raw) ? raw : 'UTC';

  if (tz === 'UTC' || tz === 'Etc/UTC') {
    const n = reference;
    const start = new Date(
      Date.UTC(n.getUTCFullYear(), n.getUTCMonth(), n.getUTCDate(), 0, 0, 0, 0)
    );
    const end = new Date(
      Date.UTC(n.getUTCFullYear(), n.getUTCMonth(), n.getUTCDate(), 23, 59, 59, 999)
    );
    return { start, end };
  }

  const refKey = formatCalendarDayInZone(reference, tz);
  const start = findMinInstantForCalendarDay(refKey, tz, reference);
  const nextDayKey = addCalendarDaysYmd(refKey, 1);
  const nextStart = findMinInstantForCalendarDay(
    nextDayKey,
    tz,
    new Date(start.getTime() + 12 * 3600000)
  );
  const end = new Date(nextStart.getTime() - 1);
  return { start, end };
}
