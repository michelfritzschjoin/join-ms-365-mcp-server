/**
 * Response Formatter - Formats API responses into structured, human-readable output
 * with server local time conversion (no UTC)
 *
 * Supports profession-based personalization:
 * - Executive: High-level bullet summaries
 * - Developer: Technical details with code examples
 * - Sales: Customer-focused narrative
 * - Default: Balanced structured output
 */

import logger from './logger.js';
import type {
  ProfessionProfile,
  DetailLevel,
  LanguageStyle,
  FormatPreference,
} from './user-profile.js';
import { getProfessionProfile } from './request-context.js';

/**
 * Structured calendar event interface
 */
export interface FormattedCalendarEvent {
  id: string;
  subject: string;
  // Server local time (primary display)
  startDate: string;
  startTime: string;
  endDate: string;
  endTime: string;
  // UTC time (for reference)
  startDateTimeUTC: string;
  endDateTimeUTC: string;
  // Combined display string: "10:30 (UTC: 09:30)"
  startTimeDisplay: string;
  endTimeDisplay: string;
  duration: string;
  isAllDay: boolean;
  location?: string;
  organizer?: string;
  organizerEmail?: string;
  attendees?: string[];
  status?: string;
  importance?: string;
  isOnlineMeeting: boolean;
  onlineMeetingUrl?: string;
  webLink?: string;
  bodyPreview?: string;
  isCancelled: boolean;
  showAs?: string;
  categories?: string[];
}

/**
 * Structured calendar response
 */
export interface FormattedCalendarResponse {
  summary: {
    totalEvents: number;
    dateRange: string;
    timezone: string;
  };
  events: FormattedCalendarEvent[];
  groupedByDate: Record<string, FormattedCalendarEvent[]>;
}

/**
 * Structured mail message interface
 */
export interface FormattedMailMessage {
  id: string;
  subject: string;
  // Server local time (primary display)
  receivedDate: string;
  receivedTime: string;
  sentDate: string;
  sentTime: string;
  // UTC time (for reference)
  receivedDateTimeUTC: string;
  sentDateTimeUTC: string;
  // Combined display string: "10:30 (UTC: 09:30)"
  receivedTimeDisplay: string;
  sentTimeDisplay: string;
  from: {
    name: string;
    email: string;
  };
  to: Array<{
    name: string;
    email: string;
  }>;
  cc?: Array<{
    name: string;
    email: string;
  }>;
  hasAttachments: boolean;
  importance: string;
  isRead: boolean;
  isDraft: boolean;
  bodyPreview?: string;
  webLink?: string;
  categories?: string[];
  flag?: {
    flagStatus: string;
  };
  conversationId?: string;
}

/**
 * Structured mail response
 */
export interface FormattedMailResponse {
  summary: {
    totalMessages: number;
    unreadCount: number;
    dateRange: string;
    timezone: string;
  };
  messages: FormattedMailMessage[];
  groupedByDate: Record<string, FormattedMailMessage[]>;
}

/**
 * Convert date string to server local time, handling timezone from Graph API
 * @param dateTimeString - The date/time string from Graph API (e.g., "2026-01-27T14:00:00.0000000")
 * @param timeZone - The timezone from Graph API (e.g., "UTC", "Europe/Berlin")
 * @returns Date object representing the time in server local timezone
 */
export function convertToLocalTime(dateTimeString: string, timeZone?: string): Date {
  // Validate input
  if (!dateTimeString || typeof dateTimeString !== 'string') {
    logger.warn(`Invalid dateTimeString provided: ${dateTimeString}`);
    return new Date();
  }

  // If no timezone specified, check if string has Z suffix (UTC indicator)
  if (!timeZone) {
    if (dateTimeString.endsWith('Z')) {
      // UTC time - JavaScript Date will parse correctly and display in local timezone
      return new Date(dateTimeString);
    }
    // No timezone info - parse as-is (will be interpreted as local time)
    return new Date(dateTimeString);
  }

  // Normalize UTC timezone identifiers
  const normalizedTimeZone = timeZone.trim();
  const isUTC =
    normalizedTimeZone === 'UTC' ||
    normalizedTimeZone === 'Etc/UTC' ||
    normalizedTimeZone === 'Etc/GMT' ||
    normalizedTimeZone === 'GMT';

  // If timezone is UTC, ensure proper UTC parsing - THIS FIXES THE UTC DISPLAY ISSUE
  if (isUTC) {
    // Ensure Z suffix for proper UTC parsing
    // This is critical: without Z, JavaScript interprets the time as local time
    const utcDateString = dateTimeString.endsWith('Z') ? dateTimeString : dateTimeString + 'Z';
    // JavaScript Date parses UTC and automatically converts to local timezone when displayed
    // The Date object stores time in UTC internally but displays in local timezone
    return new Date(utcDateString);
  }

  // For other timezones, convert from the specified timezone to local timezone
  // Graph API returns times like: "2026-01-27T14:00:00.0000000" with timeZone: "Europe/Berlin"
  // The time string represents a time in that timezone, not UTC
  try {
    // Remove any existing timezone suffix and normalize the date string
    const cleanDateTime = dateTimeString
      .replace(/Z$/, '')
      .replace(/[+-]\d{2}:\d{2}$/, '')
      .trim();

    // Parse the date components with improved regex
    const dateMatch = cleanDateTime.match(
      /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(?:\.(\d+))?$/
    );

    if (!dateMatch) {
      // Fallback to standard parsing if format doesn't match expected pattern
      logger.warn(`Unexpected date format: ${dateTimeString}, attempting standard parse`);
      return new Date(dateTimeString);
    }

    const [, year, month, day, hour, minute, second, milliseconds] = dateMatch;

    // Validate parsed components
    const yearNum = parseInt(year, 10);
    const monthNum = parseInt(month, 10);
    const dayNum = parseInt(day, 10);
    const hourNum = parseInt(hour, 10);
    const minuteNum = parseInt(minute, 10);
    const secondNum = parseInt(second, 10);

    if (
      isNaN(yearNum) ||
      isNaN(monthNum) ||
      isNaN(dayNum) ||
      isNaN(hourNum) ||
      isNaN(minuteNum) ||
      isNaN(secondNum) ||
      monthNum < 1 ||
      monthNum > 12 ||
      dayNum < 1 ||
      dayNum > 31 ||
      hourNum > 23 ||
      minuteNum > 59 ||
      secondNum > 59
    ) {
      logger.warn(`Invalid date components in: ${dateTimeString}`);
      return new Date(dateTimeString);
    }

    // Use a more reliable method: construct a date string that represents the time
    // in the specified timezone, then convert to a Date object
    // Strategy: Create a date string in ISO format, then use Intl API to convert

    // Create a date string in ISO format (without timezone)
    const isoString = `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}T${hour.padStart(2, '0')}:${minute.padStart(2, '0')}:${second.padStart(2, '0')}${milliseconds ? `.${milliseconds.padEnd(3, '0').substring(0, 3)}` : ''}`;

    // Use a more direct approach: create a date assuming it's in the target timezone
    // and calculate the UTC equivalent
    // We'll use the fact that we can format a UTC date in the target timezone
    // and work backwards

    // Create a test UTC date
    const testUTC = new Date(
      Date.UTC(yearNum, monthNum - 1, dayNum, hourNum, minuteNum, secondNum)
    );

    // Format this UTC date in the event timezone to see what time it represents there
    const formatter = new Intl.DateTimeFormat('en', {
      timeZone: normalizedTimeZone,
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: false,
    });

    const parts = formatter.formatToParts(testUTC);
    const getPart = (type: string): string => {
      const part = parts.find((p) => p.type === type);
      return part?.value || '0';
    };

    const eventYear = parseInt(getPart('year'), 10);
    const eventMonth = parseInt(getPart('month'), 10);
    const eventDay = parseInt(getPart('day'), 10);
    const eventHour = parseInt(getPart('hour'), 10);
    const eventMinute = parseInt(getPart('minute'), 10);
    const eventSecond = parseInt(getPart('second'), 10);

    // Calculate the difference between what we want and what we got
    const targetTotalMinutes = hourNum * 60 + minuteNum;
    const actualTotalMinutes = eventHour * 60 + eventMinute;

    // Calculate time difference in minutes
    let diffMinutes = targetTotalMinutes - actualTotalMinutes;

    // Handle day rollover (if the timezone conversion changed the day)
    if (eventDay !== dayNum || eventMonth !== monthNum || eventYear !== yearNum) {
      // The timezone offset caused a day change, adjust accordingly
      const dayDiff = dayNum - eventDay;
      diffMinutes += dayDiff * 24 * 60;
    }

    // Adjust the UTC time by the difference
    // If the event timezone shows an earlier time, we need to go back in UTC
    const adjustedUTC = new Date(testUTC.getTime() - diffMinutes * 60 * 1000);

    // Verify the conversion by formatting back
    const verifyFormatter = new Intl.DateTimeFormat('en', {
      timeZone: normalizedTimeZone,
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: false,
    });

    const verifyParts = verifyFormatter.formatToParts(adjustedUTC);
    const verifyHour = parseInt(verifyParts.find((p) => p.type === 'hour')?.value || '0', 10);
    const verifyMinute = parseInt(verifyParts.find((p) => p.type === 'minute')?.value || '0', 10);

    // If verification fails, log warning but return the adjusted date anyway
    if (verifyHour !== hourNum || verifyMinute !== minuteNum) {
      logger.warn(
        `Timezone conversion verification failed for ${dateTimeString} in ${normalizedTimeZone}. Expected ${hourNum}:${minuteNum}, got ${verifyHour}:${verifyMinute}`
      );
    }

    return adjustedUTC;
  } catch (error) {
    // Fallback: if timezone conversion fails, try parsing as UTC
    logger.warn(`Failed to convert timezone ${timeZone} for ${dateTimeString}: ${error}`);
    // Try to parse as UTC if it looks like a date format
    if (dateTimeString.match(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/)) {
      return new Date(dateTimeString.endsWith('Z') ? dateTimeString : dateTimeString + 'Z');
    }
    return new Date(dateTimeString);
  }
}

/**
 * @deprecated Use convertToLocalTime instead
 * Convert UTC date string to server local time
 */
export function utcToLocalTime(utcDateString: string): Date {
  // Assume UTC if no timezone specified and no Z suffix
  if (
    !utcDateString.endsWith('Z') &&
    !utcDateString.includes('+') &&
    !utcDateString.includes('-', 10)
  ) {
    return new Date(utcDateString + 'Z');
  }
  return new Date(utcDateString);
}

/**
 * Format date to local date string (DD.MM.YYYY)
 * Optimized for performance with input validation
 */
export function formatLocalDate(date: Date): string {
  // Validate input
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    logger.warn(`Invalid date provided to formatLocalDate: ${date}`);
    const now = new Date();
    return formatLocalDate(now);
  }

  const day = date.getDate().toString().padStart(2, '0');
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const year = date.getFullYear();
  return `${day}.${month}.${year}`;
}

/**
 * Format date to ISO date string (YYYY-MM-DD)
 */
export function formatISODate(date: Date): string {
  const day = date.getDate().toString().padStart(2, '0');
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const year = date.getFullYear();
  return `${year}-${month}-${day}`;
}

/**
 * Format time to local time string (HH:MM)
 * Optimized for performance with input validation
 */
export function formatLocalTime(date: Date): string {
  // Validate input
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    logger.warn(`Invalid date provided to formatLocalTime: ${date}`);
    const now = new Date();
    return formatLocalTime(now);
  }

  const hours = date.getHours().toString().padStart(2, '0');
  const minutes = date.getMinutes().toString().padStart(2, '0');
  return `${hours}:${minutes}`;
}

/**
 * Format time to UTC time string (HH:MM)
 */
export function formatUTCTime(date: Date): string {
  const hours = date.getUTCHours().toString().padStart(2, '0');
  const minutes = date.getUTCMinutes().toString().padStart(2, '0');
  return `${hours}:${minutes}`;
}

/**
 * Format date to UTC date string (DD.MM.YYYY)
 */
export function formatUTCDate(date: Date): string {
  const day = date.getUTCDate().toString().padStart(2, '0');
  const month = (date.getUTCMonth() + 1).toString().padStart(2, '0');
  const year = date.getUTCFullYear();
  return `${day}.${month}.${year}`;
}

/**
 * Format date and time to UTC datetime string (ISO format)
 */
export function formatUTCDateTime(date: Date): string {
  return date.toISOString();
}

/**
 * Format time with both local and UTC display
 * Example: "10:30 (UTC: 09:30)"
 */
export function formatTimeWithUTC(date: Date): string {
  const localTime = formatLocalTime(date);
  const utcTime = formatUTCTime(date);
  if (localTime === utcTime) {
    return localTime;
  }
  return `${localTime} (UTC: ${utcTime})`;
}

/**
 * Format date and time to local datetime string
 */
export function formatLocalDateTime(date: Date): string {
  return `${formatLocalDate(date)} ${formatLocalTime(date)}`;
}

/**
 * Calculate duration between two dates
 * Returns human-readable duration string with validation
 */
export function calculateDuration(start: Date, end: Date): string {
  // Validate inputs
  if (!(start instanceof Date) || isNaN(start.getTime())) {
    logger.warn(`Invalid start date provided to calculateDuration: ${start}`);
    return '0 min';
  }
  if (!(end instanceof Date) || isNaN(end.getTime())) {
    logger.warn(`Invalid end date provided to calculateDuration: ${end}`);
    return '0 min';
  }

  const diffMs = end.getTime() - start.getTime();

  // Handle negative durations (end before start)
  if (diffMs < 0) {
    logger.warn(`End date is before start date in calculateDuration`);
    return '0 min';
  }

  const diffMinutes = Math.floor(diffMs / (1000 * 60));
  const diffHours = Math.floor(diffMinutes / 60);
  const remainingMinutes = diffMinutes % 60;

  // Format based on duration length
  if (diffHours >= 24) {
    const days = Math.floor(diffHours / 24);
    const remainingHours = diffHours % 24;
    if (remainingHours === 0 && remainingMinutes === 0) {
      return `${days} Tag${days > 1 ? 'e' : ''}`;
    }
    return `${days} Tag${days > 1 ? 'e' : ''} ${remainingHours}h ${remainingMinutes}min`;
  }

  if (diffHours === 0) {
    return `${diffMinutes} min`;
  }

  if (remainingMinutes === 0) {
    return `${diffHours}h`;
  }

  return `${diffHours}h ${remainingMinutes}min`;
}

/**
 * Get server timezone name
 */
export function getServerTimezone(): string {
  try {
    return Intl.DateTimeFormat().resolvedOptions().timeZone;
  } catch {
    // Fallback if Intl is not available
    const offset = new Date().getTimezoneOffset();
    const offsetHours = Math.abs(Math.floor(offset / 60));
    const offsetMinutes = Math.abs(offset % 60);
    const sign = offset <= 0 ? '+' : '-';
    return `UTC${sign}${offsetHours.toString().padStart(2, '0')}:${offsetMinutes.toString().padStart(2, '0')}`;
  }
}

/**
 * Format a single calendar event from Graph API response
 */
export function formatCalendarEvent(event: Record<string, unknown>): FormattedCalendarEvent {
  // Validate input
  if (!event || typeof event !== 'object') {
    logger.warn('Invalid event object provided to formatCalendarEvent');
    const now = new Date();
    return {
      id: '',
      subject: '(Invalid Event)',
      startDate: formatLocalDate(now),
      startTime: formatLocalTime(now),
      endDate: formatLocalDate(now),
      endTime: formatLocalTime(now),
      startDateTimeUTC: formatUTCDateTime(now),
      endDateTimeUTC: formatUTCDateTime(now),
      startTimeDisplay: formatTimeWithUTC(now),
      endTimeDisplay: formatTimeWithUTC(now),
      duration: '0 min',
      isAllDay: false,
      isCancelled: false,
      isOnlineMeeting: false,
    };
  }

  // Extract start and end times with timezone info
  const startObj = event.start as { dateTime?: string; timeZone?: string } | undefined;
  const endObj = event.end as { dateTime?: string; timeZone?: string } | undefined;

  // Use convertToLocalTime with the timezone from Graph API
  // Provide fallback to current time if dateTime is missing
  let startDateTime: Date;
  let endDateTime: Date;

  try {
    startDateTime = startObj?.dateTime
      ? convertToLocalTime(startObj.dateTime, startObj.timeZone)
      : new Date();
    endDateTime = endObj?.dateTime
      ? convertToLocalTime(endObj.dateTime, endObj.timeZone)
      : new Date();

    // Validate that end is after start
    if (endDateTime < startDateTime) {
      logger.warn(
        `Event ${event.id || 'unknown'} has end time before start time. Adjusting end time.`
      );
      // Set end time to 1 hour after start if invalid
      endDateTime = new Date(startDateTime.getTime() + 60 * 60 * 1000);
    }
  } catch (error) {
    logger.error(`Error parsing event dates: ${error}`);
    const now = new Date();
    startDateTime = now;
    endDateTime = new Date(now.getTime() + 60 * 60 * 1000); // Default 1 hour duration
  }

  // Extract organizer
  const organizerObj = event.organizer as
    | {
        emailAddress?: { name?: string; address?: string };
      }
    | undefined;

  // Extract attendees
  const attendeesArray = event.attendees as
    | Array<{
        emailAddress?: { name?: string; address?: string };
        status?: { response?: string };
      }>
    | undefined;

  const attendees = attendeesArray?.map((a) => {
    const name = a.emailAddress?.name || a.emailAddress?.address || 'Unknown';
    const status = a.status?.response;
    return status ? `${name} (${status})` : name;
  });

  // Extract location
  const locationObj = event.location as { displayName?: string } | undefined;

  // Extract online meeting info
  const onlineMeetingObj = event.onlineMeeting as { joinUrl?: string } | undefined;

  // Extract categories
  const categories = event.categories as string[] | undefined;

  return {
    id: (event.id as string) || '',
    subject: (event.subject as string) || '(Kein Betreff)',
    // Server local time
    startDate: formatLocalDate(startDateTime),
    startTime: formatLocalTime(startDateTime),
    endDate: formatLocalDate(endDateTime),
    endTime: formatLocalTime(endDateTime),
    // UTC time
    startDateTimeUTC: formatUTCDateTime(startDateTime),
    endDateTimeUTC: formatUTCDateTime(endDateTime),
    // Combined display: "10:30 (UTC: 09:30)"
    startTimeDisplay: formatTimeWithUTC(startDateTime),
    endTimeDisplay: formatTimeWithUTC(endDateTime),
    duration: calculateDuration(startDateTime, endDateTime),
    isAllDay: (event.isAllDay as boolean) || false,
    location: locationObj?.displayName || undefined,
    organizer: organizerObj?.emailAddress?.name || undefined,
    organizerEmail: organizerObj?.emailAddress?.address || undefined,
    attendees: attendees || undefined,
    status: (event.showAs as string) || undefined,
    importance: (event.importance as string) || undefined,
    isOnlineMeeting: (event.isOnlineMeeting as boolean) || false,
    onlineMeetingUrl: onlineMeetingObj?.joinUrl || undefined,
    webLink: (event.webLink as string) || undefined,
    bodyPreview: (event.bodyPreview as string) || undefined,
    isCancelled: (event.isCancelled as boolean) || false,
    showAs: (event.showAs as string) || undefined,
    categories: categories || undefined,
  };
}

/**
 * Format calendar events response from Graph API
 */
export function formatCalendarResponse(
  response: Record<string, unknown>,
  startDateTime?: string,
  endDateTime?: string
): FormattedCalendarResponse {
  // Validate input
  if (!response || typeof response !== 'object') {
    logger.warn('Invalid response object provided to formatCalendarResponse');
    return {
      summary: {
        totalEvents: 0,
        dateRange: '',
        timezone: getServerTimezone(),
      },
      events: [],
      groupedByDate: {},
    };
  }

  const events = (response.value as Array<Record<string, unknown>>) || [];

  // Filter out invalid events and format
  const formattedEvents = events
    .filter((event) => {
      // Basic validation: event should have at least an id or subject
      return event && (event.id || event.subject);
    })
    .map(formatCalendarEvent)
    .filter((event) => {
      // Filter out events with invalid dates
      return event.id && event.startDate && event.startTime;
    });

  // Sort by start date/time with improved date parsing
  formattedEvents.sort((a, b) => {
    try {
      // Parse dates more reliably
      const aDateParts = a.startDate.split('.');
      const bDateParts = b.startDate.split('.');

      if (aDateParts.length !== 3 || bDateParts.length !== 3) {
        // Fallback to string comparison if date format is unexpected
        return a.startDate.localeCompare(b.startDate) || a.startTime.localeCompare(b.startTime);
      }

      const aDate = new Date(
        `${aDateParts[2]}-${aDateParts[1]}-${aDateParts[0]}T${a.startTime}:00`
      );
      const bDate = new Date(
        `${bDateParts[2]}-${bDateParts[1]}-${bDateParts[0]}T${b.startTime}:00`
      );

      const diff = aDate.getTime() - bDate.getTime();
      // If dates are equal, sort by time
      return diff !== 0 ? diff : a.startTime.localeCompare(b.startTime);
    } catch (error) {
      logger.warn(`Error sorting events: ${error}`);
      return 0;
    }
  });

  // Group by date with validation
  const groupedByDate: Record<string, FormattedCalendarEvent[]> = {};
  for (const event of formattedEvents) {
    // Ensure event has a valid startDate
    if (event.startDate) {
      if (!groupedByDate[event.startDate]) {
        groupedByDate[event.startDate] = [];
      }
      groupedByDate[event.startDate].push(event);
    } else {
      logger.warn(`Event ${event.id || 'unknown'} missing startDate, skipping grouping`);
    }
  }

  // Build date range string with improved error handling
  let dateRange = '';
  try {
    if (startDateTime && endDateTime) {
      const start = utcToLocalTime(startDateTime);
      const end = utcToLocalTime(endDateTime);
      dateRange = `${formatLocalDate(start)} - ${formatLocalDate(end)}`;
    } else if (formattedEvents.length > 0) {
      const firstDate = formattedEvents[0]?.startDate;
      const lastDate = formattedEvents[formattedEvents.length - 1]?.startDate;
      if (firstDate && lastDate) {
        dateRange = firstDate === lastDate ? firstDate : `${firstDate} - ${lastDate}`;
      }
    }
  } catch (error) {
    logger.warn(`Error building date range: ${error}`);
    dateRange = formattedEvents.length > 0 ? `${formattedEvents.length} events` : '';
  }

  logger.info(`Formatted ${formattedEvents.length} calendar events in server timezone`);

  return {
    summary: {
      totalEvents: formattedEvents.length,
      dateRange,
      timezone: getServerTimezone(),
    },
    events: formattedEvents,
    groupedByDate,
  };
}

/**
 * Convert formatted calendar response to human-readable text
 */
export function calendarResponseToText(response: FormattedCalendarResponse): string {
  const lines: string[] = [];

  // Header
  lines.push('📅 KALENDERÜBERSICHT');
  lines.push('═'.repeat(60));

  // Comprehensive summary at the top
  lines.push(`📊 Anzahl Termine: ${response.summary.totalEvents}`);
  lines.push(`📆 Zeitraum: ${response.summary.dateRange}`);
  lines.push(`🌍 Zeitzone: ${response.summary.timezone}`);

  if (response.summary.totalEvents === 0) {
    lines.push('═'.repeat(60));
    lines.push('');
    lines.push('ℹ️ Keine Termine in diesem Zeitraum gefunden.');
    return lines.join('\n');
  }

  // Calculate statistics for summary
  const now = new Date();
  const upcomingEvents = response.events.filter((e) => {
    try {
      const eventDateParts = e.startDate.split('.');
      if (eventDateParts.length === 3) {
        const eventDate = new Date(
          `${eventDateParts[2]}-${eventDateParts[1]}-${eventDateParts[0]}T${e.startTime}:00`
        );
        return eventDate >= now;
      }
      return false;
    } catch {
      return false;
    }
  });

  const pastEvents = response.events.length - upcomingEvents.length;
  const onlineMeetings = response.events.filter((e) => e.isOnlineMeeting).length;
  const cancelledEvents = response.events.filter((e) => e.isCancelled).length;
  const allDayEvents = response.events.filter((e) => e.isAllDay).length;

  // Count events by date
  const uniqueDates = Object.keys(response.groupedByDate).length;

  // Find next event
  const nextEvent = upcomingEvents.length > 0 ? upcomingEvents[0] : null;

  // Summary section
  lines.push('─'.repeat(60));
  lines.push('📋 ZUSAMMENFASSUNG:');
  lines.push(`   • ${upcomingEvents.length} bevorstehende Termine`);
  if (pastEvents > 0) {
    lines.push(`   • ${pastEvents} vergangene Termine`);
  }
  if (onlineMeetings > 0) {
    lines.push(`   • ${onlineMeetings} Online-Meetings`);
  }
  if (allDayEvents > 0) {
    lines.push(`   • ${allDayEvents} ganztägige Termine`);
  }
  if (cancelledEvents > 0) {
    lines.push(`   • ${cancelledEvents} abgesagte Termine`);
  }
  lines.push(`   • Termine verteilt auf ${uniqueDates} Tag${uniqueDates > 1 ? 'e' : ''}`);

  if (nextEvent) {
    try {
      const nextDateParts = nextEvent.startDate.split('.');
      if (nextDateParts.length === 3) {
        const nextDate = new Date(
          `${nextDateParts[2]}-${nextDateParts[1]}-${nextDateParts[0]}T${nextEvent.startTime}:00`
        );
        const hoursUntil = Math.floor((nextDate.getTime() - now.getTime()) / (1000 * 60 * 60));
        const minutesUntil = Math.floor(
          ((nextDate.getTime() - now.getTime()) % (1000 * 60 * 60)) / (1000 * 60)
        );

        if (hoursUntil >= 0 && minutesUntil >= 0) {
          if (hoursUntil === 0) {
            lines.push(`   • Nächster Termin in ${minutesUntil} Minuten: "${nextEvent.subject}"`);
          } else if (hoursUntil < 24) {
            lines.push(
              `   • Nächster Termin in ${hoursUntil}h ${minutesUntil}min: "${nextEvent.subject}"`
            );
          } else {
            const daysUntil = Math.floor(hoursUntil / 24);
            lines.push(
              `   • Nächster Termin in ${daysUntil} Tag${daysUntil > 1 ? 'en' : ''}: "${nextEvent.subject}"`
            );
          }
        }
      }
    } catch {
      // Ignore errors in next event calculation
    }
  }

  lines.push('═'.repeat(60));
  lines.push('');

  // ============================================================
  // QUICK SUMMARY LIST (so no event is overlooked)
  // ============================================================
  lines.push('📋 SCHNELLÜBERSICHT ALLER TERMINE:');
  lines.push('─'.repeat(60));

  // Sort all events by date and time
  const allEventsSorted = [...response.events].sort((a, b) => {
    const aDateStr = a.startDate.split('.').reverse().join('-') + 'T' + a.startTime;
    const bDateStr = b.startDate.split('.').reverse().join('-') + 'T' + b.startTime;
    return aDateStr.localeCompare(bDateStr);
  });

  for (let i = 0; i < allEventsSorted.length; i++) {
    const event = allEventsSorted[i];
    const statusIcon = event.isCancelled ? '❌' : event.isOnlineMeeting ? '💻' : '📍';
    const timeStr = event.isAllDay ? 'Ganztägig' : event.startTimeDisplay;
    const dateStr = event.startDate;
    lines.push(`${i + 1}. ${statusIcon} ${dateStr} ${timeStr} | ${event.subject}`);
  }

  lines.push('');
  lines.push('═'.repeat(60));
  lines.push('📖 DETAILANSICHT:');
  lines.push('═'.repeat(60));
  lines.push('');

  // Events grouped by date (detailed view)
  const sortedDates = Object.keys(response.groupedByDate).sort((a, b) => {
    const aDate = new Date(a.split('.').reverse().join('-'));
    const bDate = new Date(b.split('.').reverse().join('-'));
    return aDate.getTime() - bDate.getTime();
  });

  for (const date of sortedDates) {
    const events = response.groupedByDate[date];
    const dateObj = new Date(date.split('.').reverse().join('-'));
    const weekday = dateObj.toLocaleDateString('de-DE', { weekday: 'long' });

    lines.push(`📆 ${weekday}, ${date} (${events.length} Termin${events.length > 1 ? 'e' : ''})`);
    lines.push('─'.repeat(50));

    for (const event of events) {
      const statusIcon = event.isCancelled ? '❌' : event.isOnlineMeeting ? '💻' : '📍';
      const timeStr = event.isAllDay
        ? 'Ganztägig'
        : `${event.startTimeDisplay} - ${event.endTimeDisplay} (${event.duration})`;

      lines.push(`${statusIcon} ${event.subject}`);
      lines.push(`   ⏰ ${timeStr}`);

      if (event.location) {
        lines.push(`   📍 ${event.location}`);
      }

      if (event.organizer) {
        lines.push(`   👤 ${event.organizer}`);
      }

      if (event.attendees && event.attendees.length > 0) {
        lines.push(
          `   👥 Teilnehmer: ${event.attendees.slice(0, 5).join(', ')}${event.attendees.length > 5 ? ` (+${event.attendees.length - 5} weitere)` : ''}`
        );
      }

      if (event.onlineMeetingUrl) {
        lines.push(`   🔗 ${event.onlineMeetingUrl}`);
      }

      if (event.bodyPreview && event.bodyPreview.trim()) {
        const preview = event.bodyPreview.substring(0, 100).replace(/\n/g, ' ').trim();
        if (preview) {
          lines.push(`   📝 ${preview}${event.bodyPreview.length > 100 ? '...' : ''}`);
        }
      }

      lines.push('');
    }
  }

  return lines.join('\n');
}

/**
 * Detect if response is a calendar events response
 */
export function isCalendarResponse(response: unknown): boolean {
  if (typeof response !== 'object' || response === null) {
    return false;
  }

  const obj = response as Record<string, unknown>;
  if (!Array.isArray(obj.value)) {
    return false;
  }

  // Check if first item looks like a calendar event
  if (obj.value.length > 0) {
    const firstItem = obj.value[0] as Record<string, unknown>;
    return (
      'subject' in firstItem &&
      ('start' in firstItem || 'end' in firstItem) &&
      ('isAllDay' in firstItem || 'attendees' in firstItem)
    );
  }

  return false;
}

// ============================================================
// MAIL FORMATTING FUNCTIONS
// ============================================================

/**
 * Format a single mail message from Graph API response
 */
export function formatMailMessage(message: Record<string, unknown>): FormattedMailMessage {
  // Validate input
  if (!message || typeof message !== 'object') {
    logger.warn('Invalid message object provided to formatMailMessage');
    const now = new Date();
    return {
      id: '',
      subject: '(Invalid Message)',
      receivedDate: formatLocalDate(now),
      receivedTime: formatLocalTime(now),
      sentDate: formatLocalDate(now),
      sentTime: formatLocalTime(now),
      receivedDateTimeUTC: formatUTCDateTime(now),
      sentDateTimeUTC: formatUTCDateTime(now),
      receivedTimeDisplay: formatTimeWithUTC(now),
      sentTimeDisplay: formatTimeWithUTC(now),
      from: { name: 'Unknown', email: '' },
      to: [],
      hasAttachments: false,
      importance: 'normal',
      isRead: false,
      isDraft: false,
    };
  }

  // Extract received and sent times with error handling
  let receivedDateTime: Date;
  let sentDateTime: Date;

  try {
    receivedDateTime = message.receivedDateTime
      ? utcToLocalTime(message.receivedDateTime as string)
      : new Date();
    sentDateTime = message.sentDateTime
      ? utcToLocalTime(message.sentDateTime as string)
      : new Date();

    // Validate that dates are valid
    if (isNaN(receivedDateTime.getTime())) {
      logger.warn(`Invalid receivedDateTime: ${message.receivedDateTime}`);
      receivedDateTime = new Date();
    }
    if (isNaN(sentDateTime.getTime())) {
      logger.warn(`Invalid sentDateTime: ${message.sentDateTime}`);
      sentDateTime = new Date();
    }
  } catch (error) {
    logger.error(`Error parsing message dates: ${error}`);
    const now = new Date();
    receivedDateTime = now;
    sentDateTime = now;
  }

  // Extract from address
  const fromObj = message.from as
    | {
        emailAddress?: { name?: string; address?: string };
      }
    | undefined;

  // Extract to recipients
  const toRecipientsArray = message.toRecipients as
    | Array<{
        emailAddress?: { name?: string; address?: string };
      }>
    | undefined;

  const toRecipients =
    toRecipientsArray?.map((r) => ({
      name: r.emailAddress?.name || r.emailAddress?.address || 'Unknown',
      email: r.emailAddress?.address || '',
    })) || [];

  // Extract CC recipients
  const ccRecipientsArray = message.ccRecipients as
    | Array<{
        emailAddress?: { name?: string; address?: string };
      }>
    | undefined;

  const ccRecipients = ccRecipientsArray?.map((r) => ({
    name: r.emailAddress?.name || r.emailAddress?.address || 'Unknown',
    email: r.emailAddress?.address || '',
  }));

  // Extract flag
  const flagObj = message.flag as { flagStatus?: string } | undefined;

  // Extract categories
  const categories = message.categories as string[] | undefined;

  return {
    id: (message.id as string) || '',
    subject: (message.subject as string) || '(Kein Betreff)',
    // Server local time
    receivedDate: formatLocalDate(receivedDateTime),
    receivedTime: formatLocalTime(receivedDateTime),
    sentDate: formatLocalDate(sentDateTime),
    sentTime: formatLocalTime(sentDateTime),
    // UTC time
    receivedDateTimeUTC: formatUTCDateTime(receivedDateTime),
    sentDateTimeUTC: formatUTCDateTime(sentDateTime),
    // Combined display: "10:30 (UTC: 09:30)"
    receivedTimeDisplay: formatTimeWithUTC(receivedDateTime),
    sentTimeDisplay: formatTimeWithUTC(sentDateTime),
    from: {
      name: fromObj?.emailAddress?.name || fromObj?.emailAddress?.address || 'Unknown',
      email: fromObj?.emailAddress?.address || '',
    },
    to: toRecipients,
    cc: ccRecipients && ccRecipients.length > 0 ? ccRecipients : undefined,
    hasAttachments: (message.hasAttachments as boolean) || false,
    importance: (message.importance as string) || 'normal',
    isRead: (message.isRead as boolean) || false,
    isDraft: (message.isDraft as boolean) || false,
    bodyPreview: (message.bodyPreview as string) || undefined,
    webLink: (message.webLink as string) || undefined,
    categories: categories || undefined,
    flag: flagObj?.flagStatus ? { flagStatus: flagObj.flagStatus } : undefined,
    conversationId: (message.conversationId as string) || undefined,
  };
}

/**
 * Format mail messages response from Graph API
 */
export function formatMailResponse(response: Record<string, unknown>): FormattedMailResponse {
  // Validate input
  if (!response || typeof response !== 'object') {
    logger.warn('Invalid response object provided to formatMailResponse');
    return {
      summary: {
        totalMessages: 0,
        unreadCount: 0,
        dateRange: '',
        timezone: getServerTimezone(),
      },
      messages: [],
      groupedByDate: {},
    };
  }

  const messages = (response.value as Array<Record<string, unknown>>) || [];

  // Filter out invalid messages and format
  const formattedMessages = messages
    .filter((message) => {
      // Basic validation: message should have at least an id or subject
      return message && (message.id || message.subject);
    })
    .map(formatMailMessage)
    .filter((message) => {
      // Filter out messages with invalid dates
      return message.id && message.receivedDate && message.receivedTime;
    });

  // Sort by received date/time (newest first) with improved date parsing
  formattedMessages.sort((a, b) => {
    try {
      // Parse dates more reliably
      const aDateParts = a.receivedDate.split('.');
      const bDateParts = b.receivedDate.split('.');

      if (aDateParts.length !== 3 || bDateParts.length !== 3) {
        // Fallback to string comparison if date format is unexpected
        return (
          b.receivedDate.localeCompare(a.receivedDate) ||
          b.receivedTime.localeCompare(a.receivedTime)
        );
      }

      const aDate = new Date(
        `${aDateParts[2]}-${aDateParts[1]}-${aDateParts[0]}T${a.receivedTime}:00`
      );
      const bDate = new Date(
        `${bDateParts[2]}-${bDateParts[1]}-${bDateParts[0]}T${b.receivedTime}:00`
      );

      const diff = bDate.getTime() - aDate.getTime();
      // If dates are equal, sort by time
      return diff !== 0 ? diff : b.receivedTime.localeCompare(a.receivedTime);
    } catch (error) {
      logger.warn(`Error sorting messages: ${error}`);
      return 0;
    }
  });

  // Count unread messages
  const unreadCount = formattedMessages.filter((m) => !m.isRead).length;

  // Group by date with validation
  const groupedByDate: Record<string, FormattedMailMessage[]> = {};
  for (const message of formattedMessages) {
    // Ensure message has a valid receivedDate
    if (message.receivedDate) {
      if (!groupedByDate[message.receivedDate]) {
        groupedByDate[message.receivedDate] = [];
      }
      groupedByDate[message.receivedDate].push(message);
    } else {
      logger.warn(`Message ${message.id || 'unknown'} missing receivedDate, skipping grouping`);
    }
  }

  // Build date range string with improved error handling
  let dateRange = '';
  try {
    if (formattedMessages.length > 0) {
      const firstDate = formattedMessages[formattedMessages.length - 1]?.receivedDate;
      const lastDate = formattedMessages[0]?.receivedDate;
      if (firstDate && lastDate) {
        dateRange = firstDate === lastDate ? firstDate : `${firstDate} - ${lastDate}`;
      }
    }
  } catch (error) {
    logger.warn(`Error building date range: ${error}`);
    dateRange = formattedMessages.length > 0 ? `${formattedMessages.length} messages` : '';
  }

  logger.info(`Formatted ${formattedMessages.length} mail messages in server timezone`);

  return {
    summary: {
      totalMessages: formattedMessages.length,
      unreadCount,
      dateRange,
      timezone: getServerTimezone(),
    },
    messages: formattedMessages,
    groupedByDate,
  };
}

/**
 * Convert formatted mail response to human-readable text
 */
export function mailResponseToText(response: FormattedMailResponse): string {
  const lines: string[] = [];

  // Header
  lines.push('📧 E-MAIL ÜBERSICHT');
  lines.push('═'.repeat(60));

  // Comprehensive summary at the top
  lines.push(`📊 Anzahl E-Mails: ${response.summary.totalMessages}`);
  lines.push(`📬 Ungelesen: ${response.summary.unreadCount}`);
  lines.push(`📆 Zeitraum: ${response.summary.dateRange}`);
  lines.push(`🌍 Zeitzone: ${response.summary.timezone}`);

  if (response.summary.totalMessages === 0) {
    lines.push('═'.repeat(60));
    lines.push('');
    lines.push('ℹ️ Keine E-Mails gefunden.');
    return lines.join('\n');
  }

  // Calculate statistics for summary
  const readMessages = response.summary.totalMessages - response.summary.unreadCount;
  const highPriorityMessages = response.messages.filter((m) => m.importance === 'high').length;
  const flaggedMessages = response.messages.filter((m) => m.flag?.flagStatus === 'flagged').length;
  const messagesWithAttachments = response.messages.filter((m) => m.hasAttachments).length;
  const draftMessages = response.messages.filter((m) => m.isDraft).length;

  // Count messages by date
  const uniqueDates = Object.keys(response.groupedByDate).length;

  // Find most recent unread message
  const mostRecentUnread = response.messages.find((m) => !m.isRead);

  // Summary section
  lines.push('─'.repeat(60));
  lines.push('📋 ZUSAMMENFASSUNG:');
  lines.push(`   • ${response.summary.unreadCount} ungelesene E-Mails`);
  if (readMessages > 0) {
    lines.push(`   • ${readMessages} gelesene E-Mails`);
  }
  if (highPriorityMessages > 0) {
    lines.push(`   • ${highPriorityMessages} wichtige E-Mails`);
  }
  if (flaggedMessages > 0) {
    lines.push(`   • ${flaggedMessages} markierte E-Mails`);
  }
  if (messagesWithAttachments > 0) {
    lines.push(`   • ${messagesWithAttachments} E-Mails mit Anhängen`);
  }
  if (draftMessages > 0) {
    lines.push(`   • ${draftMessages} Entwürfe`);
  }
  lines.push(`   • E-Mails verteilt auf ${uniqueDates} Tag${uniqueDates > 1 ? 'e' : ''}`);

  if (mostRecentUnread) {
    lines.push(
      `   • Neueste ungelesene: "${mostRecentUnread.subject.substring(0, 50)}${mostRecentUnread.subject.length > 50 ? '...' : ''}"`
    );
    lines.push(
      `     Von: ${mostRecentUnread.from.name} am ${mostRecentUnread.receivedDate} um ${mostRecentUnread.receivedTimeDisplay}`
    );
  }

  lines.push('═'.repeat(60));
  lines.push('');

  // ============================================================
  // QUICK SUMMARY LIST (so no email is overlooked)
  // ============================================================
  lines.push('📋 SCHNELLÜBERSICHT ALLER E-MAILS:');
  lines.push('─'.repeat(60));

  // Sort by date/time (newest first)
  const allMessagesSorted = [...response.messages].sort((a, b) => {
    const aDateStr = a.receivedDate.split('.').reverse().join('-') + 'T' + a.receivedTime;
    const bDateStr = b.receivedDate.split('.').reverse().join('-') + 'T' + b.receivedTime;
    return bDateStr.localeCompare(aDateStr); // Newest first
  });

  for (let i = 0; i < allMessagesSorted.length; i++) {
    const msg = allMessagesSorted[i];
    const readIcon = msg.isRead ? '📭' : '📬';
    const attachIcon = msg.hasAttachments ? '📎' : '';
    const flagIcon = msg.flag?.flagStatus === 'flagged' ? '🚩' : '';
    const subjectShort =
      msg.subject.length > 40 ? msg.subject.substring(0, 40) + '...' : msg.subject;
    lines.push(
      `${i + 1}. ${readIcon}${attachIcon}${flagIcon} ${msg.receivedDate} ${msg.receivedTimeDisplay} | ${msg.from.name} | ${subjectShort}`
    );
  }

  lines.push('');
  lines.push('═'.repeat(60));
  lines.push('📖 DETAILANSICHT:');
  lines.push('═'.repeat(60));
  lines.push('');

  // Messages grouped by date (detailed view)
  const sortedDates = Object.keys(response.groupedByDate).sort((a, b) => {
    const aDate = new Date(a.split('.').reverse().join('-'));
    const bDate = new Date(b.split('.').reverse().join('-'));
    return bDate.getTime() - aDate.getTime(); // Newest first
  });

  for (const date of sortedDates) {
    const messages = response.groupedByDate[date];
    const dateObj = new Date(date.split('.').reverse().join('-'));
    const weekday = dateObj.toLocaleDateString('de-DE', { weekday: 'long' });

    lines.push(
      `📆 ${weekday}, ${date} (${messages.length} E-Mail${messages.length > 1 ? 's' : ''})`
    );
    lines.push('─'.repeat(50));

    for (const message of messages) {
      // Status icons
      const readIcon = message.isRead ? '📭' : '📬';
      const attachmentIcon = message.hasAttachments ? ' 📎' : '';
      const importanceIcon = message.importance === 'high' ? ' ❗' : '';
      const flagIcon = message.flag?.flagStatus === 'flagged' ? ' 🚩' : '';

      lines.push(`${readIcon}${attachmentIcon}${importanceIcon}${flagIcon} ${message.subject}`);
      lines.push(`   ⏰ ${message.receivedTimeDisplay}`);
      lines.push(`   👤 Von: ${message.from.name} <${message.from.email}>`);

      if (message.to.length > 0) {
        const toList = message.to
          .slice(0, 3)
          .map((r) => r.name)
          .join(', ');
        lines.push(
          `   👥 An: ${toList}${message.to.length > 3 ? ` (+${message.to.length - 3} weitere)` : ''}`
        );
      }

      if (message.cc && message.cc.length > 0) {
        const ccList = message.cc
          .slice(0, 2)
          .map((r) => r.name)
          .join(', ');
        lines.push(
          `   📋 CC: ${ccList}${message.cc.length > 2 ? ` (+${message.cc.length - 2} weitere)` : ''}`
        );
      }

      if (message.bodyPreview && message.bodyPreview.trim()) {
        const preview = message.bodyPreview.substring(0, 120).replace(/\n/g, ' ').trim();
        if (preview) {
          lines.push(`   📝 ${preview}${message.bodyPreview.length > 120 ? '...' : ''}`);
        }
      }

      if (message.categories && message.categories.length > 0) {
        lines.push(`   🏷️ ${message.categories.join(', ')}`);
      }

      lines.push('');
    }
  }

  return lines.join('\n');
}

/**
 * Detect if response is a mail messages response
 */
export function isMailResponse(response: unknown): boolean {
  if (typeof response !== 'object' || response === null) {
    return false;
  }

  const obj = response as Record<string, unknown>;
  if (!Array.isArray(obj.value)) {
    return false;
  }

  // Check if first item looks like a mail message
  if (obj.value.length > 0) {
    const firstItem = obj.value[0] as Record<string, unknown>;
    return (
      'subject' in firstItem &&
      ('from' in firstItem || 'toRecipients' in firstItem) &&
      ('receivedDateTime' in firstItem || 'sentDateTime' in firstItem)
    );
  }

  return false;
}

/**
 * Detect if response is a files/driveItems response
 */
export function isFilesResponse(response: unknown): boolean {
  if (typeof response !== 'object' || response === null) {
    return false;
  }

  const obj = response as Record<string, unknown>;
  if (!Array.isArray(obj.value)) {
    return false;
  }

  // Check if first item looks like a driveItem/file
  if (obj.value.length > 0) {
    const firstItem = obj.value[0] as Record<string, unknown>;
    return (
      ('name' in firstItem || 'displayName' in firstItem) &&
      ('size' in firstItem || 'folder' in firstItem || 'file' in firstItem) &&
      ('@odata.type' in firstItem || 'webUrl' in firstItem)
    );
  }

  return false;
}

/**
 * Format files/driveItems response from Graph API
 */
export function formatFilesResponse(response: Record<string, unknown>): {
  summary: {
    totalFiles: number;
    totalFolders: number;
    totalSize: number;
    fileTypes: Record<string, number>;
  };
  files: Array<Record<string, unknown>>;
  groupedByType: Record<string, Array<Record<string, unknown>>>;
} {
  // Validate input
  if (!response || typeof response !== 'object') {
    logger.warn('Invalid response object provided to formatFilesResponse');
    return {
      summary: {
        totalFiles: 0,
        totalFolders: 0,
        totalSize: 0,
        fileTypes: {},
      },
      files: [],
      groupedByType: {},
    };
  }

  const items = (response.value as Array<Record<string, unknown>>) || [];

  // Filter and categorize items
  const files: Array<Record<string, unknown>> = [];
  const folders: Array<Record<string, unknown>> = [];
  const fileTypes: Record<string, number> = {};
  let totalSize = 0;

  for (const item of items) {
    const odataType = (item['@odata.type'] as string) || '';
    const isFolder = odataType.includes('folder') || item.folder !== undefined;
    const isFile = odataType.includes('file') || (item.file !== undefined && !isFolder);

    if (isFolder) {
      folders.push(item);
    } else if (isFile) {
      files.push(item);

      // Calculate size
      const size = (item.size as number) || 0;
      totalSize += size;

      // Count file types
      const name = (item.name as string) || (item.displayName as string) || '';
      const extension = name.split('.').pop()?.toLowerCase() || 'unknown';
      fileTypes[extension] = (fileTypes[extension] || 0) + 1;
    } else {
      // Unknown type, treat as file
      files.push(item);
    }
  }

  // Group by file type
  const groupedByType: Record<string, Array<Record<string, unknown>>> = {};
  for (const file of files) {
    const name = (file.name as string) || (file.displayName as string) || '';
    const extension = name.split('.').pop()?.toLowerCase() || 'unknown';
    if (!groupedByType[extension]) {
      groupedByType[extension] = [];
    }
    groupedByType[extension].push(file);
  }

  return {
    summary: {
      totalFiles: files.length,
      totalFolders: folders.length,
      totalSize,
      fileTypes,
    },
    files: [...files, ...folders], // Combine files and folders
    groupedByType,
  };
}

/**
 * Format file size to human-readable string
 */
function formatFileSize(bytes: number): string {
  if (bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return `${(bytes / Math.pow(k, i)).toFixed(2)} ${sizes[i]}`;
}

/**
 * Convert formatted files response to human-readable text
 */
export function filesResponseToText(response: {
  summary: {
    totalFiles: number;
    totalFolders: number;
    totalSize: number;
    fileTypes: Record<string, number>;
  };
  files: Array<Record<string, unknown>>;
  groupedByType: Record<string, Array<Record<string, unknown>>>;
}): string {
  const lines: string[] = [];

  // Header
  lines.push('📁 DATEIEN ÜBERSICHT');
  lines.push('═'.repeat(60));

  // Comprehensive summary at the top
  lines.push(`📊 Anzahl Dateien: ${response.summary.totalFiles}`);
  lines.push(`📂 Anzahl Ordner: ${response.summary.totalFolders}`);
  lines.push(`💾 Gesamtgröße: ${formatFileSize(response.summary.totalSize)}`);

  if (response.summary.totalFiles === 0 && response.summary.totalFolders === 0) {
    lines.push('═'.repeat(60));
    lines.push('');
    lines.push('ℹ️ Keine Dateien oder Ordner gefunden.');
    return lines.join('\n');
  }

  // File type statistics
  const fileTypeEntries = Object.entries(response.summary.fileTypes).sort((a, b) => b[1] - a[1]);
  const topFileTypes = fileTypeEntries.slice(0, 5);

  // Summary section
  lines.push('─'.repeat(60));
  lines.push('📋 ZUSAMMENFASSUNG:');
  lines.push(`   • ${response.summary.totalFiles} Dateien`);
  if (response.summary.totalFolders > 0) {
    lines.push(`   • ${response.summary.totalFolders} Ordner`);
  }
  if (response.summary.totalSize > 0) {
    lines.push(`   • Gesamtgröße: ${formatFileSize(response.summary.totalSize)}`);
  }
  if (topFileTypes.length > 0) {
    lines.push(
      `   • Dateitypen: ${topFileTypes.map(([type, count]) => `${type} (${count})`).join(', ')}`
    );
    if (fileTypeEntries.length > 5) {
      lines.push(`     ... und ${fileTypeEntries.length - 5} weitere Typen`);
    }
  }

  // Find largest file
  const filesWithSize = response.files
    .filter((f) => (f.size as number) > 0)
    .sort((a, b) => ((b.size as number) || 0) - ((a.size as number) || 0));

  if (filesWithSize.length > 0) {
    const largestFile = filesWithSize[0];
    const name = (largestFile.name as string) || (largestFile.displayName as string) || 'Unbekannt';
    const size = formatFileSize((largestFile.size as number) || 0);
    lines.push(
      `   • Größte Datei: "${name.substring(0, 50)}${name.length > 50 ? '...' : ''}" (${size})`
    );
  }

  lines.push('═'.repeat(60));
  lines.push('');

  // Quick summary list
  lines.push('📋 SCHNELLÜBERSICHT:');
  lines.push('─'.repeat(60));

  for (let i = 0; i < Math.min(response.files.length, 20); i++) {
    const item = response.files[i];
    const name = (item.name as string) || (item.displayName as string) || 'Unbekannt';
    const odataType = (item['@odata.type'] as string) || '';
    const isFolder = odataType.includes('folder') || item.folder !== undefined;
    const icon = isFolder ? '📂' : '📄';
    const size = isFolder ? '' : ` (${formatFileSize((item.size as number) || 0)})`;
    const webUrl = (item.webUrl as string) || '';
    const urlPart = webUrl ? ` [🔗](${webUrl})` : '';

    lines.push(`${i + 1}. ${icon} ${name}${size}${urlPart}`);
  }

  if (response.files.length > 20) {
    lines.push(`... und ${response.files.length - 20} weitere`);
  }

  lines.push('');
  lines.push('═'.repeat(60));
  lines.push('📖 DETAILANSICHT NACH TYP:');
  lines.push('═'.repeat(60));
  lines.push('');

  // Grouped by type
  const sortedTypes = Object.keys(response.groupedByType).sort((a, b) => {
    return response.groupedByType[b].length - response.groupedByType[a].length;
  });

  for (const type of sortedTypes.slice(0, 10)) {
    const typeFiles = response.groupedByType[type];
    lines.push(
      `📄 ${type.toUpperCase()} (${typeFiles.length} Datei${typeFiles.length > 1 ? 'en' : ''})`
    );
    lines.push('─'.repeat(50));

    for (const file of typeFiles.slice(0, 10)) {
      const name = (file.name as string) || (file.displayName as string) || 'Unbekannt';
      const size = formatFileSize((file.size as number) || 0);
      const webUrl = (file.webUrl as string) || '';
      const urlPart = webUrl ? ` [🔗](${webUrl})` : '';

      lines.push(`   • ${name} (${size})${urlPart}`);
    }

    if (typeFiles.length > 10) {
      lines.push(`   ... und ${typeFiles.length - 10} weitere`);
    }

    lines.push('');
  }

  return lines.join('\n');
}

/**
 * Format any Graph API response based on its type
 */
export function formatGraphResponse(
  response: unknown,
  toolName?: string,
  params?: Record<string, unknown>
): { formatted: unknown; isFormatted: boolean; type?: 'calendar' | 'mail' | 'files' } {
  if (typeof response !== 'object' || response === null) {
    return { formatted: response, isFormatted: false };
  }

  const obj = response as Record<string, unknown>;

  // Check for calendar responses
  if (
    isCalendarResponse(response) ||
    toolName?.includes('calendar') ||
    toolName?.includes('event')
  ) {
    const startDateTime = params?.startDateTime as string | undefined;
    const endDateTime = params?.endDateTime as string | undefined;

    const formatted = formatCalendarResponse(obj, startDateTime, endDateTime);
    const textOutput = calendarResponseToText(formatted);

    // _humanReadable at the top of the response
    return {
      formatted: {
        _humanReadable: textOutput,
        ...formatted,
      },
      isFormatted: true,
      type: 'calendar',
    };
  }

  // Check for mail responses
  if (isMailResponse(response) || toolName?.includes('mail') || toolName?.includes('message')) {
    const formatted = formatMailResponse(obj);
    const textOutput = mailResponseToText(formatted);

    // _humanReadable at the top of the response
    return {
      formatted: {
        _humanReadable: textOutput,
        ...formatted,
      },
      isFormatted: true,
      type: 'mail',
    };
  }

  // Check for files/driveItems responses
  if (
    isFilesResponse(response) ||
    toolName?.includes('file') ||
    toolName?.includes('drive') ||
    toolName?.includes('driveItem')
  ) {
    const formatted = formatFilesResponse(obj);
    const textOutput = filesResponseToText(formatted);

    // _humanReadable at the top of the response
    return {
      formatted: {
        _humanReadable: textOutput,
        ...formatted,
      },
      isFormatted: true,
      type: 'files',
    };
  }

  // Return original response if no specific formatter
  return { formatted: response, isFormatted: false };
}

// ============================================================
// PROFESSION-BASED FORMATTING FUNCTIONS
// ============================================================

/**
 * Format options for profession-based customization
 */
export interface ProfessionFormatOptions {
  /** Override profession profile (if not using context) */
  professionProfile?: ProfessionProfile;
  /** Maximum items to show in summaries */
  maxItems?: number;
  /** Whether to include technical details */
  includeTechnicalDetails?: boolean;
}

/**
 * Get the active profession profile from context or options
 */
function getActiveProfessionProfile(
  options?: ProfessionFormatOptions
): ProfessionProfile | undefined {
  // Options override takes priority
  if (options?.professionProfile) {
    return options.professionProfile;
  }
  // Fall back to request context
  return getProfessionProfile();
}

/**
 * Format calendar response based on profession profile
 */
export function calendarResponseToTextByProfession(
  response: FormattedCalendarResponse,
  options?: ProfessionFormatOptions
): string {
  const profile = getActiveProfessionProfile(options);

  // No profile - use standard formatting
  if (!profile) {
    return calendarResponseToText(response);
  }

  const { detailLevel, languageStyle, formatPreference } = profile;
  const lines: string[] = [];

  // Header based on language style
  const headerText =
    languageStyle === 'formal' || languageStyle === 'professional'
      ? '📅 CALENDAR OVERVIEW'
      : languageStyle === 'customer-focused'
        ? '📅 YOUR SCHEDULE'
        : '📅 KALENDERÜBERSICHT';

  lines.push(headerText);
  lines.push('═'.repeat(60));

  // Summary based on detail level
  if (detailLevel === 'executive') {
    // Executive: Very brief summary with comprehensive overview
    lines.push(`• ${response.summary.totalEvents} appointments scheduled`);
    lines.push(`• Period: ${response.summary.dateRange}`);
    lines.push(`• Timezone: ${response.summary.timezone}`);

    // Calculate quick statistics
    const now = new Date();
    const upcomingEvents = response.events.filter((e) => {
      try {
        const eventDateParts = e.startDate.split('.');
        if (eventDateParts.length === 3) {
          const eventDate = new Date(
            `${eventDateParts[2]}-${eventDateParts[1]}-${eventDateParts[0]}T${e.startTime}:00`
          );
          return eventDate >= now;
        }
        return false;
      } catch {
        return false;
      }
    });

    const onlineMeetings = response.events.filter((e) => e.isOnlineMeeting).length;
    const uniqueDates = Object.keys(response.groupedByDate).length;

    if (upcomingEvents.length > 0) {
      lines.push(`• ${upcomingEvents.length} upcoming appointments`);
    }
    if (onlineMeetings > 0) {
      lines.push(`• ${onlineMeetings} online meetings`);
    }
    lines.push(`• Spread across ${uniqueDates} day${uniqueDates > 1 ? 's' : ''}`);

    // Just highlight key meetings (online, high-priority, etc.)
    const keyMeetings = response.events
      .filter((e) => e.isOnlineMeeting || e.importance === 'high')
      .slice(0, 5);

    if (keyMeetings.length > 0) {
      lines.push('');
      lines.push('🔑 Key Meetings:');
      for (const event of keyMeetings) {
        const timeStr = event.isAllDay ? 'All Day' : event.startTimeDisplay;
        lines.push(`  • ${event.startDate} ${timeStr}: ${event.subject}`);
      }
    }

    // Add action items
    const upcomingOnline = response.events.filter((e) => e.isOnlineMeeting).length;
    if (upcomingOnline > 0) {
      lines.push('');
      lines.push(
        `💡 ${upcomingOnline} online meeting${upcomingOnline > 1 ? 's' : ''} - prepare links`
      );
    }
  } else if (detailLevel === 'technical') {
    // Technical: Include IDs, structured data
    lines.push(`📊 Total: ${response.summary.totalEvents} events`);
    lines.push(`📆 Range: ${response.summary.dateRange}`);
    lines.push(`🌍 TZ: ${response.summary.timezone}`);
    lines.push('═'.repeat(60));
    lines.push('');

    // Code-like output for developers
    if (formatPreference === 'code-examples') {
      lines.push('```');
      lines.push('Events:');
      for (const event of response.events.slice(0, options?.maxItems || 20)) {
        lines.push(`  - id: ${event.id.substring(0, 20)}...`);
        lines.push(`    subject: "${event.subject}"`);
        lines.push(`    start: ${event.startDateTimeUTC}`);
        lines.push(`    end: ${event.endDateTimeUTC}`);
        lines.push(`    duration: ${event.duration}`);
        lines.push(`    online: ${event.isOnlineMeeting}`);
        if (event.onlineMeetingUrl) {
          lines.push(`    url: ${event.onlineMeetingUrl}`);
        }
        lines.push('');
      }
      lines.push('```');
    } else {
      // Structured technical output
      for (const event of response.events) {
        lines.push(`[${event.id.substring(0, 8)}] ${event.subject}`);
        lines.push(`  ├─ Start: ${event.startDateTimeUTC}`);
        lines.push(`  ├─ End: ${event.endDateTimeUTC}`);
        lines.push(`  ├─ Duration: ${event.duration}`);
        lines.push(`  ├─ Online: ${event.isOnlineMeeting ? 'Yes' : 'No'}`);
        if (event.location) lines.push(`  ├─ Location: ${event.location}`);
        if (event.onlineMeetingUrl) lines.push(`  └─ URL: ${event.onlineMeetingUrl}`);
        lines.push('');
      }
    }
  } else if (detailLevel === 'summary' && languageStyle === 'customer-focused') {
    // Sales/Customer-focused: Narrative style
    lines.push(`You have ${response.summary.totalEvents} upcoming appointments.`);
    lines.push(`📆 Period: ${response.summary.dateRange}`);
    lines.push('═'.repeat(60));
    lines.push('');

    // Group by importance for customer interactions
    const customerMeetings = response.events.filter(
      (e) =>
        e.subject.toLowerCase().includes('customer') ||
        e.subject.toLowerCase().includes('client') ||
        e.subject.toLowerCase().includes('meeting') ||
        e.subject.toLowerCase().includes('call')
    );

    if (customerMeetings.length > 0) {
      lines.push('🤝 Customer Interactions:');
      for (const event of customerMeetings.slice(0, 10)) {
        const timeStr = event.isAllDay ? 'All Day' : `${event.startTime} - ${event.endTime}`;
        lines.push(`  📅 ${event.startDate} | ${timeStr}`);
        lines.push(`     ${event.subject}`);
        if (event.attendees && event.attendees.length > 0) {
          lines.push(`     👥 With: ${event.attendees.slice(0, 3).join(', ')}`);
        }
        lines.push('');
      }
    }

    // Other meetings
    const otherMeetings = response.events.filter((e) => !customerMeetings.includes(e));
    if (otherMeetings.length > 0) {
      lines.push('📋 Other Appointments:');
      for (const event of otherMeetings.slice(0, 5)) {
        lines.push(`  • ${event.startDate} ${event.startTime}: ${event.subject}`);
      }
    }
  } else {
    // Default/detailed: Standard formatting
    return calendarResponseToText(response);
  }

  return lines.join('\n');
}

/**
 * Format mail response based on profession profile
 */
export function mailResponseToTextByProfession(
  response: FormattedMailResponse,
  options?: ProfessionFormatOptions
): string {
  const profile = getActiveProfessionProfile(options);

  // No profile - use standard formatting
  if (!profile) {
    return mailResponseToText(response);
  }

  const { detailLevel, languageStyle, formatPreference } = profile;
  const lines: string[] = [];

  // Header
  const headerText =
    languageStyle === 'formal' || languageStyle === 'professional'
      ? '📧 EMAIL OVERVIEW'
      : languageStyle === 'customer-focused'
        ? '📧 YOUR INBOX'
        : '📧 E-MAIL ÜBERSICHT';

  lines.push(headerText);
  lines.push('═'.repeat(60));

  if (detailLevel === 'executive') {
    // Executive: Key metrics and action items only with comprehensive summary
    lines.push(`• Total: ${response.summary.totalMessages} emails`);
    lines.push(`• Unread: ${response.summary.unreadCount} require attention`);
    lines.push(`• Period: ${response.summary.dateRange}`);
    lines.push(`• Timezone: ${response.summary.timezone}`);

    // High priority emails
    const highPriority = response.messages.filter((m) => m.importance === 'high');
    const flagged = response.messages.filter((m) => m.flag?.flagStatus === 'flagged');
    const unread = response.messages.filter((m) => !m.isRead);
    const withAttachments = response.messages.filter((m) => m.hasAttachments).length;
    const uniqueDates = Object.keys(response.groupedByDate).length;

    if (withAttachments > 0) {
      lines.push(`• ${withAttachments} emails with attachments`);
    }
    lines.push(`• Emails spread across ${uniqueDates} day${uniqueDates > 1 ? 's' : ''}`);

    if (highPriority.length > 0) {
      lines.push('');
      lines.push('❗ High Priority:');
      for (const msg of highPriority.slice(0, 5)) {
        lines.push(`  • ${msg.from.name}: ${msg.subject}`);
      }
    }

    if (flagged.length > 0) {
      lines.push('');
      lines.push('🚩 Flagged for Follow-up:');
      for (const msg of flagged.slice(0, 5)) {
        lines.push(`  • ${msg.from.name}: ${msg.subject}`);
      }
    }

    // Action summary
    lines.push('');
    lines.push('💡 Recommended Actions:');
    if (unread.length > 5) {
      lines.push(`  • Review ${unread.length} unread emails`);
    }
    if (flagged.length > 0) {
      lines.push(`  • Address ${flagged.length} flagged item${flagged.length > 1 ? 's' : ''}`);
    }
  } else if (detailLevel === 'technical') {
    // Technical: Structured data output
    lines.push(
      `📊 Total: ${response.summary.totalMessages} | Unread: ${response.summary.unreadCount}`
    );
    lines.push(`📆 Range: ${response.summary.dateRange}`);
    lines.push(`🌍 TZ: ${response.summary.timezone}`);
    lines.push('═'.repeat(60));
    lines.push('');

    if (formatPreference === 'code-examples') {
      lines.push('```');
      lines.push('Messages:');
      for (const msg of response.messages.slice(0, options?.maxItems || 15)) {
        lines.push(`  - id: ${msg.id.substring(0, 20)}...`);
        lines.push(`    subject: "${msg.subject}"`);
        lines.push(`    from: ${msg.from.email}`);
        lines.push(`    received: ${msg.receivedDateTimeUTC}`);
        lines.push(`    read: ${msg.isRead}`);
        lines.push(`    hasAttachments: ${msg.hasAttachments}`);
        if (msg.conversationId) {
          lines.push(`    conversationId: ${msg.conversationId.substring(0, 20)}...`);
        }
        lines.push('');
      }
      lines.push('```');
    } else {
      // Structured output
      for (const msg of response.messages.slice(0, 20)) {
        const status = msg.isRead ? '📭' : '📬';
        lines.push(`${status} [${msg.id.substring(0, 8)}] ${msg.subject}`);
        lines.push(`  ├─ From: ${msg.from.email}`);
        lines.push(`  ├─ Received: ${msg.receivedDateTimeUTC}`);
        lines.push(`  ├─ Attachments: ${msg.hasAttachments}`);
        lines.push(`  └─ Importance: ${msg.importance}`);
        lines.push('');
      }
    }
  } else if (languageStyle === 'customer-focused') {
    // Sales: Focus on customer communications
    lines.push(`You have ${response.summary.totalMessages} emails in your inbox.`);
    if (response.summary.unreadCount > 0) {
      lines.push(
        `📬 ${response.summary.unreadCount} unread message${response.summary.unreadCount > 1 ? 's' : ''} awaiting your attention.`
      );
    }
    lines.push('═'.repeat(60));
    lines.push('');

    // Prioritize unread
    const unread = response.messages.filter((m) => !m.isRead);
    if (unread.length > 0) {
      lines.push('📬 Unread Messages:');
      for (const msg of unread.slice(0, 10)) {
        lines.push(`  📧 ${msg.receivedDate} ${msg.receivedTime}`);
        lines.push(`     From: ${msg.from.name}`);
        lines.push(`     Subject: ${msg.subject}`);
        if (msg.bodyPreview) {
          const preview = msg.bodyPreview.substring(0, 80).replace(/\n/g, ' ').trim();
          lines.push(`     Preview: ${preview}...`);
        }
        lines.push('');
      }
    }

    // Recent read messages
    const read = response.messages.filter((m) => m.isRead).slice(0, 5);
    if (read.length > 0) {
      lines.push('📭 Recent Messages:');
      for (const msg of read) {
        lines.push(`  • ${msg.from.name}: ${msg.subject}`);
      }
    }
  } else {
    // Default: Standard formatting
    return mailResponseToText(response);
  }

  return lines.join('\n');
}

/**
 * Format any data structure based on profession profile
 * This is a generic formatter for arbitrary data
 */
export function formatDataByProfession(
  data: unknown,
  label: string,
  options?: ProfessionFormatOptions
): string {
  const profile = getActiveProfessionProfile(options);
  const lines: string[] = [];

  if (!profile || typeof data !== 'object' || data === null) {
    return JSON.stringify(data, null, 2);
  }

  const { detailLevel, formatPreference } = profile;

  if (detailLevel === 'executive') {
    // Executive summary - extract key points only
    lines.push(`📊 ${label} Summary`);
    lines.push('─'.repeat(40));

    if (Array.isArray(data)) {
      lines.push(`• ${data.length} item${data.length !== 1 ? 's' : ''} found`);
      // Show first 3 items briefly
      for (const item of data.slice(0, 3)) {
        if (typeof item === 'object' && item !== null) {
          const obj = item as Record<string, unknown>;
          const name = obj.displayName || obj.name || obj.subject || obj.title || 'Item';
          lines.push(`  • ${name}`);
        }
      }
      if (data.length > 3) {
        lines.push(`  ... and ${data.length - 3} more`);
      }
    } else if (typeof data === 'object') {
      const obj = data as Record<string, unknown>;
      const keys = Object.keys(obj).slice(0, 5);
      for (const key of keys) {
        const value = obj[key];
        if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
          lines.push(`• ${key}: ${value}`);
        } else if (Array.isArray(value)) {
          lines.push(`• ${key}: ${value.length} items`);
        }
      }
    }
  } else if (detailLevel === 'technical' && formatPreference === 'code-examples') {
    // Code block format
    lines.push(`// ${label}`);
    lines.push('```json');
    lines.push(JSON.stringify(data, null, 2));
    lines.push('```');
  } else {
    // Default: Pretty printed JSON
    lines.push(`📋 ${label}`);
    lines.push('─'.repeat(40));
    lines.push(JSON.stringify(data, null, 2));
  }

  return lines.join('\n');
}

/**
 * Get appropriate greeting/introduction based on profession
 */
export function getProfessionGreeting(
  context: 'calendar' | 'mail' | 'search' | 'general',
  options?: ProfessionFormatOptions
): string {
  const profile = getActiveProfessionProfile(options);

  if (!profile) {
    return '';
  }

  const { detailLevel, languageStyle } = profile;

  if (languageStyle === 'formal') {
    switch (context) {
      case 'calendar':
        return 'Here is your schedule overview:';
      case 'mail':
        return 'Here is your email summary:';
      case 'search':
        return 'Search results:';
      default:
        return 'Results:';
    }
  } else if (languageStyle === 'customer-focused') {
    switch (context) {
      case 'calendar':
        return 'Here are your upcoming appointments - let me highlight the key ones:';
      case 'mail':
        return "I've organized your emails by priority:";
      case 'search':
        return 'I found the following information for you:';
      default:
        return "Here's what I found:";
    }
  } else if (languageStyle === 'technical') {
    switch (context) {
      case 'calendar':
        return 'Calendar query results:';
      case 'mail':
        return 'Mail query results:';
      case 'search':
        return 'Query results:';
      default:
        return 'Data:';
    }
  }

  return '';
}

/**
 * Format Graph API response with profession-based personalization
 */
export function formatGraphResponseByProfession(
  response: unknown,
  toolName?: string,
  params?: Record<string, unknown>,
  options?: ProfessionFormatOptions
): {
  formatted: unknown;
  isFormatted: boolean;
  type?: 'calendar' | 'mail' | 'files';
  humanReadable?: string;
} {
  if (typeof response !== 'object' || response === null) {
    return { formatted: response, isFormatted: false };
  }

  const obj = response as Record<string, unknown>;

  // Check for calendar responses
  if (
    isCalendarResponse(response) ||
    toolName?.includes('calendar') ||
    toolName?.includes('event')
  ) {
    const startDateTime = params?.startDateTime as string | undefined;
    const endDateTime = params?.endDateTime as string | undefined;

    const formatted = formatCalendarResponse(obj, startDateTime, endDateTime);
    const textOutput = calendarResponseToTextByProfession(formatted, options);

    return {
      formatted: {
        _humanReadable: textOutput,
        ...formatted,
      },
      isFormatted: true,
      type: 'calendar',
      humanReadable: textOutput,
    };
  }

  // Check for mail responses
  if (isMailResponse(response) || toolName?.includes('mail') || toolName?.includes('message')) {
    const formatted = formatMailResponse(obj);
    const textOutput = mailResponseToTextByProfession(formatted, options);

    return {
      formatted: {
        _humanReadable: textOutput,
        ...formatted,
      },
      isFormatted: true,
      type: 'mail',
      humanReadable: textOutput,
    };
  }

  // Check for files/driveItems responses
  if (
    isFilesResponse(response) ||
    toolName?.includes('file') ||
    toolName?.includes('drive') ||
    toolName?.includes('driveItem')
  ) {
    const formatted = formatFilesResponse(obj);
    const textOutput = filesResponseToText(formatted);

    return {
      formatted: {
        _humanReadable: textOutput,
        ...formatted,
      },
      isFormatted: true,
      type: 'files',
      humanReadable: textOutput,
    };
  }

  // Return original response if no specific formatter
  return { formatted: response, isFormatted: false };
}

export default {
  // Calendar functions
  formatCalendarEvent,
  formatCalendarResponse,
  calendarResponseToText,
  calendarResponseToTextByProfession,
  isCalendarResponse,
  // Mail functions
  formatMailMessage,
  formatMailResponse,
  mailResponseToText,
  mailResponseToTextByProfession,
  isMailResponse,
  // Files functions
  formatFilesResponse,
  filesResponseToText,
  isFilesResponse,
  // Profession-based formatting
  formatDataByProfession,
  getProfessionGreeting,
  formatGraphResponseByProfession,
  // General functions
  formatGraphResponse,
  convertToLocalTime,
  utcToLocalTime,
  formatLocalDate,
  formatLocalTime,
  formatLocalDateTime,
  getServerTimezone,
  calculateDuration,
};
