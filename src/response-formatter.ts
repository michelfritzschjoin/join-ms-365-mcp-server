/**
 * Response Formatter - Formats API responses into structured, human-readable output
 * with server local time conversion (no UTC)
 */

import logger from './logger.js';

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
 */
export function convertToLocalTime(dateTimeString: string, timeZone?: string): Date {
  // If no timezone specified or already has Z suffix, parse directly
  if (!timeZone || dateTimeString.endsWith('Z')) {
    return new Date(dateTimeString);
  }

  // If timezone is UTC, append Z to ensure correct parsing
  if (timeZone === 'UTC' || timeZone === 'Etc/UTC' || timeZone === 'Etc/GMT') {
    // Append Z to indicate UTC
    const utcDateString = dateTimeString.endsWith('Z') ? dateTimeString : dateTimeString + 'Z';
    return new Date(utcDateString);
  }

  // For other timezones, we need to handle the conversion properly
  // The dateTime string is in the specified timezone, not UTC
  try {
    // Try to use Intl.DateTimeFormat to get the offset for the specified timezone
    const date = new Date(dateTimeString);

    // Get the offset difference between the event's timezone and local timezone
    const eventTzDate = new Date(date.toLocaleString('en-US', { timeZone: timeZone }));
    const localTzDate = new Date(date.toLocaleString('en-US'));

    // Calculate offset difference in milliseconds
    const offsetDiff = localTzDate.getTime() - eventTzDate.getTime();

    // Adjust the date by the offset difference
    return new Date(date.getTime() + offsetDiff);
  } catch {
    // Fallback: if timezone conversion fails, return the date as-is
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
 */
export function formatLocalDate(date: Date): string {
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
 */
export function formatLocalTime(date: Date): string {
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
 */
export function calculateDuration(start: Date, end: Date): string {
  const diffMs = end.getTime() - start.getTime();
  const diffMinutes = Math.floor(diffMs / (1000 * 60));
  const diffHours = Math.floor(diffMinutes / 60);
  const remainingMinutes = diffMinutes % 60;

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
  // Extract start and end times with timezone info
  const startObj = event.start as { dateTime?: string; timeZone?: string } | undefined;
  const endObj = event.end as { dateTime?: string; timeZone?: string } | undefined;

  // Use convertToLocalTime with the timezone from Graph API
  const startDateTime = startObj?.dateTime
    ? convertToLocalTime(startObj.dateTime, startObj.timeZone)
    : new Date();
  const endDateTime = endObj?.dateTime
    ? convertToLocalTime(endObj.dateTime, endObj.timeZone)
    : new Date();

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
  const events = (response.value as Array<Record<string, unknown>>) || [];
  const formattedEvents = events.map(formatCalendarEvent);

  // Sort by start date/time
  formattedEvents.sort((a, b) => {
    const aDate = new Date(`${a.startDate.split('.').reverse().join('-')} ${a.startTime}`);
    const bDate = new Date(`${b.startDate.split('.').reverse().join('-')} ${b.startTime}`);
    return aDate.getTime() - bDate.getTime();
  });

  // Group by date
  const groupedByDate: Record<string, FormattedCalendarEvent[]> = {};
  for (const event of formattedEvents) {
    if (!groupedByDate[event.startDate]) {
      groupedByDate[event.startDate] = [];
    }
    groupedByDate[event.startDate].push(event);
  }

  // Build date range string
  let dateRange = '';
  if (startDateTime && endDateTime) {
    const start = utcToLocalTime(startDateTime);
    const end = utcToLocalTime(endDateTime);
    dateRange = `${formatLocalDate(start)} - ${formatLocalDate(end)}`;
  } else if (formattedEvents.length > 0) {
    const firstDate = formattedEvents[0].startDate;
    const lastDate = formattedEvents[formattedEvents.length - 1].startDate;
    dateRange = firstDate === lastDate ? firstDate : `${firstDate} - ${lastDate}`;
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
  lines.push(`📊 Anzahl Termine: ${response.summary.totalEvents}`);
  lines.push(`📆 Zeitraum: ${response.summary.dateRange}`);
  lines.push(`🌍 Zeitzone: ${response.summary.timezone}`);
  lines.push('═'.repeat(60));
  lines.push('');

  if (response.summary.totalEvents === 0) {
    lines.push('ℹ️ Keine Termine in diesem Zeitraum gefunden.');
    return lines.join('\n');
  }

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
  // Extract received and sent times
  const receivedDateTime = message.receivedDateTime
    ? utcToLocalTime(message.receivedDateTime as string)
    : new Date();
  const sentDateTime = message.sentDateTime
    ? utcToLocalTime(message.sentDateTime as string)
    : new Date();

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
  const messages = (response.value as Array<Record<string, unknown>>) || [];
  const formattedMessages = messages.map(formatMailMessage);

  // Sort by received date/time (newest first)
  formattedMessages.sort((a, b) => {
    const aDate = new Date(`${a.receivedDate.split('.').reverse().join('-')} ${a.receivedTime}`);
    const bDate = new Date(`${b.receivedDate.split('.').reverse().join('-')} ${b.receivedTime}`);
    return bDate.getTime() - aDate.getTime();
  });

  // Count unread messages
  const unreadCount = formattedMessages.filter((m) => !m.isRead).length;

  // Group by date
  const groupedByDate: Record<string, FormattedMailMessage[]> = {};
  for (const message of formattedMessages) {
    if (!groupedByDate[message.receivedDate]) {
      groupedByDate[message.receivedDate] = [];
    }
    groupedByDate[message.receivedDate].push(message);
  }

  // Build date range string
  let dateRange = '';
  if (formattedMessages.length > 0) {
    const firstDate = formattedMessages[formattedMessages.length - 1].receivedDate;
    const lastDate = formattedMessages[0].receivedDate;
    dateRange = firstDate === lastDate ? firstDate : `${firstDate} - ${lastDate}`;
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
  lines.push(`📊 Anzahl E-Mails: ${response.summary.totalMessages}`);
  lines.push(`📬 Ungelesen: ${response.summary.unreadCount}`);
  lines.push(`📆 Zeitraum: ${response.summary.dateRange}`);
  lines.push(`🌍 Zeitzone: ${response.summary.timezone}`);
  lines.push('═'.repeat(60));
  lines.push('');

  if (response.summary.totalMessages === 0) {
    lines.push('ℹ️ Keine E-Mails gefunden.');
    return lines.join('\n');
  }

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
 * Format any Graph API response based on its type
 */
export function formatGraphResponse(
  response: unknown,
  toolName?: string,
  params?: Record<string, unknown>
): { formatted: unknown; isFormatted: boolean; type?: 'calendar' | 'mail' } {
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

  // Return original response if no specific formatter
  return { formatted: response, isFormatted: false };
}

export default {
  // Calendar functions
  formatCalendarEvent,
  formatCalendarResponse,
  calendarResponseToText,
  isCalendarResponse,
  // Mail functions
  formatMailMessage,
  formatMailResponse,
  mailResponseToText,
  isMailResponse,
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
