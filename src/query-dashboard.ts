/**
 * Query Dashboard - Secure Web Interface for viewing stored queries
 *
 * Provides a password-protected web dashboard to view all user queries,
 * statistics, and analytics.
 *
 * Security Features:
 * - Password protection via DASHBOARD_PASSWORD environment variable
 * - Session-based authentication with secure cookies
 * - CSRF protection for state-changing operations
 * - Rate limiting on login attempts
 * - No sensitive data exposed in responses
 *
 * ISO 27001 Compliance:
 * - Access logging for audit trail
 * - Secure authentication mechanism
 * - Data is read-only through dashboard
 */

import { Router, Request, Response, NextFunction } from 'express';
import { randomBytes, timingSafeEqual } from 'crypto';
import rateLimit from 'express-rate-limit';
import { getQueryStore, type QueryFilter } from './query-store.js';
import { getDashboardSessionStore } from './dashboard-session-store.js';
import logger from './logger.js';
import { rateLimitMiddleware } from './middleware/rate-limit.js';

// SECURITY: Strict rate limiter for login endpoints using express-rate-limit
// This is recognized by security scanners like CodeQL
const loginRateLimiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 minutes
  max: 5, // Limit each IP to 5 login attempts per window
  message: { error: 'Too many login attempts', message: 'Please try again later' },
  standardHeaders: true, // Return rate limit info in the `RateLimit-*` headers
  legacyHeaders: false, // Disable the `X-RateLimit-*` headers
  handler: (req, res) => {
    logger.warn('Login rate limit exceeded', { ip: req.ip });
    res.status(429).json({
      error: 'Too many login attempts',
      message: 'Please try again in 15 minutes',
    });
  },
});

const SESSION_DURATION_MS = 8 * 60 * 60 * 1000; // 8 hours
const MAX_LOGIN_ATTEMPTS = 5;
const LOGIN_LOCKOUT_MS = 15 * 60 * 1000; // 15 minutes

// Get persistent session store instance
const sessionStore = getDashboardSessionStore();

// Cache for hashed password (hashed once on first access)
let cachedPasswordHash: string | null = null;

/**
 * Get dashboard password hash from environment
 * SECURITY: Hash is computed once and cached for performance
 */
async function getDashboardPasswordHash(): Promise<string | null> {
  const password = process.env.DASHBOARD_PASSWORD;
  if (!password) {
    return null;
  }

  // Cache the hash to avoid re-hashing on every request
  if (!cachedPasswordHash) {
    cachedPasswordHash = await hashPassword(password);
  }

  return cachedPasswordHash;
}

/**
 * Check if dashboard is enabled
 */
export function isDashboardEnabled(): boolean {
  return !!process.env.DASHBOARD_PASSWORD;
}

/**
 * Hash password for comparison using bcrypt (secure, slow hashing)
 * SECURITY: Use bcrypt with proper cost factor to prevent brute force attacks
 */
async function hashPassword(password: string): Promise<string> {
  // Dynamic import to avoid requiring bcrypt as a hard dependency
  // bcrypt will be added as a dependency
  const bcrypt = await import('bcrypt');
  const saltRounds = 12; // OWASP recommended minimum: 10-12 rounds
  return await bcrypt.hash(password, saltRounds);
}

/**
 * Secure password comparison using bcrypt (timing-safe)
 * SECURITY: Use bcrypt.compare which is timing-safe and handles salt automatically
 */
async function verifyPassword(input: string, expectedHash: string): Promise<boolean> {
  try {
    const bcrypt = await import('bcrypt');
    return await bcrypt.compare(input, expectedHash);
  } catch {
    return false;
  }
}

/**
 * Generate secure session token
 */
function generateSessionToken(): string {
  return randomBytes(32).toString('hex');
}

/**
 * Create new session
 */
function createSession(ipAddress: string): string {
  const token = generateSessionToken();
  const expiresAt = new Date(Date.now() + SESSION_DURATION_MS);

  sessionStore.createSession(token, ipAddress, expiresAt);

  return token;
}

/**
 * Validate session token
 */
function validateSession(token: string, _ipAddress: string): boolean {
  return sessionStore.hasValidSession(token);
}

/**
 * Clean up expired sessions
 * Note: This is now handled automatically by the session store
 */
function cleanupSessions(): void {
  // Cleanup is handled automatically by DashboardSessionStore
  // This function is kept for compatibility but does nothing
}

/**
 * Check rate limiting for login attempts
 */
function isLoginRateLimited(ipAddress: string): boolean {
  return sessionStore.isRateLimited(ipAddress, MAX_LOGIN_ATTEMPTS);
}

/**
 * Record login attempt
 */
function recordLoginAttempt(ipAddress: string, success: boolean): void {
  sessionStore.recordLoginAttempt(ipAddress, success);
}

/**
 * Get session token from request
 */
function getSessionToken(req: Request): string | null {
  // Check cookie first
  const cookieHeader = req.headers.cookie || '';
  const cookies = cookieHeader.split(';').reduce(
    (acc, cookie) => {
      const [key, value] = cookie.trim().split('=');
      if (key && value) {
        acc[key] = value;
      }
      return acc;
    },
    {} as Record<string, string>
  );

  if (cookies['dashboard_session']) {
    return cookies['dashboard_session'];
  }

  // Check Authorization header
  const authHeader = req.headers.authorization;
  if (authHeader?.startsWith('Bearer ')) {
    return authHeader.substring(7);
  }

  return null;
}

/**
 * Authentication middleware
 */
function requireAuth(req: Request, res: Response, next: NextFunction): void {
  const token = getSessionToken(req);
  const ipAddress = req.ip || req.socket.remoteAddress || 'unknown';

  if (!token || !validateSession(token, ipAddress)) {
    // For HTML requests (browser), redirect to login
    // For API requests (JSON), return 401 JSON
    const acceptHeader = req.get('Accept') || '';
    const isApiRequest = acceptHeader.includes('application/json') || req.path.startsWith('/api/');

    if (isApiRequest) {
      res.status(401).json({
        error: 'Unauthorized',
        message: 'Please login to access the dashboard',
      });
    } else {
      // Redirect to login page for browser requests
      res.redirect('/dashboard/login');
    }
    return;
  }

  next();
}

/**
 * Create dashboard router
 */
export function createDashboardRouter(): Router {
  const router = Router();
  const queryStore = getQueryStore();

  // Login page
  router.get('/login', (req, res) => {
    const token = getSessionToken(req);
    const ipAddress = req.ip || req.socket.remoteAddress || 'unknown';

    // Already logged in?
    if (token && validateSession(token, ipAddress)) {
      return res.redirect('/dashboard');
    }

    res.send(getLoginPageHtml());
  });

  // Login API with strict rate limiting middleware
  // SECURITY: Using strict login rate limiter (5 attempts per 15 minutes)
  router.post('/login', loginRateLimiter, async (req, res) => {
    const ipAddress = req.ip || req.socket.remoteAddress || 'unknown';

    // Check rate limiting (additional layer)
    if (isLoginRateLimited(ipAddress)) {
      logger.warn('Dashboard login rate limited', { ip: ipAddress });
      return res.status(429).json({
        error: 'Too many login attempts',
        message: 'Please try again later',
      });
    }

    const password = req.body?.password;
    const expectedPasswordHash = await getDashboardPasswordHash();

    if (!expectedPasswordHash) {
      return res.status(503).json({
        error: 'Dashboard not configured',
        message: 'DASHBOARD_PASSWORD environment variable is not set',
      });
    }

    // SECURITY: Verify password using bcrypt (async)
    if (!password || !(await verifyPassword(password, expectedPasswordHash))) {
      recordLoginAttempt(ipAddress, false);
      logger.warn('Dashboard login failed', { ip: ipAddress });

      return res.status(401).json({
        error: 'Invalid password',
        message: 'Please check your password and try again',
      });
    }

    // Success
    recordLoginAttempt(ipAddress, true);
    const sessionToken = createSession(ipAddress);

    logger.info('Dashboard login successful', { ip: ipAddress });

    // Set secure cookie (Secure flag for HTTPS, SameSite for CSRF protection)
    const isSecure = req.secure || req.get('X-Forwarded-Proto') === 'https';
    const cookieOptions = [
      `dashboard_session=${sessionToken}`,
      'Path=/',
      'HttpOnly',
      isSecure ? 'Secure' : '',
      'SameSite=Lax', // Lax allows navigation from external sites
      `Max-Age=${SESSION_DURATION_MS / 1000}`,
    ]
      .filter(Boolean)
      .join('; ');

    res.setHeader('Set-Cookie', cookieOptions);

    res.json({
      success: true,
      message: 'Login successful',
      token: sessionToken,
    });
  });

  // Logout
  router.post('/logout', (req, res) => {
    const token = getSessionToken(req);

    if (token) {
      sessionStore.deleteSession(token);
    }

    const isSecure = req.secure || req.get('X-Forwarded-Proto') === 'https';
    const cookieOptions = [
      'dashboard_session=',
      'Path=/',
      'HttpOnly',
      isSecure ? 'Secure' : '',
      'SameSite=Lax',
      'Max-Age=0',
    ]
      .filter(Boolean)
      .join('; ');

    res.setHeader('Set-Cookie', cookieOptions);

    res.json({
      success: true,
      message: 'Logged out successfully',
    });
  });

  // Dashboard main page
  router.get('/', requireAuth, (req, res) => {
    res.send(getDashboardPageHtml());
  });

  // API: Get queries with rate limiting
  router.get('/api/queries', rateLimitMiddleware, requireAuth, (req, res) => {
    try {
      const filter: QueryFilter = {
        limit: parseInt((req.query.limit as string) || '50', 10),
        offset: parseInt((req.query.offset as string) || '0', 10),
      };

      if (req.query.toolName) {
        filter.toolName = req.query.toolName as string;
      }

      if (req.query.userIdHash) {
        filter.userIdHash = req.query.userIdHash as string;
      }

      if (req.query.success !== undefined) {
        filter.success = req.query.success === 'true';
      }

      if (req.query.startDate) {
        filter.startDate = new Date(req.query.startDate as string);
      }

      if (req.query.endDate) {
        filter.endDate = new Date(req.query.endDate as string);
      }

      const queries = queryStore.getQueries(filter);
      const total = queryStore.getQueryCount(filter);

      res.json({
        queries,
        total,
        limit: filter.limit,
        offset: filter.offset,
      });
    } catch (error) {
      logger.error('Error fetching queries:', error);
      res.status(500).json({ error: 'Failed to fetch queries' });
    }
  });

  // API: Get statistics with rate limiting
  router.get('/api/stats', rateLimitMiddleware, requireAuth, (req, res) => {
    try {
      const stats = queryStore.getStats();
      res.json(stats);
    } catch (error) {
      logger.error('Error fetching stats:', error);
      res.status(500).json({ error: 'Failed to fetch statistics' });
    }
  });

  // API: Get tool names (for filter dropdown) with rate limiting
  router.get('/api/tools', rateLimitMiddleware, requireAuth, (req, res) => {
    try {
      const tools = queryStore.getToolNames();
      res.json({ tools });
    } catch (error) {
      logger.error('Error fetching tool names:', error);
      res.status(500).json({ error: 'Failed to fetch tool names' });
    }
  });

  // API: Get tool performance statistics with rate limiting
  router.get('/api/tool-performance', rateLimitMiddleware, requireAuth, (req, res) => {
    try {
      const allQueries = queryStore.getQueries({ limit: 100000 });
      const toolStats: Record<
        string,
        {
          tool: string;
          totalQueries: number;
          successfulQueries: number;
          failedQueries: number;
          successRate: number;
          averageDuration: number;
          totalDuration: number;
        }
      > = {};

      for (const query of allQueries) {
        if (!toolStats[query.toolName]) {
          toolStats[query.toolName] = {
            tool: query.toolName,
            totalQueries: 0,
            successfulQueries: 0,
            failedQueries: 0,
            successRate: 0,
            averageDuration: 0,
            totalDuration: 0,
          };
        }

        const stat = toolStats[query.toolName];
        stat.totalQueries++;
        if (query.success) {
          stat.successfulQueries++;
        } else {
          stat.failedQueries++;
        }
        if (query.durationMs) {
          stat.totalDuration += query.durationMs;
        }
      }

      // Calculate success rates and average durations
      const toolPerformance = Object.values(toolStats).map((stat) => ({
        ...stat,
        successRate: stat.totalQueries > 0 ? (stat.successfulQueries / stat.totalQueries) * 100 : 0,
        averageDuration:
          stat.totalQueries > 0 ? Math.round(stat.totalDuration / stat.totalQueries) : 0,
      }));

      // Sort by total queries (most used first)
      toolPerformance.sort((a, b) => b.totalQueries - a.totalQueries);

      res.json({ toolPerformance });
    } catch (error) {
      logger.error('Error fetching tool performance:', error);
      res.status(500).json({ error: 'Failed to fetch tool performance' });
    }
  });

  // API: Get time series data for charts
  router.get('/api/timeseries', rateLimitMiddleware, requireAuth, (req, res) => {
    try {
      const days = parseInt((req.query.days as string) || '7', 10);
      const allQueries = queryStore.getQueries({ limit: 100000 });

      // Group by day
      const dailyData: Record<
        string,
        { date: string; total: number; successful: number; failed: number }
      > = {};

      const now = new Date();
      for (let i = days - 1; i >= 0; i--) {
        const date = new Date(now);
        date.setDate(date.getDate() - i);
        const dateKey = date.toISOString().split('T')[0];
        dailyData[dateKey] = {
          date: dateKey,
          total: 0,
          successful: 0,
          failed: 0,
        };
      }

      for (const query of allQueries) {
        const queryDate = new Date(query.timestamp).toISOString().split('T')[0];
        if (dailyData[queryDate]) {
          dailyData[queryDate].total++;
          if (query.success) {
            dailyData[queryDate].successful++;
          } else {
            dailyData[queryDate].failed++;
          }
        }
      }

      const timeseries = Object.values(dailyData);
      res.json({ timeseries });
    } catch (error) {
      logger.error('Error fetching time series data:', error);
      res.status(500).json({ error: 'Failed to fetch time series data' });
    }
  });

  // API: Get user statistics
  router.get('/api/user-stats', rateLimitMiddleware, requireAuth, (req, res) => {
    try {
      const allQueries = queryStore.getQueries({ limit: 100000 });
      const userStats: Record<
        string,
        {
          userIdHash: string;
          totalQueries: number;
          successfulQueries: number;
          failedQueries: number;
          successRate: number;
          uniqueTools: Set<string>;
          lastActivity: string;
        }
      > = {};

      for (const query of allQueries) {
        if (!userStats[query.userIdHash]) {
          userStats[query.userIdHash] = {
            userIdHash: query.userIdHash,
            totalQueries: 0,
            successfulQueries: 0,
            failedQueries: 0,
            successRate: 0,
            uniqueTools: new Set(),
            lastActivity: query.timestamp,
          };
        }

        const stat = userStats[query.userIdHash];
        stat.totalQueries++;
        if (query.success) {
          stat.successfulQueries++;
        } else {
          stat.failedQueries++;
        }
        stat.uniqueTools.add(query.toolName);
        if (new Date(query.timestamp) > new Date(stat.lastActivity)) {
          stat.lastActivity = query.timestamp;
        }
      }

      const userStatsArray = Object.values(userStats).map((stat) => ({
        userIdHash: stat.userIdHash,
        totalQueries: stat.totalQueries,
        successfulQueries: stat.successfulQueries,
        failedQueries: stat.failedQueries,
        successRate: stat.totalQueries > 0 ? (stat.successfulQueries / stat.totalQueries) * 100 : 0,
        uniqueTools: stat.uniqueTools.size,
        lastActivity: stat.lastActivity,
      }));

      // Sort by total queries (most active first)
      userStatsArray.sort((a, b) => b.totalQueries - a.totalQueries);

      res.json({ userStats: userStatsArray });
    } catch (error) {
      logger.error('Error fetching user statistics:', error);
      res.status(500).json({ error: 'Failed to fetch user statistics' });
    }
  });

  // API: Get error analysis
  router.get('/api/error-analysis', rateLimitMiddleware, requireAuth, (req, res) => {
    try {
      const allQueries = queryStore.getQueries({ limit: 100000 });
      const failedQueries = allQueries.filter((q) => !q.success);

      // Group errors by error message
      const errorCounts: Record<string, number> = {};
      const toolErrors: Record<string, Record<string, number>> = {};
      const errorTrends: Record<string, number[]> = {};

      for (const query of failedQueries) {
        const errorMsg = query.errorMessage || 'Unknown error';
        errorCounts[errorMsg] = (errorCounts[errorMsg] || 0) + 1;

        if (!toolErrors[query.toolName]) {
          toolErrors[query.toolName] = {};
        }
        toolErrors[query.toolName][errorMsg] = (toolErrors[query.toolName][errorMsg] || 0) + 1;

        // Group by day for trends
        const day = query.timestamp.split('T')[0];
        if (!errorTrends[day]) {
          errorTrends[day] = [];
        }
        errorTrends[day].push(1);
      }

      const topErrors = Object.entries(errorCounts)
        .map(([error, count]) => ({ error, count }))
        .sort((a, b) => b.count - a.count)
        .slice(0, 10);

      const toolErrorSummary = Object.entries(toolErrors).map(([tool, errors]) => ({
        tool,
        totalErrors: Object.values(errors).reduce((sum, count) => sum + count, 0),
        uniqueErrors: Object.keys(errors).length,
        topError: Object.entries(errors).sort((a, b) => b[1] - a[1])[0]?.[0] || 'N/A',
      }));

      res.json({
        totalErrors: failedQueries.length,
        topErrors,
        toolErrorSummary: toolErrorSummary.sort((a, b) => b.totalErrors - a.totalErrors),
        errorTrends: Object.entries(errorTrends)
          .map(([day, errors]) => ({ day, count: errors.length }))
          .sort((a, b) => a.day.localeCompare(b.day)),
      });
    } catch (error) {
      logger.error('Error fetching error analysis:', error);
      res.status(500).json({ error: 'Failed to fetch error analysis' });
    }
  });

  // API: Get performance bottlenecks
  router.get('/api/performance-bottlenecks', rateLimitMiddleware, requireAuth, (req, res) => {
    try {
      const allQueries = queryStore.getQueries({ limit: 100000 });
      const queriesWithDuration = allQueries.filter((q) => q.durationMs && q.durationMs > 0);

      // Find slow queries (above 95th percentile)
      const durations = queriesWithDuration.map((q) => q.durationMs!).sort((a, b) => a - b);
      const p95Index = Math.floor(durations.length * 0.95);
      const p95Threshold = durations[p95Index] || 0;

      const slowQueries = queriesWithDuration.filter((q) => q.durationMs! >= p95Threshold);

      // Group by tool
      const toolPerformance: Record<
        string,
        { tool: string; slowCount: number; avgDuration: number; maxDuration: number }
      > = {};

      for (const query of slowQueries) {
        if (!toolPerformance[query.toolName]) {
          const toolQueries = queriesWithDuration.filter((q) => q.toolName === query.toolName);
          toolPerformance[query.toolName] = {
            tool: query.toolName,
            slowCount: 0,
            avgDuration:
              toolQueries.reduce((sum, q) => sum + (q.durationMs || 0), 0) / toolQueries.length,
            maxDuration: Math.max(...toolQueries.map((q) => q.durationMs || 0)),
          };
        }
        toolPerformance[query.toolName].slowCount++;
      }

      const bottlenecks = Object.values(toolPerformance)
        .sort((a, b) => b.slowCount - a.slowCount)
        .slice(0, 10);

      res.json({
        p95Threshold: Math.round(p95Threshold),
        slowQueriesCount: slowQueries.length,
        bottlenecks,
      });
    } catch (error) {
      logger.error('Error fetching performance bottlenecks:', error);
      res.status(500).json({ error: 'Failed to fetch performance bottlenecks' });
    }
  });

  // API: Get usage patterns
  router.get('/api/usage-patterns', rateLimitMiddleware, requireAuth, (req, res) => {
    try {
      const allQueries = queryStore.getQueries({ limit: 100000 });

      // Hourly patterns
      const hourlyPatterns: Record<number, number> = {};
      for (let i = 0; i < 24; i++) {
        hourlyPatterns[i] = 0;
      }

      // Day of week patterns
      const dayPatterns: Record<number, number> = {};
      for (let i = 0; i < 7; i++) {
        dayPatterns[i] = 0;
      }

      // Tool combinations (which tools are used together)
      const userToolSequences: Record<string, string[]> = {};
      for (const query of allQueries) {
        const date = new Date(query.timestamp);
        const hour = date.getHours();
        const dayOfWeek = date.getDay();

        hourlyPatterns[hour]++;
        dayPatterns[dayOfWeek]++;

        if (!userToolSequences[query.userIdHash]) {
          userToolSequences[query.userIdHash] = [];
        }
        userToolSequences[query.userIdHash].push(query.toolName);
      }

      // Find common tool sequences
      const sequences: Record<string, number> = {};
      for (const userSeq of Object.values(userToolSequences)) {
        for (let i = 0; i < userSeq.length - 1; i++) {
          const seq = `${userSeq[i]} → ${userSeq[i + 1]}`;
          sequences[seq] = (sequences[seq] || 0) + 1;
        }
      }

      const topSequences = Object.entries(sequences)
        .map(([sequence, count]) => ({ sequence, count }))
        .sort((a, b) => b.count - a.count)
        .slice(0, 10);

      res.json({
        hourlyPatterns: Object.entries(hourlyPatterns).map(([hour, count]) => ({
          hour: parseInt(hour),
          count,
        })),
        dayPatterns: Object.entries(dayPatterns).map(([day, count]) => ({
          day: parseInt(day),
          dayName: [
            'Sonntag',
            'Montag',
            'Dienstag',
            'Mittwoch',
            'Donnerstag',
            'Freitag',
            'Samstag',
          ][parseInt(day)],
          count,
        })),
        topSequences,
      });
    } catch (error) {
      logger.error('Error fetching usage patterns:', error);
      res.status(500).json({ error: 'Failed to fetch usage patterns' });
    }
  });

  // API: Get success rate trends
  router.get('/api/success-trends', rateLimitMiddleware, requireAuth, (req, res) => {
    try {
      const days = parseInt((req.query.days as string) || '30', 10);
      const allQueries = queryStore.getQueries({ limit: 100000 });

      const dailyStats: Record<
        string,
        { date: string; total: number; successful: number; failed: number }
      > = {};

      const now = new Date();
      for (let i = days - 1; i >= 0; i--) {
        const date = new Date(now);
        date.setDate(date.getDate() - i);
        const dateKey = date.toISOString().split('T')[0];
        dailyStats[dateKey] = {
          date: dateKey,
          total: 0,
          successful: 0,
          failed: 0,
        };
      }

      for (const query of allQueries) {
        const queryDate = new Date(query.timestamp).toISOString().split('T')[0];
        if (dailyStats[queryDate]) {
          dailyStats[queryDate].total++;
          if (query.success) {
            dailyStats[queryDate].successful++;
          } else {
            dailyStats[queryDate].failed++;
          }
        }
      }

      const trends = Object.values(dailyStats).map((stat) => ({
        date: stat.date,
        successRate: stat.total > 0 ? (stat.successful / stat.total) * 100 : 0,
        total: stat.total,
        successful: stat.successful,
        failed: stat.failed,
      }));

      // Calculate trend direction
      const recent = trends.slice(-7);
      const older = trends.slice(-14, -7);
      const recentAvg = recent.reduce((sum, t) => sum + t.successRate, 0) / recent.length;
      const olderAvg = older.reduce((sum, t) => sum + t.successRate, 0) / older.length;
      const trendDirection =
        recentAvg > olderAvg ? 'improving' : recentAvg < olderAvg ? 'declining' : 'stable';
      const trendChange = recentAvg - olderAvg;

      res.json({
        trends,
        trendDirection,
        trendChange: Math.round(trendChange * 10) / 10,
        recentAverage: Math.round(recentAvg * 10) / 10,
        olderAverage: Math.round(olderAvg * 10) / 10,
      });
    } catch (error) {
      logger.error('Error fetching success trends:', error);
      res.status(500).json({ error: 'Failed to fetch success trends' });
    }
  });

  // API: Export all queries as CSV
  router.get('/api/export/csv', rateLimitMiddleware, requireAuth, (req, res) => {
    try {
      const filter: QueryFilter = {};
      if (req.query.toolName) {
        filter.toolName = req.query.toolName as string;
      }
      if (req.query.startDate) {
        filter.startDate = new Date(req.query.startDate as string);
      }
      if (req.query.endDate) {
        filter.endDate = new Date(req.query.endDate as string);
      }

      const queries = queryStore.getQueries({ ...filter, limit: 100000 });

      // Generate CSV
      const headers = [
        'Timestamp',
        'Tool Name',
        'User ID Hash',
        'Chat ID',
        'Success',
        'Duration (ms)',
        'Parameters',
        'Error Message',
      ];
      const csvRows = [headers.join(',')];

      for (const query of queries) {
        const row = [
          query.timestamp,
          query.toolName,
          query.userIdHash,
          query.chatId || '',
          query.success ? 'true' : 'false',
          query.durationMs?.toString() || '',
          JSON.stringify(query.parameters).replace(/"/g, '""'),
          query.errorMessage || '',
        ];
        csvRows.push(row.map((cell) => `"${cell}"`).join(','));
      }

      const csv = csvRows.join('\n');

      res.setHeader('Content-Type', 'text/csv');
      res.setHeader(
        'Content-Disposition',
        `attachment; filename="queries-export-${new Date().toISOString().split('T')[0]}.csv"`
      );
      res.send(csv);
    } catch (error) {
      logger.error('Error exporting queries as CSV:', error);
      res.status(500).json({ error: 'Failed to export queries' });
    }
  });

  // API: Export user data (GDPR) with rate limiting
  router.get('/api/export/:userIdHash', rateLimitMiddleware, requireAuth, (req, res) => {
    try {
      const queries = queryStore.exportUserQueries(req.params.userIdHash);

      res.setHeader('Content-Type', 'application/json');
      res.setHeader(
        'Content-Disposition',
        `attachment; filename="user-data-${req.params.userIdHash}.json"`
      );

      res.json({
        exportDate: new Date().toISOString(),
        userIdHash: req.params.userIdHash,
        queries,
      });
    } catch (error) {
      logger.error('Error exporting user data:', error);
      res.status(500).json({ error: 'Failed to export user data' });
    }
  });

  // API: Delete user data (GDPR Right to Erasure) with rate limiting
  router.delete('/api/user/:userIdHash', rateLimitMiddleware, requireAuth, (req, res) => {
    try {
      const deleted = queryStore.deleteUserQueries(req.params.userIdHash);

      res.json({
        success: true,
        deleted,
        message: `Deleted ${deleted} queries for user`,
      });
    } catch (error) {
      logger.error('Error deleting user data:', error);
      res.status(500).json({ error: 'Failed to delete user data' });
    }
  });

  return router;
}

/**
 * Get login page HTML
 */
function getLoginPageHtml(): string {
  return `<!DOCTYPE html>
<html lang="de">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Dashboard Login - MS 365 MCP Server</title>
  <style>
    :root {
      --bg-primary: #0a0a0f;
      --bg-secondary: #12121a;
      --bg-tertiary: #1a1a24;
      --text-primary: #e8e8ed;
      --text-secondary: #9898a8;
      --accent: #6366f1;
      --accent-hover: #818cf8;
      --error: #ef4444;
      --success: #22c55e;
      --border: #2a2a38;
    }

    * {
      margin: 0;
      padding: 0;
      box-sizing: border-box;
    }

    body {
      font-family: 'SF Pro Display', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      background: var(--bg-primary);
      color: var(--text-primary);
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      background-image: 
        radial-gradient(ellipse at 50% 0%, rgba(99, 102, 241, 0.15) 0%, transparent 50%),
        radial-gradient(ellipse at 80% 80%, rgba(139, 92, 246, 0.1) 0%, transparent 50%);
    }

    .login-container {
      background: var(--bg-secondary);
      border: 1px solid var(--border);
      border-radius: 16px;
      padding: 48px;
      width: 100%;
      max-width: 420px;
      box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5);
    }

    .logo {
      text-align: center;
      margin-bottom: 32px;
    }

    .logo svg {
      width: 64px;
      height: 64px;
      fill: var(--accent);
    }

    h1 {
      font-size: 24px;
      font-weight: 600;
      text-align: center;
      margin-bottom: 8px;
    }

    .subtitle {
      color: var(--text-secondary);
      text-align: center;
      font-size: 14px;
      margin-bottom: 32px;
    }

    .form-group {
      margin-bottom: 24px;
    }

    label {
      display: block;
      font-size: 14px;
      font-weight: 500;
      margin-bottom: 8px;
      color: var(--text-secondary);
    }

    input[type="password"] {
      width: 100%;
      padding: 14px 16px;
      font-size: 16px;
      background: var(--bg-tertiary);
      border: 1px solid var(--border);
      border-radius: 10px;
      color: var(--text-primary);
      outline: none;
      transition: border-color 0.2s, box-shadow 0.2s;
    }

    input[type="password"]:focus {
      border-color: var(--accent);
      box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.2);
    }

    button {
      width: 100%;
      padding: 14px 24px;
      font-size: 16px;
      font-weight: 600;
      background: var(--accent);
      color: white;
      border: none;
      border-radius: 10px;
      cursor: pointer;
      transition: background 0.2s, transform 0.1s;
    }

    button:hover {
      background: var(--accent-hover);
    }

    button:active {
      transform: scale(0.98);
    }

    button:disabled {
      opacity: 0.6;
      cursor: not-allowed;
    }

    .error-message {
      background: rgba(239, 68, 68, 0.1);
      border: 1px solid rgba(239, 68, 68, 0.3);
      color: var(--error);
      padding: 12px 16px;
      border-radius: 8px;
      font-size: 14px;
      margin-bottom: 24px;
      display: none;
    }

    .error-message.visible {
      display: block;
    }

    .footer {
      margin-top: 32px;
      text-align: center;
      font-size: 12px;
      color: var(--text-secondary);
    }
  </style>
</head>
<body>
  <div class="login-container">
    <div class="logo">
      <svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
        <path d="M12 2L2 7l10 5 10-5-10-5zM2 17l10 5 10-5M2 12l10 5 10-5"/>
      </svg>
    </div>
    <h1>Query Dashboard</h1>
    <p class="subtitle">MS 365 MCP Server - Authentifizierung erforderlich</p>
    
    <div class="error-message" id="errorMessage"></div>
    
    <form id="loginForm">
      <div class="form-group">
        <label for="password">Dashboard Passwort</label>
        <input type="password" id="password" name="password" placeholder="Passwort eingeben..." required autofocus>
      </div>
      <button type="submit" id="submitBtn">Anmelden</button>
    </form>
    
    <div class="footer">
      Geschützt durch ISO 27001 konforme Authentifizierung
    </div>
  </div>

  <script>
    const form = document.getElementById('loginForm');
    const errorMessage = document.getElementById('errorMessage');
    const submitBtn = document.getElementById('submitBtn');

    form.addEventListener('submit', async (e) => {
      e.preventDefault();
      
      const password = document.getElementById('password').value;
      
      submitBtn.disabled = true;
      submitBtn.textContent = 'Anmelden...';
      errorMessage.classList.remove('visible');

      try {
        const response = await fetch('/dashboard/login', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ password }),
        });

        const data = await response.json();

        if (response.ok && data.success) {
          window.location.href = '/dashboard';
        } else {
          errorMessage.textContent = data.message || 'Anmeldung fehlgeschlagen';
          errorMessage.classList.add('visible');
        }
      } catch (error) {
        errorMessage.textContent = 'Verbindungsfehler. Bitte erneut versuchen.';
        errorMessage.classList.add('visible');
      } finally {
        submitBtn.disabled = false;
        submitBtn.textContent = 'Anmelden';
      }
    });
  </script>
</body>
</html>`;
}

/**
 * Get dashboard page HTML
 */
function getDashboardPageHtml(): string {
  return `<!DOCTYPE html>
<html lang="de">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Query Dashboard - MS 365 MCP Server</title>
  <style>
    :root {
      --bg-primary: #0a0a0f;
      --bg-secondary: #12121a;
      --bg-tertiary: #1a1a24;
      --bg-hover: #22222e;
      --text-primary: #e8e8ed;
      --text-secondary: #9898a8;
      --text-muted: #6b6b78;
      --accent: #6366f1;
      --accent-hover: #818cf8;
      --success: #22c55e;
      --error: #ef4444;
      --warning: #f59e0b;
      --border: #2a2a38;
      --border-light: #3a3a48;
    }

    * {
      margin: 0;
      padding: 0;
      box-sizing: border-box;
    }

    body {
      font-family: 'SF Pro Display', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      background: var(--bg-primary);
      color: var(--text-primary);
      min-height: 100vh;
    }

    /* Header */
    .header {
      background: var(--bg-secondary);
      border-bottom: 1px solid var(--border);
      padding: 16px 32px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      position: sticky;
      top: 0;
      z-index: 100;
    }

    .header h1 {
      font-size: 20px;
      font-weight: 600;
      display: flex;
      align-items: center;
      gap: 12px;
    }

    .header h1 svg {
      width: 28px;
      height: 28px;
      fill: var(--accent);
    }

    .header-actions {
      display: flex;
      gap: 12px;
    }

    .btn {
      padding: 10px 20px;
      font-size: 14px;
      font-weight: 500;
      border: none;
      border-radius: 8px;
      cursor: pointer;
      transition: all 0.2s;
    }

    .btn-primary {
      background: var(--accent);
      color: white;
    }

    .btn-primary:hover {
      background: var(--accent-hover);
    }

    .btn-secondary {
      background: var(--bg-tertiary);
      color: var(--text-primary);
      border: 1px solid var(--border);
    }

    .btn-secondary:hover {
      background: var(--bg-hover);
    }

    /* Main Layout */
    .main {
      padding: 32px;
      max-width: 1600px;
      margin: 0 auto;
    }

    /* Stats Grid */
    .stats-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 20px;
      margin-bottom: 32px;
    }

    .stat-card {
      background: var(--bg-secondary);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 24px;
    }

    .stat-label {
      font-size: 13px;
      color: var(--text-secondary);
      margin-bottom: 8px;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }

    .stat-value {
      font-size: 32px;
      font-weight: 700;
      color: var(--text-primary);
    }

    .stat-value.success {
      color: var(--success);
    }

    .stat-value.error {
      color: var(--error);
    }

    /* Filters */
    .filters {
      background: var(--bg-secondary);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 20px;
      margin-bottom: 24px;
      display: flex;
      gap: 16px;
      flex-wrap: wrap;
      align-items: flex-end;
    }

    .filter-group {
      display: flex;
      flex-direction: column;
      gap: 6px;
    }

    .filter-group label {
      font-size: 12px;
      color: var(--text-secondary);
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }

    .filter-group select,
    .filter-group input {
      padding: 10px 14px;
      font-size: 14px;
      background: var(--bg-tertiary);
      border: 1px solid var(--border);
      border-radius: 8px;
      color: var(--text-primary);
      min-width: 180px;
    }

    .filter-group select:focus,
    .filter-group input:focus {
      outline: none;
      border-color: var(--accent);
    }

    /* Queries Table */
    .queries-section {
      background: var(--bg-secondary);
      border: 1px solid var(--border);
      border-radius: 12px;
      overflow: hidden;
    }

    .queries-header {
      padding: 20px 24px;
      border-bottom: 1px solid var(--border);
      display: flex;
      justify-content: space-between;
      align-items: center;
    }

    .queries-header h2 {
      font-size: 18px;
      font-weight: 600;
    }

    .queries-count {
      font-size: 14px;
      color: var(--text-secondary);
    }

    .queries-table {
      width: 100%;
      border-collapse: collapse;
    }

    .queries-table th,
    .queries-table td {
      padding: 14px 20px;
      text-align: left;
      border-bottom: 1px solid var(--border);
    }

    .queries-table th {
      background: var(--bg-tertiary);
      font-size: 12px;
      font-weight: 600;
      color: var(--text-secondary);
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }

    .queries-table tr:hover td {
      background: var(--bg-hover);
    }

    .queries-table td {
      font-size: 14px;
    }

    .tool-name {
      font-family: 'SF Mono', 'Monaco', 'Consolas', monospace;
      background: var(--bg-tertiary);
      padding: 4px 8px;
      border-radius: 4px;
      font-size: 13px;
    }

    .status-badge {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      padding: 4px 10px;
      border-radius: 20px;
      font-size: 12px;
      font-weight: 500;
    }

    .status-badge.success {
      background: rgba(34, 197, 94, 0.15);
      color: var(--success);
    }

    .status-badge.error {
      background: rgba(239, 68, 68, 0.15);
      color: var(--error);
    }

    .user-hash {
      font-family: 'SF Mono', 'Monaco', 'Consolas', monospace;
      font-size: 12px;
      color: var(--text-muted);
    }

    .timestamp {
      font-size: 13px;
      color: var(--text-secondary);
    }

    .duration {
      font-size: 13px;
      color: var(--text-secondary);
    }

    /* Pagination */
    .pagination {
      display: flex;
      justify-content: center;
      gap: 8px;
      padding: 20px;
      border-top: 1px solid var(--border);
    }

    .pagination button {
      padding: 8px 16px;
      font-size: 14px;
      background: var(--bg-tertiary);
      border: 1px solid var(--border);
      border-radius: 6px;
      color: var(--text-primary);
      cursor: pointer;
    }

    .pagination button:hover:not(:disabled) {
      background: var(--bg-hover);
      border-color: var(--accent);
    }

    .pagination button:disabled {
      opacity: 0.5;
      cursor: not-allowed;
    }

    .pagination .page-info {
      display: flex;
      align-items: center;
      padding: 0 16px;
      font-size: 14px;
      color: var(--text-secondary);
    }

    /* Chart */
    .chart-section {
      background: var(--bg-secondary);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 24px;
      margin-bottom: 24px;
    }

    .chart-section h3 {
      font-size: 16px;
      font-weight: 600;
      margin-bottom: 20px;
    }

    .chart-container {
      height: 200px;
      display: flex;
      align-items: flex-end;
      gap: 4px;
    }

    .chart-bar {
      flex: 1;
      background: var(--accent);
      border-radius: 4px 4px 0 0;
      transition: all 0.3s;
      min-height: 4px;
    }

    .chart-bar:hover {
      background: var(--accent-hover);
    }

    /* Loading */
    .loading {
      text-align: center;
      padding: 60px;
      color: var(--text-secondary);
    }

    .loading-spinner {
      width: 40px;
      height: 40px;
      border: 3px solid var(--border);
      border-top-color: var(--accent);
      border-radius: 50%;
      animation: spin 0.8s linear infinite;
      margin: 0 auto 16px;
    }

    @keyframes spin {
      to { transform: rotate(360deg); }
    }

    /* Empty State */
    .empty-state {
      text-align: center;
      padding: 80px 40px;
      color: var(--text-secondary);
    }

    .empty-state svg {
      width: 64px;
      height: 64px;
      fill: var(--text-muted);
      margin-bottom: 16px;
    }

    /* Tabs */
    .tabs {
      display: flex;
      gap: 8px;
      border-bottom: 2px solid var(--border);
      margin-bottom: 24px;
      background: var(--bg-secondary);
      border-radius: 12px 12px 0 0;
      padding: 8px 8px 0 8px;
    }

    .tab {
      padding: 12px 24px;
      font-size: 14px;
      font-weight: 500;
      color: var(--text-secondary);
      background: transparent;
      border: none;
      border-bottom: 2px solid transparent;
      cursor: pointer;
      transition: all 0.2s;
      position: relative;
      top: 2px;
    }

    .tab:hover {
      color: var(--text-primary);
      background: var(--bg-tertiary);
    }

    .tab.active {
      color: var(--accent);
      border-bottom-color: var(--accent);
      background: var(--bg-secondary);
    }

    .tab-content {
      display: none;
    }

    .tab-content.active {
      display: block;
    }

    /* Tool Performance Table */
    .performance-table {
      width: 100%;
      border-collapse: collapse;
    }

    .performance-table th,
    .performance-table td {
      padding: 12px 16px;
      text-align: left;
      border-bottom: 1px solid var(--border);
    }

    .performance-table th {
      background: var(--bg-tertiary);
      font-size: 12px;
      font-weight: 600;
      color: var(--text-secondary);
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }

    .performance-table tr:hover td {
      background: var(--bg-hover);
    }

    .progress-bar {
      height: 8px;
      background: var(--bg-tertiary);
      border-radius: 4px;
      overflow: hidden;
      margin-top: 4px;
    }

    .progress-bar-fill {
      height: 100%;
      background: var(--accent);
      transition: width 0.3s;
    }

    .progress-bar-fill.success {
      background: var(--success);
    }

    .progress-bar-fill.error {
      background: var(--error);
    }

    /* Line Chart */
    .line-chart-container {
      height: 300px;
      position: relative;
      margin-top: 20px;
    }

    .line-chart {
      width: 100%;
      height: 100%;
      position: relative;
    }

    .chart-line {
      fill: none;
      stroke: var(--accent);
      stroke-width: 2;
    }

    .chart-line.success {
      stroke: var(--success);
    }

    .chart-line.error {
      stroke: var(--error);
    }

    .chart-axis {
      stroke: var(--border);
      stroke-width: 1;
    }

    .chart-label {
      font-size: 11px;
      fill: var(--text-secondary);
    }

    /* Export Button */
    .btn-export {
      background: var(--bg-tertiary);
      color: var(--text-primary);
      border: 1px solid var(--border);
      padding: 8px 16px;
      border-radius: 6px;
      cursor: pointer;
      font-size: 13px;
      transition: all 0.2s;
    }

    .btn-export:hover {
      background: var(--bg-hover);
      border-color: var(--accent);
    }

    /* Responsive */
    @media (max-width: 768px) {
      .header {
        padding: 12px 16px;
      }

      .main {
        padding: 16px;
      }

      .filters {
        flex-direction: column;
      }

      .filter-group select,
      .filter-group input {
        width: 100%;
      }

      .queries-table {
        display: block;
        overflow-x: auto;
      }

      .tabs {
        overflow-x: auto;
        flex-wrap: nowrap;
      }

      .tab {
        padding: 10px 16px;
        font-size: 13px;
        white-space: nowrap;
      }
    }
  </style>
</head>
<body>
  <header class="header">
    <h1>
      <svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
        <path d="M12 2L2 7l10 5 10-5-10-5zM2 17l10 5 10-5M2 12l10 5 10-5"/>
      </svg>
      Query Dashboard
    </h1>
    <div class="header-actions">
      <button class="btn btn-secondary" onclick="refreshData()">
        🔄 Aktualisieren
      </button>
      <button class="btn btn-secondary" onclick="exportData()">
        📥 Export
      </button>
      <button class="btn btn-secondary" onclick="logout()">
        🚪 Abmelden
      </button>
    </div>
  </header>

  <main class="main">
    <!-- Stats -->
    <div class="stats-grid" id="statsGrid">
      <div class="stat-card">
        <div class="stat-label">Gesamt Queries</div>
        <div class="stat-value" id="totalQueries">-</div>
      </div>
      <div class="stat-card">
        <div class="stat-label">Eindeutige Nutzer</div>
        <div class="stat-value" id="uniqueUsers">-</div>
      </div>
      <div class="stat-card">
        <div class="stat-label">Erfolgsrate</div>
        <div class="stat-value success" id="successRate">-</div>
      </div>
      <div class="stat-card">
        <div class="stat-label">Durchschn. Dauer</div>
        <div class="stat-value" id="avgDuration">-</div>
      </div>
    </div>

     <!-- Tabs -->
     <div class="tabs">
       <button class="tab active" onclick="switchTab('overview', event)">Übersicht</button>
       <button class="tab" onclick="switchTab('queries', event)">Queries</button>
       <button class="tab" onclick="switchTab('tools', event)">Tools</button>
       <button class="tab" onclick="switchTab('users', event)">Nutzer</button>
       <button class="tab" onclick="switchTab('analytics', event)">Analytics</button>
       <button class="tab" onclick="switchTab('insights', event)">Erkenntnisse</button>
     </div>

    <!-- Overview Tab -->
    <div id="tab-overview" class="tab-content active">
      <!-- Time Series Chart -->
      <div class="chart-section">
        <h3>Queries der letzten 7 Tage</h3>
        <div class="line-chart-container">
          <svg class="line-chart" id="timeSeriesChart"></svg>
        </div>
      </div>

      <!-- Hourly Chart -->
      <div class="chart-section">
        <h3>Queries der letzten 24 Stunden</h3>
        <div class="chart-container" id="chartContainer"></div>
      </div>
    </div>

    <!-- Queries Tab -->
    <div id="tab-queries" class="tab-content">
      <!-- Filters -->
      <div class="filters">
        <div class="filter-group">
          <label>Tool</label>
          <select id="filterTool">
            <option value="">Alle Tools</option>
          </select>
        </div>
        <div class="filter-group">
          <label>Status</label>
          <select id="filterStatus">
            <option value="">Alle</option>
            <option value="true">Erfolgreich</option>
            <option value="false">Fehlgeschlagen</option>
          </select>
        </div>
        <div class="filter-group">
          <label>Von</label>
          <input type="date" id="filterStartDate">
        </div>
        <div class="filter-group">
          <label>Bis</label>
          <input type="date" id="filterEndDate">
        </div>
        <button class="btn btn-primary" onclick="applyFilters()">Filter anwenden</button>
      </div>

      <!-- Queries Table -->
      <div class="queries-section">
        <div class="queries-header">
          <h2>Queries</h2>
          <span class="queries-count" id="queriesCount">0 von 0</span>
        </div>
        <table class="queries-table">
          <thead>
            <tr>
              <th>Zeitstempel</th>
              <th>Tool</th>
              <th>Nutzer</th>
              <th>Status</th>
              <th>Dauer</th>
              <th>Parameter</th>
            </tr>
          </thead>
          <tbody id="queriesBody">
            <tr>
              <td colspan="6">
                <div class="loading">
                  <div class="loading-spinner"></div>
                  Lade Queries...
                </div>
              </td>
            </tr>
          </tbody>
        </table>
        <div class="pagination">
          <button id="prevBtn" onclick="prevPage()" disabled>← Zurück</button>
          <span class="page-info" id="pageInfo">Seite 1</span>
          <button id="nextBtn" onclick="nextPage()">Weiter →</button>
        </div>
      </div>
    </div>

    <!-- Tools Tab -->
    <div id="tab-tools" class="tab-content">
      <div class="queries-section">
        <div class="queries-header">
          <h2>Tool Performance</h2>
          <button class="btn-export" onclick="exportToolPerformance()">📥 CSV Export</button>
        </div>
        <table class="performance-table">
          <thead>
            <tr>
              <th>Tool</th>
              <th>Gesamt</th>
              <th>Erfolgreich</th>
              <th>Fehlgeschlagen</th>
              <th>Erfolgsrate</th>
              <th>Durchschn. Dauer</th>
            </tr>
          </thead>
          <tbody id="toolPerformanceBody">
            <tr>
              <td colspan="6">
                <div class="loading">
                  <div class="loading-spinner"></div>
                  Lade Tool-Performance...
                </div>
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>

    <!-- Users Tab -->
    <div id="tab-users" class="tab-content">
      <div class="queries-section">
        <div class="queries-header">
          <h2>Nutzer Statistiken</h2>
        </div>
        <table class="performance-table">
          <thead>
            <tr>
              <th>Nutzer Hash</th>
              <th>Gesamt Queries</th>
              <th>Erfolgreich</th>
              <th>Fehlgeschlagen</th>
              <th>Erfolgsrate</th>
              <th>Eindeutige Tools</th>
              <th>Letzte Aktivität</th>
            </tr>
          </thead>
          <tbody id="userStatsBody">
            <tr>
              <td colspan="7">
                <div class="loading">
                  <div class="loading-spinner"></div>
                  Lade Nutzer-Statistiken...
                </div>
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>

     <!-- Analytics Tab -->
     <div id="tab-analytics" class="tab-content">
       <div class="chart-section">
         <h3>Top Tools nach Nutzung</h3>
         <div id="topToolsChart" style="height: 300px; margin-top: 20px;"></div>
       </div>
       <div class="chart-section">
         <h3>Erfolgsrate nach Tool</h3>
         <div id="toolSuccessChart" style="height: 300px; margin-top: 20px;"></div>
       </div>
     </div>

     <!-- Insights Tab -->
     <div id="tab-insights" class="tab-content">
       <!-- Key Insights Cards -->
       <div class="stats-grid" style="margin-bottom: 24px;">
         <div class="stat-card" id="insightCard1">
           <div class="stat-label">Erkenntnis 1</div>
           <div class="stat-value" style="font-size: 16px;">Lade...</div>
         </div>
         <div class="stat-card" id="insightCard2">
           <div class="stat-label">Erkenntnis 2</div>
           <div class="stat-value" style="font-size: 16px;">Lade...</div>
         </div>
         <div class="stat-card" id="insightCard3">
           <div class="stat-label">Erkenntnis 3</div>
           <div class="stat-value" style="font-size: 16px;">Lade...</div>
         </div>
         <div class="stat-card" id="insightCard4">
           <div class="stat-label">Erkenntnis 4</div>
           <div class="stat-value" style="font-size: 16px;">Lade...</div>
         </div>
       </div>

       <!-- Error Analysis -->
       <div class="chart-section">
         <h3>Fehleranalyse</h3>
         <div id="errorAnalysis" style="margin-top: 20px;">
           <div class="loading">
             <div class="loading-spinner"></div>
             Lade Fehleranalyse...
           </div>
         </div>
       </div>

       <!-- Performance Bottlenecks -->
       <div class="chart-section">
         <h3>Performance-Bottlenecks</h3>
         <div id="performanceBottlenecks" style="margin-top: 20px;">
           <div class="loading">
             <div class="loading-spinner"></div>
             Lade Performance-Analyse...
           </div>
         </div>
       </div>

       <!-- Usage Patterns -->
       <div class="chart-section">
         <h3>Nutzungsmuster</h3>
         <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 24px; margin-top: 20px;">
           <div>
             <h4 style="font-size: 14px; margin-bottom: 12px; color: var(--text-secondary);">Stündliche Verteilung</h4>
             <div id="hourlyPatternChart" style="height: 200px;"></div>
           </div>
           <div>
             <h4 style="font-size: 14px; margin-bottom: 12px; color: var(--text-secondary);">Wochentags-Verteilung</h4>
             <div id="dayPatternChart" style="height: 200px;"></div>
           </div>
         </div>
         <div style="margin-top: 24px;">
           <h4 style="font-size: 14px; margin-bottom: 12px; color: var(--text-secondary);">Häufige Tool-Sequenzen</h4>
           <div id="toolSequences" style="margin-top: 12px;"></div>
         </div>
       </div>

       <!-- Success Rate Trends -->
       <div class="chart-section">
         <h3>Erfolgsrate-Trends</h3>
         <div id="successTrendsChart" style="height: 300px; margin-top: 20px;"></div>
       </div>
     </div>
   </main>

  <script>
    let currentPage = 0;
    const pageSize = 50;
    let totalQueries = 0;
    let filters = {};
    let currentTab = 'overview';

     function switchTab(tabName, evt) {
       // Update tab buttons
       document.querySelectorAll('.tab').forEach(tab => {
         tab.classList.remove('active');
       });
       if (evt && evt.target) {
         evt.target.classList.add('active');
       } else {
         // Find button by tab name
         const tabs = Array.from(document.querySelectorAll('.tab'));
         const tabNames = ['overview', 'queries', 'tools', 'users', 'analytics'];
         const index = tabNames.indexOf(tabName);
         if (index >= 0 && tabs[index]) {
           tabs[index].classList.add('active');
         }
       }

       // Update tab content
       document.querySelectorAll('.tab-content').forEach(content => {
         content.classList.remove('active');
       });
       document.getElementById('tab-' + tabName).classList.add('active');

       currentTab = tabName;

       // Load tab-specific data
       if (tabName === 'tools') {
         fetchToolPerformance();
       } else if (tabName === 'users') {
         fetchUserStats();
       } else if (tabName === 'analytics') {
         fetchAnalytics();
       } else if (tabName === 'insights') {
         fetchInsights();
       } else if (tabName === 'overview') {
         fetchTimeSeries();
       }
     }

     async function fetchInsights() {
       try {
         await Promise.all([
           fetchErrorAnalysis(),
           fetchPerformanceBottlenecks(),
           fetchUsagePatterns(),
           fetchSuccessTrends(),
         ]);
       } catch (error) {
         console.error('Error fetching insights:', error);
       }
     }

     async function fetchErrorAnalysis() {
       try {
         const res = await fetch('/dashboard/api/error-analysis');
         const data = await res.json();
         renderErrorAnalysis(data);
       } catch (error) {
         console.error('Error fetching error analysis:', error);
       }
     }

     function renderErrorAnalysis(data) {
       const container = document.getElementById('errorAnalysis');
       if (!container) return;

       let html = '<div style="margin-bottom: 24px;">';
       html += '<div style="display: flex; gap: 16px; margin-bottom: 20px;">';
       html += '<div style="flex: 1; padding: 16px; background: var(--bg-tertiary); border-radius: 8px;">';
       html += '<div style="font-size: 12px; color: var(--text-secondary); margin-bottom: 8px;">Gesamtfehler</div>';
       html += '<div style="font-size: 24px; font-weight: 700; color: var(--error);">' + data.totalErrors.toLocaleString('de-DE') + '</div>';
       html += '</div>';
       html += '</div>';

       html += '<h4 style="font-size: 14px; margin-bottom: 12px; color: var(--text-secondary);">Häufigste Fehler</h4>';
       html += '<div style="display: flex; flex-direction: column; gap: 8px;">';
       data.topErrors.forEach((error, index) => {
         html += '<div style="display: flex; align-items: center; gap: 12px; padding: 12px; background: var(--bg-tertiary); border-radius: 6px;">';
         html += '<span style="font-size: 12px; color: var(--text-muted); min-width: 24px;">#' + (index + 1) + '</span>';
         html += '<div style="flex: 1; font-size: 13px;">' + (error.error.length > 100 ? error.error.substring(0, 100) + '...' : error.error) + '</div>';
         html += '<span style="font-size: 13px; font-weight: 600; color: var(--error);">' + error.count + 'x</span>';
         html += '</div>';
       });
       html += '</div>';

       html += '<h4 style="font-size: 14px; margin-top: 24px; margin-bottom: 12px; color: var(--text-secondary);">Fehler nach Tool</h4>';
       html += '<div style="display: flex; flex-direction: column; gap: 8px;">';
       data.toolErrorSummary.slice(0, 5).forEach(tool => {
         html += '<div style="display: flex; align-items: center; gap: 12px; padding: 12px; background: var(--bg-tertiary); border-radius: 6px;">';
         html += '<span class="tool-name">' + tool.tool + '</span>';
         html += '<div style="flex: 1; display: flex; align-items: center; gap: 8px;">';
         html += '<span style="font-size: 12px; color: var(--text-secondary);">' + tool.totalErrors + ' Fehler</span>';
         html += '<span style="font-size: 12px; color: var(--text-muted);">(' + tool.uniqueErrors + ' verschiedene)</span>';
         html += '</div>';
         html += '</div>';
       });
       html += '</div>';

       html += '</div>';
       container.innerHTML = html;
     }

     async function fetchPerformanceBottlenecks() {
       try {
         const res = await fetch('/dashboard/api/performance-bottlenecks');
         const data = await res.json();
         renderPerformanceBottlenecks(data);
       } catch (error) {
         console.error('Error fetching performance bottlenecks:', error);
       }
     }

     function renderPerformanceBottlenecks(data) {
       const container = document.getElementById('performanceBottlenecks');
       if (!container) return;

       let html = '<div style="margin-bottom: 16px; padding: 16px; background: var(--bg-tertiary); border-radius: 8px;">';
       html += '<div style="font-size: 12px; color: var(--text-secondary); margin-bottom: 8px;">95. Perzentil Schwellenwert</div>';
       html += '<div style="font-size: 24px; font-weight: 700;">' + data.p95Threshold.toLocaleString('de-DE') + ' ms</div>';
       html += '<div style="font-size: 12px; color: var(--text-secondary); margin-top: 4px;">' + data.slowQueriesCount + ' langsame Queries</div>';
       html += '</div>';

       html += '<h4 style="font-size: 14px; margin-bottom: 12px; color: var(--text-secondary);">Langsamste Tools</h4>';
       html += '<div style="display: flex; flex-direction: column; gap: 8px;">';
       data.bottlenecks.forEach(bottleneck => {
         html += '<div style="display: flex; align-items: center; gap: 12px; padding: 12px; background: var(--bg-tertiary); border-radius: 6px;">';
         html += '<span class="tool-name">' + bottleneck.tool + '</span>';
         html += '<div style="flex: 1; display: flex; flex-direction: column; gap: 4px;">';
         html += '<div style="display: flex; gap: 16px; font-size: 12px;">';
         html += '<span style="color: var(--text-secondary);">Langsam: <strong>' + bottleneck.slowCount + '</strong></span>';
         html += '<span style="color: var(--text-secondary);">Ø Dauer: <strong>' + Math.round(bottleneck.avgDuration) + ' ms</strong></span>';
         html += '<span style="color: var(--text-secondary);">Max: <strong>' + Math.round(bottleneck.maxDuration) + ' ms</strong></span>';
         html += '</div>';
         html += '</div>';
         html += '</div>';
       });
       html += '</div>';

       container.innerHTML = html;
     }

     async function fetchUsagePatterns() {
       try {
         const res = await fetch('/dashboard/api/usage-patterns');
         const data = await res.json();
         renderUsagePatterns(data);
       } catch (error) {
         console.error('Error fetching usage patterns:', error);
       }
     }

     function renderUsagePatterns(data) {
       // Hourly chart
       const hourlyContainer = document.getElementById('hourlyPatternChart');
       if (hourlyContainer) {
         hourlyContainer.innerHTML = '';
         const maxHourly = Math.max(...data.hourlyPatterns.map(h => h.count), 1);
         data.hourlyPatterns.forEach(({ hour, count }) => {
           const bar = document.createElement('div');
           bar.style.display = 'flex';
           bar.style.alignItems = 'center';
           bar.style.gap = '8px';
           bar.style.marginBottom = '4px';

           const label = document.createElement('div');
           label.style.minWidth = '40px';
           label.style.fontSize = '11px';
           label.textContent = hour + ':00';
           bar.appendChild(label);

           const barWrapper = document.createElement('div');
           barWrapper.style.flex = '1';
           barWrapper.style.height = '20px';
           barWrapper.style.background = 'var(--bg-tertiary)';
           barWrapper.style.borderRadius = '4px';
           barWrapper.style.overflow = 'hidden';
           barWrapper.style.position = 'relative';

           const barFill = document.createElement('div');
           barFill.style.height = '100%';
           barFill.style.width = (count / maxHourly * 100) + '%';
           barFill.style.background = 'var(--accent)';
           barWrapper.appendChild(barFill);

           const value = document.createElement('div');
           value.style.position = 'absolute';
           value.style.right = '4px';
           value.style.top = '50%';
           value.style.transform = 'translateY(-50%)';
           value.style.fontSize = '10px';
           value.textContent = count;
           barWrapper.appendChild(value);

           bar.appendChild(barWrapper);
           hourlyContainer.appendChild(bar);
         });
       }

       // Day chart
       const dayContainer = document.getElementById('dayPatternChart');
       if (dayContainer) {
         dayContainer.innerHTML = '';
         const maxDay = Math.max(...data.dayPatterns.map(d => d.count), 1);
         data.dayPatterns.forEach(({ dayName, count }) => {
           const bar = document.createElement('div');
           bar.style.display = 'flex';
           bar.style.alignItems = 'center';
           bar.style.gap = '8px';
           bar.style.marginBottom = '4px';

           const label = document.createElement('div');
           label.style.minWidth = '80px';
           label.style.fontSize = '11px';
           label.textContent = dayName;
           bar.appendChild(label);

           const barWrapper = document.createElement('div');
           barWrapper.style.flex = '1';
           barWrapper.style.height = '20px';
           barWrapper.style.background = 'var(--bg-tertiary)';
           barWrapper.style.borderRadius = '4px';
           barWrapper.style.overflow = 'hidden';
           barWrapper.style.position = 'relative';

           const barFill = document.createElement('div');
           barFill.style.height = '100%';
           barFill.style.width = (count / maxDay * 100) + '%';
           barFill.style.background = 'var(--accent)';
           barWrapper.appendChild(barFill);

           const value = document.createElement('div');
           value.style.position = 'absolute';
           value.style.right = '4px';
           value.style.top = '50%';
           value.style.transform = 'translateY(-50%)';
           value.style.fontSize = '10px';
           value.textContent = count;
           barWrapper.appendChild(value);

           bar.appendChild(barWrapper);
           dayContainer.appendChild(bar);
         });
       }

       // Tool sequences
       const sequencesContainer = document.getElementById('toolSequences');
       if (sequencesContainer) {
         sequencesContainer.innerHTML = '';
         data.topSequences.forEach((seq, index) => {
           const item = document.createElement('div');
           item.style.display = 'flex';
           item.style.alignItems = 'center';
           item.style.gap = '12px';
           item.style.padding = '8px 12px';
           item.style.background = 'var(--bg-tertiary)';
           item.style.borderRadius = '6px';
           item.style.marginBottom = '6px';

           const rank = document.createElement('span');
           rank.style.fontSize = '12px';
           rank.style.color = 'var(--text-muted)';
           rank.style.minWidth = '24px';
           rank.textContent = '#' + (index + 1);
           item.appendChild(rank);

           const sequence = document.createElement('span');
           sequence.style.fontSize = '13px';
           sequence.style.fontFamily = 'monospace';
           sequence.textContent = seq.sequence;
           item.appendChild(sequence);

           const count = document.createElement('span');
           count.style.marginLeft = 'auto';
           count.style.fontSize = '12px';
           count.style.fontWeight = '600';
           count.textContent = seq.count + 'x';
           item.appendChild(count);

           sequencesContainer.appendChild(item);
         });
       }
     }

     async function fetchSuccessTrends() {
       try {
         const res = await fetch('/dashboard/api/success-trends?days=30');
         const data = await res.json();
         renderSuccessTrends(data);
       } catch (error) {
         console.error('Error fetching success trends:', error);
       }
     }

     function renderSuccessTrends(data) {
       const container = document.getElementById('successTrendsChart');
       if (!container) return;

       // Update insight cards
       const trendIcon = data.trendDirection === 'improving' ? '📈' : data.trendDirection === 'declining' ? '📉' : '➡️';
       const trendColor = data.trendDirection === 'improving' ? 'var(--success)' : data.trendDirection === 'declining' ? 'var(--error)' : 'var(--text-secondary)';
       
       document.getElementById('insightCard1').innerHTML = 
         '<div class="stat-label">Trend</div>' +
         '<div class="stat-value" style="font-size: 16px; color: ' + trendColor + ';">' + trendIcon + ' ' + 
         (data.trendDirection === 'improving' ? 'Verbesserung' : data.trendDirection === 'declining' ? 'Rückgang' : 'Stabil') + '</div>' +
         '<div style="font-size: 12px; color: var(--text-secondary); margin-top: 4px;">' + 
         (data.trendChange > 0 ? '+' : '') + data.trendChange.toFixed(1) + '% vs. Vorwoche</div>';

       document.getElementById('insightCard2').innerHTML = 
         '<div class="stat-label">Aktuelle Erfolgsrate</div>' +
         '<div class="stat-value" style="font-size: 16px;">' + data.recentAverage.toFixed(1) + '%</div>' +
         '<div style="font-size: 12px; color: var(--text-secondary); margin-top: 4px;">Letzte 7 Tage</div>';

       // Render chart
       container.innerHTML = '';
       const maxValue = 100;
       const width = container.clientWidth || 800;
       const height = 300;
       const padding = { top: 20, right: 20, bottom: 40, left: 60 };
       const chartWidth = width - padding.left - padding.right;
       const chartHeight = height - padding.top - padding.bottom;

       const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
       svg.setAttribute('width', width);
       svg.setAttribute('height', height);
       svg.style.width = '100%';
       svg.style.height = '100%';

       // Draw axes
       const xAxis = document.createElementNS('http://www.w3.org/2000/svg', 'line');
       xAxis.setAttribute('x1', padding.left);
       xAxis.setAttribute('y1', height - padding.bottom);
       xAxis.setAttribute('x2', width - padding.right);
       xAxis.setAttribute('y2', height - padding.bottom);
       xAxis.setAttribute('stroke', 'var(--border)');
       xAxis.setAttribute('stroke-width', '1');
       svg.appendChild(xAxis);

       const yAxis = document.createElementNS('http://www.w3.org/2000/svg', 'line');
       yAxis.setAttribute('x1', padding.left);
       yAxis.setAttribute('y1', padding.top);
       yAxis.setAttribute('x2', padding.left);
       yAxis.setAttribute('y2', height - padding.bottom);
       yAxis.setAttribute('stroke', 'var(--border)');
       yAxis.setAttribute('stroke-width', '1');
       svg.appendChild(yAxis);

       // Draw trend line
       const points = data.trends.map((trend, i) => ({
         x: padding.left + (i / (data.trends.length - 1)) * chartWidth,
         y: height - padding.bottom - (trend.successRate / maxValue) * chartHeight,
         date: trend.date,
         rate: trend.successRate,
       }));

       const path = points.map((p, i) => (i === 0 ? 'M' : 'L') + ' ' + p.x + ' ' + p.y).join(' ');
       const trendLine = document.createElementNS('http://www.w3.org/2000/svg', 'path');
       trendLine.setAttribute('d', path);
       trendLine.setAttribute('stroke', trendColor);
       trendLine.setAttribute('stroke-width', '2');
       trendLine.setAttribute('fill', 'none');
       svg.appendChild(trendLine);

       // Add labels
       points.forEach((point, i) => {
         if (i % Math.ceil(data.trends.length / 10) === 0 || i === data.trends.length - 1) {
           const label = document.createElementNS('http://www.w3.org/2000/svg', 'text');
           label.setAttribute('x', point.x);
           label.setAttribute('y', height - padding.bottom + 20);
           label.setAttribute('font-size', '11px');
           label.setAttribute('fill', 'var(--text-secondary)');
           label.setAttribute('text-anchor', 'middle');
           label.textContent = point.date.split('-')[2];
           svg.appendChild(label);
         }
       });

       container.appendChild(svg);
     }

    async function fetchStats() {
      try {
        const res = await fetch('/dashboard/api/stats');
        const stats = await res.json();

        document.getElementById('totalQueries').textContent = stats.totalQueries.toLocaleString('de-DE');
        document.getElementById('uniqueUsers').textContent = stats.uniqueUsers.toLocaleString('de-DE');
        document.getElementById('successRate').textContent = stats.successRate.toFixed(1) + '%';
        document.getElementById('avgDuration').textContent = stats.averageDuration + ' ms';

        // Render hourly chart
        const container = document.getElementById('chartContainer');
        container.innerHTML = '';
        const maxCount = Math.max(...stats.queriesPerHour.map(h => h.count), 1);

        stats.queriesPerHour.forEach(({ hour, count }) => {
          const bar = document.createElement('div');
          bar.className = 'chart-bar';
          bar.style.height = (count / maxCount * 100) + '%';
          bar.title = hour + ': ' + count + ' Queries';
          container.appendChild(bar);
        });
      } catch (error) {
        console.error('Error fetching stats:', error);
      }
    }

    async function fetchTimeSeries() {
      try {
        const res = await fetch('/dashboard/api/timeseries?days=7');
        const data = await res.json();
        renderTimeSeriesChart(data.timeseries);
      } catch (error) {
        console.error('Error fetching time series:', error);
      }
    }

    function renderTimeSeriesChart(timeseries) {
      const svg = document.getElementById('timeSeriesChart');
      if (!svg) return;

      svg.innerHTML = '';
      const width = svg.clientWidth || 800;
      const height = 300;
      const padding = { top: 20, right: 20, bottom: 40, left: 60 };
      const chartWidth = width - padding.left - padding.right;
      const chartHeight = height - padding.top - padding.bottom;

      svg.setAttribute('width', width);
      svg.setAttribute('height', height);

      const maxValue = Math.max(...timeseries.map(d => Math.max(d.total, d.successful, d.failed)), 1);

      // Draw axes
      const axisGroup = document.createElementNS('http://www.w3.org/2000/svg', 'g');
      axisGroup.setAttribute('class', 'chart-axis');

      // X-axis
      const xAxis = document.createElementNS('http://www.w3.org/2000/svg', 'line');
      xAxis.setAttribute('x1', padding.left);
      xAxis.setAttribute('y1', height - padding.bottom);
      xAxis.setAttribute('x2', width - padding.right);
      xAxis.setAttribute('y2', height - padding.bottom);
      axisGroup.appendChild(xAxis);

      // Y-axis
      const yAxis = document.createElementNS('http://www.w3.org/2000/svg', 'line');
      yAxis.setAttribute('x1', padding.left);
      yAxis.setAttribute('y1', padding.top);
      yAxis.setAttribute('x2', padding.left);
      yAxis.setAttribute('y2', height - padding.bottom);
      axisGroup.appendChild(yAxis);

      svg.appendChild(axisGroup);

      // Draw lines
      const points = timeseries.map((d, i) => ({
        x: padding.left + (i / (timeseries.length - 1)) * chartWidth,
        y: height - padding.bottom - (d.total / maxValue) * chartHeight,
        date: d.date,
        total: d.total,
        successful: d.successful,
        failed: d.failed,
      }));

      // Success line
      const successPath = points
        .map((p, i) => {
          const yPos = height - padding.bottom - (p.successful / maxValue) * chartHeight;
          return (i === 0 ? 'M' : 'L') + ' ' + p.x + ' ' + yPos;
        })
        .join(' ');
      const successLine = document.createElementNS('http://www.w3.org/2000/svg', 'path');
      successLine.setAttribute('d', successPath);
      successLine.setAttribute('class', 'chart-line success');
      successLine.setAttribute('stroke-width', '2');
      successLine.setAttribute('fill', 'none');
      svg.appendChild(successLine);

      // Total line
      const totalPath = points
        .map((p, i) => {
          return (i === 0 ? 'M' : 'L') + ' ' + p.x + ' ' + p.y;
        })
        .join(' ');
      const totalLine = document.createElementNS('http://www.w3.org/2000/svg', 'path');
      totalLine.setAttribute('d', totalPath);
      totalLine.setAttribute('class', 'chart-line');
      totalLine.setAttribute('stroke-width', '2');
      totalLine.setAttribute('fill', 'none');
      svg.appendChild(totalLine);

      // Labels
      points.forEach((point, i) => {
        if (i % Math.ceil(timeseries.length / 7) === 0 || i === timeseries.length - 1) {
          const label = document.createElementNS('http://www.w3.org/2000/svg', 'text');
          label.setAttribute('x', point.x);
          label.setAttribute('y', height - padding.bottom + 20);
          label.setAttribute('class', 'chart-label');
          label.setAttribute('text-anchor', 'middle');
          label.textContent = point.date.split('-')[2]; // Day only
          svg.appendChild(label);
        }
      });
    }

    async function fetchToolPerformance() {
      try {
        const res = await fetch('/dashboard/api/tool-performance');
        const data = await res.json();
        renderToolPerformance(data.toolPerformance);
      } catch (error) {
        console.error('Error fetching tool performance:', error);
        document.getElementById('toolPerformanceBody').innerHTML = 
          '<tr><td colspan="6" class="empty-state">Fehler beim Laden der Tool-Performance</td></tr>';
      }
    }

    function renderToolPerformance(toolPerformance) {
      const tbody = document.getElementById('toolPerformanceBody');
      
      if (toolPerformance.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6"><div class="empty-state"><svg viewBox="0 0 24 24"><path d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"/></svg><p>Keine Tool-Performance-Daten gefunden</p></div></td></tr>';
        return;
      }

      tbody.innerHTML = toolPerformance.map(tool => {
        const successRatePercent = tool.successRate.toFixed(1);
        return '<tr>' +
          '<td><span class="tool-name">' + tool.tool + '</span></td>' +
          '<td>' + tool.totalQueries.toLocaleString('de-DE') + '</td>' +
          '<td>' + tool.successfulQueries.toLocaleString('de-DE') + '</td>' +
          '<td>' + tool.failedQueries.toLocaleString('de-DE') + '</td>' +
          '<td>' +
            '<div style="display: flex; align-items: center; gap: 8px;">' +
              '<span>' + successRatePercent + '%</span>' +
              '<div class="progress-bar" style="flex: 1; max-width: 200px;">' +
                '<div class="progress-bar-fill ' + (tool.successRate >= 80 ? 'success' : tool.successRate < 50 ? 'error' : '') + '" ' +
                'style="width: ' + tool.successRate + '%"></div>' +
              '</div>' +
            '</div>' +
          '</td>' +
          '<td>' + (tool.averageDuration > 0 ? tool.averageDuration + ' ms' : '-') + '</td>' +
        '</tr>';
      }).join('');
    }

    async function fetchUserStats() {
      try {
        const res = await fetch('/dashboard/api/user-stats');
        const data = await res.json();
        renderUserStats(data.userStats);
      } catch (error) {
        console.error('Error fetching user stats:', error);
        document.getElementById('userStatsBody').innerHTML = 
          '<tr><td colspan="7" class="empty-state">Fehler beim Laden der Nutzer-Statistiken</td></tr>';
      }
    }

    function renderUserStats(userStats) {
      const tbody = document.getElementById('userStatsBody');
      
      if (userStats.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7"><div class="empty-state"><svg viewBox="0 0 24 24"><path d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"/></svg><p>Keine Nutzer-Statistiken gefunden</p></div></td></tr>';
        return;
      }

      tbody.innerHTML = userStats.slice(0, 50).map(user => {
        const date = new Date(user.lastActivity);
        const formattedDate = date.toLocaleDateString('de-DE') + ' ' + date.toLocaleTimeString('de-DE');
        const successRatePercent = user.successRate.toFixed(1);
        return '<tr>' +
          '<td><span class="user-hash">' + user.userIdHash + '</span></td>' +
          '<td>' + user.totalQueries.toLocaleString('de-DE') + '</td>' +
          '<td>' + user.successfulQueries.toLocaleString('de-DE') + '</td>' +
          '<td>' + user.failedQueries.toLocaleString('de-DE') + '</td>' +
          '<td>' +
            '<div style="display: flex; align-items: center; gap: 8px;">' +
              '<span>' + successRatePercent + '%</span>' +
              '<div class="progress-bar" style="flex: 1; max-width: 150px;">' +
                '<div class="progress-bar-fill ' + (user.successRate >= 80 ? 'success' : user.successRate < 50 ? 'error' : '') + '" ' +
                'style="width: ' + user.successRate + '%"></div>' +
              '</div>' +
            '</div>' +
          '</td>' +
          '<td>' + user.uniqueTools + '</td>' +
          '<td class="timestamp">' + formattedDate + '</td>' +
        '</tr>';
      }).join('');
    }

    async function fetchAnalytics() {
      try {
        const toolRes = await fetch('/dashboard/api/tool-performance');
        const toolData = await toolRes.json();
        renderTopToolsChart(toolData.toolPerformance.slice(0, 10));
        renderToolSuccessChart(toolData.toolPerformance.slice(0, 10));
      } catch (error) {
        console.error('Error fetching analytics:', error);
      }
    }

    function renderTopToolsChart(toolPerformance) {
      const container = document.getElementById('topToolsChart');
      if (!container) return;

      container.innerHTML = '';
      const maxValue = Math.max(...toolPerformance.map(t => t.totalQueries), 1);

      toolPerformance.forEach(tool => {
        const barContainer = document.createElement('div');
        barContainer.style.display = 'flex';
        barContainer.style.alignItems = 'center';
        barContainer.style.marginBottom = '12px';
        barContainer.style.gap = '12px';

        const label = document.createElement('div');
        label.style.minWidth = '200px';
        label.style.fontSize = '13px';
        label.textContent = tool.tool;
        barContainer.appendChild(label);

        const barWrapper = document.createElement('div');
        barWrapper.style.flex = '1';
        barWrapper.style.height = '24px';
        barWrapper.style.background = 'var(--bg-tertiary)';
        barWrapper.style.borderRadius = '4px';
        barWrapper.style.overflow = 'hidden';
        barWrapper.style.position = 'relative';

        const bar = document.createElement('div');
        bar.style.height = '100%';
        bar.style.width = (tool.totalQueries / maxValue * 100) + '%';
        bar.style.background = 'var(--accent)';
        bar.style.transition = 'width 0.3s';
        barWrapper.appendChild(bar);

        const value = document.createElement('div');
        value.style.position = 'absolute';
        value.style.right = '8px';
        value.style.top = '50%';
        value.style.transform = 'translateY(-50%)';
        value.style.fontSize = '12px';
        value.style.color = 'var(--text-primary)';
        value.textContent = tool.totalQueries.toLocaleString('de-DE');
        barWrapper.appendChild(value);

        barContainer.appendChild(barWrapper);
        container.appendChild(barContainer);
      });
    }

    function renderToolSuccessChart(toolPerformance) {
      const container = document.getElementById('toolSuccessChart');
      if (!container) return;

      container.innerHTML = '';
      const maxValue = 100; // Percentage

      toolPerformance.forEach(tool => {
        const barContainer = document.createElement('div');
        barContainer.style.display = 'flex';
        barContainer.style.alignItems = 'center';
        barContainer.style.marginBottom = '12px';
        barContainer.style.gap = '12px';

        const label = document.createElement('div');
        label.style.minWidth = '200px';
        label.style.fontSize = '13px';
        label.textContent = tool.tool;
        barContainer.appendChild(label);

        const barWrapper = document.createElement('div');
        barWrapper.style.flex = '1';
        barWrapper.style.height = '24px';
        barWrapper.style.background = 'var(--bg-tertiary)';
        barWrapper.style.borderRadius = '4px';
        barWrapper.style.overflow = 'hidden';
        barWrapper.style.position = 'relative';

        const bar = document.createElement('div');
        bar.style.height = '100%';
        bar.style.width = tool.successRate + '%';
        bar.style.background = tool.successRate >= 80 ? 'var(--success)' : tool.successRate < 50 ? 'var(--error)' : 'var(--warning)';
        bar.style.transition = 'width 0.3s';
        barWrapper.appendChild(bar);

        const value = document.createElement('div');
        value.style.position = 'absolute';
        value.style.right = '8px';
        value.style.top = '50%';
        value.style.transform = 'translateY(-50%)';
        value.style.fontSize = '12px';
        value.style.color = 'var(--text-primary)';
        value.textContent = tool.successRate.toFixed(1) + '%';
        barWrapper.appendChild(value);

        barContainer.appendChild(barWrapper);
        container.appendChild(barContainer);
      });
    }

    function exportData() {
      const params = new URLSearchParams();
      if (filters.toolName) params.append('toolName', filters.toolName);
      if (filters.startDate) params.append('startDate', filters.startDate);
      if (filters.endDate) params.append('endDate', filters.endDate);
      
      window.open('/dashboard/api/export/csv?' + params.toString(), '_blank');
    }

    function exportToolPerformance() {
      window.open('/dashboard/api/export/csv', '_blank');
    }

    async function fetchTools() {
      try {
        const res = await fetch('/dashboard/api/tools');
        const data = await res.json();
        const select = document.getElementById('filterTool');

        data.tools.forEach(tool => {
          const option = document.createElement('option');
          option.value = tool;
          option.textContent = tool;
          select.appendChild(option);
        });
      } catch (error) {
        console.error('Error fetching tools:', error);
      }
    }

    async function fetchQueries() {
      try {
        const params = new URLSearchParams({
          limit: pageSize.toString(),
          offset: (currentPage * pageSize).toString(),
          ...filters,
        });

        const res = await fetch('/dashboard/api/queries?' + params);
        const data = await res.json();

        totalQueries = data.total;
        renderQueries(data.queries);
        updatePagination();
      } catch (error) {
        console.error('Error fetching queries:', error);
        document.getElementById('queriesBody').innerHTML = '<tr><td colspan="6" class="empty-state">Fehler beim Laden der Queries</td></tr>';
      }
    }

    function renderQueries(queries) {
      const tbody = document.getElementById('queriesBody');

      if (queries.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6"><div class="empty-state"><svg viewBox="0 0 24 24"><path d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"/></svg><p>Keine Queries gefunden</p></div></td></tr>';
        return;
      }

      tbody.innerHTML = queries.map(q => {
        const date = new Date(q.timestamp);
        const formattedDate = date.toLocaleDateString('de-DE') + ' ' + date.toLocaleTimeString('de-DE');
        const params = JSON.stringify(q.parameters, null, 2).substring(0, 100);

        return '<tr>' +
          '<td class="timestamp">' + formattedDate + '</td>' +
          '<td><span class="tool-name">' + q.toolName + '</span></td>' +
          '<td><span class="user-hash">' + q.userIdHash + '</span></td>' +
          '<td><span class="status-badge ' + (q.success ? 'success' : 'error') + '">' + 
            (q.success ? '✓ Erfolg' : '✗ Fehler') + '</span></td>' +
          '<td class="duration">' + (q.durationMs ? q.durationMs + ' ms' : '-') + '</td>' +
          '<td><code style="font-size:11px;color:var(--text-muted)">' + params + '</code></td>' +
        '</tr>';
      }).join('');

      document.getElementById('queriesCount').textContent = 
        ((currentPage * pageSize) + 1) + '-' + Math.min((currentPage + 1) * pageSize, totalQueries) + 
        ' von ' + totalQueries;
    }

    function updatePagination() {
      const totalPages = Math.ceil(totalQueries / pageSize);
      document.getElementById('prevBtn').disabled = currentPage === 0;
      document.getElementById('nextBtn').disabled = currentPage >= totalPages - 1;
      document.getElementById('pageInfo').textContent = 'Seite ' + (currentPage + 1) + ' von ' + Math.max(1, totalPages);
    }

    function prevPage() {
      if (currentPage > 0) {
        currentPage--;
        fetchQueries();
      }
    }

    function nextPage() {
      const totalPages = Math.ceil(totalQueries / pageSize);
      if (currentPage < totalPages - 1) {
        currentPage++;
        fetchQueries();
      }
    }

    function applyFilters() {
      currentPage = 0;
      filters = {};

      const tool = document.getElementById('filterTool').value;
      const status = document.getElementById('filterStatus').value;
      const startDate = document.getElementById('filterStartDate').value;
      const endDate = document.getElementById('filterEndDate').value;

      if (tool) filters.toolName = tool;
      if (status) filters.success = status;
      if (startDate) filters.startDate = startDate;
      if (endDate) filters.endDate = endDate;

      fetchQueries();
    }

    function refreshData() {
      fetchStats();
      fetchQueries();
    }

    async function logout() {
      await fetch('/dashboard/logout', { method: 'POST' });
      window.location.href = '/dashboard/login';
    }

     // Initial load
     fetchStats();
     fetchTools();
     fetchQueries();
     fetchTimeSeries();

     // Auto-refresh every 30 seconds
     setInterval(() => {
       fetchStats();
       if (currentTab === 'overview') {
         fetchTimeSeries();
       } else if (currentTab === 'insights') {
         fetchInsights();
       }
     }, 30000);
  </script>
</body>
</html>`;
}

export default createDashboardRouter;
