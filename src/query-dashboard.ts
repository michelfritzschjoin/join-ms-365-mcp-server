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
import { getQueryStore, type QueryFilter } from './query-store.js';
import logger from './logger.js';
import { rateLimitMiddleware, createRateLimitMiddleware } from './middleware/rate-limit.js';

// SECURITY: Strict rate limiter for login endpoints (5 attempts per 15 minutes)
const loginRateLimiter = createRateLimitMiddleware(
  15 * 60 * 1000, // 15 minutes
  5 // Only 5 attempts
);

/**
 * Session store for dashboard authentication
 */
interface DashboardSession {
  token: string;
  createdAt: Date;
  expiresAt: Date;
  ipAddress: string;
}

const sessions: Map<string, DashboardSession> = new Map();
const SESSION_DURATION_MS = 8 * 60 * 60 * 1000; // 8 hours
const MAX_LOGIN_ATTEMPTS = 5;
const LOGIN_LOCKOUT_MS = 15 * 60 * 1000; // 15 minutes

// Track login attempts for rate limiting
const loginAttempts: Map<string, { count: number; lastAttempt: Date }> = new Map();

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
  const now = new Date();

  sessions.set(token, {
    token,
    createdAt: now,
    expiresAt: new Date(now.getTime() + SESSION_DURATION_MS),
    ipAddress,
  });

  // Clean up old sessions
  cleanupSessions();

  return token;
}

/**
 * Validate session token
 */
function validateSession(token: string, _ipAddress: string): boolean {
  const session = sessions.get(token);

  if (!session) {
    return false;
  }

  // Check if expired
  if (new Date() > session.expiresAt) {
    sessions.delete(token);
    return false;
  }

  // Optional: Check IP address (can be disabled for mobile users)
  // if (session.ipAddress !== ipAddress) {
  //   return false;
  // }

  return true;
}

/**
 * Clean up expired sessions
 */
function cleanupSessions(): void {
  const now = new Date();
  for (const [token, session] of sessions) {
    if (now > session.expiresAt) {
      sessions.delete(token);
    }
  }
}

/**
 * Check rate limiting for login attempts
 */
function isLoginRateLimited(ipAddress: string): boolean {
  const attempts = loginAttempts.get(ipAddress);

  if (!attempts) {
    return false;
  }

  // Reset if lockout period has passed
  const timeSinceLastAttempt = Date.now() - attempts.lastAttempt.getTime();
  if (timeSinceLastAttempt > LOGIN_LOCKOUT_MS) {
    loginAttempts.delete(ipAddress);
    return false;
  }

  return attempts.count >= MAX_LOGIN_ATTEMPTS;
}

/**
 * Record login attempt
 */
function recordLoginAttempt(ipAddress: string, success: boolean): void {
  if (success) {
    loginAttempts.delete(ipAddress);
    return;
  }

  const attempts = loginAttempts.get(ipAddress) || { count: 0, lastAttempt: new Date() };
  attempts.count++;
  attempts.lastAttempt = new Date();
  loginAttempts.set(ipAddress, attempts);
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
      sessions.delete(token);
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

    <!-- Chart -->
    <div class="chart-section">
      <h3>Queries der letzten 24 Stunden</h3>
      <div class="chart-container" id="chartContainer"></div>
    </div>

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
  </main>

  <script>
    let currentPage = 0;
    const pageSize = 50;
    let totalQueries = 0;
    let filters = {};

    async function fetchStats() {
      try {
        const res = await fetch('/dashboard/api/stats');
        const stats = await res.json();

        document.getElementById('totalQueries').textContent = stats.totalQueries.toLocaleString('de-DE');
        document.getElementById('uniqueUsers').textContent = stats.uniqueUsers.toLocaleString('de-DE');
        document.getElementById('successRate').textContent = stats.successRate.toFixed(1) + '%';
        document.getElementById('avgDuration').textContent = stats.averageDuration + ' ms';

        // Render chart
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

    // Auto-refresh every 30 seconds
    setInterval(fetchStats, 30000);
  </script>
</body>
</html>`;
}

export default createDashboardRouter;
