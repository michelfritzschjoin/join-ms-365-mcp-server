/**
 * Dashboard Session Store - Persistent storage for dashboard sessions and login attempts
 *
 * Stores dashboard authentication sessions and login attempt tracking
 * to survive server restarts.
 *
 * Security Features:
 * - Sessions are persisted to disk
 * - Expired sessions are automatically cleaned up
 * - Login attempts are tracked for rate limiting
 * - Data is stored in /app/data directory
 *
 * ISO 27001 Compliance:
 * - Session data is stored securely
 * - Expired sessions are automatically removed
 * - Login attempt tracking for security monitoring
 */

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import logger from './logger.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

/**
 * Dashboard session data
 */
export interface DashboardSession {
  token: string;
  createdAt: string; // ISO 8601 timestamp
  expiresAt: string; // ISO 8601 timestamp
  ipAddress: string;
}

/**
 * Login attempt tracking data
 */
export interface LoginAttempt {
  count: number;
  lastAttempt: string; // ISO 8601 timestamp
}

/**
 * Dashboard session store data structure
 */
interface DashboardSessionStoreData {
  sessions: Record<string, DashboardSession>;
  loginAttempts: Record<string, LoginAttempt>;
  lastCleanup: string; // ISO 8601 timestamp
}

/**
 * Dashboard Session Store class - manages persistent session storage
 */
export class DashboardSessionStore {
  private dataDir: string;
  private sessionsFile: string;
  private sessions: Map<string, DashboardSession> = new Map();
  private loginAttempts: Map<string, LoginAttempt> = new Map();
  private saveDebounceTimer: ReturnType<typeof setTimeout> | null = null;
  private cleanupInterval: ReturnType<typeof setInterval> | null = null;

  constructor() {
    // Use same data directory as query store
    this.dataDir = process.env.QUERY_STORE_DIR || path.join(__dirname, '..', 'data');
    this.sessionsFile = path.join(this.dataDir, 'dashboard-sessions.json');

    this.ensureDataDir();
    this.loadData();
    this.startPeriodicCleanup();
  }

  /**
   * Ensure data directory exists
   */
  private ensureDataDir(): void {
    if (!fs.existsSync(this.dataDir)) {
      fs.mkdirSync(this.dataDir, { recursive: true });
      logger.info('Dashboard session store data directory created', { path: this.dataDir });
    }
  }

  /**
   * Load sessions and login attempts from disk
   */
  private loadData(): void {
    try {
      if (fs.existsSync(this.sessionsFile)) {
        const data = fs.readFileSync(this.sessionsFile, 'utf-8');
        const parsed: DashboardSessionStoreData = JSON.parse(data);

        // Load sessions
        if (parsed.sessions) {
          for (const [token, session] of Object.entries(parsed.sessions)) {
            // Only load non-expired sessions
            if (new Date(session.expiresAt) > new Date()) {
              this.sessions.set(token, session);
            }
          }
          logger.info('Dashboard sessions loaded', { count: this.sessions.size });
        }

        // Load login attempts
        if (parsed.loginAttempts) {
          for (const [ip, attempt] of Object.entries(parsed.loginAttempts)) {
            // Only load recent attempts (within lockout period)
            const lockoutMs = 15 * 60 * 1000; // 15 minutes
            const timeSinceLastAttempt = Date.now() - new Date(attempt.lastAttempt).getTime();
            if (timeSinceLastAttempt < lockoutMs) {
              this.loginAttempts.set(ip, attempt);
            }
          }
          logger.info('Dashboard login attempts loaded', { count: this.loginAttempts.size });
        }
      }
    } catch (error) {
      logger.error('Failed to load dashboard session store:', error);
      this.sessions = new Map();
      this.loginAttempts = new Map();
    }
  }

  /**
   * Save sessions and login attempts to disk (debounced)
   */
  private saveData(): void {
    if (this.saveDebounceTimer) {
      clearTimeout(this.saveDebounceTimer);
    }

    this.saveDebounceTimer = setTimeout(() => {
      try {
        const data: DashboardSessionStoreData = {
          sessions: Object.fromEntries(this.sessions),
          loginAttempts: Object.fromEntries(this.loginAttempts),
          lastCleanup: new Date().toISOString(),
        };

        fs.writeFileSync(this.sessionsFile, JSON.stringify(data, null, 2), 'utf-8');
        logger.debug('Dashboard session store saved', {
          sessions: this.sessions.size,
          loginAttempts: this.loginAttempts.size,
        });
      } catch (error) {
        logger.error('Failed to save dashboard session store:', error);
      }
    }, 1000); // Debounce for 1 second
  }

  /**
   * Start periodic cleanup of expired sessions
   */
  private startPeriodicCleanup(): void {
    // Run cleanup every hour
    this.cleanupInterval = setInterval(
      () => {
        this.cleanupExpiredSessions();
      },
      60 * 60 * 1000
    );

    // Also run on startup
    this.cleanupExpiredSessions();
  }

  /**
   * Remove expired sessions and old login attempts
   */
  private cleanupExpiredSessions(): void {
    const now = new Date();
    let removedSessions = 0;
    let removedAttempts = 0;

    // Remove expired sessions
    for (const [token, session] of this.sessions.entries()) {
      if (new Date(session.expiresAt) <= now) {
        this.sessions.delete(token);
        removedSessions++;
      }
    }

    // Remove old login attempts (older than lockout period)
    const lockoutMs = 15 * 60 * 1000; // 15 minutes
    for (const [ip, attempt] of this.loginAttempts.entries()) {
      const timeSinceLastAttempt = Date.now() - new Date(attempt.lastAttempt).getTime();
      if (timeSinceLastAttempt >= lockoutMs) {
        this.loginAttempts.delete(ip);
        removedAttempts++;
      }
    }

    if (removedSessions > 0 || removedAttempts > 0) {
      logger.debug('Dashboard session store cleanup completed', {
        removedSessions,
        removedAttempts,
        remainingSessions: this.sessions.size,
        remainingAttempts: this.loginAttempts.size,
      });
      this.saveData();
    }
  }

  /**
   * Create a new session
   */
  public createSession(token: string, ipAddress: string, expiresAt: Date): void {
    const session: DashboardSession = {
      token,
      createdAt: new Date().toISOString(),
      expiresAt: expiresAt.toISOString(),
      ipAddress,
    };

    this.sessions.set(token, session);
    this.saveData();
  }

  /**
   * Get a session by token
   */
  public getSession(token: string): DashboardSession | undefined {
    const session = this.sessions.get(token);
    if (!session) {
      return undefined;
    }

    // Check if expired
    if (new Date(session.expiresAt) <= new Date()) {
      this.sessions.delete(token);
      this.saveData();
      return undefined;
    }

    return session;
  }

  /**
   * Delete a session
   */
  public deleteSession(token: string): void {
    if (this.sessions.delete(token)) {
      this.saveData();
    }
  }

  /**
   * Check if a session exists and is valid
   */
  public hasValidSession(token: string): boolean {
    const session = this.getSession(token);
    return session !== undefined;
  }

  /**
   * Get login attempt data for an IP address
   */
  public getLoginAttempt(ipAddress: string): LoginAttempt | undefined {
    const attempt = this.loginAttempts.get(ipAddress);
    if (!attempt) {
      return undefined;
    }

    // Check if lockout period has passed
    const lockoutMs = 15 * 60 * 1000; // 15 minutes
    const timeSinceLastAttempt = Date.now() - new Date(attempt.lastAttempt).getTime();
    if (timeSinceLastAttempt >= lockoutMs) {
      this.loginAttempts.delete(ipAddress);
      this.saveData();
      return undefined;
    }

    return attempt;
  }

  /**
   * Record a login attempt
   */
  public recordLoginAttempt(ipAddress: string, success: boolean): void {
    if (success) {
      // Clear attempts on successful login
      if (this.loginAttempts.delete(ipAddress)) {
        this.saveData();
      }
      return;
    }

    // Increment failed attempt
    const existing = this.loginAttempts.get(ipAddress);
    const attempt: LoginAttempt = {
      count: existing ? existing.count + 1 : 1,
      lastAttempt: new Date().toISOString(),
    };

    this.loginAttempts.set(ipAddress, attempt);
    this.saveData();
  }

  /**
   * Check if an IP address is rate limited
   */
  public isRateLimited(ipAddress: string, maxAttempts: number): boolean {
    const attempt = this.getLoginAttempt(ipAddress);
    if (!attempt) {
      return false;
    }

    return attempt.count >= maxAttempts;
  }

  /**
   * Get all active sessions (for debugging/admin purposes)
   */
  public getAllSessions(): DashboardSession[] {
    return Array.from(this.sessions.values());
  }

  /**
   * Clear all sessions (admin function)
   */
  public clearAllSessions(): void {
    this.sessions.clear();
    this.saveData();
    logger.warn('All dashboard sessions cleared');
  }

  /**
   * Clear all login attempts (admin function)
   */
  public clearAllLoginAttempts(): void {
    this.loginAttempts.clear();
    this.saveData();
    logger.warn('All dashboard login attempts cleared');
  }

  /**
   * Cleanup and save before shutdown
   */
  public shutdown(): void {
    if (this.cleanupInterval) {
      clearInterval(this.cleanupInterval);
    }
    if (this.saveDebounceTimer) {
      clearTimeout(this.saveDebounceTimer);
    }
    this.cleanupExpiredSessions();
    this.saveData();
  }
}

// Singleton instance
let dashboardSessionStoreInstance: DashboardSessionStore | null = null;

/**
 * Get the dashboard session store singleton
 */
export function getDashboardSessionStore(): DashboardSessionStore {
  if (!dashboardSessionStoreInstance) {
    dashboardSessionStoreInstance = new DashboardSessionStore();
  }
  return dashboardSessionStoreInstance;
}

export default getDashboardSessionStore;
