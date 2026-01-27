#!/usr/bin/env node

// Load environment variables from .env file (optional)
// Environment variables can come from .env file, system environment, Azure Key Vault, Docker, etc.
import { config } from 'dotenv';
import { resolve } from 'path';
import { existsSync } from 'fs';

// Try to load .env file from project root if it exists
const envPath = resolve(process.cwd(), '.env');
if (existsSync(envPath)) {
  const result = config({ path: envPath });
  if (result.error) {
    console.warn(`Warning: Error loading .env file: ${result.error.message}`);
  }
} else {
  // Fallback: try to load from default location (silently, no warning if not found)
  config();
  // Note: Environment variables may be set from other sources (system env, Docker, Key Vault, etc.)
  // No warning needed if .env file doesn't exist
}

import { parseArgs } from './cli.js';
import logger from './logger.js';
import AuthManager, { buildScopesFromEndpoints } from './auth.js';
import MicrosoftGraphServer from './server.js';
import { version } from './version.js';
import { initializeLicenseCheck } from './license-check.js';

/**
 * Masks sensitive values for display purposes
 * Shows first 4 and last 4 characters, masks the rest
 */
function maskSensitiveValue(value: string | undefined): string {
  if (!value) {
    return '(not set)';
  }

  // For short values, mask completely
  if (value.length <= 12) {
    return '***';
  }

  // Show first 4 and last 4 characters
  const prefix = value.substring(0, 4);
  const suffix = value.substring(value.length - 4);
  const maskedLength = value.length - 8;
  const masked = '*'.repeat(Math.min(maskedLength, 20)); // Max 20 asterisks

  return `${prefix}${masked}${suffix}`;
}

/**
 * Lists all relevant environment variables with masked sensitive values
 */
function displayEnvironmentVariables(): void {
  // List of all relevant environment variables
  const relevantVars = [
    // Authentication & Azure Configuration
    'MS365_MCP_CLIENT_ID',
    'MS365_MCP_TENANT_ID',
    'MS365_MCP_CLIENT_SECRET', // Sensitive
    'MS365_MCP_CLOUD_TYPE',
    'MS365_MCP_OAUTH_TOKEN', // Sensitive
    'MS365_MCP_KEYVAULT_URL',
    // Server Mode & Behavior
    'READ_ONLY',
    'MS365_MCP_ORG_MODE',
    'ENABLED_TOOLS',
    'MS365_MCP_OUTPUT_FORMAT',
    'MS365_MCP_ENABLE_DISCOVERY_TOOLS',
    'MS365_MCP_FORCE_WORK_SCOPES',
    // Performance & Limits
    'MS365_MCP_MAX_RESULTS',
    'MS365_MCP_MAX_PAGES',
    'MS365_MCP_MAX_AGGREGATE_ITEMS',
    'MS365_MCP_MAX_CONCURRENT_TOOLS',
    'MS365_MCP_MAX_REPAIR_ATTEMPTS',
    // Learning & AI Features
    'MS365_MCP_LEARNING_ENABLED',
    'MS365_MCP_LEARNING_DECAY_DAYS',
    'MS365_MCP_LEARNING_DECAY_FACTOR',
    'MS365_MCP_LEARNING_CLUSTER_ENABLED',
    'MS365_MCP_LEARNING_NLP_ENABLED',
    'MS365_MCP_KNOWLEDGE_BASE_PATH',
    // Deep Research
    'MS365_MCP_DEEP_RESEARCH_MAX_DEPTH',
    'MS365_MCP_MAX_RESEARCH_ITERATIONS',
    'MS365_MCP_DEEP_RESEARCH_ITEMS_PER_ITERATION',
    // CORS & Security
    'MS365_MCP_CORS_ORIGINS',
    'MS365_MCP_CORS_ORIGIN',
    'MS365_MCP_CORS_METHODS',
    'MS365_MCP_CORS_HEADERS',
    'MS365_MCP_X_FRAME_OPTIONS',
    'MS365_MCP_REFERRER_POLICY',
    'MS365_MCP_CSP',
    'MS365_MCP_HSTS_MAX_AGE',
    // Rate Limiting
    'MS365_MCP_RATE_LIMIT_WINDOW_MS',
    'MS365_MCP_RATE_LIMIT_MAX_REQUESTS',
    // Self-Repair
    'MS365_MCP_ENABLE_SELF_REPAIR',
    'MS365_MCP_REPAIR_STRATEGIES',
    'MS365_MCP_STOP_ON_ERROR',
    // Query & Search
    'MS365_MCP_MAX_QUERY_VARIANTS',
    // Logging
    'LOG_LEVEL',
    'LOG_FORMAT',
    'SILENT',
    'DEBUG_REQUESTS',
    // General
    'NODE_ENV',
    'TRUST_PROXY_COUNT',
    // License (Sensitive)
    'CGPT_JOIN_LICENSE', // Sensitive
  ];

  // Sensitive variables that should be masked
  const sensitiveVars = new Set([
    'MS365_MCP_CLIENT_SECRET',
    'MS365_MCP_OAUTH_TOKEN',
    'CGPT_JOIN_LICENSE',
  ]);

  // CRITICAL: MCP STDIO servers MUST NOT write to stdout
  // Use console.error (stderr) for all diagnostic output
  console.error('\n\x1b[33m═══════════════════════════════════════════════════════════════\x1b[0m');
  console.error('\x1b[33m  Environment Variables (Masked)\x1b[0m');
  console.error('\x1b[33m═══════════════════════════════════════════════════════════════\x1b[0m\n');

  let hasVariables = false;
  for (const varName of relevantVars) {
    const value = process.env[varName];
    if (value !== undefined) {
      hasVariables = true;
      const displayValue = sensitiveVars.has(varName) ? maskSensitiveValue(value) : value;
      console.error(`  \x1b[36m${varName.padEnd(40)}\x1b[0m = ${displayValue}`);
    }
  }

  if (!hasVariables) {
    console.error('  \x1b[90m(No relevant environment variables set)\x1b[0m');
  }

  console.error(
    '\n\x1b[33m═══════════════════════════════════════════════════════════════\x1b[0m\n'
  );
}

/**
 * Display ASCII art banner for ki.join.de
 */
function displayBanner(): void {
  const versionLine = `Version ${version}`.padEnd(63);
  const banner = `
╔═══════════════════════════════════════════════════════════════╗
║                                                                 ║
║     ██╗  ██╗██╗      ██╗ ██████╗ ██╗███╗   ██╗      ██████╗    ║
║     ██║ ██╔╝██║      ██║██╔═══██╗██║████╗  ██║     ██╔═══██╗   ║
║     █████╔╝ ██║█████╗██║██║   ██║██║██╔██╗ ██║     ██║   ██║   ║
║     ██╔═██╗ ██║╚════╝██║██║   ██║██║██║╚██╗██║     ██║   ██║   ║
║     ██║  ██╗██║      ██║╚██████╔╝██║██║ ╚████║     ╚██████╔╝   ║
║     ╚═╝  ╚═╝╚═╝      ╚═╝ ╚═════╝ ╚═╝╚═╝  ╚═══╝      ╚═════╝    ║
║                                                                 ║
║                    Microsoft 365 MCP Server                     ║
║                    Join GmbH - ki.join.de                         ║
║                         ${versionLine}║
║                                                                 ║
╚═══════════════════════════════════════════════════════════════╝
`;

  // CRITICAL: MCP STDIO servers MUST NOT write to stdout
  // Use console.error (stderr) for all diagnostic output
  console.error('\x1b[36m%s\x1b[0m', banner); // Cyan color
}

async function main(): Promise<void> {
  try {
    // Initialize license validation (internal)
    initializeLicenseCheck();

    // Display banner on startup
    displayBanner();

    // Display environment variables (masked)
    displayEnvironmentVariables();

    const args = parseArgs();

    const includeWorkScopes = args.orgMode || false;
    if (includeWorkScopes) {
      logger.info('Organization mode enabled - including work account scopes');
    }

    const scopes = buildScopesFromEndpoints(includeWorkScopes, args.enabledTools);
    const authManager = await AuthManager.create(scopes);
    await authManager.loadTokenCache();

    if (args.login) {
      await authManager.acquireTokenByDeviceCode();
      logger.info('Login completed, testing connection with Graph API...');
      const result = await authManager.testLogin();
      console.log(JSON.stringify(result));
      process.exit(0);
    }

    if (args.verifyLogin) {
      logger.info('Verifying login...');
      const result = await authManager.testLogin();
      console.log(JSON.stringify(result));
      process.exit(0);
    }

    if (args.logout) {
      await authManager.logout();
      console.log(JSON.stringify({ message: 'Logged out successfully' }));
      process.exit(0);
    }

    if (args.listAccounts) {
      const accounts = await authManager.listAccounts();
      const selectedAccountId = authManager.getSelectedAccountId();
      const result = accounts.map((account) => ({
        id: account.homeAccountId,
        username: account.username,
        name: account.name,
        selected: account.homeAccountId === selectedAccountId,
      }));
      console.log(JSON.stringify({ accounts: result }));
      process.exit(0);
    }

    if (args.selectAccount) {
      const success = await authManager.selectAccount(args.selectAccount);
      if (success) {
        console.log(JSON.stringify({ message: `Selected account: ${args.selectAccount}` }));
      } else {
        console.log(JSON.stringify({ error: `Account not found: ${args.selectAccount}` }));
        process.exit(1);
      }
      process.exit(0);
    }

    if (args.removeAccount) {
      const success = await authManager.removeAccount(args.removeAccount);
      if (success) {
        console.log(JSON.stringify({ message: `Removed account: ${args.removeAccount}` }));
      } else {
        console.log(JSON.stringify({ error: `Account not found: ${args.removeAccount}` }));
        process.exit(1);
      }
      process.exit(0);
    }

    const server = new MicrosoftGraphServer(authManager, args);
    await server.initialize(version);
    await server.start();
  } catch (error) {
    logger.error(`Startup error: ${error}`);
    process.exit(1);
  }
}

main();
