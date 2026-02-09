/**
 * API Change Detection Utilities
 *
 * These utilities help detect changes in Graph API endpoints and configurations.
 * They can be used in CI/CD pipelines to detect breaking changes.
 */

import { readFileSync, existsSync } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

export interface EndpointConfig {
  pathPattern: string;
  method: string;
  toolName: string;
  scopes?: string[];
  workScopes?: string[];
  returnDownloadUrl?: boolean;
  supportsTimezone?: boolean;
  llmTip?: string;
}

export interface EndpointChange {
  type: 'added' | 'removed' | 'modified';
  endpoint: EndpointConfig;
  oldEndpoint?: EndpointConfig;
  changes?: string[];
}

export interface ApiChangeReport {
  timestamp: number;
  added: EndpointConfig[];
  removed: EndpointConfig[];
  modified: EndpointChange[];
  summary: {
    totalEndpoints: number;
    addedCount: number;
    removedCount: number;
    modifiedCount: number;
  };
}

/**
 * Load endpoints from endpoints.json
 */
export function loadEndpoints(): EndpointConfig[] {
  const endpointsPath = path.join(__dirname, '../../src/endpoints.json');
  if (!existsSync(endpointsPath)) {
    throw new Error(`Endpoints file not found: ${endpointsPath}`);
  }
  return JSON.parse(readFileSync(endpointsPath, 'utf8')) as EndpointConfig[];
}

/**
 * Compare two endpoint configurations and detect changes
 */
export function compareEndpoints(
  oldEndpoints: EndpointConfig[],
  newEndpoints: EndpointConfig[]
): ApiChangeReport {
  const oldMap = new Map<string, EndpointConfig>();
  const newMap = new Map<string, EndpointConfig>();

  // Create maps for easy lookup
  for (const endpoint of oldEndpoints) {
    const key = `${endpoint.method}:${endpoint.pathPattern}`;
    oldMap.set(key, endpoint);
  }

  for (const endpoint of newEndpoints) {
    const key = `${endpoint.method}:${endpoint.pathPattern}`;
    newMap.set(key, endpoint);
  }

  const added: EndpointConfig[] = [];
  const removed: EndpointConfig[] = [];
  const modified: EndpointChange[] = [];

  // Find added endpoints
  for (const [key, endpoint] of newMap.entries()) {
    if (!oldMap.has(key)) {
      added.push(endpoint);
    }
  }

  // Find removed endpoints
  for (const [key, endpoint] of oldMap.entries()) {
    if (!newMap.has(key)) {
      removed.push(endpoint);
    }
  }

  // Find modified endpoints
  for (const [key, newEndpoint] of newMap.entries()) {
    const oldEndpoint = oldMap.get(key);
    if (oldEndpoint) {
      const changes = detectEndpointChanges(oldEndpoint, newEndpoint);
      if (changes.length > 0) {
        modified.push({
          type: 'modified',
          endpoint: newEndpoint,
          oldEndpoint,
          changes,
        });
      }
    }
  }

  return {
    timestamp: Date.now(),
    added,
    removed,
    modified,
    summary: {
      totalEndpoints: newEndpoints.length,
      addedCount: added.length,
      removedCount: removed.length,
      modifiedCount: modified.length,
    },
  };
}

/**
 * Detect specific changes between two endpoint configurations
 */
function detectEndpointChanges(oldEndpoint: EndpointConfig, newEndpoint: EndpointConfig): string[] {
  const changes: string[] = [];

  if (oldEndpoint.toolName !== newEndpoint.toolName) {
    changes.push(`toolName changed from "${oldEndpoint.toolName}" to "${newEndpoint.toolName}"`);
  }

  if (oldEndpoint.method !== newEndpoint.method) {
    changes.push(`method changed from "${oldEndpoint.method}" to "${newEndpoint.method}"`);
  }

  if (oldEndpoint.pathPattern !== newEndpoint.pathPattern) {
    changes.push(
      `pathPattern changed from "${oldEndpoint.pathPattern}" to "${newEndpoint.pathPattern}"`
    );
  }

  const oldScopes = JSON.stringify(oldEndpoint.scopes || []);
  const newScopes = JSON.stringify(newEndpoint.scopes || []);
  if (oldScopes !== newScopes) {
    changes.push('scopes changed');
  }

  const oldWorkScopes = JSON.stringify(oldEndpoint.workScopes || []);
  const newWorkScopes = JSON.stringify(newEndpoint.workScopes || []);
  if (oldWorkScopes !== newWorkScopes) {
    changes.push('workScopes changed');
  }

  if (oldEndpoint.returnDownloadUrl !== newEndpoint.returnDownloadUrl) {
    changes.push(
      `returnDownloadUrl changed from ${oldEndpoint.returnDownloadUrl} to ${newEndpoint.returnDownloadUrl}`
    );
  }

  if (oldEndpoint.supportsTimezone !== newEndpoint.supportsTimezone) {
    changes.push(
      `supportsTimezone changed from ${oldEndpoint.supportsTimezone} to ${newEndpoint.supportsTimezone}`
    );
  }

  return changes;
}

/**
 * Validate endpoint configuration
 */
export function validateEndpoint(endpoint: EndpointConfig): {
  valid: boolean;
  errors: string[];
} {
  const errors: string[] = [];

  if (!endpoint.pathPattern) {
    errors.push('pathPattern is required');
  } else {
    if (!endpoint.pathPattern.startsWith('/')) {
      errors.push('pathPattern must start with /');
    }
    if (endpoint.pathPattern.includes('//')) {
      errors.push('pathPattern must not contain //');
    }
    if (endpoint.pathPattern.includes('/v1.0/') || endpoint.pathPattern.includes('/beta/')) {
      errors.push('pathPattern must not include API version prefix');
    }
  }

  if (!endpoint.method) {
    errors.push('method is required');
  } else {
    const validMethods = ['get', 'post', 'patch', 'put', 'delete'];
    if (!validMethods.includes(endpoint.method.toLowerCase())) {
      errors.push(`method must be one of: ${validMethods.join(', ')}`);
    }
  }

  if (!endpoint.toolName) {
    errors.push('toolName is required');
  } else {
    if (!/^[a-z0-9-]+$/.test(endpoint.toolName)) {
      errors.push('toolName must be lowercase alphanumeric with hyphens');
    }
  }

  const allScopes = [...(endpoint.scopes || []), ...(endpoint.workScopes || [])];
  if (allScopes.length === 0) {
    errors.push('endpoint must have at least one scope (scopes or workScopes)');
  }

  return {
    valid: errors.length === 0,
    errors,
  };
}

/**
 * Validate all endpoints
 */
export function validateAllEndpoints(endpoints: EndpointConfig[]): {
  valid: boolean;
  errors: Array<{ endpoint: EndpointConfig; errors: string[] }>;
} {
  const allErrors: Array<{ endpoint: EndpointConfig; errors: string[] }> = [];

  for (const endpoint of endpoints) {
    const validation = validateEndpoint(endpoint);
    if (!validation.valid) {
      allErrors.push({
        endpoint,
        errors: validation.errors,
      });
    }
  }

  return {
    valid: allErrors.length === 0,
    errors: allErrors,
  };
}

/**
 * Generate a change report as markdown
 */
export function generateChangeReportMarkdown(report: ApiChangeReport): string {
  const lines: string[] = [];

  lines.push('# Graph API Change Report');
  lines.push(`Generated: ${new Date(report.timestamp).toISOString()}`);
  lines.push('');

  lines.push('## Summary');
  lines.push(`- Total Endpoints: ${report.summary.totalEndpoints}`);
  lines.push(`- Added: ${report.summary.addedCount}`);
  lines.push(`- Removed: ${report.summary.removedCount}`);
  lines.push(`- Modified: ${report.summary.modifiedCount}`);
  lines.push('');

  if (report.added.length > 0) {
    lines.push('## Added Endpoints');
    for (const endpoint of report.added) {
      lines.push(
        `- \`${endpoint.method.toUpperCase()} ${endpoint.pathPattern}\` (${endpoint.toolName})`
      );
    }
    lines.push('');
  }

  if (report.removed.length > 0) {
    lines.push('## Removed Endpoints');
    for (const endpoint of report.removed) {
      lines.push(
        `- \`${endpoint.method.toUpperCase()} ${endpoint.pathPattern}\` (${endpoint.toolName})`
      );
    }
    lines.push('');
  }

  if (report.modified.length > 0) {
    lines.push('## Modified Endpoints');
    for (const change of report.modified) {
      lines.push(`### ${change.endpoint.toolName}`);
      lines.push(`- Path: \`${change.endpoint.pathPattern}\``);
      lines.push(`- Method: \`${change.endpoint.method.toUpperCase()}\``);
      lines.push('- Changes:');
      for (const changeDesc of change.changes || []) {
        lines.push(`  - ${changeDesc}`);
      }
      lines.push('');
    }
  }

  return lines.join('\n');
}
