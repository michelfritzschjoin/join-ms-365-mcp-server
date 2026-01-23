/**
 * Configuration for Graph API Self-Repair System
 */

/**
 * Endpoint mappings for alternative paths
 */
export interface EndpointMapping {
  pattern: RegExp;
  alternatives: Array<{
    endpoint: string;
    condition?: (params: Record<string, unknown>) => boolean;
  }>;
}

/**
 * Scope hierarchy for automatic scope expansion
 */
export interface ScopeHierarchy {
  [key: string]: string[];
}

/**
 * Parameter validation rules
 */
export interface ParameterRule {
  name: string;
  pattern?: RegExp;
  validator?: (value: unknown) => boolean;
  fixer?: (value: unknown) => unknown;
}

/**
 * Endpoint mappings - maps patterns to alternative endpoints
 */
export const ENDPOINT_MAPPINGS: EndpointMapping[] = [
  // /me/* can be replaced with /users/{userId}/* if userId is available
  {
    pattern: /^\/me\/(.+)$/,
    alternatives: [
      {
        endpoint: '/users/{userId}/$1',
        condition: (params) => !!params.userId || !!params['user-id'],
      },
    ],
  },
  // /users/{id} can be replaced with /me if id matches current user
  {
    pattern: /^\/users\/([^/]+)(?:\/(.+))?$/,
    alternatives: [
      {
        endpoint: '/me/$2',
        condition: (params) => {
          const userId = params.userId || params['user-id'] || params.id;
          // If we have a way to check if this is the current user, use it
          // For now, we'll try /me as an alternative
          return true;
        },
      },
    ],
  },
  // /v1.0/* can fallback to /beta/*
  {
    pattern: /^\/v1\.0\/(.+)$/,
    alternatives: [
      {
        endpoint: '/beta/$1',
      },
    ],
  },
];

/**
 * Scope hierarchy - defines which scopes include others
 */
export const SCOPE_HIERARCHY: ScopeHierarchy = {
  'Mail.ReadWrite': ['Mail.Read', 'Mail.Send'],
  'Mail.Read': [],
  'Mail.Send': [],
  'Calendars.ReadWrite': ['Calendars.Read'],
  'Calendars.Read': [],
  'Files.ReadWrite': ['Files.Read'],
  'Files.Read': [],
  'Tasks.ReadWrite': ['Tasks.Read'],
  'Tasks.Read': [],
  'Contacts.ReadWrite': ['Contacts.Read'],
  'Contacts.Read': [],
  'User.ReadWrite': ['User.Read'],
  'User.Read': [],
  'Mail.ReadWrite.Shared': ['Mail.Read.Shared'],
  'Mail.Read.Shared': [],
};

/**
 * Get all scopes that are included in a given scope
 */
export function getIncludedScopes(scope: string): string[] {
  return SCOPE_HIERARCHY[scope] || [];
}

/**
 * Get all parent scopes that include the given scope
 */
export function getParentScopes(scope: string): string[] {
  const parents: string[] = [];
  for (const [parent, children] of Object.entries(SCOPE_HIERARCHY)) {
    if (children.includes(scope)) {
      parents.push(parent);
    }
  }
  return parents;
}

/**
 * Parameter validation rules
 */
export const PARAMETER_RULES: ParameterRule[] = [
  // UUID validation and fixing
  {
    name: 'id',
    pattern: /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i,
    validator: (value) => {
      if (typeof value !== 'string') return false;
      return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value);
    },
    fixer: (value) => {
      if (typeof value === 'string') {
        // Try to fix common UUID issues
        const cleaned = value.replace(/[^0-9a-f-]/gi, '');
        if (cleaned.length === 36) return cleaned;
      }
      return value;
    },
  },
  // Date format validation
  {
    name: 'date',
    validator: (value) => {
      if (typeof value !== 'string') return false;
      // ISO 8601 date format
      return /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2}(\.\d{3})?Z?)?$/.test(value);
    },
    fixer: (value) => {
      if (value instanceof Date) {
        return value.toISOString();
      }
      if (typeof value === 'string') {
        try {
          const date = new Date(value);
          if (!isNaN(date.getTime())) {
            return date.toISOString();
          }
        } catch {
          // Invalid date
        }
      }
      return value;
    },
  },
];

/**
 * OData parameter normalization
 */
export function normalizeODataParam(name: string): string {
  const odataParams = [
    'filter',
    'select',
    'expand',
    'orderby',
    'skip',
    'top',
    'count',
    'search',
    'format',
  ];

  const normalized = name.startsWith('$') ? name.slice(1) : name;
  if (odataParams.includes(normalized.toLowerCase())) {
    return `$${normalized.toLowerCase()}`;
  }
  return name;
}

/**
 * Get alternative endpoint for a given endpoint pattern
 */
export function getAlternativeEndpoint(
  endpoint: string,
  params: Record<string, unknown>
): string | null {
  for (const mapping of ENDPOINT_MAPPINGS) {
    const match = endpoint.match(mapping.pattern);
    if (match) {
      for (const alternative of mapping.alternatives) {
        if (!alternative.condition || alternative.condition(params)) {
          // Replace placeholders in alternative endpoint
          let altEndpoint = alternative.endpoint;
          for (let i = 1; i < match.length; i++) {
            altEndpoint = altEndpoint.replace(`$${i}`, match[i] || '');
          }
          // Replace parameter placeholders
          altEndpoint = altEndpoint.replace(/{(\w+)}/g, (_, key) => {
            const paramValue = params[key] || params[key.replace(/-/g, '_')];
            return paramValue ? String(paramValue) : `{${key}}`;
          });
          // Only return if all placeholders are replaced
          if (!altEndpoint.includes('{')) {
            return altEndpoint;
          }
        }
      }
    }
  }
  return null;
}

/**
 * Repair configuration
 */
export interface RepairConfig {
  enabled: boolean;
  maxRepairAttempts: number;
  enabledStrategies: string[];
  endpointRepair: boolean;
  parameterRepair: boolean;
  scopeRepair: boolean;
  versionRepair: boolean;
  rateLimitRepair: boolean;
}

/**
 * Get repair configuration from environment variables
 */
export function getRepairConfig(): RepairConfig {
  const enabled = process.env.MS365_MCP_ENABLE_SELF_REPAIR === 'true';
  const maxAttempts = parseInt(process.env.MS365_MCP_MAX_REPAIR_ATTEMPTS || '3', 10);
  const strategiesEnv =
    process.env.MS365_MCP_REPAIR_STRATEGIES || 'endpoint,parameter,scope,version,ratelimit';
  const enabledStrategies = strategiesEnv.split(',').map((s) => s.trim());

  return {
    enabled,
    maxRepairAttempts: maxAttempts,
    enabledStrategies,
    endpointRepair: enabledStrategies.includes('endpoint'),
    parameterRepair: enabledStrategies.includes('parameter'),
    scopeRepair: enabledStrategies.includes('scope'),
    versionRepair: enabledStrategies.includes('version'),
    rateLimitRepair: enabledStrategies.includes('ratelimit'),
  };
}
