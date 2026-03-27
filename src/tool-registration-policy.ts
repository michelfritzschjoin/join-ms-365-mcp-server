export type ToolRegistrationMode = 'classic' | 'super' | 'hybrid';

function normalizeMode(value: string | undefined): ToolRegistrationMode | null {
  if (!value) {
    return null;
  }

  const normalized = value.trim().toLowerCase();
  if (normalized === 'classic' || normalized === 'super' || normalized === 'hybrid') {
    return normalized;
  }

  return null;
}

/**
 * Resolves the effective MCP tool registration mode.
 * Priority:
 * 1) MS365_MCP_TOOL_MODE=classic|super|hybrid
 * 2) MS365_MCP_USE_SUPER_TOOLS=true|1 -> super
 * 3) fallback -> classic
 */
export function resolveToolRegistrationMode(
  environment: Record<string, string | undefined>
): ToolRegistrationMode {
  const configuredMode = normalizeMode(environment.MS365_MCP_TOOL_MODE);
  if (configuredMode) {
    return configuredMode;
  }

  const useSuperTools =
    environment.MS365_MCP_USE_SUPER_TOOLS === 'true' ||
    environment.MS365_MCP_USE_SUPER_TOOLS === '1';

  return useSuperTools ? 'super' : 'classic';
}
