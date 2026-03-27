/**
 * OData filter helpers for safe string literal handling.
 * Microsoft Graph OData strings must escape single quotes by doubling them.
 */

/**
 * Escapes a value for safe OData string literal usage.
 */
export function escapeODataStringLiteral(value: string): string {
  return value.replace(/'/g, "''");
}

/**
 * Builds a safe OData eq expression for a string value.
 */
export function buildEqFilter(field: string, value: string): string {
  return `${field} eq '${escapeODataStringLiteral(value)}'`;
}

/**
 * Builds a safe OData contains expression for a string value.
 */
export function buildContainsFilter(field: string, value: string): string {
  return `contains(${field},'${escapeODataStringLiteral(value)}')`;
}

/**
 * Builds a safe OData startswith expression for a string value.
 */
export function buildStartsWithFilter(field: string, value: string): string {
  return `startswith(${field},'${escapeODataStringLiteral(value)}')`;
}

/**
 * Joins multiple filter clauses with OR.
 */
export function joinOrFilters(filters: string[]): string {
  return filters.filter((filter) => filter.trim().length > 0).join(' or ');
}
