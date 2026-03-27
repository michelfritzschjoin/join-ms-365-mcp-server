import { buildContainsFilter, buildEqFilter, joinOrFilters } from '../utils/odata-filter.js';

export function buildTaskTitleSearchFilter(search: string): string {
  return buildContainsFilter('title', search);
}

export function buildGroupSearchFilter(query: string): string {
  return joinOrFilters([
    buildContainsFilter('displayName', query),
    buildContainsFilter('mail', query),
    buildContainsFilter('description', query),
  ]);
}

export function buildPersonInfoMailFilter(person: string): string {
  return joinOrFilters([
    buildEqFilter('from/emailAddress/address', person),
    buildContainsFilter('from/emailAddress/name', person),
  ]);
}
