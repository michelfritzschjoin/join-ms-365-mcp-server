import { describe, expect, it } from 'vitest';
import {
  escapeODataStringLiteral,
  buildContainsFilter,
  buildEqFilter,
  buildStartsWithFilter,
  joinOrFilters,
} from '../src/utils/odata-filter.js';

describe('OData query sanitization', () => {
  it('escapes single quotes in literals', () => {
    expect(escapeODataStringLiteral("O'Reilly")).toBe("O''Reilly");
  });

  it('builds safe contains/eq/startswith filters', () => {
    expect(buildContainsFilter('title', "Q1's plan")).toBe("contains(title,'Q1''s plan')");
    expect(buildEqFilter('mail', "john.o'connor@contoso.com")).toBe(
      "mail eq 'john.o''connor@contoso.com'"
    );
    expect(buildStartsWithFilter('displayName', "D'Angelo")).toBe(
      "startswith(displayName,'D''Angelo')"
    );
  });

  it('joins filters with OR while ignoring blanks', () => {
    const result = joinOrFilters([
      buildEqFilter('mail', 'alice@contoso.com'),
      '  ',
      buildContainsFilter('displayName', 'Alice'),
    ]);
    expect(result).toBe("mail eq 'alice@contoso.com' or contains(displayName,'Alice')");
  });
});
