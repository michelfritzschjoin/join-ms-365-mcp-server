import { describe, expect, it } from 'vitest';
import {
  buildGroupSearchFilter,
  buildPersonInfoMailFilter,
  buildTaskTitleSearchFilter,
} from '../src/super-tools/search-domain.js';

describe('super-tools domain helpers', () => {
  it('builds safe task title filter', () => {
    expect(buildTaskTitleSearchFilter("Roadmap '26")).toBe("contains(title,'Roadmap ''26')");
  });

  it('builds safe group search filter', () => {
    const filter = buildGroupSearchFilter("O'Reilly");
    expect(filter).toContain("contains(displayName,'O''Reilly')");
    expect(filter).toContain("contains(mail,'O''Reilly')");
    expect(filter).toContain("contains(description,'O''Reilly')");
  });

  it('builds safe person-info mail filter', () => {
    const filter = buildPersonInfoMailFilter("anne.o'connor@contoso.com");
    expect(filter).toContain("from/emailAddress/address eq 'anne.o''connor@contoso.com'");
    expect(filter).toContain("contains(from/emailAddress/name,'anne.o''connor@contoso.com')");
  });
});
