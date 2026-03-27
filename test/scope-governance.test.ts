import { describe, expect, it } from 'vitest';
import { readFileSync } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { INTENT_REQUIRED_SCOPES } from '../src/intent-scope-policy.js';

interface EndpointConfig {
  scopes?: string[];
  workScopes?: string[];
}

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

function loadEndpointScopes(): Set<string> {
  const endpointsPath = path.join(__dirname, '../src/endpoints.json');
  const endpoints = JSON.parse(readFileSync(endpointsPath, 'utf-8')) as EndpointConfig[];
  const endpointScopes = new Set<string>();

  for (const endpoint of endpoints) {
    for (const scope of endpoint.scopes ?? []) {
      endpointScopes.add(scope);
    }
    for (const workScope of endpoint.workScopes ?? []) {
      endpointScopes.add(workScope);
    }
  }

  return endpointScopes;
}

describe('scope governance', () => {
  it('uses only delegated-style scopes in intent scope policy', () => {
    const forbiddenApplicationStyle = ['.Default', 'Application', 'AppRole', 'Role'];

    for (const [intent, requiredScopes] of Object.entries(INTENT_REQUIRED_SCOPES)) {
      const allScopes = [...requiredScopes.scopes, ...requiredScopes.workScopes];
      for (const scope of allScopes) {
        for (const forbidden of forbiddenApplicationStyle) {
          expect(
            scope.includes(forbidden),
            `Intent "${intent}" contains forbidden scope marker "${forbidden}" in "${scope}"`
          ).toBe(false);
        }
      }
    }
  });

  it('keeps intent scope policy aligned with endpoint scopes', () => {
    const endpointScopes = loadEndpointScopes();
    const allowedExtraScopes = new Set<string>(['Sites.Selected', 'Bookings.Read.All']);

    for (const [intent, requiredScopes] of Object.entries(INTENT_REQUIRED_SCOPES)) {
      const allScopes = [...requiredScopes.scopes, ...requiredScopes.workScopes];
      for (const scope of allScopes) {
        const isDeclaredByEndpoint = endpointScopes.has(scope);
        const isAllowedExtra = allowedExtraScopes.has(scope);
        expect(
          isDeclaredByEndpoint || isAllowedExtra,
          `Intent "${intent}" references undeclared scope "${scope}" (not found in endpoints.json)`
        ).toBe(true);
      }
    }
  });
});
