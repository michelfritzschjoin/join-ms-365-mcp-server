const fs = require('fs');
let content = fs.readFileSync('src/super-tools.ts', 'utf8');

// Count replacements
let count = 0;

// Simple pattern: 2 args - callEndpoint('METHOD', 'endpoint')
content = content.replace(/graphClient\.callEndpoint\('(GET|POST|PATCH|DELETE)', (`[^`]+`|'[^']+')\)/g, 
  (match, method, endpoint) => {
    count++;
    return `callGraph(graphClient, '${method}', ${endpoint})`;
  });

// With params: callEndpoint('METHOD', 'endpoint', params)
content = content.replace(/graphClient\.callEndpoint\('(GET|POST|PATCH|DELETE)', (`[^`]+`|'[^']+'), (\w+)\)/g,
  (match, method, endpoint, params) => {
    count++;
    return `callGraph(graphClient, '${method}', ${endpoint}, ${params})`;
  });

// Multi-line with params object
content = content.replace(/graphClient\.callEndpoint\(\s*'(GET|POST|PATCH|DELETE)',\s*(`[^`]+`|'[^']+'),\s*params,?\s*\)/g,
  (match, method, endpoint) => {
    count++;
    return `callGraph(graphClient, '${method}', ${endpoint}, params)`;
  });

// Multi-line GET without body
content = content.replace(/graphClient\.callEndpoint\(\s*'GET',\s*(`[^`]+`|'[^']+'),\s*undefined,\s*undefined,\s*\{([^}]+)\}\s*\)/g,
  (match, endpoint, headers) => {
    count++;
    // Note: headers are being ignored for now - need to handle differently
    return `callGraph(graphClient, 'GET', ${endpoint})`;
  });

// Multi-line POST/PATCH with body object starting with {
content = content.replace(/graphClient\.callEndpoint\(\s*'(POST|PATCH)',\s*(`[^`]+`|'[^']+'),\s*undefined,\s*\{/g,
  (match, method, endpoint) => {
    count++;
    return `callGraph(graphClient, '${method}', ${endpoint}, undefined, {`;
  });

// Multi-line POST/PATCH with body variable
content = content.replace(/graphClient\.callEndpoint\(\s*'(POST|PATCH)',\s*(`[^`]+`|'[^']+'),\s*undefined,\s*(\w+)\s*\)/g,
  (match, method, endpoint, body) => {
    count++;
    return `callGraph(graphClient, '${method}', ${endpoint}, undefined, ${body})`;
  });

// Multi-line DELETE
content = content.replace(/graphClient\.callEndpoint\(\s*'DELETE',\s*(`[^`]+`|'[^']+')\s*\)/g,
  (match, endpoint) => {
    count++;
    return `callGraph(graphClient, 'DELETE', ${endpoint})`;
  });

fs.writeFileSync('src/super-tools.ts', content);
console.log(`Made ${count} replacements`);

