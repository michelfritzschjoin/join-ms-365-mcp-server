const fs = require('fs');
let content = fs.readFileSync('src/super-tools.ts', 'utf8');

let count = 0;

// Pattern: Multi-line callEndpoint with params on next line then body/undefined
// graphClient.callEndpoint(
//   'GET',
//   '/endpoint',
//   params,
//   ...
// )
content = content.replace(
  /graphClient\.callEndpoint\(\s*'(GET|POST|PATCH|DELETE)',\s*([`'][^`']+[`']),\s*(params|\{[^}]+\}|undefined)/gs,
  (match, method, endpoint, third) => {
    count++;
    if (third === 'params') {
      return `callGraph(graphClient, '${method}', ${endpoint}, params`;
    } else if (third === 'undefined') {
      return `callGraph(graphClient, '${method}', ${endpoint}, undefined`;
    } else {
      return `callGraph(graphClient, '${method}', ${endpoint}, ${third}`;
    }
  }
);

// Pattern: calls with 5th arg for headers
// graphClient.callEndpoint('GET', '/users', params, undefined, { ConsistencyLevel: 'eventual' })
content = content.replace(
  /graphClient\.callEndpoint\('(GET)', '([^']+)', (\w+), undefined, \{([^}]+)\}\)/g,
  (match, method, endpoint, params) => {
    count++;
    // Headers are ignored - need to handle separately if needed
    return `callGraph(graphClient, '${method}', '${endpoint}', ${params})`;
  }
);

// Pattern: simple call with template literal endpoint 
content = content.replace(
  /graphClient\.callEndpoint\('(GET|POST|PATCH|DELETE)', (`[^`]+`)\)/g,
  (match, method, endpoint) => {
    count++;
    return `callGraph(graphClient, '${method}', ${endpoint})`;
  }
);

// Pattern: call with single quote endpoint and params
content = content.replace(
  /graphClient\.callEndpoint\('(GET)', '([^']+)', \{([^}]+)\}\)/g,
  (match, method, endpoint, params) => {
    count++;
    return `callGraph(graphClient, '${method}', '${endpoint}', {${params}})`;
  }
);

fs.writeFileSync('src/super-tools.ts', content);
console.log(`Made ${count} additional replacements`);
