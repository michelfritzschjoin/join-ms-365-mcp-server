#!/usr/bin/env node
/**
 * Patches generated client.ts so that response type "binary" is "json"
 * (fixes TS2322 when openapi-zod-client emits "binary" where only "json" is allowed).
 */

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, '..');
const clientPath = path.join(rootDir, 'src', 'generated', 'client.ts');

if (!fs.existsSync(clientPath)) {
  console.log('patch-generated-client: src/generated/client.ts not found, skipping.');
  process.exit(0);
}

let content = fs.readFileSync(clientPath, 'utf-8');
const original = content;

// Fix: Type '"binary"' is not assignable to type '"json"' (openapi-zod-client endpoint responseType)
content = content.replace(/: "binary"/g, ': "json"');
content = content.replace(/: 'binary'/g, ": 'json'");

if (content !== original) {
  fs.writeFileSync(clientPath, content);
  console.log('patch-generated-client: patched response type "binary" -> "json" in client.ts');
} else {
  console.log('patch-generated-client: no "binary" response type found, nothing to patch.');
}
