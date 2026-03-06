#!/usr/bin/env node

import { existsSync } from 'fs';
import { join } from 'path';
import { fileURLToPath } from 'url';
import { dirname } from 'path';
import { execSync } from 'child_process';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const rootDir = join(__dirname, '..');
const clientPath = join(rootDir, 'src', 'generated', 'client.ts');

if (!existsSync(clientPath)) {
  console.log('⚠️  Generated client files not found.');
  console.log('📦 Running generate script...\n');
  try {
    execSync('npm run generate', { stdio: 'inherit', cwd: rootDir });
    console.log('\n✅ Generation complete!\n');
  } catch (error) {
    console.error('\n❌ Generation failed. Please run "npm run generate" manually.\n');
    process.exit(1);
  }
} else {
  // Patch existing client so response type "binary" -> "json" (fixes TS2322)
  execSync('node bin/patch-generated-client.mjs', { stdio: 'inherit', cwd: rootDir });
}
