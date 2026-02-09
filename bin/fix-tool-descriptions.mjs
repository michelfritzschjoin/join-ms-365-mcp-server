#!/usr/bin/env node

/**
 * Fix Tool Descriptions Script
 *
 * Validates and fixes tool descriptions in the generated client:
 * - Ensures all endpoints have non-empty descriptions
 * - Removes trailing/leading whitespace
 * - Ensures proper capitalization
 * - Fixes common formatting issues
 */

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, '..');
const clientFile = path.join(rootDir, 'src', 'generated', 'client.ts');

function fixDescription(description) {
  if (!description || typeof description !== 'string') {
    return null;
  }

  // Remove leading/trailing whitespace
  let fixed = description.trim();

  // Skip if empty after trimming
  if (fixed.length === 0) {
    return null;
  }

  // Ensure first letter is capitalized
  if (fixed.length > 0) {
    fixed = fixed.charAt(0).toUpperCase() + fixed.slice(1);
  }

  // Remove multiple consecutive spaces
  fixed = fixed.replace(/\s+/g, ' ');

  // Remove trailing periods if there are multiple
  fixed = fixed.replace(/\.+$/, '.');

  // Ensure description ends with proper punctuation (if it's a sentence)
  if (fixed.length > 10 && !/[.!?]$/.test(fixed)) {
    fixed += '.';
  }

  return fixed;
}

function fixToolDescriptions() {
  console.log('🔧 Fixing tool descriptions...\n');

  if (!fs.existsSync(clientFile)) {
    console.error(`❌ Client file not found: ${clientFile}`);
    console.error('   Run "npm run generate" first to generate the client file.');
    process.exit(1);
  }

  let clientCode = fs.readFileSync(clientFile, 'utf8');
  let fixedCount = 0;
  let issues = [];

  // Find all endpoint descriptions in the makeApi array
  // Pattern: description: `...` (template literals are used in the generated file)
  // We need to match across multiple lines for template literals
  const descriptionPattern = /description:\s*`((?:(?:[^`\\]|\\.|`[^`])*))`/g;

  clientCode = clientCode.replace(descriptionPattern, (match, description) => {
    const original = description;
    const fixed = fixDescription(description);

    if (fixed === null) {
      issues.push({
        type: 'empty',
        original: original.substring(0, 50),
      });
      return match; // Keep original if fixing failed
    }

    if (fixed !== original) {
      fixedCount++;
      // Escape backticks in the fixed description if needed
      const escaped = fixed.replace(/`/g, '\\`');
      return `description: \`${escaped}\``;
    }

    return match;
  });

  if (fixedCount > 0) {
    fs.writeFileSync(clientFile, clientCode, 'utf8');
    console.log(`✅ Fixed ${fixedCount} tool description(s)\n`);
  } else {
    console.log('✅ All tool descriptions are already properly formatted\n');
  }

  if (issues.length > 0) {
    console.warn(`⚠️  Found ${issues.length} description(s) with issues:`);
    issues.forEach((issue, index) => {
      console.warn(`   ${index + 1}. ${issue.type}: "${issue.original}..."`);
    });
    console.warn(
      '\n   Consider running "npm run generate" to update descriptions from OpenAPI spec.\n'
    );
  }

  return fixedCount > 0 || issues.length > 0;
}

try {
  const hasChanges = fixToolDescriptions();
  process.exit(hasChanges ? 0 : 0); // Exit with 0 regardless (non-breaking)
} catch (error) {
  console.error('❌ Error fixing tool descriptions:', error.message);
  if (error.stack) {
    console.error(error.stack);
  }
  process.exit(1);
}
