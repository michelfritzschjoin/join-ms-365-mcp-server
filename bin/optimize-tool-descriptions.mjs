#!/usr/bin/env node

/**
 * Optimize Tool Descriptions Script
 *
 * Optimizes tool descriptions for better LLM understanding and user experience:
 * - Ensures descriptions are clear and actionable
 * - Optimizes length (not too short, not too long)
 * - Improves clarity and removes redundancy
 * - Adds context where helpful
 */

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, '..');
const clientFile = path.join(rootDir, 'src', 'generated', 'client.ts');

// Minimum and maximum recommended description lengths
const MIN_LENGTH = 20;
const MAX_LENGTH = 500;

// Common patterns to improve
const OPTIMIZATION_PATTERNS = [
  // Remove redundant phrases
  { pattern: /\bRetrieve a single\b/gi, replacement: 'Get' },
  { pattern: /\bRetrieve\b/gi, replacement: 'Get' },
  { pattern: /\bObtain\b/gi, replacement: 'Get' },
  { pattern: /\bFetch\b/gi, replacement: 'Get' },

  // Improve clarity
  {
    pattern: /\bThe messages in a mailbox or folder\b/gi,
    replacement: 'List messages in a mailbox or folder',
  },
  { pattern: /\bThe\s+(\w+)\s+in\b/gi, replacement: 'List $1 in' },

  // Remove unnecessary words
  { pattern: /\bThis method\b/gi, replacement: 'This' },
  { pattern: /\bThis operation\b/gi, replacement: 'This' },

  // Fix common issues
  { pattern: /\s+/g, replacement: ' ' }, // Multiple spaces
  { pattern: /\.{2,}/g, replacement: '.' }, // Multiple periods
];

function optimizeDescription(description, alias = '', method = 'get', path = '') {
  if (!description || typeof description !== 'string') {
    return null;
  }

  let optimized = description.trim();

  // Skip if empty
  if (optimized.length === 0) {
    return null;
  }

  // Apply optimization patterns
  for (const { pattern, replacement } of OPTIMIZATION_PATTERNS) {
    optimized = optimized.replace(pattern, replacement);
  }

  // Ensure proper capitalization
  if (optimized.length > 0) {
    optimized = optimized.charAt(0).toUpperCase() + optimized.slice(1);
  }

  // Trim again after replacements
  optimized = optimized.trim();

  // Check length and provide feedback
  if (optimized.length < MIN_LENGTH && alias) {
    // Try to enhance short descriptions with context from alias/path
    const aliasWords = alias.split('-').filter((w) => w.length > 2);
    if (aliasWords.length > 0 && !optimized.toLowerCase().includes(aliasWords[0].toLowerCase())) {
      const action =
        method.toUpperCase() === 'GET'
          ? 'Get'
          : method.toUpperCase() === 'POST'
            ? 'Create'
            : method.toUpperCase() === 'PATCH'
              ? 'Update'
              : method.toUpperCase() === 'DELETE'
                ? 'Delete'
                : 'Execute';
      optimized = `${action} ${aliasWords.join(' ')}. ${optimized}`;
    }
  }

  // Truncate if too long (but preserve sentence structure)
  if (optimized.length > MAX_LENGTH) {
    const truncated = optimized.substring(0, MAX_LENGTH);
    const lastPeriod = truncated.lastIndexOf('.');
    if (lastPeriod > MAX_LENGTH * 0.7) {
      optimized = truncated.substring(0, lastPeriod + 1);
    } else {
      optimized = truncated + '...';
    }
  }

  // Ensure proper ending punctuation
  if (optimized.length > 10 && !/[.!?]$/.test(optimized)) {
    optimized += '.';
  }

  return optimized;
}

function optimizeToolDescriptions() {
  console.log('✨ Optimizing tool descriptions...\n');

  if (!fs.existsSync(clientFile)) {
    console.error(`❌ Client file not found: ${clientFile}`);
    console.error('   Run "npm run generate" first to generate the client file.');
    process.exit(1);
  }

  let clientCode = fs.readFileSync(clientFile, 'utf8');
  let optimizedCount = 0;
  let stats = {
    tooShort: 0,
    tooLong: 0,
    optimized: 0,
  };

  // Find all endpoint descriptions (template literals)
  // Extract context from the code before each description match
  const descriptionPattern = /description:\s*`((?:(?:[^`\\]|\\.|`[^`])*))`/g;

  // We need to extract context (method, alias, path) for each description
  // Strategy: find each description, then look backwards for method/alias/path
  let match;
  const descriptions = [];

  while ((match = descriptionPattern.exec(clientCode)) !== null) {
    const descStart = match.index;
    const descEnd = match.index + match[0].length;

    // Extract context from code before this description
    const codeBefore = clientCode.substring(Math.max(0, descStart - 2000), descStart);

    // Find the most recent method, alias, and path before this description
    const methodMatch = codeBefore.match(/method:\s*['"](get|post|patch|put|delete)['"][^]*$/);
    const aliasMatch = codeBefore.match(/alias:\s*['"]([^'"]+)['"][^]*$/);
    const pathMatch = codeBefore.match(/path:\s*['"]([^'"]+)['"][^]*$/);

    descriptions.push({
      fullMatch: match[0],
      description: match[1],
      method: methodMatch ? methodMatch[1] : 'get',
      alias: aliasMatch ? aliasMatch[1] : '',
      path: pathMatch ? pathMatch[1] : '',
      index: match.index,
    });
  }

  // Process descriptions in reverse order to maintain indices
  for (let i = descriptions.length - 1; i >= 0; i--) {
    const info = descriptions[i];
    const original = info.description;
    const optimized = optimizeDescription(original, info.alias, info.method, info.path);

    if (optimized === null || optimized === original) {
      continue;
    }

    // Track statistics
    if (original.length < MIN_LENGTH && optimized.length >= MIN_LENGTH) {
      stats.tooShort++;
    }
    if (original.length > MAX_LENGTH && optimized.length <= MAX_LENGTH) {
      stats.tooLong++;
    }

    optimizedCount++;
    stats.optimized++;

    // Replace in the code
    const escaped = optimized.replace(/`/g, '\\`');
    const replacement = `description: \`${escaped}\``;
    clientCode =
      clientCode.substring(0, info.index) +
      replacement +
      clientCode.substring(info.index + info.fullMatch.length);
  }

  if (optimizedCount > 0) {
    fs.writeFileSync(clientFile, clientCode, 'utf8');
    console.log(`✅ Optimized ${optimizedCount} tool description(s)`);
    if (stats.tooShort > 0) {
      console.log(`   - Enhanced ${stats.tooShort} description(s) that were too short`);
    }
    if (stats.tooLong > 0) {
      console.log(`   - Truncated ${stats.tooLong} description(s) that were too long`);
    }
    console.log('');
  } else {
    console.log('✅ All tool descriptions are already optimized\n');
  }

  return optimizedCount > 0;
}

try {
  const hasChanges = optimizeToolDescriptions();
  process.exit(hasChanges ? 0 : 0); // Exit with 0 regardless (non-breaking)
} catch (error) {
  console.error('❌ Error optimizing tool descriptions:', error.message);
  if (error.stack) {
    console.error(error.stack);
  }
  process.exit(1);
}
