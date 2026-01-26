#!/usr/bin/env node
/**
 * Update version in various files during release.
 * Called by semantic-release via @semantic-release/exec plugin.
 *
 * Usage: node scripts/update-version.mjs <version>
 */

import { readFileSync, writeFileSync, existsSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const rootDir = join(__dirname, '..');

const version = process.argv[2];

if (!version) {
  console.error('Error: Version argument required');
  console.error('Usage: node scripts/update-version.mjs <version>');
  process.exit(1);
}

console.log(`📦 Updating version to: ${version}`);

/**
 * Update a JSON file with the new version
 * @param {string} filePath - Path to the JSON file
 */
function updateJsonFile(filePath) {
  const fullPath = join(rootDir, filePath);

  if (!existsSync(fullPath)) {
    console.log(`  ⏭️  Skipping ${filePath} (file not found)`);
    return;
  }

  try {
    const content = JSON.parse(readFileSync(fullPath, 'utf8'));
    content.version = version;
    writeFileSync(fullPath, JSON.stringify(content, null, 2) + '\n');
    console.log(`  ✅ Updated ${filePath}`);
  } catch (error) {
    console.error(`  ❌ Failed to update ${filePath}: ${error.message}`);
  }
}

/**
 * Update glama.json with version in title
 * @param {string} filePath - Path to glama.json
 */
function updateGlamaJson(filePath) {
  const fullPath = join(rootDir, filePath);

  if (!existsSync(fullPath)) {
    console.log(`  ⏭️  Skipping ${filePath} (file not found)`);
    return;
  }

  try {
    const content = JSON.parse(readFileSync(fullPath, 'utf8'));
    // Update version in title if it exists
    if (content.title) {
      content.title = content.title.replace(/v[\d.]+(-\w+)?/, `v${version}`);
    }
    writeFileSync(fullPath, JSON.stringify(content, null, 2) + '\n');
    console.log(`  ✅ Updated ${filePath}`);
  } catch (error) {
    console.error(`  ❌ Failed to update ${filePath}: ${error.message}`);
  }
}

/**
 * Update version badge in README.md
 * @param {string} filePath - Path to README.md
 */
function updateReadme(filePath) {
  const fullPath = join(rootDir, filePath);

  if (!existsSync(fullPath)) {
    console.log(`  ⏭️  Skipping ${filePath} (file not found)`);
    return;
  }

  try {
    let content = readFileSync(fullPath, 'utf8');

    // Update version badge if exists
    content = content.replace(/version-[\d.]+(-\w+)?-/g, `version-${version}-`);

    // Update any explicit version references
    content = content.replace(
      /join-ms-365-mcp-server@[\d.]+(-\w+)?/g,
      `join-ms-365-mcp-server@${version}`
    );

    writeFileSync(fullPath, content);
    console.log(`  ✅ Updated ${filePath}`);
  } catch (error) {
    console.error(`  ❌ Failed to update ${filePath}: ${error.message}`);
  }
}

// Update files
updateJsonFile('package.json');
updateJsonFile('package-lock.json');
updateGlamaJson('glama.json');
updateReadme('README.md');

console.log(`\n✨ Version update complete: v${version}`);
