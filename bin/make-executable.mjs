#!/usr/bin/env node

/**
 * Cross-platform script to make a file executable
 * On Unix/Linux: uses chmod
 * On Windows: no-op (files are executable by default)
 */

import { chmod } from 'fs/promises';
import { platform } from 'os';
import { fileURLToPath } from 'url';
import { dirname, resolve } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const projectRoot = resolve(__dirname, '..');
const targetFile = resolve(projectRoot, 'dist/index.js');

if (platform() !== 'win32') {
  try {
    await chmod(targetFile, 0o755);
    console.log(`Made ${targetFile} executable`);
  } catch (error) {
    // File might not exist yet, that's okay
    if (error?.code !== 'ENOENT') {
      console.warn(`Warning: Could not make file executable: ${error}`);
    }
  }
} else {
  // Windows doesn't need chmod - files are executable by default
  // Just verify the file exists
  try {
    const { access, constants } = await import('fs/promises');
    await access(targetFile, constants.F_OK);
  } catch {
    // File doesn't exist yet, that's okay
  }
}

