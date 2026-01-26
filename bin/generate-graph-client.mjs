#!/usr/bin/env node

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { downloadGraphOpenAPI } from './modules/download-openapi.mjs';
import { generateMcpTools } from './modules/generate-mcp-tools.mjs';
import { createAndSaveSimplifiedOpenAPI } from './modules/simplified-openapi.mjs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, '..');
const openapiDir = path.join(rootDir, 'openapi');
const srcDir = path.join(rootDir, 'src');

const openapiFile = path.join(openapiDir, 'openapi.yaml');
const openapiTrimmedFile = path.join(openapiDir, 'openapi-trimmed.yaml');
const endpointsFile = path.join(srcDir, 'endpoints.json');

const generatedDir = path.join(srcDir, 'generated');

const args = process.argv.slice(2);
const forceDownload = args.includes('--force');

async function main() {
  console.log('Microsoft Graph API OpenAPI Processor');
  console.log('------------------------------------');

  try {
    console.log('\n📥 Step 1: Downloading OpenAPI specification');
    const downloaded = await downloadGraphOpenAPI(
      openapiDir,
      openapiFile,
      undefined,
      forceDownload
    );

    if (downloaded) {
      console.log('\n✅ OpenAPI specification successfully downloaded');
    } else {
      console.log('\n⏭️ Download skipped (file exists)');
    }

    console.log('\n🔧 Step 2: Creating simplified OpenAPI specification');
    createAndSaveSimplifiedOpenAPI(endpointsFile, openapiFile, openapiTrimmedFile);
    console.log('✅ Successfully created simplified OpenAPI specification');

    console.log('\n🚀 Step 3: Generating client code using openapi-zod-client');
    try {
      generateMcpTools(null, generatedDir);
      console.log('✅ Successfully generated client code');
    } catch (genError) {
      // Check if we're in Docker/emulated environment
      const isDocker = fs.existsSync('/.dockerenv');
      const isArm64 = process.arch === 'arm64';

      if (isDocker && isArm64) {
        console.error('\n❌ Generation failed in Docker ARM64 emulated environment');
        console.error('   This is a known issue with qemu emulation and native dependencies');
        console.error('   Solutions:');
        console.error('   1. Commit the generated src/generated/client.ts file to the repository');
        console.error('   2. Use a native ARM64 build environment (not emulated)');
        console.error('   3. Build on x86_64/amd64 platform instead');
        console.error(`\n   Error: ${genError.message}`);
        process.exit(1);
      } else {
        throw genError;
      }
    }
  } catch (error) {
    console.error('\n❌ Error processing OpenAPI specification:', error.message);
    if (error.stack) {
      console.error(error.stack);
    }
    process.exit(1);
  }
}

main().catch((error) => {
  console.error('Fatal error:', error);
  process.exit(1);
});
