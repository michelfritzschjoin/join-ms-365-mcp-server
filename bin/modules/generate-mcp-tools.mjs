import fs from 'fs';
import path from 'path';
import { execSync } from 'child_process';

export function generateMcpTools(openApiSpec, outputDir) {
  try {
    const rootDir = path.resolve(outputDir, '../..');
    const openapiDir = path.join(rootDir, 'openapi');
    const openapiTrimmedFile = path.join(openapiDir, 'openapi-trimmed.yaml');
    const clientFilePath = path.join(outputDir, 'client.ts');

    // Check if client file already exists - skip generation if it does
    if (fs.existsSync(clientFilePath)) {
      console.log('Generated client file already exists, skipping generation...');
      return true;
    }

    if (!fs.existsSync(openapiTrimmedFile)) {
      throw new Error(`OpenAPI trimmed file not found: ${openapiTrimmedFile}`);
    }

    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
      console.log(`Created directory: ${outputDir}`);
    }

    console.log('Generating client code from OpenAPI spec using openapi-zod-client...');

    // Check if we're in Docker/CI environment
    const isDocker = fs.existsSync('/.dockerenv') || process.env.CI === 'true';
    const isArm64 = process.arch === 'arm64';

    if (isDocker && isArm64) {
      console.warn('⚠️  Warning: Running on ARM64 in Docker - generation may fail in emulated environments');
      console.warn('   If generation fails, ensure generated files are committed or use native ARM64 build');
    }

    try {
      execSync(
        `npx -y openapi-zod-client "${openapiTrimmedFile}" -o "${clientFilePath}" --with-description --strict-objects --additional-props-default-value=false`,
        {
          stdio: 'inherit',
          timeout: 300000, // 5 minute timeout
        }
      );
    } catch (execError) {
      // If generation fails and we're in Docker, check if file was partially created
      if (isDocker && fs.existsSync(clientFilePath)) {
        console.warn('⚠️  Generation command failed but file exists - using existing file');
        console.warn('   This may indicate a partial generation or emulation issue');
      } else {
        throw execError;
      }
    }

    if (!fs.existsSync(clientFilePath)) {
      throw new Error('Client file was not generated');
    }

    console.log(`Generated client code at: ${clientFilePath}`);

    let clientCode = fs.readFileSync(clientFilePath, 'utf-8');
    clientCode = clientCode.replace(/'@zodios\/core';/, "'./hack.js';");

    clientCode = clientCode.replace(
      /const microsoft_graph_attachment = z\s+\.object\({[\s\S]*?}\)\s+\.strict\(\);/,
      (match) => match.replace(/\.strict\(\);/, '.passthrough();')
    );

    console.log('Stripping unused errors arrays from endpoint definitions...');
    // I didn't make up this crazy regex myself; you know who did. It seems works though.
    clientCode = clientCode.replace(/,?\s*errors:\s*\[[\s\S]*?],?(?=\s*})/g, '');

    fs.writeFileSync(clientFilePath, clientCode);

    return true;
  } catch (error) {
    throw new Error(`Error generating client code: ${error.message}`);
  }
}

