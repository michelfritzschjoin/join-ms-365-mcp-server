import fs from 'fs';

const DEFAULT_OPENAPI_URL =
  'https://raw.githubusercontent.com/microsoftgraph/msgraph-metadata/refs/heads/master/openapi/v1.0/openapi.yaml';

export async function downloadGraphOpenAPI(
  targetDir,
  targetFile,
  openapiUrl = DEFAULT_OPENAPI_URL,
  forceDownload = false
) {
  if (!fs.existsSync(targetDir)) {
    console.log(`Creating directory: ${targetDir}`);
    fs.mkdirSync(targetDir, { recursive: true });
  }

  if (fs.existsSync(targetFile) && !forceDownload) {
    console.log(`OpenAPI specification already exists at ${targetFile}`);
    console.log('Use --force to download again');
    return false;
  }

  console.log(`Downloading OpenAPI specification from ${openapiUrl}`);

  const maxRetries = 3;
  const retryDelay = 2000; // 2 seconds

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      if (attempt > 1) {
        console.log(`Retry attempt ${attempt}/${maxRetries}...`);
        await new Promise((resolve) => setTimeout(resolve, retryDelay * attempt));
      }

      const response = await fetch(openapiUrl);

      if (!response.ok) {
        // Retry on 5xx server errors
        if (response.status >= 500 && attempt < maxRetries) {
          console.warn(`Server error ${response.status} ${response.statusText}, retrying...`);
          continue;
        }
        throw new Error(`Failed to download: ${response.status} ${response.statusText}`);
      }

      const content = await response.text();
      fs.writeFileSync(targetFile, content);
      console.log(`OpenAPI specification downloaded to ${targetFile}`);
      return true;
    } catch (error) {
      // Retry on network errors or 5xx errors
      if (attempt < maxRetries && (error.message.includes('503') || error.message.includes('50'))) {
        console.warn(`Error downloading (attempt ${attempt}/${maxRetries}): ${error.message}`);
        continue;
      }
      console.error('Error downloading OpenAPI specification:', error.message);
      throw error;
    }
  }
}
