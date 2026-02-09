/**
 * Download Link Generator for Files (OneDrive, SharePoint) with direct download URLs
 */

import GraphClient from './graph-client.js';
import logger from './logger.js';
import { getCloudEndpoints } from './cloud-config.js';
import type { AppSecrets } from './secrets.js';
import { getRequestTokens } from './request-context.js';

export interface DownloadLink {
  fileId: string;
  fileName: string;
  downloadUrl: string;
  webUrl?: string;
  size?: number;
  mimeType?: string;
  expiresIn?: number; // seconds until expiration
}

export class DownloadLinkGenerator {
  private graphClient: GraphClient;
  private secrets: AppSecrets;

  constructor(graphClient: GraphClient, secrets: AppSecrets) {
    this.graphClient = graphClient;
    this.secrets = secrets;
  }

  /**
   * Generate download link for a file
   */
  async generateDownloadLink(
    driveId: string,
    itemId: string,
    accessToken?: string
  ): Promise<DownloadLink | null> {
    try {
      // Get file metadata
      const fileMetadata = await this.graphClient.makeRequest(
        `/drives/${driveId}/items/${itemId}`,
        {
          accessToken,
        }
      );

      if (!fileMetadata || typeof fileMetadata !== 'object') {
        return null;
      }

      const metadata = fileMetadata as Record<string, unknown>;

      // Get download URL
      const downloadUrl = await this.getDownloadUrl(driveId, itemId, accessToken);

      if (!downloadUrl) {
        return null;
      }

      const fileInfo = metadata['file'] as Record<string, unknown> | undefined;
      return {
        fileId: itemId,
        fileName: (metadata['name'] as string) || 'unknown',
        downloadUrl,
        webUrl: (metadata['webUrl'] as string) || undefined,
        size: (metadata['size'] as number) || undefined,
        mimeType: (fileInfo?.['mimeType'] as string) || undefined,
        expiresIn: 3600, // Default 1 hour
      };
    } catch (error) {
      logger.error(`Failed to generate download link: ${error}`);
      return null;
    }
  }

  /**
   * Generate download link from SharePoint site
   */
  async generateDownloadLinkFromSite(
    siteId: string,
    itemId: string,
    accessToken?: string
  ): Promise<DownloadLink | null> {
    try {
      // Get file metadata
      const fileMetadata = await this.graphClient.makeRequest(`/sites/${siteId}/items/${itemId}`, {
        accessToken,
      });

      if (!fileMetadata || typeof fileMetadata !== 'object') {
        return null;
      }

      const metadata = fileMetadata as Record<string, unknown>;
      const parentRef = metadata['parentReference'] as Record<string, unknown> | undefined;
      const driveId = parentRef?.['driveId'] as string | undefined;

      if (!driveId) {
        return null;
      }

      // Use drive-based download
      return this.generateDownloadLink(driveId, itemId, accessToken);
    } catch (error) {
      logger.error(`Failed to generate download link from site: ${error}`);
      return null;
    }
  }

  /**
   * Generate download link from webUrl
   */
  async generateDownloadLinkFromWebUrl(
    webUrl: string,
    accessToken?: string
  ): Promise<DownloadLink | null> {
    try {
      // Extract site and item ID from webUrl
      const siteMatch = webUrl.match(/\/sites\/([^/]+)/);
      const itemMatch = webUrl.match(/\/items\/([^/?]+)/);

      if (siteMatch && itemMatch) {
        const siteId = siteMatch[1];
        const itemId = itemMatch[1];
        return this.generateDownloadLinkFromSite(siteId, itemId, accessToken);
      }

      // Try to get item by URL
      const encodedUrl = encodeURIComponent(webUrl);
      const itemResponse = await this.graphClient.makeRequest(
        `/sites/getByPath(path='${encodedUrl}')/driveItem`,
        {
          accessToken,
        }
      );

      if (itemResponse && typeof itemResponse === 'object') {
        const item = itemResponse as Record<string, unknown>;
        const parentRef = item['parentReference'] as Record<string, unknown> | undefined;
        const driveId = parentRef?.['driveId'] as string | undefined;
        const itemId = item['id'] as string | undefined;

        if (driveId && itemId) {
          return this.generateDownloadLink(driveId, itemId, accessToken);
        }
      }

      return null;
    } catch (error) {
      logger.error(`Failed to generate download link from webUrl: ${error}`);
      return null;
    }
  }

  /**
   * Get download URL for a file
   */
  private async getDownloadUrl(
    driveId: string,
    itemId: string,
    accessToken?: string
  ): Promise<string | null> {
    try {
      // Get download URL from content endpoint
      const cloudEndpoints = getCloudEndpoints(this.secrets.cloudType);

      // Use access token from context if not provided
      const token = accessToken || getRequestTokens()?.accessToken;
      if (!token) {
        logger.warn('No access token available for download URL generation');
        return null;
      }

      // Create direct download URL with access token
      // Note: In production, you might want to create a temporary download link via Graph API
      const downloadUrl = `${cloudEndpoints.graphApi}/v1.0/drives/${driveId}/items/${itemId}/content`;

      return downloadUrl;
    } catch (error) {
      logger.error(`Failed to get download URL: ${error}`);
      return null;
    }
  }

  /**
   * Generate multiple download links
   */
  async generateDownloadLinks(
    files: Array<{ driveId?: string; siteId?: string; itemId: string; webUrl?: string }>,
    accessToken?: string
  ): Promise<DownloadLink[]> {
    const links: DownloadLink[] = [];

    for (const file of files) {
      let link: DownloadLink | null = null;

      if (file.driveId && file.itemId) {
        link = await this.generateDownloadLink(file.driveId, file.itemId, accessToken);
      } else if (file.siteId && file.itemId) {
        link = await this.generateDownloadLinkFromSite(file.siteId, file.itemId, accessToken);
      } else if (file.webUrl) {
        link = await this.generateDownloadLinkFromWebUrl(file.webUrl, accessToken);
      }

      if (link) {
        links.push(link);
      }
    }

    return links;
  }

  /**
   * Add download links to search results
   */
  async addDownloadLinksToResults(
    results: unknown[],
    accessToken?: string
  ): Promise<Array<unknown & { downloadLink?: DownloadLink }>> {
    const enriched: Array<unknown & { downloadLink?: DownloadLink }> = [];

    for (const result of results) {
      if (typeof result === 'object' && result !== null) {
        const obj = result as Record<string, unknown>;

        // Check if it's a file/driveItem
        const entityType = obj['@odata.type'] as string | undefined;
        const isFile =
          entityType?.includes('driveItem') ||
          obj['file'] !== undefined ||
          obj['webUrl']?.toString().includes('/sites/') ||
          obj['webUrl']?.toString().includes('/drives/');

        if (isFile) {
          const webUrl = obj['webUrl'] as string | undefined;
          const parentRef = obj['parentReference'] as Record<string, unknown> | undefined;
          const driveId = parentRef?.['driveId'] as string | undefined;
          const itemId = obj['id'] as string | undefined;

          if (webUrl || (driveId && itemId)) {
            let downloadLink: DownloadLink | null = null;

            if (driveId && itemId) {
              downloadLink = await this.generateDownloadLink(driveId, itemId, accessToken);
            } else if (webUrl) {
              downloadLink = await this.generateDownloadLinkFromWebUrl(webUrl, accessToken);
            }

            if (downloadLink) {
              enriched.push({
                ...result,
                downloadLink,
              });
              continue;
            }
          }
        }
      }

      // Not a file or couldn't generate link
      enriched.push(result as unknown & { downloadLink?: DownloadLink });
    }

    return enriched;
  }
}

export default DownloadLinkGenerator;
