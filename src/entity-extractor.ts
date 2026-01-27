/**
 * Entity Extractor for extracting entities from search results
 */

export interface ExtractedEntity {
  type: string;
  value: string;
  confidence: number;
  source: string;
  metadata?: Record<string, unknown>;
}

export interface ExtractedInfo {
  entities: ExtractedEntity[];
  keywords: string[];
  ids: Record<string, string[]>; // entityType -> array of IDs
  sites: string[];
  drives: string[];
  teams: string[];
  users: string[];
  files: string[];
}

export class EntityExtractor {
  /**
   * Extract entities and information from search results
   */
  extractFromResults(results: unknown[]): ExtractedInfo {
    const extracted: ExtractedInfo = {
      entities: [],
      keywords: [],
      ids: {},
      sites: [],
      drives: [],
      teams: [],
      users: [],
      files: [],
    };

    for (const item of results) {
      if (typeof item !== 'object' || item === null) {
        continue;
      }

      const obj = item as Record<string, unknown>;

      // Extract based on entity type
      const entityType = this.getEntityType(obj);
      if (!entityType) {
        continue;
      }

      // Extract ID
      const id = this.extractId(obj, entityType);
      if (id) {
        if (!extracted.ids[entityType]) {
          extracted.ids[entityType] = [];
        }
        if (!extracted.ids[entityType].includes(id)) {
          extracted.ids[entityType].push(id);
        }
      }

      // Extract keywords from text fields
      const keywords = this.extractKeywords(obj);
      for (const keyword of keywords) {
        if (!extracted.keywords.includes(keyword)) {
          extracted.keywords.push(keyword);
        }
      }

      // Extract specific entity types
      if (entityType === 'site' || entityType === 'driveItem') {
        const siteId = this.extractSiteId(obj);
        if (siteId && !extracted.sites.includes(siteId)) {
          extracted.sites.push(siteId);
        }

        const driveId = this.extractDriveId(obj);
        if (driveId && !extracted.drives.includes(driveId)) {
          extracted.drives.push(driveId);
        }

        const fileId = this.extractFileId(obj);
        if (fileId && !extracted.files.includes(fileId)) {
          extracted.files.push(fileId);
        }
      }

      if (entityType === 'chatMessage' || entityType === 'message') {
        const teamId = this.extractTeamId(obj);
        if (teamId && !extracted.teams.includes(teamId)) {
          extracted.teams.push(teamId);
        }
      }

      if (entityType === 'person' || entityType === 'user') {
        const userId = this.extractUserId(obj);
        if (userId && !extracted.users.includes(userId)) {
          extracted.users.push(userId);
        }
      }

      // Create entity record
      extracted.entities.push({
        type: entityType,
        value: this.extractDisplayName(obj) || id || '',
        confidence: this.calculateConfidence(obj),
        source: this.extractSource(obj),
        metadata: this.extractMetadata(obj),
      });
    }

    return extracted;
  }

  /**
   * Get entity type from result item
   */
  private getEntityType(obj: Record<string, unknown>): string | undefined {
    // Check explicit entity type
    if (typeof obj['@odata.type'] === 'string') {
      const type = obj['@odata.type'] as string;
      if (type.includes('site')) return 'site';
      if (type.includes('driveItem')) return 'driveItem';
      if (type.includes('message')) return 'message';
      if (type.includes('chatMessage')) return 'chatMessage';
      if (type.includes('person')) return 'person';
      if (type.includes('event')) return 'event';
    }

    // Infer from structure
    if (obj['webUrl'] && typeof obj['webUrl'] === 'string') {
      const url = obj['webUrl'] as string;
      if (url.includes('/sites/')) return 'site';
      if (url.includes('/drives/')) return 'driveItem';
    }

    if (obj['siteId']) return 'site';
    if (obj['driveId']) return 'driveItem';
    if (obj['teamId']) return 'chatMessage';
    if (obj['userId'] || obj['userPrincipalName']) return 'person';
    if (obj['start'] || obj['end']) return 'event';
    if (obj['subject'] || obj['from']) return 'message';

    return undefined;
  }

  /**
   * Extract ID from entity
   */
  private extractId(obj: Record<string, unknown>, entityType: string): string | undefined {
    // Try common ID fields
    const idFields = ['id', 'objectId', 'itemId', 'messageId', 'eventId', 'siteId', 'driveId'];
    for (const field of idFields) {
      if (typeof obj[field] === 'string') {
        return obj[field] as string;
      }
    }

    // Try entity-specific fields
    if (entityType === 'site' && obj['id']) {
      return obj['id'] as string;
    }
    if (entityType === 'driveItem' && obj['id']) {
      return obj['id'] as string;
    }
    if (entityType === 'message' && obj['id']) {
      return obj['id'] as string;
    }

    return undefined;
  }

  /**
   * Extract keywords from text fields
   */
  private extractKeywords(obj: Record<string, unknown>): string[] {
    const keywords: string[] = [];
    const textFields = [
      'name',
      'title',
      'subject',
      'displayName',
      'content',
      'body',
      'description',
    ];

    for (const field of textFields) {
      if (typeof obj[field] === 'string') {
        const text = obj[field] as string;
        // Extract words (simple tokenization)
        const words = text
          .toLowerCase()
          .split(/\s+/)
          .filter((w) => w.length > 3) // Only meaningful words
          .slice(0, 10); // Limit
        keywords.push(...words);
      }
    }

    return keywords;
  }

  /**
   * Extract site ID
   */
  private extractSiteId(obj: Record<string, unknown>): string | undefined {
    if (typeof obj['siteId'] === 'string') {
      return obj['siteId'] as string;
    }
    if (typeof obj['parentReference'] === 'object' && obj['parentReference'] !== null) {
      const parent = obj['parentReference'] as Record<string, unknown>;
      if (typeof parent['siteId'] === 'string') {
        return parent['siteId'] as string;
      }
    }
    return undefined;
  }

  /**
   * Extract drive ID
   */
  private extractDriveId(obj: Record<string, unknown>): string | undefined {
    if (typeof obj['driveId'] === 'string') {
      return obj['driveId'] as string;
    }
    if (typeof obj['parentReference'] === 'object' && obj['parentReference'] !== null) {
      const parent = obj['parentReference'] as Record<string, unknown>;
      if (typeof parent['driveId'] === 'string') {
        return parent['driveId'] as string;
      }
    }
    return undefined;
  }

  /**
   * Extract file ID
   */
  private extractFileId(obj: Record<string, unknown>): string | undefined {
    if (typeof obj['id'] === 'string' && this.getEntityType(obj) === 'driveItem') {
      return obj['id'] as string;
    }
    return undefined;
  }

  /**
   * Extract team ID
   */
  private extractTeamId(obj: Record<string, unknown>): string | undefined {
    if (typeof obj['teamId'] === 'string') {
      return obj['teamId'] as string;
    }
    if (typeof obj['channelIdentity'] === 'object' && obj['channelIdentity'] !== null) {
      const channel = obj['channelIdentity'] as Record<string, unknown>;
      if (typeof channel['teamId'] === 'string') {
        return channel['teamId'] as string;
      }
    }
    return undefined;
  }

  /**
   * Extract user ID
   */
  private extractUserId(obj: Record<string, unknown>): string | undefined {
    if (typeof obj['id'] === 'string') {
      return obj['id'] as string;
    }
    if (typeof obj['userPrincipalName'] === 'string') {
      return obj['userPrincipalName'] as string;
    }
    if (typeof obj['mail'] === 'string') {
      return obj['mail'] as string;
    }
    return undefined;
  }

  /**
   * Extract display name
   */
  private extractDisplayName(obj: Record<string, unknown>): string | undefined {
    const nameFields = ['displayName', 'name', 'title', 'subject'];
    for (const field of nameFields) {
      if (typeof obj[field] === 'string') {
        return obj[field] as string;
      }
    }
    return undefined;
  }

  /**
   * Calculate confidence score
   */
  private calculateConfidence(obj: Record<string, unknown>): number {
    let score = 0.5; // Base confidence

    // Higher confidence if has ID
    if (obj['id']) score += 0.2;

    // Higher confidence if has display name
    if (obj['displayName'] || obj['name'] || obj['title']) score += 0.2;

    // Higher confidence if has webUrl
    if (obj['webUrl']) score += 0.1;

    return Math.min(score, 1.0);
  }

  /**
   * Extract source information
   */
  private extractSource(obj: Record<string, unknown>): string {
    if (obj['webUrl'] && typeof obj['webUrl'] === 'string') {
      const url = obj['webUrl'] as string;
      // SECURITY: Use exact hostname matching with whitelist instead of substring matching
      try {
        const urlObj = new URL(url);
        const hostname = urlObj.hostname.toLowerCase();
        // Use exact matching or endsWith for specific domains to prevent bypass
        if (hostname === 'sharepoint.com' || hostname.endsWith('.sharepoint.com')) {
          return 'sharepoint';
        }
        if (hostname === 'teams.microsoft.com' || hostname.endsWith('.teams.microsoft.com')) {
          return 'teams';
        }
        if (hostname === 'outlook.office.com' || hostname.endsWith('.outlook.office.com')) {
          return 'outlook';
        }
      } catch {
        // Invalid URL, skip
      }
    }

    const entityType = this.getEntityType(obj);
    return entityType || 'unknown';
  }

  /**
   * Extract metadata
   */
  private extractMetadata(obj: Record<string, unknown>): Record<string, unknown> {
    const metadata: Record<string, unknown> = {};

    // Extract relevant metadata fields
    const metadataFields = [
      'createdDateTime',
      'lastModifiedDateTime',
      'createdBy',
      'lastModifiedBy',
      'webUrl',
      'size',
      'mimeType',
    ];

    for (const field of metadataFields) {
      if (obj[field] !== undefined) {
        metadata[field] = obj[field];
      }
    }

    return metadata;
  }
}

export default EntityExtractor;
