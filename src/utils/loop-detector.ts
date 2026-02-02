/**
 * Microsoft Loop File Detection Utility
 *
 * Detects and handles Microsoft Loop files stored in OneDrive and SharePoint.
 *
 * Loop files are stored in various locations:
 * - OneDrive: In folders like `OneNote Loop files`, `Microsoft Teams Chat files`, `Meetings`, `Attachments`
 * - SharePoint: In team channels and meeting folders
 * - SharePoint Embedded: In user-owned or shared containers
 *
 * Reference: https://learn.microsoft.com/en-us/microsoft-365/loop/loop-storage
 */

import logger from '../logger.js';

/**
 * Known Loop-specific folder names where Loop components are stored
 */
const LOOP_FOLDER_PATTERNS: readonly string[] = [
  'OneNote Loop files',
  'Microsoft Teams Chat files',
  'Meetings',
  'Attachments',
  'Whiteboard/Components',
  'Loop',
] as const;

/**
 * Known Loop-specific URL patterns
 */
const LOOP_URL_PATTERNS: readonly string[] = [
  'loop.microsoft.com',
  'loop.cloud.microsoft',
  '.loop.',
] as const;

/**
 * Known Loop file extensions
 */
const LOOP_FILE_EXTENSIONS: readonly string[] = ['.loop', '.fluid'] as const;

/**
 * Known MIME types for Loop/Fluid files
 */
const LOOP_MIME_TYPES: readonly string[] = [
  'application/vnd.ms-loop',
  'application/fluid',
  'application/vnd.microsoft-fluid',
] as const;

/**
 * Result of Loop file detection with details about why it was detected
 */
export interface LoopFileDetectionResult {
  /** Whether the file is a Loop file */
  isLoopFile: boolean;
  /** The detection method used */
  detectionMethod?: 'url' | 'folder' | 'extension' | 'mimeType' | 'name';
  /** The specific pattern that matched */
  matchedPattern?: string;
  /** Confidence level of the detection (high, medium, low) */
  confidence: 'high' | 'medium' | 'low';
}

/**
 * Parsed Loop content structure
 */
export interface ParsedLoopContent {
  /** Whether parsing was successful */
  success: boolean;
  /** The raw content */
  rawContent: string;
  /** Extracted text content (if available) */
  textContent?: string;
  /** Content type detected */
  contentType?: 'json' | 'fluid' | 'text' | 'unknown';
  /** Any extracted metadata */
  metadata?: Record<string, unknown>;
  /** Error message if parsing failed */
  error?: string;
}

/**
 * Detects if a DriveItem is a Microsoft Loop file
 *
 * @param driveItem - The DriveItem object from Microsoft Graph API
 * @returns Detection result with details
 */
export function detectLoopFile(driveItem: Record<string, unknown>): LoopFileDetectionResult {
  // Check WebUrl for Loop-specific patterns
  const webUrl = driveItem.webUrl as string | undefined;
  if (webUrl) {
    for (const pattern of LOOP_URL_PATTERNS) {
      if (webUrl.toLowerCase().includes(pattern.toLowerCase())) {
        logger.debug(`Loop file detected by URL pattern: ${pattern}`);
        return {
          isLoopFile: true,
          detectionMethod: 'url',
          matchedPattern: pattern,
          confidence: 'high',
        };
      }
    }
  }

  // Check file extension
  const fileName = (driveItem.name as string) || '';
  const fileNameLower = fileName.toLowerCase();
  for (const ext of LOOP_FILE_EXTENSIONS) {
    if (fileNameLower.endsWith(ext)) {
      logger.debug(`Loop file detected by extension: ${ext}`);
      return {
        isLoopFile: true,
        detectionMethod: 'extension',
        matchedPattern: ext,
        confidence: 'high',
      };
    }
  }

  // Check MIME type
  const file = driveItem.file as Record<string, unknown> | undefined;
  const mimeType = file?.mimeType as string | undefined;
  if (mimeType) {
    for (const loopMime of LOOP_MIME_TYPES) {
      if (mimeType.toLowerCase().includes(loopMime.toLowerCase())) {
        logger.debug(`Loop file detected by MIME type: ${loopMime}`);
        return {
          isLoopFile: true,
          detectionMethod: 'mimeType',
          matchedPattern: loopMime,
          confidence: 'high',
        };
      }
    }
  }

  // Check parent folder path
  const parentReference = driveItem.parentReference as Record<string, unknown> | undefined;
  const parentPath = parentReference?.path as string | undefined;
  if (parentPath) {
    for (const folderPattern of LOOP_FOLDER_PATTERNS) {
      if (parentPath.includes(folderPattern)) {
        logger.debug(`Loop file detected by folder pattern: ${folderPattern}`);
        return {
          isLoopFile: true,
          detectionMethod: 'folder',
          matchedPattern: folderPattern,
          confidence: 'medium',
        };
      }
    }
  }

  // Check if file name contains "Loop" or specific Loop patterns
  if (
    fileNameLower.includes('loop') ||
    fileNameLower.includes('fluid') ||
    fileNameLower.includes('collaborative note')
  ) {
    logger.debug(`Loop file detected by name pattern: ${fileName}`);
    return {
      isLoopFile: true,
      detectionMethod: 'name',
      matchedPattern: fileName,
      confidence: 'low',
    };
  }

  return {
    isLoopFile: false,
    confidence: 'high',
  };
}

/**
 * Simple check if a DriveItem is a Loop file
 *
 * @param driveItem - The DriveItem object from Microsoft Graph API
 * @returns true if the file is a Loop file
 */
export function isLoopFile(driveItem: Record<string, unknown>): boolean {
  return detectLoopFile(driveItem).isLoopFile;
}

/**
 * Parses Loop file content and extracts readable text
 *
 * Loop files may be stored in various formats including JSON-based Fluid framework format.
 * This function attempts to extract readable text content from the raw file data.
 *
 * @param content - The raw content of the Loop file
 * @returns Parsed content with extracted text
 */
export function parseLoopContent(content: string): ParsedLoopContent {
  if (!content || typeof content !== 'string') {
    return {
      success: false,
      rawContent: '',
      error: 'No content provided',
    };
  }

  // Try to detect content type
  const trimmedContent = content.trim();

  // Check if it's JSON (Fluid format)
  if (trimmedContent.startsWith('{') || trimmedContent.startsWith('[')) {
    try {
      const parsed = JSON.parse(trimmedContent);
      const textContent = extractTextFromFluidJson(parsed);

      return {
        success: true,
        rawContent: content,
        textContent: textContent || undefined,
        contentType: 'fluid',
        metadata: extractMetadataFromFluidJson(parsed),
      };
    } catch {
      // Not valid JSON, continue with other methods
      logger.debug('Loop content is not valid JSON');
    }
  }

  // Check if it looks like plain text
  if (isPrintableText(trimmedContent)) {
    return {
      success: true,
      rawContent: content,
      textContent: content,
      contentType: 'text',
    };
  }

  // Unknown format - return raw content
  return {
    success: true,
    rawContent: content,
    contentType: 'unknown',
  };
}

/**
 * Extracts text content from Fluid JSON structure
 *
 * @param fluidJson - The parsed Fluid JSON object
 * @returns Extracted text content
 */
function extractTextFromFluidJson(fluidJson: unknown): string | null {
  if (!fluidJson || typeof fluidJson !== 'object') {
    return null;
  }

  const textParts: string[] = [];

  // Recursive function to extract text from nested structures
  function extractText(obj: unknown): void {
    if (typeof obj === 'string') {
      textParts.push(obj);
      return;
    }

    if (Array.isArray(obj)) {
      for (const item of obj) {
        extractText(item);
      }
      return;
    }

    if (typeof obj === 'object' && obj !== null) {
      const record = obj as Record<string, unknown>;

      // Common text field names in Fluid format
      const textFields = [
        'text',
        'content',
        'value',
        'title',
        'description',
        'body',
        'data',
        'str',
        'string',
      ];

      for (const field of textFields) {
        if (typeof record[field] === 'string') {
          textParts.push(record[field] as string);
        }
      }

      // Recursively process all properties
      for (const value of Object.values(record)) {
        extractText(value);
      }
    }
  }

  extractText(fluidJson);

  // Clean up and deduplicate
  const uniqueTexts = [...new Set(textParts.filter((t) => t.trim().length > 0))];

  if (uniqueTexts.length === 0) {
    return null;
  }

  return uniqueTexts.join('\n\n');
}

/**
 * Extracts metadata from Fluid JSON structure
 *
 * @param fluidJson - The parsed Fluid JSON object
 * @returns Extracted metadata
 */
function extractMetadataFromFluidJson(fluidJson: unknown): Record<string, unknown> | undefined {
  if (!fluidJson || typeof fluidJson !== 'object') {
    return undefined;
  }

  const record = fluidJson as Record<string, unknown>;
  const metadata: Record<string, unknown> = {};

  // Common metadata field names
  const metadataFields = [
    'id',
    'version',
    'type',
    'schema',
    'created',
    'modified',
    'author',
    'lastModifiedBy',
    'title',
    'name',
  ];

  for (const field of metadataFields) {
    if (record[field] !== undefined) {
      metadata[field] = record[field];
    }
  }

  return Object.keys(metadata).length > 0 ? metadata : undefined;
}

/**
 * Checks if a string contains mostly printable text
 *
 * @param content - The content to check
 * @returns true if the content is mostly printable text
 */
function isPrintableText(content: string): boolean {
  if (!content || content.length === 0) {
    return false;
  }

  // Count printable characters
  let printableCount = 0;
  for (const char of content) {
    const code = char.charCodeAt(0);
    // Printable ASCII (32-126) plus common control chars (newline, tab, carriage return)
    if ((code >= 32 && code <= 126) || code === 10 || code === 13 || code === 9 || code > 127) {
      printableCount++;
    }
  }

  // If more than 90% is printable, consider it text
  return printableCount / content.length > 0.9;
}

/**
 * Formats Loop file information for display
 *
 * @param driveItem - The DriveItem object
 * @param detectionResult - The detection result
 * @returns Formatted string with Loop file information
 */
export function formatLoopFileInfo(
  driveItem: Record<string, unknown>,
  detectionResult: LoopFileDetectionResult
): string {
  const name = (driveItem.name as string) || 'Unknown';
  const webUrl = (driveItem.webUrl as string) || '';
  const size = (driveItem.size as number) || 0;
  const lastModified = (driveItem.lastModifiedDateTime as string) || '';

  const parts = [
    `📋 Loop File: ${name}`,
    `   Detection: ${detectionResult.detectionMethod} (${detectionResult.confidence} confidence)`,
  ];

  if (detectionResult.matchedPattern) {
    parts.push(`   Pattern: ${detectionResult.matchedPattern}`);
  }

  if (size > 0) {
    parts.push(`   Size: ${formatFileSize(size)}`);
  }

  if (lastModified) {
    parts.push(`   Modified: ${new Date(lastModified).toLocaleString()}`);
  }

  if (webUrl) {
    parts.push(`   URL: ${webUrl}`);
  }

  return parts.join('\n');
}

/**
 * Formats file size for display
 *
 * @param bytes - File size in bytes
 * @returns Formatted file size string
 */
function formatFileSize(bytes: number): string {
  if (bytes === 0) return '0 B';

  const units = ['B', 'KB', 'MB', 'GB', 'TB'];
  const k = 1024;
  const i = Math.floor(Math.log(bytes) / Math.log(k));

  return `${parseFloat((bytes / Math.pow(k, i)).toFixed(2))} ${units[i]}`;
}

/**
 * Gets the list of known Loop folder patterns
 *
 * @returns Array of folder patterns
 */
export function getLoopFolderPatterns(): readonly string[] {
  return LOOP_FOLDER_PATTERNS;
}

/**
 * Gets the list of known Loop file extensions
 *
 * @returns Array of file extensions
 */
export function getLoopFileExtensions(): readonly string[] {
  return LOOP_FILE_EXTENSIONS;
}

/**
 * Gets the list of known Loop MIME types
 *
 * @returns Array of MIME types
 */
export function getLoopMimeTypes(): readonly string[] {
  return LOOP_MIME_TYPES;
}
