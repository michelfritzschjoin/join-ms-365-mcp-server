/**
 * Office file type detection for Word, Excel, and PowerPoint.
 * Used to mark driveItems so tools (e.g. get-content) can offer Office content extraction.
 */

/** File extensions for Word documents. */
const WORD_EXTENSIONS = new Set(['.docx', '.doc', '.dot', '.dotx', '.dotm', '.rtf']);

/** File extensions for Excel workbooks. */
const EXCEL_EXTENSIONS = new Set(['.xlsx', '.xls', '.xlsm', '.xlsb']);

/** File extensions for PowerPoint presentations. */
const POWERPOINT_EXTENSIONS = new Set(['.pptx', '.ppt', '.pps', '.ppsx', '.pot', '.potx']);

/** MIME type prefixes for Word. */
const WORD_MIME_PREFIXES = [
  'application/vnd.openxmlformats-officedocument.wordprocessingml',
  'application/msword',
  'application/rtf',
];

/** MIME type prefixes for Excel. */
const EXCEL_MIME_PREFIXES = [
  'application/vnd.openxmlformats-officedocument.spreadsheetml',
  'application/vnd.ms-excel',
];

/** MIME type prefixes for PowerPoint. */
const POWERPOINT_MIME_PREFIXES = [
  'application/vnd.openxmlformats-officedocument.presentationml',
  'application/vnd.ms-powerpoint',
];

export type OfficeType = 'word' | 'excel' | 'powerpoint';

export interface OfficeFileDetectionResult {
  isWord: boolean;
  isExcel: boolean;
  isPowerPoint: boolean;
  officeType?: OfficeType;
}

/**
 * Detects if a DriveItem is a Word, Excel, or PowerPoint file.
 * Uses fileExtension (from file.fileExtension or name) and file.mimeType.
 *
 * @param driveItem - The DriveItem object from Microsoft Graph API
 * @returns Detection result with isWord, isExcel, isPowerPoint, and officeType
 */
export function detectOfficeFile(driveItem: Record<string, unknown>): OfficeFileDetectionResult {
  const result: OfficeFileDetectionResult = {
    isWord: false,
    isExcel: false,
    isPowerPoint: false,
  };

  const file = driveItem.file as Record<string, unknown> | undefined;
  const name = (driveItem.name as string) || '';
  const nameLower = name.toLowerCase();
  const ext = nameLower.includes('.') ? nameLower.slice(nameLower.lastIndexOf('.')) : '';
  const mimeType = (file?.mimeType as string) || '';
  const mimeLower = mimeType.toLowerCase();
  const fileExtension = (file?.fileExtension as string) || ext || '';

  const extToCheck = fileExtension ? `.${fileExtension.replace(/^\./, '')}` : ext;

  if (extToCheck && WORD_EXTENSIONS.has(extToCheck)) {
    result.isWord = true;
    result.officeType = 'word';
    return result;
  }
  if (extToCheck && EXCEL_EXTENSIONS.has(extToCheck)) {
    result.isExcel = true;
    result.officeType = 'excel';
    return result;
  }
  if (extToCheck && POWERPOINT_EXTENSIONS.has(extToCheck)) {
    result.isPowerPoint = true;
    result.officeType = 'powerpoint';
    return result;
  }

  for (const prefix of WORD_MIME_PREFIXES) {
    if (mimeLower.startsWith(prefix)) {
      result.isWord = true;
      result.officeType = 'word';
      return result;
    }
  }
  for (const prefix of EXCEL_MIME_PREFIXES) {
    if (mimeLower.startsWith(prefix)) {
      result.isExcel = true;
      result.officeType = 'excel';
      return result;
    }
  }
  for (const prefix of POWERPOINT_MIME_PREFIXES) {
    if (mimeLower.startsWith(prefix)) {
      result.isPowerPoint = true;
      result.officeType = 'powerpoint';
      return result;
    }
  }

  return result;
}

/**
 * Applies Office detection to a driveItem by mutating it with isWord, isExcel, isPowerPoint, officeType.
 */
export function applyOfficeDetection(driveItem: Record<string, unknown>): void {
  const detection = detectOfficeFile(driveItem);
  if (detection.officeType) {
    driveItem.isWord = detection.isWord;
    driveItem.isExcel = detection.isExcel;
    driveItem.isPowerPoint = detection.isPowerPoint;
    driveItem.officeType = detection.officeType;
  }
}
