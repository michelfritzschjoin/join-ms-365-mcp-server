/**
 * Office document content extraction (Word, Excel, PowerPoint).
 * Excel uses Graph Workbook API; Word uses mammoth; PowerPoint uses JSZip + XML text extraction.
 */

import type GraphClient from './graph-client.js';
import logger from './logger.js';

const DEFAULT_MAX_CONTENT_CHARS = 100_000;
const DEFAULT_MAX_EXCEL_ROWS = 500;

/** Result shape for get-content tool. */
export interface OfficeContentResult {
  type: 'word' | 'excel' | 'powerpoint';
  fileName?: string;
  textContent: string;
  structuredContent?: unknown;
  error?: string;
}

/** Graph client interface for Excel (makeRequest returns parsed JSON). */
interface GraphClientForExcel {
  makeRequest(
    endpoint: string,
    options?: { method?: string; queryParams?: Record<string, string> }
  ): Promise<unknown>;
}

/** Graph client interface for binary (getBinaryContent). */
interface GraphClientForBinary {
  getBinaryContent(endpoint: string): Promise<ArrayBuffer>;
}

/**
 * Get Excel workbook content via Graph Workbook API (worksheets + usedRange).
 * Only works for OneDrive for Business / SharePoint; not Consumer.
 */
export async function getExcelContent(
  graphClient: GraphClientForExcel,
  driveId: string,
  itemId: string,
  options?: { sheetName?: string; maxRows?: number }
): Promise<OfficeContentResult> {
  const basePath =
    driveId === 'me'
      ? `/me/drive/items/${itemId}/workbook`
      : `/drives/${driveId}/items/${itemId}/workbook`;
  const maxRows = options?.maxRows ?? DEFAULT_MAX_EXCEL_ROWS;

  try {
    const worksheetsRes = (await graphClient.makeRequest(`${basePath}/worksheets`, {
      method: 'GET',
    })) as { value?: Array<{ id: string; name: string; position: number }> };
    const worksheets = worksheetsRes?.value ?? [];
    if (worksheets.length === 0) {
      return {
        type: 'excel',
        textContent: '',
        structuredContent: { sheets: [], message: 'No worksheets found.' },
      };
    }

    const sheetsContent: Array<{ name: string; values: unknown[][]; text?: string[][] }> = [];
    let fullText = '';

    for (const sheet of worksheets) {
      if (options?.sheetName && sheet.name !== options.sheetName) continue;
      try {
        const rangeRes = (await graphClient.makeRequest(
          `${basePath}/worksheets/${sheet.id}/usedRange`,
          { method: 'GET' }
        )) as { values?: unknown[][]; text?: string[][] };
        const values = (rangeRes?.values ?? []) as unknown[][];
        const text = (rangeRes?.text ?? []) as string[][];
        const limitedValues = values.slice(0, maxRows);
        const limitedText = text.slice(0, maxRows);
        sheetsContent.push({
          name: sheet.name,
          values: limitedValues,
          text: limitedText.length ? limitedText : undefined,
        });
        if (limitedText.length) {
          fullText += `\n### ${sheet.name}\n`;
          for (const row of limitedText) {
            fullText += row.join('\t') + '\n';
          }
        } else if (limitedValues.length) {
          fullText += `\n### ${sheet.name}\n`;
          for (const row of limitedValues) {
            fullText += row.map((c) => String(c ?? '')).join('\t') + '\n';
          }
        }
      } catch (err) {
        logger.warn(`Excel sheet "${sheet.name}" failed: ${(err as Error).message}`);
      }
    }

    return {
      type: 'excel',
      textContent: fullText.trim() || '(No cell content)',
      structuredContent: { sheets: sheetsContent },
    };
  } catch (err) {
    const message = (err as Error).message;
    logger.warn(`Excel content failed: ${message}`);
    return {
      type: 'excel',
      textContent: '',
      error: message,
    };
  }
}

/**
 * Get Word document text via download + mammoth.
 */
export async function getWordContent(
  graphClient: GraphClientForBinary,
  driveId: string,
  itemId: string,
  maxLength = DEFAULT_MAX_CONTENT_CHARS
): Promise<OfficeContentResult> {
  const endpoint =
    driveId === 'me'
      ? `/me/drive/items/${itemId}/content`
      : `/drives/${driveId}/items/${itemId}/content`;
  try {
    const arrayBuffer = await graphClient.getBinaryContent(endpoint);
    const mammoth = await import('mammoth');
    const result = await mammoth.extractRawText({ arrayBuffer });
    let text = result.value?.trim() ?? '';
    if (text.length > maxLength) {
      text = text.slice(0, maxLength) + '\n\n[... truncated]';
    }
    return {
      type: 'word',
      textContent: text || '(No text content)',
    };
  } catch (err) {
    const message = (err as Error).message;
    logger.warn(`Word content failed: ${message}`);
    return {
      type: 'word',
      textContent: '',
      error: message,
    };
  }
}

/** Extract text from a single slide XML string (a:t elements). */
function extractTextFromSlideXml(xml: string): string {
  const matches = xml.match(/<a:t>([^<]*)<\/a:t>/g);
  if (!matches) return '';
  return matches
    .map((m) => m.replace(/<\/?a:t>/g, ''))
    .join(' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Get PowerPoint text via download + JSZip (ppt/slides/slide*.xml).
 */
export async function getPowerPointContent(
  graphClient: GraphClientForBinary,
  driveId: string,
  itemId: string,
  maxLength = DEFAULT_MAX_CONTENT_CHARS
): Promise<OfficeContentResult> {
  const endpoint =
    driveId === 'me'
      ? `/me/drive/items/${itemId}/content`
      : `/drives/${driveId}/items/${itemId}/content`;
  try {
    const arrayBuffer = await graphClient.getBinaryContent(endpoint);
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(arrayBuffer);
    const slides: Array<{ index: number; text: string }> = [];
    let fullText = '';
    const slideFiles = Object.keys(zip.files).filter(
      (n) => n.startsWith('ppt/slides/slide') && n.endsWith('.xml')
    );
    slideFiles.sort((a, b) => {
      const an = parseInt(a.replace(/\D/g, ''), 10);
      const bn = parseInt(b.replace(/\D/g, ''), 10);
      return an - bn;
    });
    for (let i = 0; i < slideFiles.length; i++) {
      const name = slideFiles[i];
      const xml = await zip.files[name].async('string');
      const text = extractTextFromSlideXml(xml);
      slides.push({ index: i + 1, text });
      fullText += `\n--- Slide ${i + 1} ---\n${text}\n`;
    }
    fullText = fullText.trim();
    if (fullText.length > maxLength) {
      fullText = fullText.slice(0, maxLength) + '\n\n[... truncated]';
    }
    return {
      type: 'powerpoint',
      textContent: fullText || '(No text content)',
      structuredContent: { slides },
    };
  } catch (err) {
    const message = (err as Error).message;
    logger.warn(`PowerPoint content failed: ${message}`);
    return {
      type: 'powerpoint',
      textContent: '',
      error: message,
    };
  }
}
