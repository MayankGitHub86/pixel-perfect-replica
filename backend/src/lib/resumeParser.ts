// backend/src/lib/resumeParser.ts
// PDF and DOCX resume text extraction with scanned PDF detection.

import { PDFParse } from 'pdf-parse';
import mammoth from 'mammoth';
import path from 'path';
import { logger } from '../logger';

export interface ResumeParseResult {
  text: string;
  isScanned: boolean;
  wordCount: number;
  format: 'pdf' | 'docx' | 'unknown';
  error?: string;
}

const SCANNED_TEXT_THRESHOLD_CHARS = 50;

/**
 * Extract plain text from PDF buffer.
 */
export async function parsePdfBuffer(buffer: Buffer): Promise<ResumeParseResult> {
  let parser: PDFParse | null = null;
  try {
    parser = new PDFParse({ data: buffer });
    const textResult = await parser.getText();
    const rawText = textResult.text || '';
    const trimmedText = rawText.trim();
    const wordCount = trimmedText ? trimmedText.split(/\s+/).length : 0;
    const numPages = textResult.total || 0;

    // If PDF has pages but virtually no selectable text, it's likely a scanned image
    const isScanned = numPages > 0 && trimmedText.length < SCANNED_TEXT_THRESHOLD_CHARS;

    return {
      text: trimmedText,
      isScanned,
      wordCount,
      format: 'pdf',
    };
  } catch (err: any) {
    logger.warn({ err: err.message }, 'Failed to parse PDF resume buffer');
    return {
      text: '',
      isScanned: false,
      wordCount: 0,
      format: 'pdf',
      error: `PDF parse failure: ${err.message}`,
    };
  } finally {
    if (parser) {
      await parser.destroy().catch(() => {});
    }
  }
}

/**
 * Extract plain text from DOCX buffer.
 */
export async function parseDocxBuffer(buffer: Buffer): Promise<ResumeParseResult> {
  try {
    const result = await mammoth.extractRawText({ buffer });
    const rawText = result.value || '';
    const trimmedText = rawText.trim();
    const wordCount = trimmedText ? trimmedText.split(/\s+/).length : 0;

    return {
      text: trimmedText,
      isScanned: false,
      wordCount,
      format: 'docx',
    };
  } catch (err: any) {
    logger.warn({ err: err.message }, 'Failed to parse DOCX resume buffer');
    return {
      text: '',
      isScanned: false,
      wordCount: 0,
      format: 'docx',
      error: `DOCX parse failure: ${err.message}`,
    };
  }
}

/**
 * Generic resume buffer parser based on file extension.
 */
export async function parseResumeBuffer(buffer: Buffer, filename: string): Promise<ResumeParseResult> {
  const ext = path.extname(filename).toLowerCase();

  if (ext === '.pdf') {
    return parsePdfBuffer(buffer);
  } else if (ext === '.docx' || ext === '.doc') {
    return parseDocxBuffer(buffer);
  }

  return {
    text: '',
    isScanned: false,
    wordCount: 0,
    format: 'unknown',
    error: `Unsupported resume file extension: ${ext}`,
  };
}
