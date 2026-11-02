// backend/src/__tests__/attachmentService.test.ts
import { describe, it, expect } from 'vitest';
import {
  validateAttachment,
  sanitizeFilename,
  getStoragePath,
  MAX_ATTACHMENT_SIZE_BYTES,
  ATTACHMENTS_STORAGE_DIR,
} from '../lib/attachmentService';
import path from 'path';

describe('Attachment Service', () => {
  describe('validateAttachment', () => {
    it('accepts valid PDF file', () => {
      const res = validateAttachment('resume.pdf', 'application/pdf', 1024 * 50);
      expect(res.valid).toBe(true);
      expect(res.extension).toBe('.pdf');
    });

    it('accepts valid DOCX file', () => {
      const res = validateAttachment(
        'candidate_cv.docx',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        1024 * 100
      );
      expect(res.valid).toBe(true);
      expect(res.extension).toBe('.docx');
    });

    it('rejects executable file (.exe)', () => {
      const res = validateAttachment('malware.exe', 'application/octet-stream', 500);
      expect(res.valid).toBe(false);
      expect(res.reason).toMatch(/security/i);
    });

    it('rejects zip archive (.zip)', () => {
      const res = validateAttachment('archive.zip', 'application/zip', 5000);
      expect(res.valid).toBe(false);
      expect(res.reason).toMatch(/security/i);
    });

    it('rejects unsupported extensions (.txt)', () => {
      const res = validateAttachment('notes.txt', 'text/plain', 500);
      expect(res.valid).toBe(false);
      expect(res.reason).toMatch(/Unsupported file extension/);
    });

    it('rejects files larger than 10MB', () => {
      const res = validateAttachment('large_cv.pdf', 'application/pdf', MAX_ATTACHMENT_SIZE_BYTES + 1);
      expect(res.valid).toBe(false);
      expect(res.reason).toMatch(/exceeds maximum size/i);
    });

    it('rejects empty (0 byte) files', () => {
      const res = validateAttachment('empty.pdf', 'application/pdf', 0);
      expect(res.valid).toBe(false);
      expect(res.reason).toMatch(/empty/i);
    });
  });

  describe('sanitizeFilename', () => {
    it('strips directory traversal paths', () => {
      const clean = sanitizeFilename('../../etc/passwd');
      expect(clean).toBe('passwd');
      expect(clean).not.toContain('..');
    });

    it('replaces dangerous symbols with underscores', () => {
      const clean = sanitizeFilename('my<resume>?*:"|.pdf');
      expect(clean).toBe('my_resume______.pdf');
    });
  });

  describe('getStoragePath', () => {
    it('creates path under ATTACHMENTS_STORAGE_DIR', () => {
      const storagePath = getStoragePath('email-123', 'att-456', 'cv.pdf');
      expect(storagePath.startsWith(ATTACHMENTS_STORAGE_DIR)).toBe(true);
      expect(storagePath).toContain('email-123');
      expect(storagePath).toContain('att-456_cv.pdf');
    });
  });
});
