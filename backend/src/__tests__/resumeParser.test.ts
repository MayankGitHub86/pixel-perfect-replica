// backend/src/__tests__/resumeParser.test.ts
import { describe, it, expect } from 'vitest';
import { parseResumeBuffer, parsePdfBuffer, parseDocxBuffer } from '../lib/resumeParser';

describe('Resume Parser', () => {
  it('returns format=unknown for unsupported file extensions', async () => {
    const res = await parseResumeBuffer(Buffer.from('hello'), 'resume.txt');
    expect(res.format).toBe('unknown');
    expect(res.error).toMatch(/Unsupported resume file extension/);
  });

  it('handles corrupted PDF buffer gracefully without crashing', async () => {
    const corruptBuffer = Buffer.from('corrupted non-pdf data');
    const res = await parsePdfBuffer(corruptBuffer);
    expect(res.format).toBe('pdf');
    expect(res.error).toBeDefined();
    expect(res.text).toBe('');
  });

  it('handles corrupted DOCX buffer gracefully without crashing', async () => {
    const corruptBuffer = Buffer.from('corrupted non-docx data');
    const res = await parseDocxBuffer(corruptBuffer);
    expect(res.format).toBe('docx');
    expect(res.error).toBeDefined();
    expect(res.text).toBe('');
  });
});

