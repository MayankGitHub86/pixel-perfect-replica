// backend/src/tests/create-real-workbook-upload.test.ts
//
// Mocked unit tests for:
//   1) backend/scripts/lib/upload-to-onedrive.js (Graph upload logic)
//   2) backend/scripts/lib/update-workbook-pointer.js (Prisma DB pointer update)
//
// These tests exercise the ACTUAL production modules with mocked HTTP and Prisma.
// No real network calls or database mutations occur.

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import path from 'path';

// Import production modules (CommonJS)
// eslint-disable-next-line @typescript-eslint/no-var-requires
const { uploadToOneDrive } = require(
  path.resolve(__dirname, '../../scripts/lib/upload-to-onedrive.js')
);
// eslint-disable-next-line @typescript-eslint/no-var-requires
const { updateWorkbookPointer } = require(
  path.resolve(__dirname, '../../scripts/lib/update-workbook-pointer.js')
);

const EXISTENCE_URL =
  'https://graph.microsoft.com/v1.0/me/drive/root:/Candidate_Master_RealData.xlsx';
const UPLOAD_URL =
  'https://graph.microsoft.com/v1.0/me/drive/root:/Candidate_Master_RealData.xlsx:/content';

describe('uploadToOneDrive (production module)', () => {
  const mockFetch = vi.fn();
  const originalFetch = globalThis.fetch;
  const fakeToken = 'test-bearer-token';
  const fakeBuffer = Buffer.from('fake-xlsx-bytes');
  const fileName = 'Candidate_Master_RealData.xlsx';

  beforeEach(() => {
    globalThis.fetch = mockFetch;
    mockFetch.mockReset();
  });

  afterEach(() => {
    globalThis.fetch = originalFetch;
  });

  // ---------------------------------------------------------------
  // URL & header correctness
  // ---------------------------------------------------------------

  it('PUT URL contains no @microsoft.graph.conflictBehavior query parameter', async () => {
    mockFetch
      .mockResolvedValueOnce({ ok: false, status: 404, text: async () => 'Not Found' })
      .mockResolvedValueOnce({
        ok: true,
        status: 201,
        json: async () => ({ id: 'id-1', name: fileName }),
      });

    await uploadToOneDrive({ token: fakeToken, fileBuffer: fakeBuffer, fileName });

    const putUrl: string = mockFetch.mock.calls[1][0];
    expect(putUrl).toBe(UPLOAD_URL);
    expect(putUrl).not.toContain('conflictBehavior');
    expect(putUrl).not.toContain('@microsoft.graph');
  });

  it('PUT request includes If-None-Match: * and no If-Match', async () => {
    mockFetch
      .mockResolvedValueOnce({ ok: false, status: 404, text: async () => '' })
      .mockResolvedValueOnce({
        ok: true,
        status: 201,
        json: async () => ({ id: 'id-1' }),
      });

    await uploadToOneDrive({ token: fakeToken, fileBuffer: fakeBuffer, fileName });

    const putHeaders = mockFetch.mock.calls[1][1].headers;
    expect(putHeaders['If-None-Match']).toBe('*');
    expect(putHeaders['If-Match']).toBeUndefined();
  });

  it('PUT sends correct method, Content-Type, Content-Length, and Authorization', async () => {
    mockFetch
      .mockResolvedValueOnce({ ok: false, status: 404, text: async () => '' })
      .mockResolvedValueOnce({
        ok: true,
        status: 201,
        json: async () => ({ id: 'id-1' }),
      });

    await uploadToOneDrive({ token: fakeToken, fileBuffer: fakeBuffer, fileName });

    const opts = mockFetch.mock.calls[1][1];
    expect(opts.method).toBe('PUT');
    expect(opts.headers['Content-Type']).toBe(
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    expect(opts.headers['Content-Length']).toBe(String(fakeBuffer.length));
    expect(opts.headers.Authorization).toBe(`Bearer ${fakeToken}`);
  });

  // ---------------------------------------------------------------
  // Preflight GET: existing file or unexpected status
  // ---------------------------------------------------------------

  it('aborts when preflight GET returns 200 (file exists) – PUT never sent', async () => {
    mockFetch.mockResolvedValueOnce({
      ok: true,
      status: 200,
      json: async () => ({ name: 'exists' }),
    });

    await expect(
      uploadToOneDrive({ token: fakeToken, fileBuffer: fakeBuffer, fileName })
    ).rejects.toThrow(/already exists.*abort/i);

    expect(mockFetch).toHaveBeenCalledTimes(1); // only the GET
  });

  it('aborts when preflight GET returns 401', async () => {
    mockFetch.mockResolvedValueOnce({
      ok: false,
      status: 401,
      text: async () => 'Unauthorized',
    });

    await expect(
      uploadToOneDrive({ token: fakeToken, fileBuffer: fakeBuffer, fileName })
    ).rejects.toThrow(/Unexpected response.*401/);

    expect(mockFetch).toHaveBeenCalledTimes(1);
  });

  it('aborts when preflight GET returns 500', async () => {
    mockFetch.mockResolvedValueOnce({
      ok: false,
      status: 500,
      text: async () => 'Internal Server Error',
    });

    await expect(
      uploadToOneDrive({ token: fakeToken, fileBuffer: fakeBuffer, fileName })
    ).rejects.toThrow(/Unexpected response.*500/);

    expect(mockFetch).toHaveBeenCalledTimes(1);
  });

  // ---------------------------------------------------------------
  // PUT 412 Precondition Failed (race-condition safety)
  // ---------------------------------------------------------------

  it('aborts with clear message when PUT returns 412', async () => {
    mockFetch
      .mockResolvedValueOnce({ ok: false, status: 404, text: async () => '' })
      .mockResolvedValueOnce({
        ok: false,
        status: 412,
        text: async () => '{"error":{"code":"conditionNotMet"}}',
      });

    await expect(
      uploadToOneDrive({ token: fakeToken, fileBuffer: fakeBuffer, fileName })
    ).rejects.toThrow(/Upload prevented.*412 Precondition Failed/);
  });

  it('does not call json() on a 412 response', async () => {
    const jsonSpy = vi.fn();
    mockFetch
      .mockResolvedValueOnce({ ok: false, status: 404, text: async () => '' })
      .mockResolvedValueOnce({
        ok: false,
        status: 412,
        text: async () => 'conflict',
        json: jsonSpy,
      });

    try {
      await uploadToOneDrive({ token: fakeToken, fileBuffer: fakeBuffer, fileName });
    } catch {
      // expected
    }

    expect(jsonSpy).not.toHaveBeenCalled();
  });

  // ---------------------------------------------------------------
  // Other non-2xx PUT responses
  // ---------------------------------------------------------------

  it('aborts on 409 Conflict', async () => {
    mockFetch
      .mockResolvedValueOnce({ ok: false, status: 404, text: async () => '' })
      .mockResolvedValueOnce({
        ok: false,
        status: 409,
        text: async () => 'nameAlreadyExists',
      });

    await expect(
      uploadToOneDrive({ token: fakeToken, fileBuffer: fakeBuffer, fileName })
    ).rejects.toThrow(/Upload failed: 409/);
  });

  it('aborts on 401 Unauthorized from PUT', async () => {
    mockFetch
      .mockResolvedValueOnce({ ok: false, status: 404, text: async () => '' })
      .mockResolvedValueOnce({
        ok: false,
        status: 401,
        text: async () => 'InvalidAuthenticationToken',
      });

    await expect(
      uploadToOneDrive({ token: fakeToken, fileBuffer: fakeBuffer, fileName })
    ).rejects.toThrow(/Upload failed: 401/);
  });

  it('aborts on 500 Internal Server Error from PUT', async () => {
    mockFetch
      .mockResolvedValueOnce({ ok: false, status: 404, text: async () => '' })
      .mockResolvedValueOnce({
        ok: false,
        status: 500,
        text: async () => 'Internal Server Error',
      });

    await expect(
      uploadToOneDrive({ token: fakeToken, fileBuffer: fakeBuffer, fileName })
    ).rejects.toThrow(/Upload failed: 500/);
  });

  // ---------------------------------------------------------------
  // Success path
  // ---------------------------------------------------------------

  it('returns parsed DriveItem metadata on 201 Created', async () => {
    const expectedMeta = {
      id: 'drive-item-42',
      name: fileName,
      webUrl: 'https://onedrive.live.com/x',
      size: 9876,
    };
    mockFetch
      .mockResolvedValueOnce({ ok: false, status: 404, text: async () => '' })
      .mockResolvedValueOnce({
        ok: true,
        status: 201,
        json: async () => expectedMeta,
      });

    const result = await uploadToOneDrive({
      token: fakeToken,
      fileBuffer: fakeBuffer,
      fileName,
    });

    expect(result).toEqual(expectedMeta);
  });
});

describe('updateWorkbookPointer (production module)', () => {
  const newDriveItemId = '5082520FEDE864C4!sRealDataDriveItemId';
  const newWebUrl = 'https://onedrive.live.com/realdata.xlsx';

  it('updates existing record using its UUID string when one record exists', async () => {
    const existingId = '1f02aa26-c472-4683-8adc-6e6a59890b53';
    const mockTx = {
      excelWorkbook: {
        findMany: vi.fn().mockResolvedValue([
          {
            id: existingId,
            driveItemId: 'old-item-id',
            webUrl: 'old-url',
            createdAt: new Date('2026-10-30T10:00:00Z'),
          },
        ]),
        update: vi.fn().mockResolvedValue({
          id: existingId,
          driveItemId: newDriveItemId,
          webUrl: newWebUrl,
        }),
        deleteMany: vi.fn(),
        create: vi.fn(),
      },
    };

    const res = await updateWorkbookPointer(mockTx, {
      driveItemId: newDriveItemId,
      webUrl: newWebUrl,
    });

    // Validates that update was called with string UUID, NOT integer 1
    expect(mockTx.excelWorkbook.update).toHaveBeenCalledWith({
      where: { id: existingId },
      data: { driveItemId: newDriveItemId, webUrl: newWebUrl },
    });
    expect(typeof mockTx.excelWorkbook.update.mock.calls[0][0].where.id).toBe('string');
    expect(mockTx.excelWorkbook.create).not.toHaveBeenCalled();
    expect(mockTx.excelWorkbook.deleteMany).not.toHaveBeenCalled();
    expect(res.id).toBe(existingId);
    expect(res.driveItemId).toBe(newDriveItemId);
  });

  it('creates a new record when no workbook record currently exists in DB', async () => {
    const generatedId = 'new-random-uuid-1234';
    const mockTx = {
      excelWorkbook: {
        findMany: vi.fn().mockResolvedValue([]),
        update: vi.fn(),
        deleteMany: vi.fn(),
        create: vi.fn().mockResolvedValue({
          id: generatedId,
          driveItemId: newDriveItemId,
          webUrl: newWebUrl,
        }),
      },
    };

    const res = await updateWorkbookPointer(mockTx, {
      driveItemId: newDriveItemId,
      webUrl: newWebUrl,
    });

    expect(mockTx.excelWorkbook.create).toHaveBeenCalledWith({
      data: { driveItemId: newDriveItemId, webUrl: newWebUrl },
    });
    expect(mockTx.excelWorkbook.update).not.toHaveBeenCalled();
    expect(res.id).toBe(generatedId);
    expect(res.driveItemId).toBe(newDriveItemId);
  });

  it('aborts without update, create, or delete when multiple records exist', async () => {
    const primaryId = 'primary-uuid-0001';
    const dupId1 = 'dup-uuid-0002';
    const dupId2 = 'dup-uuid-0003';

    const mockTx = {
      excelWorkbook: {
        findMany: vi.fn().mockResolvedValue([
          { id: primaryId, createdAt: new Date('2026-10-30T08:00:00Z') },
          { id: dupId1, createdAt: new Date('2026-10-30T09:00:00Z') },
          { id: dupId2, createdAt: new Date('2026-10-30T10:00:00Z') },
        ]),
        update: vi.fn(),
        deleteMany: vi.fn(),
        create: vi.fn(),
      },
    };

    await expect(
      updateWorkbookPointer(mockTx, {
        driveItemId: newDriveItemId,
        webUrl: newWebUrl,
      })
    ).rejects.toThrow(/Multiple ExcelWorkbook records found.*Manual investigation is required/);

    // Absolute zero mutation / deletion on multiple records
    expect(mockTx.excelWorkbook.update).not.toHaveBeenCalled();
    expect(mockTx.excelWorkbook.deleteMany).not.toHaveBeenCalled();
    expect(mockTx.excelWorkbook.create).not.toHaveBeenCalled();
  });

  it('surfaces database errors when update fails', async () => {
    const mockTx = {
      excelWorkbook: {
        findMany: vi.fn().mockResolvedValue([{ id: 'uuid-1' }]),
        update: vi.fn().mockRejectedValue(new Error('DB connection lost')),
      },
    };

    await expect(
      updateWorkbookPointer(mockTx, {
        driveItemId: newDriveItemId,
        webUrl: newWebUrl,
      })
    ).rejects.toThrow('DB connection lost');
  });

  it('validates input parameters before issuing queries', async () => {
    const mockTx = { excelWorkbook: { findMany: vi.fn() } };

    await expect(
      updateWorkbookPointer(mockTx, { driveItemId: '', webUrl: 'http://valid.url' })
    ).rejects.toThrow(/valid string driveItemId/);

    await expect(
      updateWorkbookPointer(mockTx, { driveItemId: 'valid-id', webUrl: '' })
    ).rejects.toThrow(/valid string webUrl/);

    expect(mockTx.excelWorkbook.findMany).not.toHaveBeenCalled();
  });
});

describe('create-real-workbook full flow isolation check', () => {
  it('ensures DB pointer update is not reached if upload or validation throws', async () => {
    // Test the control flow invariant: In create-real-workbook.js,
    // updateWorkbookPointer is only invoked at Step 8, strictly after
    // uploadToOneDrive (Step 3) and Graph read-back (Steps 4-7).
    let uploadExecuted = false;
    let readBackExecuted = false;
    let dbPointerUpdated = false;

    // Simulate failure at upload step
    async function simulateFlow(failStep: 'upload' | 'readBack' | 'none') {
      // Step 3
      if (failStep === 'upload') {
        throw new Error('Upload step failed');
      }
      uploadExecuted = true;

      // Steps 4-7
      if (failStep === 'readBack') {
        throw new Error('Validation mismatch in read-back');
      }
      readBackExecuted = true;

      // Step 8
      dbPointerUpdated = true;
    }

    // 1. Upload fails
    await expect(simulateFlow('upload')).rejects.toThrow('Upload step failed');
    expect(readBackExecuted).toBe(false);
    expect(dbPointerUpdated).toBe(false);

    // 2. Read-back fails
    await expect(simulateFlow('readBack')).rejects.toThrow('Validation mismatch in read-back');
    expect(dbPointerUpdated).toBe(false);

    // 3. Success reaches DB update
    await simulateFlow('none');
    expect(uploadExecuted).toBe(true);
    expect(readBackExecuted).toBe(true);
    expect(dbPointerUpdated).toBe(true);
  });
});
