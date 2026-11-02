// backend/src/__tests__/excelEligibility.test.ts
//
// Tests the production candidate eligibility guard and Excel sync pipeline guards.
// Uses Vitest with fully mocked Prisma and Graph fetch calls.
// Zero remote operations, zero live database mutations.

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { checkCandidateExcelEligibility, upsertCandidateRow, CANDIDATE_TABLE_COLUMNS } from '../lib/excel';
import { prisma } from '../lib/db';

vi.mock('../lib/db', () => ({
  prisma: {
    candidate: {
      findUnique: vi.fn(),
      findMany: vi.fn(),
    },
    emailMessage: {
      findFirst: vi.fn(),
    },
    excelWorkbook: {
      findFirst: vi.fn(),
    },
    mailboxConnection: {
      findFirst: vi.fn(),
    },
    excelSyncRun: {
      create: vi.fn(),
    },
  },
}));

vi.mock('../lib/crypto', () => ({
  decrypt: vi.fn(() => 'decrypted-test-token'),
}));

vi.mock('../lib/tokenRefresh', () => ({
  refreshAccessToken: vi.fn(),
}));

describe('checkCandidateExcelEligibility', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('marks verified genuine candidate with real Outlook provenance as eligible', async () => {
    (prisma.candidate.findUnique as any).mockResolvedValue({
      id: '106945ea-6ae3-4c64-b5bb-9c1148d72985',
      name: 'CHIRAYU SHARMA',
      email: 'chirayusharma202004@gmail.com',
      sourceEmailId: 'ca1341f7-7987-4cea-a3db-3ce2347425b8',
    });

    (prisma.emailMessage.findFirst as any).mockResolvedValue({
      id: 'ca1341f7-7987-4cea-a3db-3ce2347425b8',
      messageId: 'AQMkADAwATM0MDAAMS03YmNlLWEyNTktMDACLTAwCgBGAAADayCNNu6OLEqxm8U-3ZindwcAr14ZxUAXkUuUTG80VuIFHQAAAgEMAAAAr14ZxUAXkUuUTG80VuIFHQACDRuvEQAAAA==',
      fromEmail: 'singhamagain123@gmail.com',
      subject: 'Resume Submission',
    });

    const result = await checkCandidateExcelEligibility('106945ea-6ae3-4c64-b5bb-9c1148d72985');
    expect(result.eligible).toBe(true);
    expect(result.reason).toContain('Verified genuine candidate');
  });

  it('rejects candidate with [DEV-SYNTHETIC] in name', async () => {
    (prisma.candidate.findUnique as any).mockResolvedValue({
      id: '562c936d-fa4d-48f3-8dc7-25fe7c4ad925',
      name: '[DEV-SYNTHETIC] Sarah Jenkins',
      email: 'sarah.jenkins@synthetic.dev',
      sourceEmailId: 'SYNTH-MSG-001',
    });

    const result = await checkCandidateExcelEligibility('562c936d-fa4d-48f3-8dc7-25fe7c4ad925');
    expect(result.eligible).toBe(false);
    expect(result.reason).toContain('synthetic test name prefix');
  });

  it('rejects candidate with @synthetic.dev email domain even if name lacks prefix', async () => {
    (prisma.candidate.findUnique as any).mockResolvedValue({
      id: 'synth-custom-id',
      name: 'John Doe',
      email: 'john.doe@synthetic.dev',
      sourceEmailId: 'some-msg-id',
    });

    const result = await checkCandidateExcelEligibility('synth-custom-id');
    expect(result.eligible).toBe(false);
    expect(result.reason).toContain('synthetic test domain');
  });

  it('rejects candidate with SYNTH- sourceEmailId', async () => {
    (prisma.candidate.findUnique as any).mockResolvedValue({
      id: 'synth-source-id',
      name: 'Alex Mercer',
      email: 'alex@example.com',
      sourceEmailId: 'SYNTH-MSG-999',
    });

    const result = await checkCandidateExcelEligibility('synth-source-id');
    expect(result.eligible).toBe(false);
    expect(result.reason).toContain('synthetic sourceEmailId (SYNTH-*)');
  });

  it('rejects candidate with missing sourceEmailId provenance', async () => {
    (prisma.candidate.findUnique as any).mockResolvedValue({
      id: 'no-source-id',
      name: 'Direct Input Candidate',
      email: 'direct@example.com',
      sourceEmailId: null,
    });

    const result = await checkCandidateExcelEligibility('no-source-id');
    expect(result.eligible).toBe(false);
    expect(result.reason).toContain('lacks sourceEmailId provenance link');
  });

  it('rejects candidate whose sourceEmailId does not resolve to an EmailMessage in PostgreSQL', async () => {
    (prisma.candidate.findUnique as any).mockResolvedValue({
      id: 'dangling-source-id',
      name: 'Orphaned Candidate',
      email: 'orphan@example.com',
      sourceEmailId: 'non-existent-email-uuid',
    });

    (prisma.emailMessage.findFirst as any).mockResolvedValue(null);

    const result = await checkCandidateExcelEligibility('dangling-source-id');
    expect(result.eligible).toBe(false);
    expect(result.reason).toContain('does not resolve to an EmailMessage record');
  });

  it('rejects candidate if linked EmailMessage is synthetic test data', async () => {
    (prisma.candidate.findUnique as any).mockResolvedValue({
      id: 'sneaky-candidate-id',
      name: 'Sneaky Name',
      email: 'sneaky@gmail.com',
      sourceEmailId: 'valid-looking-uuid',
    });

    (prisma.emailMessage.findFirst as any).mockResolvedValue({
      id: 'valid-looking-uuid',
      messageId: 'SYNTH-MSG-005',
      fromEmail: 'priya.sharma@synthetic.dev',
      subject: 'Application for Staff Backend Architect - [DEV-SYNTHETIC] Priya Sharma',
    });

    const result = await checkCandidateExcelEligibility('sneaky-candidate-id');
    expect(result.eligible).toBe(false);
    expect(result.reason).toContain('Linked EmailMessage record is marked as synthetic test data');
  });
});

describe('upsertCandidateRow guard enforcement', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('skips Graph write completely when candidate is not eligible', async () => {
    const fetchSpy = vi.fn();
    const originalFetch = globalThis.fetch;
    globalThis.fetch = fetchSpy as any;

    try {
      (prisma.candidate.findUnique as any).mockResolvedValue({
        id: '562c936d-fa4d-48f3-8dc7-25fe7c4ad925',
        name: '[DEV-SYNTHETIC] Sarah Jenkins',
        email: 'sarah.jenkins@synthetic.dev',
        sourceEmailId: 'SYNTH-MSG-001',
      });

      const synced = await upsertCandidateRow('562c936d-fa4d-48f3-8dc7-25fe7c4ad925');
      expect(synced).toBe(false);
      // Zero calls to Graph API fetch
      expect(fetchSpy).not.toHaveBeenCalled();
      // Zero calls to excelSyncRun.create
      expect(prisma.excelSyncRun.create).not.toHaveBeenCalled();
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('executes Graph sync and writes audit row when candidate is eligible', async () => {
    const mockFetch = vi.fn();
    const originalFetch = globalThis.fetch;
    globalThis.fetch = mockFetch as any;

    try {
      // 1. Eligibility check passes
      (prisma.candidate.findUnique as any).mockResolvedValue({
        id: '106945ea-6ae3-4c64-b5bb-9c1148d72985',
        name: 'CHIRAYU SHARMA',
        email: 'chirayusharma202004@gmail.com',
        sourceEmailId: 'ca1341f7-7987-4cea-a3db-3ce2347425b8',
        phone: '+919876543210',
        appliedRole: 'Business Analyst',
        degree: 'B.Tech',
        college: 'RTU',
        graduationYear: 2024,
        experienceYears: 1.5,
        skills: ['SQL', 'Excel', 'Python'],
        receivedAt: new Date(),
        updatedAt: new Date(),
      });

      (prisma.emailMessage.findFirst as any).mockResolvedValue({
        id: 'ca1341f7-7987-4cea-a3db-3ce2347425b8',
        messageId: 'real-outlook-msg-id',
        fromEmail: 'singhamagain123@gmail.com',
        subject: 'Resume Submission',
      });

      // 2. Mock token & workbook
      (prisma.mailboxConnection.findFirst as any).mockResolvedValue({
        id: 'mailbox-1',
        accessToken: 'enc-token',
        tokenExpiresAt: new Date(Date.now() + 3600000),
      });

      (prisma.excelWorkbook.findFirst as any).mockResolvedValue({
        id: 'wb-1',
        driveItemId: 'item-123',
        webUrl: 'https://onedrive.live.com/test',
      });

      // 3. Mock Graph responses:
      // a) Worksheet check
      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({ value: [{ name: 'Candidates' }] }),
      });
      // b) Table check
      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({ value: [{ name: 'CandidateTable' }] }),
      });
      // c) Columns check
      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({ value: CANDIDATE_TABLE_COLUMNS.map(name => ({ name })) }),
      });
      // d) Rows check (empty table -> append)
      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({ value: [] }),
      });
      // e) Rows add
      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({ index: 0, values: [] }),
      });

      const synced = await upsertCandidateRow('106945ea-6ae3-4c64-b5bb-9c1148d72985');
      expect(synced).toBe(true);
      expect(mockFetch).toHaveBeenCalled();
      expect(prisma.excelSyncRun.create).toHaveBeenCalledWith({
        data: {
          candidateId: '106945ea-6ae3-4c64-b5bb-9c1148d72985',
          status: 'success',
          result: 'success',
          rows: 1,
          message: 'inserted',
        },
      });
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

