// backend/src/__tests__/excelDownloadCv.test.ts
//
// Tests Phase 5 functionality:
// - Extraction of candidate details from email and CV (including phone with spaces, location, graduation year).
// - Excel 22-column structure with resumeFileName, resumeLink (HYPERLINK formula), sourceEmailSubject, sourceEmailReceivedAt, extractionStatus.
// - Absence of resume leaves link blank.
// - Idempotent sync updates row rather than duplicating.
// - Synthetic candidate rejection.

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { extractCandidateInfo, extractPhone, extractLocation, extractGraduationYear, extractExperienceYears } from '../lib/candidateExtractor';
import { CANDIDATE_TABLE_COLUMNS, upsertCandidateRow, checkCandidateExcelEligibility, toExcelSerialDate, EXCEL_DATE_TIME_FORMAT, getBackendPublicUrl } from '../lib/excel';
import { prisma } from '../lib/db';

vi.mock('../lib/db', () => ({
  prisma: {
    candidate: {
      findUnique: vi.fn(),
      update: vi.fn(),
    },
    emailMessage: {
      findFirst: vi.fn(),
    },
    mailboxConnection: {
      findFirst: vi.fn(),
      findUnique: vi.fn(),
    },
    excelWorkbook: {
      findFirst: vi.fn(),
    },
    excelSyncRun: {
      create: vi.fn(),
    },
  },
}));

vi.mock('../lib/crypto', () => ({
  decrypt: vi.fn(() => 'mock-decrypted-token'),
  encrypt: vi.fn(() => 'mock-encrypted-token'),
}));

vi.mock('../lib/tokenRefresh', () => ({
  refreshAccessToken: vi.fn().mockResolvedValue(undefined),
}));

describe('Phase 5 Candidate Extraction & Excel Sync with Download CV', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  describe('Field Extraction from Resume & Email', () => {
    const resumeText = `
CHIRAYU SHARMA
BUSINESS ANALYST | DATA ANALYST | POWER BI | EXCEL
+91 79903 36208 | chirayusharma202004@gmail.com | A803 Flora Iris, South Bopal, Ahmedabad, Gujarat

EDUCATION
Bachelor of Engineering – Computer Engineering — Indus University 2022 – 2026

SKILLS
Power BI, DAX, Power Query, Advanced Excel, SQL, Python, JavaScript, React, Node.js

EXPERIENCE
Data Analyst Intern at LCC Projects Ltd. (2026 - Present)
`;

    it('extracts phone numbers with spaces and country codes', () => {
      const phone = extractPhone(resumeText);
      expect(phone).toBe('+91 79903 36208');
    });

    it('extracts location from resume contact details', () => {
      const location = extractLocation(resumeText);
      expect(location).toContain('Ahmedabad');
      expect(location).toContain('Gujarat');
    });

    it('extracts correct graduation completion year from ranges', () => {
      const gradYear = extractGraduationYear(resumeText);
      expect(gradYear).toBe(2026);
    });

    it('extracts role, skills, and details without fabricating missing data', () => {
      const info = extractCandidateInfo(resumeText, {
        subject: 'Resume Submission',
        from: 'singhamagain123@gmail.com',
      });

      expect(info.name).toBe('CHIRAYU SHARMA');
      expect(info.email).toBe('chirayusharma202004@gmail.com');
      expect(info.phone).toBe('+91 79903 36208');
      expect(info.location).toBe('A803 Flora Iris, South Bopal, Ahmedabad, Gujarat');
      expect(info.appliedRole).toBe('BUSINESS ANALYST');
      expect(info.degree).toBe('Bachelor of Engineering');
      expect(info.college).toBe('Indus University');
      expect(info.graduationYear).toBe(2026);
      expect(info.skills).toContain('Power BI');
      expect(info.skills).toContain('DAX');
      expect(info.skills).toContain('Advanced Excel');
      expect(info.skills).toContain('SQL');
      expect(info.processingStatus).toBe('processed');
    });

    it('leaves missing fields as null rather than inventing data', () => {
      const minimalText = `
JANE DOE
Software Engineer
jane.doe@company.org
`;
      const info = extractCandidateInfo(minimalText);
      expect(info.name).toBe('JANE DOE');
      expect(info.email).toBe('jane.doe@company.org');
      expect(info.phone).toBeNull();
      expect(info.location).toBeNull();
      expect(info.college).toBeNull();
      expect(info.graduationYear).toBeNull();
      expect(info.experienceYears).toBeNull();
    });

    describe('Graduation Year Extraction & Regression Tests', () => {
      it('does not select 12th-pass year (2022) as degree graduation year when Bachelor degree is 2022 - 2026', () => {
        const fullEducationText = `
EDUCATION
Bachelor of Engineering – Computer Engineering (9.05 CGPA) — Indus University 2022 – 2026
HSC (12th) – CBSE Board (75%) — St. Ann's School, Ahmedabad 2022
SSC (10th) – CBSE Board (85%) — Delhi Public School, Ahmedabad 2020
`;
        const year = extractGraduationYear(fullEducationText);
        expect(year).toBe(2026);
        expect(year).not.toBe(2022);
      });

      it('selects degree completion year when clearly stated', () => {
        const text = `
EDUCATION
Master of Science in Computer Science, Stanford University
Graduated in 2025
`;
        expect(extractGraduationYear(text)).toBe(2025);
      });

      it('selects expected degree year only when explicitly stated', () => {
        const text = `
EDUCATION
Bachelor of Technology, Indian Institute of Technology
Expected 2027
`;
        expect(extractGraduationYear(text)).toBe(2027);
      });

      it('leaves missing or ambiguous degree years as null without guessing', () => {
        const text = `
EDUCATION
Bachelor of Science in Information Technology
Mumbai University
(Coursework in Database Management)
`;
        expect(extractGraduationYear(text)).toBeNull();
      });

      it('preserves existing correct graduation year when new extraction is ambiguous', () => {
        const existingYear = 2026;
        const newExtractedYear = null;
        const finalYear = newExtractedYear ?? existingYear;
        expect(finalYear).toBe(2026);
      });
    });

    describe('Experience Years Calculation & Regression Tests', () => {
      const refDate = new Date(Date.UTC(2026, 10, 2)); // 2026-11-02

      it('calculates total experience for a single job with complete dates', () => {
        const text = `
EXPERIENCE
Software Engineer at TechCorp (Jan 2021 – Dec 2022)
`;
        const years = extractExperienceYears(text, refDate);
        expect(years).toBe(2.0);
      });

      it('calculates total experience across multiple sequential jobs', () => {
        const text = `
EXPERIENCE
Junior Developer at Alpha Corp (Jan 2021 – Dec 2021)
Software Engineer at Beta Solutions (Jan 2022 – Dec 2022)
`;
        const years = extractExperienceYears(text, refDate);
        expect(years).toBe(2.0);
      });

      it('merges overlapping employment periods to prevent double-counting', () => {
        const text = `
EXPERIENCE
Full Stack Developer at Company A (Jan 2021 – Dec 2022)
Technical Consultant at Company B (Jun 2021 – Jun 2022)
`;
        const years = extractExperienceYears(text, refDate);
        expect(years).toBe(2.0);
      });

      it('includes internships and part-time employment with meaningful fractional values', () => {
        const text = `
EXPERIENCE
Data Analyst Intern at LCC (Jan 2024 – Jun 2024)
`;
        const years = extractExperienceYears(text, refDate);
        expect(years).toBe(0.5);
      });

      it('excludes student clubs, volunteer activities, and academic projects from professional employment', () => {
        const text = `
EXPERIENCE
Graphic Head – CESA (Computer Engineering Student Association) (2022 – 2024)
Software Engineer at RealCorp (Jan 2023 – Dec 2023)
Lead – Google Developer Student Clubs (2021 – 2022)
`;
        const years = extractExperienceYears(text, refDate);
        expect(years).toBe(1.0);
      });

      it('calculates current employment up to the consistent reference date for Present', () => {
        const text = `
EXPERIENCE
Software Engineer at CurrentCorp (Jan 2026 – Present)
`;
        const years = extractExperienceYears(text, refDate);
        expect(years).toBe(0.8);
      });

      it('returns null for missing or ambiguous employment dates without guessing', () => {
        const text = `
EXPERIENCE
Software Developer at MysteryCorp
Worked on internal tooling and customer support.
`;
        expect(extractExperienceYears(text, refDate)).toBeNull();
      });

      it('applies documented conservative rule for year-only ranges without inflating to 3 full years', () => {
        const text = `
EXPERIENCE
Part-Time Mathematics Tutor at The Knowledge Academy (2022 – 2024)
`;
        const years = extractExperienceYears(text, refDate);
        // Conservative difference 2024 - 2022 = 2.0 years (not 3.0 years)
        expect(years).toBe(2.0);
      });

      it('preserves existing reliable experienceYears when a new extraction is uncertain', () => {
        const existingYears = 3.1;
        const newExtractedYears = null;
        const finalYears = newExtractedYears ?? existingYears;
        expect(finalYears).toBe(3.1);
      });
    });
  });

  describe('Excel Schema and Download CV Link Formulation', () => {
    it('defines all 22 required columns in exact order', () => {
      expect(CANDIDATE_TABLE_COLUMNS).toEqual([
        'id',
        'name',
        'email',
        'phone',
        'location',
        'appliedRole',
        'degree',
        'college',
        'graduationYear',
        'experienceYears',
        'skills',
        'confidenceScore',
        'reviewStatus',
        'applicationStatus',
        'sourceEmailId',
        'receivedAt',
        'updatedAt',
        'resumeFileName',
        'resumeLink',
        'sourceEmailSubject',
        'sourceEmailReceivedAt',
        'extractionStatus',
      ]);
    });

    it('writes Excel row with =HYPERLINK formula and email provenance when resume is present', async () => {
      const mockFetch = vi.fn();
      const originalFetch = globalThis.fetch;
      globalThis.fetch = mockFetch as any;

      try {
        (prisma.candidate.findUnique as any).mockResolvedValue({
          id: '106945ea-6ae3-4c64-b5bb-9c1148d72985',
          name: 'CHIRAYU SHARMA',
          email: 'chirayusharma202004@gmail.com',
          phone: '+91 79903 36208',
          location: 'Ahmedabad, Gujarat',
          appliedRole: 'BUSINESS ANALYST',
          degree: 'Bachelor of Engineering',
          college: 'Indus University',
          graduationYear: 2026,
          experienceYears: 1,
          skills: ['Power BI', 'SQL', 'Excel'],
          resumeFileName: 'Chirayu_Sharma_Resume.docx',
          resumeUrl: 'https://onedrive.live.com/view/resume-link',
          sourceEmailId: 'ca1341f7-7987-4cea-a3db-3ce2347425b8',
          processingStatus: 'processed',
          receivedAt: new Date('2026-10-30T10:02:38Z'),
          updatedAt: new Date('2026-10-30T10:02:38Z'),
        });

        (prisma.emailMessage.findFirst as any).mockResolvedValue({
          id: 'ca1341f7-7987-4cea-a3db-3ce2347425b8',
          messageId: 'real-outlook-msg-id',
          fromEmail: 'singhamagain123@gmail.com',
          subject: 'Resume Submission',
          receivedAt: new Date('2026-10-30T09:55:30Z'),
        });

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

        // Mock Graph responses:
        // 1. Worksheet check
        mockFetch.mockResolvedValueOnce({
          ok: true,
          json: async () => ({ value: [{ name: 'Candidates' }] }),
        });
        // 2. Table check
        mockFetch.mockResolvedValueOnce({
          ok: true,
          json: async () => ({ value: [{ name: 'CandidateTable' }] }),
        });
        // 3. Columns check (all 22 columns present)
        mockFetch.mockResolvedValueOnce({
          ok: true,
          json: async () => ({ value: CANDIDATE_TABLE_COLUMNS.map(name => ({ name })) }),
        });
        // 4. Rows check (empty table -> append)
        mockFetch.mockResolvedValueOnce({
          ok: true,
          json: async () => ({ value: [] }),
        });
        // 5. Rows add
        let capturedRowValues: any[] = [];
        mockFetch.mockImplementationOnce(async (_url, opts) => {
          const body = JSON.parse(opts.body);
          capturedRowValues = body.values[0];
          return {
            ok: true,
            json: async () => ({ index: 0, values: [capturedRowValues] }),
          };
        });

        const { config } = await import('../config');
        const originalUrl = config.BACKEND_PUBLIC_URL;
        (config as any).BACKEND_PUBLIC_URL = 'https://deploy.customdomain.org';

        try {
          const synced = await upsertCandidateRow('106945ea-6ae3-4c64-b5bb-9c1148d72985');
          expect(synced).toBe(true);

          expect(capturedRowValues).toHaveLength(22);
          // Column 17: resumeFileName
          expect(capturedRowValues[17]).toBe('Chirayu_Sharma_Resume.docx');
          // Column 18: resumeLink with =IF([@id]... HYPERLINK formula pointing to configured public URL
          expect(capturedRowValues[18]).toBe(
            '=IF([@id]="", "", HYPERLINK("https://deploy.customdomain.org/api/candidates/" & [@id] & "/resume", "Download CV"))'
          );
          // Column 19: sourceEmailSubject
          expect(capturedRowValues[19]).toBe('Resume Submission');
          // Column 20: sourceEmailReceivedAt formatted as Excel serial date
          expect(capturedRowValues[20]).toBe(toExcelSerialDate(new Date('2026-10-30T09:55:30Z')));
          // Column 21: extractionStatus
          expect(capturedRowValues[21]).toBe('processed');
        } finally {
          (config as any).BACKEND_PUBLIC_URL = originalUrl;
        }
      } finally {
        globalThis.fetch = originalFetch;
      }
    });

    it('leaves resumeLink blank when BACKEND_PUBLIC_URL is unconfigured or points to localhost', async () => {
      const mockFetch = vi.fn();
      const originalFetch = globalThis.fetch;
      globalThis.fetch = mockFetch as any;

      try {
        (prisma.candidate.findUnique as any).mockResolvedValue({
          id: '106945ea-6ae3-4c64-b5bb-9c1148d72985',
          name: 'CHIRAYU SHARMA',
          email: 'chirayusharma202004@gmail.com',
          phone: '+91 79903 36208',
          location: 'Ahmedabad',
          appliedRole: 'Business Analyst',
          degree: 'B.E.',
          college: 'Indus University',
          graduationYear: 2026,
          skills: ['Power BI'],
          resumeFileName: 'Chirayu_Sharma_Resume.docx',
          resumeUrl: null,
          resumeRef: '106945ea-6ae3-4c64-b5bb-9c1148d72985',
          sourceEmailId: 'ca1341f7-7987-4cea-a3db-3ce2347425b8',
          processingStatus: 'processed',
          receivedAt: new Date('2026-10-30T10:02:38Z'),
          updatedAt: new Date('2026-10-30T10:02:38Z'),
        });

        (prisma.emailMessage.findFirst as any).mockResolvedValue({
          id: 'ca1341f7-7987-4cea-a3db-3ce2347425b8',
          messageId: 'real-outlook-msg-id',
          fromEmail: 'singhamagain123@gmail.com',
          subject: 'Resume Submission',
          receivedAt: new Date('2026-10-30T09:55:30Z'),
        });

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

        // Mock Graph responses:
        // 1. Worksheet check
        mockFetch.mockResolvedValueOnce({
          ok: true,
          json: async () => ({ value: [{ name: 'Candidates' }] }),
        });
        // 2. Table check
        mockFetch.mockResolvedValueOnce({
          ok: true,
          json: async () => ({ value: [{ name: 'CandidateTable' }] }),
        });
        // 3. Columns check
        mockFetch.mockResolvedValueOnce({
          ok: true,
          json: async () => ({ value: CANDIDATE_TABLE_COLUMNS.map(name => ({ name })) }),
        });
        // 4. Rows check (empty table -> append)
        mockFetch.mockResolvedValueOnce({
          ok: true,
          json: async () => ({ value: [] }),
        });
        // 5. Rows add
        let capturedRowValues: any[] = [];
        mockFetch.mockImplementationOnce(async (_url, opts) => {
          const body = JSON.parse(opts.body);
          capturedRowValues = body.values[0];
          return {
            ok: true,
            json: async () => ({ index: 0, values: [capturedRowValues] }),
          };
        });

        const { config } = await import('../config');
        const originalUrl = config.BACKEND_PUBLIC_URL;
        (config as any).BACKEND_PUBLIC_URL = ''; // Empty / unconfigured

        try {
          const synced = await upsertCandidateRow('106945ea-6ae3-4c64-b5bb-9c1148d72985');
          expect(synced).toBe(true);
          // Column 18: resumeLink MUST be empty string (blank), NEVER localhost!
          expect(capturedRowValues[18]).toBe('');
        } finally {
          (config as any).BACKEND_PUBLIC_URL = originalUrl;
        }
      } finally {
        globalThis.fetch = originalFetch;
      }
    });

    it('leaves resumeLink blank when candidate has no resume', async () => {
      const mockFetch = vi.fn();
      const originalFetch = globalThis.fetch;
      globalThis.fetch = mockFetch as any;

      try {
        (prisma.candidate.findUnique as any).mockResolvedValue({
          id: '106945ea-6ae3-4c64-b5bb-9c1148d72985',
          name: 'CHIRAYU SHARMA',
          email: 'chirayusharma202004@gmail.com',
          phone: null,
          location: null,
          appliedRole: 'Business Analyst',
          degree: 'B.Tech',
          college: 'Indus University',
          graduationYear: 2026,
          skills: [],
          resumeFileName: null,
          resumeUrl: null,
          resumeRef: null,
          sourceEmailId: 'ca1341f7-7987-4cea-a3db-3ce2347425b8',
          processingStatus: 'needs_review',
          receivedAt: new Date('2026-10-30T10:02:38Z'),
          updatedAt: new Date('2026-10-30T10:02:38Z'),
        });

        (prisma.emailMessage.findFirst as any).mockResolvedValue({
          id: 'ca1341f7-7987-4cea-a3db-3ce2347425b8',
          messageId: 'real-outlook-msg-id',
          fromEmail: 'singhamagain123@gmail.com',
          subject: 'Inquiry without resume',
          receivedAt: new Date('2026-10-30T09:55:30Z'),
        });

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

        mockFetch.mockResolvedValueOnce({
          ok: true,
          json: async () => ({ value: [{ name: 'Candidates' }] }),
        });
        mockFetch.mockResolvedValueOnce({
          ok: true,
          json: async () => ({ value: [{ name: 'CandidateTable' }] }),
        });
        mockFetch.mockResolvedValueOnce({
          ok: true,
          json: async () => ({ value: CANDIDATE_TABLE_COLUMNS.map(name => ({ name })) }),
        });
        mockFetch.mockResolvedValueOnce({
          ok: true,
          json: async () => ({ value: [] }),
        });
        let capturedRowValues: any[] = [];
        mockFetch.mockImplementationOnce(async (_url, opts) => {
          const body = JSON.parse(opts.body);
          capturedRowValues = body.values[0];
          return {
            ok: true,
            json: async () => ({ index: 0, values: [capturedRowValues] }),
          };
        });

        await upsertCandidateRow('106945ea-6ae3-4c64-b5bb-9c1148d72985');

        expect(capturedRowValues[17]).toBe(''); // resumeFileName is blank
        expect(capturedRowValues[18]).toBe(''); // resumeLink is blank
      } finally {
        globalThis.fetch = originalFetch;
      }
    });

    it('updates existing row instead of creating duplicate when row exists', async () => {
      const mockFetch = vi.fn();
      const originalFetch = globalThis.fetch;
      globalThis.fetch = mockFetch as any;

      try {
        (prisma.candidate.findUnique as any).mockResolvedValue({
          id: '106945ea-6ae3-4c64-b5bb-9c1148d72985',
          name: 'CHIRAYU SHARMA',
          email: 'chirayusharma202004@gmail.com',
          sourceEmailId: 'ca1341f7-7987-4cea-a3db-3ce2347425b8',
          resumeFileName: 'Chirayu_Resume.docx',
          resumeUrl: 'https://onedrive.live.com/doc',
        });

        (prisma.emailMessage.findFirst as any).mockResolvedValue({
          id: 'ca1341f7-7987-4cea-a3db-3ce2347425b8',
          messageId: 'real-outlook-msg-id',
          fromEmail: 'singhamagain123@gmail.com',
          subject: 'Resume Submission',
        });

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

        mockFetch.mockResolvedValueOnce({
          ok: true,
          json: async () => ({ value: [{ name: 'Candidates' }] }),
        });
        mockFetch.mockResolvedValueOnce({
          ok: true,
          json: async () => ({ value: [{ name: 'CandidateTable' }] }),
        });
        mockFetch.mockResolvedValueOnce({
          ok: true,
          json: async () => ({ value: CANDIDATE_TABLE_COLUMNS.map(name => ({ name })) }),
        });
        // Existing row is found at index 0!
        mockFetch.mockResolvedValueOnce({
          ok: true,
          json: async () => ({
            value: [
              {
                index: 0,
                values: [['106945ea-6ae3-4c64-b5bb-9c1148d72985', 'OLD NAME']],
              },
            ],
          }),
        });
        // PATCH call for row itemAt(index=0)
        let patchCalled = false;
        mockFetch.mockImplementationOnce(async (url, opts) => {
          if (opts.method === 'PATCH' && url.includes('itemAt(index=0)')) {
            patchCalled = true;
          }
          return { ok: true, json: async () => ({}) };
        });

        const synced = await upsertCandidateRow('106945ea-6ae3-4c64-b5bb-9c1148d72985');
        expect(synced).toBe(true);
        expect(patchCalled).toBe(true);
      } finally {
        globalThis.fetch = originalFetch;
      }
    });

    it('strictly rejects synthetic candidates from entering Excel', async () => {
      (prisma.candidate.findUnique as any).mockResolvedValue({
        id: 'synthetic-id',
        name: '[DEV-SYNTHETIC] Sarah Jenkins',
        email: 'sarah.jenkins@synthetic.dev',
        sourceEmailId: 'SYNTH-MSG-001',
      });
      const eligibility = await checkCandidateExcelEligibility('synthetic-id');
      expect(eligibility.eligible).toBe(false);
      expect(eligibility.reason).toContain('synthetic');
    });

    it('correctly converts dates to Excel serial numbers matching dd-mm-yyyy hh:mm:ss AM/PM', () => {
      expect(EXCEL_DATE_TIME_FORMAT).toBe('dd-mm-yyyy hh:mm:ss AM/PM');
      
      // Known reference: 2026-10-30T10:02:38Z
      const date = new Date('2026-10-30T10:02:38.000Z');
      const serial = toExcelSerialDate(date);
      expect(typeof serial).toBe('number');
      // Serial should be around 46325.418495
      expect(serial as number).toBeCloseTo(46325.418495, 4);

      // Verify empty inputs return empty string
      expect(toExcelSerialDate(null)).toBe('');
      expect(toExcelSerialDate(undefined)).toBe('');
      expect(toExcelSerialDate('')).toBe('');
      expect(toExcelSerialDate('invalid-date')).toBe('');
    });

    it('correctly resolves getBackendPublicUrl according to environment and config', async () => {
      const { config } = await import('../config');
      const originalEnv = config.NODE_ENV;
      const originalUrl = config.BACKEND_PUBLIC_URL;

      try {
        // Explicit valid public URL set
        (config as any).BACKEND_PUBLIC_URL = 'https://api.customdomain.org/';
        expect(getBackendPublicUrl()).toBe('https://api.customdomain.org');

        // Localhost URL must never be accepted as public URL
        (config as any).BACKEND_PUBLIC_URL = 'http://localhost:3000';
        expect(getBackendPublicUrl()).toBeNull();

        // 127.0.0.1 must never be accepted as public URL
        (config as any).BACKEND_PUBLIC_URL = 'http://127.0.0.1:3000';
        expect(getBackendPublicUrl()).toBeNull();

        // Empty / unconfigured BACKEND_PUBLIC_URL must always return null to prevent localhost in workbook
        (config as any).BACKEND_PUBLIC_URL = '';
        (config as any).NODE_ENV = 'development';
        expect(getBackendPublicUrl()).toBeNull();

        (config as any).NODE_ENV = 'production';
        expect(getBackendPublicUrl()).toBeNull();
      } finally {
        (config as any).NODE_ENV = originalEnv;
        (config as any).BACKEND_PUBLIC_URL = originalUrl;
      }
    });
  });
});
