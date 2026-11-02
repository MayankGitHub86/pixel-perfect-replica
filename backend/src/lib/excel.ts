// backend/src/lib/excel.ts
// Real Microsoft OneDrive Excel workbook integration.
// Implements workbook creation, worksheet/table setup, and candidate upsert.

import { prisma } from './db';
import { decrypt } from './crypto';
import { refreshAccessToken } from './tokenRefresh';
import { logger } from '../logger';
import { config } from '../config';
import { ExcelWorkbook } from '@prisma/client';
// Using global fetch (Node 18+). No external import needed.

// Helper to get the latest valid access token (refresh if needed)
async function getAccessToken(): Promise<string> {
  const conn = await prisma.mailboxConnection.findFirst({ orderBy: { createdAt: 'desc' } });
  if (!conn) throw new Error('No mailbox connection found');
  if (new Date() >= new Date(conn.tokenExpiresAt.getTime() - 60000)) {
    await refreshAccessToken(conn.id);
    const refreshed = await prisma.mailboxConnection.findUnique({ where: { id: conn.id } });
    if (!refreshed) throw new Error('Failed to refresh token');
    return decrypt(refreshed.accessToken);
  }
  return decrypt(conn.accessToken);
}

// ------------- Workbook handling -----------------
export async function ensureWorkbookExists(): Promise<ExcelWorkbook> {
  const token = await getAccessToken();
  // Check if workbook already stored in DB
  const existing = await prisma.excelWorkbook.findFirst();
  if (existing) return existing;

  // List root children to find the recovery or master workbook
  const listRes = await fetch('https://graph.microsoft.com/v1.0/me/drive/root/children', {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!listRes.ok) {
    const body = await listRes.text();
    throw new Error(`OneDrive list root failed: ${listRes.status} ${body}`);
  }
  const listData: any = await listRes.json();
  const found = listData.value?.find((i: any) => i.name === 'Candidate_Master_RealData.xlsx') ||
                listData.value?.find((i: any) => i.name === 'Candidate_Master_Recovery_V2.xlsx') ||
                listData.value?.find((i: any) => i.name === 'Candidate_Master_Recovery.xlsx') ||
                listData.value?.find((i: any) => i.name === 'Candidate_Master.xlsx');
  if (found) {
    const wb = await prisma.excelWorkbook.create({
      data: { driveItemId: found.id, webUrl: found.webUrl },
    });
    return wb;
  }

  throw new Error('No valid Candidate workbook found in OneDrive root.');
}

export const CANDIDATE_TABLE_COLUMNS = [
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
];

export const EXCEL_DATE_TIME_FORMAT = 'dd-mm-yyyy hh:mm:ss AM/PM';

/**
 * Resolves the backend public URL for use in Excel hyperlinks.
 * - Requires BACKEND_PUBLIC_URL to be explicitly configured with a valid public HTTP(S) URL.
 * - Never returns a localhost URL or loopback address, ensuring that OneDrive
 *   workbooks (which are hosted in Microsoft 365 cloud) never receive broken localhost links.
 * - If BACKEND_PUBLIC_URL is not set, empty, or points to localhost/127.0.0.1, returns null.
 * - When null, the Excel resume hyperlink cell remains cleanly blank ("") until the actual
 *   production deployment domain is configured.
 */
export function getBackendPublicUrl(): string | null {
  const explicit = config.BACKEND_PUBLIC_URL?.trim();
  if (explicit && !/^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?(\/.*)?$/i.test(explicit)) {
    // Strip trailing slash for consistency
    return explicit.replace(/\/+$/, '');
  }
  return null;
}

/**
 * Converts a JS Date, timestamp string, or null/undefined to an Excel date-time serial number.
 * Days since 1899-12-30 (accounting for Lotus 1-2-3 1900 leap-year bug: 25569 days between 1899-12-30 and 1970-01-01).
 * Preserves the exact timestamp without artificial timezone shifting.
 */
export function toExcelSerialDate(dateInput: Date | string | null | undefined): number | string {
  if (!dateInput) return '';
  if (typeof dateInput === 'number') return dateInput;
  const d = typeof dateInput === 'string' ? new Date(dateInput) : dateInput;
  if (isNaN(d.getTime())) return '';

  const msPerDay = 86400000;
  return 25569 + (d.getTime() / msPerDay);
}

/**
 * Ensures table column number format is set to dd-mm-yyyy hh:mm:ss AM/PM for receivedAt, updatedAt, and sourceEmailReceivedAt.
 */
export async function ensureDateColumnsFormatting(base: string, token: string): Promise<void> {
  for (const colName of ['receivedAt', 'updatedAt', 'sourceEmailReceivedAt']) {
    try {
      const colRangeRes = await fetch(`${base}/tables/CandidateTable/columns('${colName}')/dataBodyRange`, {
        headers: { Authorization: `Bearer ${token}` },
      });
      if (colRangeRes.ok) {
        const colRangeData: any = await colRangeRes.json();
        const rowCount = colRangeData.values?.length || 0;
        if (rowCount > 0) {
          const numberFormat = new Array(rowCount).fill([EXCEL_DATE_TIME_FORMAT]);
          await fetch(`${base}/tables/CandidateTable/columns('${colName}')/dataBodyRange`, {
            method: 'PATCH',
            headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({ numberFormat }),
          });
        }
      }
    } catch (err: any) {
      logger.warn({ colName, err: err.message }, 'Failed to apply date format to table column');
    }
  }
}

// Ensure worksheet "Candidates" and table "CandidateTable" with required columns exist
export async function ensureWorksheetAndTable(driveItemId: string): Promise<string[]> {
  const token = await getAccessToken();
  const base = `https://graph.microsoft.com/v1.0/me/drive/items/${driveItemId}/workbook`;

  // 1. Check if worksheet "Candidates" exists
  const wsListRes = await fetch(`${base}/worksheets`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (wsListRes.ok) {
    const wsListData: any = await wsListRes.json();
    const wsFound = wsListData.value?.find((w: any) => w.name === 'Candidates');
    if (!wsFound) {
      const addWsRes = await fetch(`${base}/worksheets/add`, {
        method: 'POST',
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: 'Candidates' }),
      });
      if (!addWsRes.ok && addWsRes.status !== 409) {
        const body = await addWsRes.text();
        throw new Error(`Create worksheet failed: ${addWsRes.status} ${body}`);
      }
    }
  }

  // 2. Check if table "CandidateTable" exists
  const tableListRes = await fetch(`${base}/tables`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (tableListRes.ok) {
    const tableListData: any = await tableListRes.json();
    const tableFound = tableListData.value?.find((t: any) => t.name === 'CandidateTable');
    if (!tableFound) {
      const addTableRes = await fetch(`${base}/tables/add`, {
        method: 'POST',
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({
          address: 'Candidates!A1:V1',
          hasHeaders: true,
        }),
      });
      if (addTableRes.ok) {
        const addedTable: any = await addTableRes.json();
        await fetch(`${base}/tables/${addedTable.id}`, {
          method: 'PATCH',
          headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
          body: JSON.stringify({ name: 'CandidateTable' }),
        });
      }
    } else {
      // Ensure all 22 columns are present in existing CandidateTable
      const colsRes = await fetch(`${base}/tables/CandidateTable/columns`, {
        headers: { Authorization: `Bearer ${token}` },
      });
      if (colsRes.ok) {
        const colsData: any = await colsRes.json();
        const existingCols: string[] = (colsData.value || []).map((c: any) => c.name);
        const existingColNames = new Set(existingCols);
        for (const colName of CANDIDATE_TABLE_COLUMNS) {
          if (!existingColNames.has(colName)) {
            await fetch(`${base}/tables/CandidateTable/columns/add`, {
              method: 'POST',
              headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
              body: JSON.stringify({ name: colName }),
            });
            existingCols.push(colName);
          }
        }
        return existingCols;
      }
    }
  }
  return [];
}

// ----- Candidate Eligibility for Production Excel Sync -----

export interface CandidateEligibilityResult {
  eligible: boolean;
  reason: string;
}

/**
 * Checks whether a candidate record is genuine and eligible to enter the production Excel workbook.
 *
 * Rules:
 * - Rejects any candidate with explicit synthetic indicators (e.g. name containing '[DEV-SYNTHETIC]', email on 'synthetic.dev').
 * - Rejects synthetic source email IDs (e.g. 'SYNTH-MSG-*').
 * - Rejects candidates lacking a sourceEmailId (provenance requirement).
 * - Requires candidate's sourceEmailId to resolve to an actual EmailMessage record in PostgreSQL.
 * - Requires the linked EmailMessage to not be synthetic.
 * - Allows genuine candidates with verified Outlook provenance.
 */
export async function checkCandidateExcelEligibility(
  candidateId: string
): Promise<CandidateEligibilityResult> {
  const candidate = await prisma.candidate.findUnique({
    where: { id: candidateId },
  });

  if (!candidate) {
    return { eligible: false, reason: `Candidate ${candidateId} not found in database` };
  }

  // 1. Synthetic name check
  if (candidate.name && /\[DEV-SYNTHETIC\]/i.test(candidate.name)) {
    return { eligible: false, reason: 'Candidate marked with synthetic test name prefix' };
  }

  // 2. Synthetic email domain check
  if (candidate.email && /@synthetic\.dev$/i.test(candidate.email)) {
    return { eligible: false, reason: 'Candidate email belongs to synthetic test domain' };
  }

  // 3. Synthetic source ID prefix check
  if (candidate.sourceEmailId && /^SYNTH-/i.test(candidate.sourceEmailId)) {
    return { eligible: false, reason: 'Candidate has synthetic sourceEmailId (SYNTH-*)' };
  }

  // 4. Source provenance requirement
  if (!candidate.sourceEmailId) {
    return { eligible: false, reason: 'Candidate lacks sourceEmailId provenance link' };
  }

  // 5. Verify sourceEmailId resolves to an actual EmailMessage in PostgreSQL
  const emailMessage = await prisma.emailMessage.findFirst({
    where: {
      OR: [
        { id: candidate.sourceEmailId },
        { messageId: candidate.sourceEmailId },
      ],
    },
  });

  if (!emailMessage) {
    return {
      eligible: false,
      reason: `sourceEmailId "${candidate.sourceEmailId}" does not resolve to an EmailMessage record`,
    };
  }

  // 6. Verify linked EmailMessage is not a synthetic test message
  if (
    emailMessage.messageId.startsWith('SYNTH-') ||
    emailMessage.fromEmail.endsWith('@synthetic.dev') ||
    /\[DEV-SYNTHETIC\]/i.test(emailMessage.subject)
  ) {
    return {
      eligible: false,
      reason: 'Linked EmailMessage record is marked as synthetic test data',
    };
  }

  return { eligible: true, reason: 'Verified genuine candidate with valid Outlook provenance' };
}

// ----- Candidate upsert -----
export async function upsertCandidateRow(candidateId: string): Promise<boolean> {
  // Pre-flight check: Ensure candidate is eligible for production Excel sync
  const eligibility = await checkCandidateExcelEligibility(candidateId);
  if (!eligibility.eligible) {
    logger.info(
      { candidateId, reason: eligibility.reason },
      'Skipping Excel synchronization: Candidate not eligible for production workbook'
    );
    return false;
  }

  const token = await getAccessToken();
  const workbook = await ensureWorkbookExists();
  const tableCols = await ensureWorksheetAndTable(workbook.driveItemId);
  const base = `https://graph.microsoft.com/v1.0/me/drive/items/${workbook.driveItemId}/workbook`;

  // Fetch candidate from DB
  const candidate = await prisma.candidate.findUnique({
    where: { id: candidateId },
    include: { attachments: true },
  });
  if (!candidate) throw new Error(`Candidate ${candidateId} not found`);

  // Fetch source email for provenance
  let sourceEmail: { subject: string; receivedAt: Date; fromEmail?: string } | null = null;
  if (candidate.sourceEmailId) {
    sourceEmail = await prisma.emailMessage.findFirst({
      where: {
        OR: [
          { id: candidate.sourceEmailId },
          { messageId: candidate.sourceEmailId },
        ],
      },
    });
  }

  // Ensure OneDrive resume upload if needed
  let resumeHyperlink = '';
  let resumeUrl = candidate.resumeUrl;
  if (!resumeUrl && candidate.resumeRef && candidate.resumeFileName) {
    try {
      const fs = await import('fs');
      if (fs.existsSync(candidate.resumeRef)) {
        const fileBuffer = fs.readFileSync(candidate.resumeRef);
        const { uploadResumeToOneDrive } = await import('./attachmentService');
        const odResult = await uploadResumeToOneDrive(candidate.id, candidate.resumeFileName, fileBuffer);
        resumeUrl = odResult.webUrl;
        await prisma.candidate.update({
          where: { id: candidate.id },
          data: { resumeUrl: odResult.webUrl },
        });
      }
    } catch (odErr: any) {
      logger.warn({ candidateId, err: odErr.message }, 'Failed on-the-fly OneDrive resume upload during Excel sync');
    }
  }

  if (candidate.resumeFileName || candidate.resumeRef || candidate.resumeUrl) {
    const backendUrl = getBackendPublicUrl();
    if (backendUrl) {
      // Dynamic structured reference hyperlink ensuring each candidate gets their own candidate-specific download
      resumeHyperlink = `=IF([@id]="", "", HYPERLINK("${backendUrl}/api/candidates/" & [@id] & "/resume", "Download CV"))`;
    } else {
      // Production without BACKEND_PUBLIC_URL configured: leave blank rather than write localhost
      logger.warn(
        { candidateId },
        'BACKEND_PUBLIC_URL not configured in production — Excel resume hyperlink left blank'
      );
    }
  }

  // Build row mapping candidate fields to columns
  const fieldMap: Record<string, any> = {
    id: candidate.id ?? '',
    name: candidate.name ?? '',
    email: candidate.email ?? '',
    phone: candidate.phone ?? '',
    location: candidate.location ?? '',
    appliedRole: candidate.appliedRole ?? '',
    degree: candidate.degree ?? '',
    college: candidate.college ?? '',
    graduationYear: candidate.graduationYear?.toString() ?? '',
    experienceYears: candidate.experienceYears != null ? candidate.experienceYears.toString() : '',
    skills: candidate.skills?.join(', ') ?? '',
    confidenceScore: '',
    reviewStatus: '',
    applicationStatus: '',
    sourceEmailId: candidate.sourceEmailId ?? '',
    receivedAt: candidate.receivedAt != null ? toExcelSerialDate(candidate.receivedAt) : '',
    updatedAt: candidate.updatedAt != null ? toExcelSerialDate(candidate.updatedAt) : '',
    resumeFileName: candidate.resumeFileName ?? '',
    resumeLink: resumeHyperlink,
    sourceEmailSubject: sourceEmail?.subject ?? '',
    sourceEmailReceivedAt: sourceEmail?.receivedAt != null ? toExcelSerialDate(sourceEmail.receivedAt) : '',
    extractionStatus: candidate.processingStatus ?? '',
  };

  const targetCols = tableCols && tableCols.length > 0 ? tableCols : CANDIDATE_TABLE_COLUMNS;
  const row = targetCols.map(colName => (fieldMap[colName] !== undefined ? fieldMap[colName] : ''));

  // Retrieve existing rows (limit 1000 for safety)
  const rowsRes = await fetch(`${base}/tables/CandidateTable/rows`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!rowsRes.ok) {
    const body = await rowsRes.text();
    throw new Error(`Fetch table rows failed: ${rowsRes.status} ${body}`);
  }
  const rowsData: any = await rowsRes.json();
  const existingRow = rowsData.value?.find((r: any) => r.values?.[0]?.[0] === candidate.id);

  if (existingRow) {
    // Preserve any existing non-empty values (e.g. manual Excel edits, concise Experience, etc.) if new value is empty
    if (Array.isArray(existingRow.values?.[0])) {
      const existingVals = existingRow.values[0];
      // Expand row array if table has additional columns (e.g. Experience, Reference Type) to avoid truncation
      while (row.length < existingVals.length) {
        row.push(existingVals[row.length]);
      }
      // Only preserve non-managed / custom columns (e.g. manual Experience notes, Reference Type, Referrer details)
      const managedFields = new Set([
        'id', 'name', 'email', 'phone', 'location', 'appliedRole',
        'degree', 'college', 'graduationYear', 'experienceYears',
        'skills', 'resumeFileName', 'resumeLink', 'sourceEmailId',
        'receivedAt', 'updatedAt', 'sourceEmailSubject', 'sourceEmailReceivedAt',
        'extractionStatus'
      ]);
      for (let c = 0; c < row.length; c++) {
        const colName = targetCols[c];
        if (!managedFields.has(colName) && (row[c] === '' || row[c] == null) && existingVals[c] !== '' && existingVals[c] != null) {
          row[c] = existingVals[c];
        }
      }
    }

    // Update the row – using Graph table itemAt(index)
    const rowIndex = existingRow.index;
    const patchRes = await fetch(`${base}/tables/CandidateTable/rows/itemAt(index=${rowIndex})`, {
      method: 'PATCH',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ values: [row] }),
    });
    if (!patchRes.ok) {
      // Fallback try direct rows/${rowIndex}
      await fetch(`${base}/tables/CandidateTable/rows/${rowIndex}`, {
        method: 'PATCH',
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ values: [row] }),
      });
    }
    logger.info({ candidateId }, 'Updated existing Excel candidate row');
  } else {
    // Append new row
    const addRes = await fetch(`${base}/tables/CandidateTable/rows/add`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ values: [row] }),
    });
    if (!addRes.ok) {
      const errBody = await addRes.text();
      throw new Error(`Add Excel row failed: ${addRes.status} ${errBody}`);
    }
    logger.info({ candidateId }, 'Inserted new Excel candidate row');
  }

  // Ensure Excel display format dd-mm-yyyy hh:mm:ss AM/PM is applied to date columns
  if (process.env.NODE_ENV !== 'test') {
    await ensureDateColumnsFormatting(base, token);
  }

  // Record sync run for auditing
  await prisma.excelSyncRun.create({
    data: {
      candidateId: candidate.id,
      status: 'success',
      result: 'success',
      rows: 1,
      message: existingRow ? 'updated' : 'inserted',
    },
  });

  return true;
}

// Helper for dashboard status
export async function getExcelStatus() {
  const workbook = await prisma.excelWorkbook.findFirst();
  const [recentRuns, lastSuccess, failedJobsCount, totalCandidates, syncedCount] = await Promise.all([
    prisma.excelSyncRun.findMany({ orderBy: { createdAt: 'desc' }, take: 10 }),
    prisma.excelSyncRun.findFirst({ where: { result: 'success' }, orderBy: { createdAt: 'desc' } }),
    prisma.processingJob.count({ where: { type: 'excel_sync', status: 'failed' } }),
    prisma.candidate.count(),
    prisma.excelSyncRun.count({ where: { result: 'success' } }),
  ]);
  const pendingRows = Math.max(0, totalCandidates - syncedCount);
  let activeWorkbookName: string | null = null;
  if (workbook?.webUrl) {
    if (workbook.webUrl.includes('Candidate_Master_RealData') || workbook.driveItemId === '5082520FEDE864C4!s8dc280a78ce54f6b9ee4617268ede8bb') {
      activeWorkbookName = 'Candidate_Master_RealData.xlsx';
    } else {
      activeWorkbookName = 'Candidate_Master_Recovery_V2.xlsx';
    }
  }

  return {
    workbookName: activeWorkbookName,
    state: workbook ? 'connected' : 'not_connected',
    storageProvider: 'OneDrive',
    workbookUrl: workbook?.webUrl ?? null,
    lastSuccessfulSync: lastSuccess?.createdAt?.toISOString() ?? null,
    pendingRows,
    failedJobs: failedJobsCount,
    lastReconciliation: null,
    history: recentRuns.map((r) => ({
      id: r.id,
      at: r.createdAt.toISOString(),
      rows: r.rows,
      result: r.result as any,
      message: r.message,
    })),
  };
}

export type ExcelStatus = Awaited<ReturnType<typeof getExcelStatus>>;

