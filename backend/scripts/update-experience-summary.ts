// backend/scripts/update-experience-summary.ts
// Updates the Experience cell in CandidateTable with concise, professional summaries.

import crypto from 'crypto';
import path from 'path';
import { PrismaClient } from '@prisma/client';
import dotenv from 'dotenv';

dotenv.config({ path: path.resolve(process.cwd(), '.env') });
const prisma = new PrismaClient();

function decryptToken(payload: string): string {
  const [ivB64, encryptedB64, tagB64] = payload.split('.');
  const iv = Buffer.from(ivB64, 'base64');
  const encrypted = Buffer.from(encryptedB64, 'base64');
  const tag = Buffer.from(tagB64, 'base64');
  const key = Buffer.from(process.env.TOKEN_ENCRYPTION_KEY!, 'base64');
  const decipher = crypto.createDecipheriv('aes-256-gcm', key, iv, { authTagLength: 16 });
  decipher.setAuthTag(tag);
  return Buffer.concat([decipher.update(encrypted), decipher.final()]).toString('utf8');
}

async function getValidAccessToken(): Promise<string> {
  const conn = await prisma.mailboxConnection.findFirst({ orderBy: { createdAt: 'desc' } });
  if (!conn) throw new Error('No mailbox connection found');
  const decryptedAccess = decryptToken(conn.accessToken);
  const [, payloadB64] = decryptedAccess.split('.');
  let payload: any = {};
  try { payload = JSON.parse(Buffer.from(payloadB64, 'base64url').toString('utf8')); } catch { payload = {}; }
  const expiresAt = payload.exp ? payload.exp * 1000 : (conn.tokenExpiresAt as any)?.getTime?.();
  const now = Date.now();
  if (!expiresAt || expiresAt - now < 5 * 60 * 1000) {
    const refreshToken = decryptToken(conn.refreshToken);
    const tokenRes = await fetch(
      `https://login.microsoftonline.com/${process.env.MICROSOFT_TENANT_ID}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id: process.env.MICROSOFT_CLIENT_ID!,
          client_secret: process.env.MICROSOFT_CLIENT_SECRET!,
          grant_type: 'refresh_token',
          refresh_token: refreshToken,
          redirect_uri: process.env.MICROSOFT_REDIRECT_URI || 'http://localhost:3000/api/integrations/microsoft/callback',
        }).toString(),
      }
    );
    if (!tokenRes.ok) throw new Error(`Token refresh failed: ${tokenRes.status}`);
    const tokenData: any = await tokenRes.json();
    return tokenData.access_token;
  }
  return decryptedAccess;
}

async function graphGet(url: string, token: string) {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const text = await res.text();
  let data: any;
  try { data = JSON.parse(text); } catch { data = text; }
  return { status: res.status, data };
}

async function graphPatch(url: string, token: string, body: any) {
  const res = await fetch(url, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  const text = await res.text();
  let data: any;
  try { data = JSON.parse(text); } catch { data = text; }
  return { status: res.status, data };
}

async function main() {
  console.log('=== UPDATE EXPERIENCE CELL WITH CONCISE SUMMARY ===\n');

  const token = await getValidAccessToken();
  console.log('[1/4] Graph token acquired');

  const wb = await prisma.excelWorkbook.findFirst();
  if (!wb) throw new Error('Excel workbook record not found');
  const base = `https://graph.microsoft.com/v1.0/me/drive/items/${wb.driveItemId}/workbook`;

  // 1. Get CandidateTable info
  const tblRes = await graphGet(`${base}/tables`, token);
  const tables = tblRes.data?.value || [];
  const candTable = tables.find((t: any) => t.name === 'CandidateTable');
  if (!candTable) throw new Error('CandidateTable not found');

  const colsRes = await graphGet(`${base}/tables/${candTable.id}/columns`, token);
  const cols: string[] = (colsRes.data?.value || []).map((c: any) => c.name);
  console.log(`[2/4] CandidateTable columns (${cols.length})`);

  const expColIdx = cols.indexOf('Experience');
  const expYearsColIdx = cols.indexOf('experienceYears');
  const idColIdx = cols.indexOf('id');
  const nameColIdx = cols.indexOf('name');

  if (expColIdx === -1) throw new Error('"Experience" column not found');
  if (expYearsColIdx === -1) throw new Error('"experienceYears" column not found');

  // 2. Get current row
  const rowsRes = await graphGet(`${base}/tables/${candTable.id}/rows`, token);
  const rows = rowsRes.data?.value || [];
  console.log(`[3/4] Current table rows count: ${rows.length}`);

  if (rows.length === 0) throw new Error('No rows found in CandidateTable');

  // Concise professional summary strictly adhering to the requested format:
  // Company Name | Job Title | Employment Dates/Duration
  // - 1–2 short bullet points summarizing the main responsibilities and relevant work.
  const conciseExperienceSummary = [
    'LCC Projects Ltd. | Data Analyst Intern | Apr 2026 – Present',
    '- Built Power BI dashboards for daily log monitoring and equipment productivity/idle tracking across plant sites.',
    '- Cleaned and transformed operational datasets using Power Query and DAX, automating recurring reporting workflows.',
    '',
    'Sparks to Ideas | MERN Stack Intern | Oct 2025 – Mar 2026',
    '- Developed full-stack web application features using MongoDB, Express, React, and Node.js during a 6-month internship.',
    '',
    'CESA | Graphic Head | 2025 – 2026',
    '- Coordinated creative branding assets, documentation, and design deliverables for technical events.',
    '',
    'The Knowledge Academy | Part-Time Mathematics Tutor | 2022 – 2024',
    '- Designed structured lesson plans and tracked student performance metrics to identify and address learning gaps.',
  ].join('\n');

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const vals = [...(row.values?.[0] || [])];
    const candidateId = vals[idColIdx];
    const candidateName = vals[nameColIdx];
    const previousYears = vals[expYearsColIdx];

    console.log(`\nUpdating Candidate: ${candidateName} (UUID: ${candidateId})`);
    console.log(`  experienceYears (preserving unchanged): ${JSON.stringify(previousYears)}`);

    // Update ONLY the Experience column in this row
    vals[expColIdx] = conciseExperienceSummary;
    // Ensure experienceYears is preserved exactly as it is
    vals[expYearsColIdx] = previousYears;

    const rowIndex = row.index;
    const patchRes = await graphPatch(
      `${base}/tables/${candTable.id}/rows/itemAt(index=${rowIndex})`,
      token,
      { values: [vals] }
    );

    if (patchRes.status === 200) {
      console.log(`  ✓ Successfully updated row index ${rowIndex}`);
    } else {
      throw new Error(`Failed to update row: ${patchRes.status} ${JSON.stringify(patchRes.data)}`);
    }
  }

  // 3. Verification read-back
  console.log('\n[4/4] VERIFICATION READ-BACK FROM MICROSOFT GRAPH:');
  console.log('----------------------------------------------------');

  const verifyRowsRes = await graphGet(`${base}/tables/${candTable.id}/rows`, token);
  const verifyRows = verifyRowsRes.data?.value || [];

  for (let i = 0; i < verifyRows.length; i++) {
    const vr = verifyRows[i].values?.[0] || [];
    console.log(`Row ${i + 1}:`);
    console.log(`UUID:               ${vr[idColIdx]}`);
    console.log(`Name:               ${vr[nameColIdx]}`);
    console.log(`experienceYears:    ${JSON.stringify(vr[expYearsColIdx])}`);
    console.log(`Experience Cell:\n"""\n${vr[expColIdx]}\n"""`);
  }

  console.log('\nDone.');
  await prisma.$disconnect();
}

main().catch(err => {
  console.error('Error:', err);
  process.exit(1);
});

