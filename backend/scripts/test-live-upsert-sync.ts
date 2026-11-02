import { PrismaClient } from '@prisma/client';
import dotenv from 'dotenv';
import path from 'path';
import crypto from 'crypto';
import { upsertCandidateRow, toExcelSerialDate, EXCEL_DATE_TIME_FORMAT } from '../src/lib/excel';

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

async function main() {
  console.log('=== TEST LIVE EXCEL SYNC / WRITE & READ-BACK ===\n');

  const candidateId = '106945ea-6ae3-4c64-b5bb-9c1148d72985';

  // Step 1: Run upsertCandidateRow
  console.log(`[Step 1] Executing upsertCandidateRow for ${candidateId}...`);
  const success = await upsertCandidateRow(candidateId);
  console.log('upsertCandidateRow result:', success);
  if (!success) throw new Error('upsertCandidateRow failed');

  // Step 2: Read back from Microsoft Graph
  console.log('\n[Step 2] Reading back workbook data via Microsoft Graph...');
  const token = await getValidAccessToken();
  const wb = await prisma.excelWorkbook.findFirst();
  if (!wb) throw new Error('Workbook record not found');
  const base = `https://graph.microsoft.com/v1.0/me/drive/items/${wb.driveItemId}/workbook`;

  // Check columns
  const colsRes = await fetch(`${base}/tables/CandidateTable/columns`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const colsData = await colsRes.json();
  const cols: string[] = (colsData.value || []).map((c: any) => c.name);
  console.log('Columns count:', cols.length);

  // Check rows
  const rowsRes = await fetch(`${base}/tables/CandidateTable/rows`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const rowsData = await rowsRes.json();
  const rows = rowsData.value || [];
  console.log('Rows count:', rows.length);
  if (rows.length !== 1) {
    throw new Error(`Expected exactly 1 row, found ${rows.length}`);
  }

  const row0 = rows[0].values[0];
  const idColIdx = cols.indexOf('id');
  const nameColIdx = cols.indexOf('name');
  const emailColIdx = cols.indexOf('email');
  const expYearsColIdx = cols.indexOf('experienceYears');
  const expColIdx = cols.indexOf('Experience');
  const recAtColIdx = cols.indexOf('receivedAt');
  const updAtColIdx = cols.indexOf('updatedAt');
  const refTypeColIdx = cols.indexOf('Reference Type');
  const refNameColIdx = cols.indexOf('Referrer Name');
  const refEmailColIdx = cols.indexOf('Referrer Email');
  const resumeLinkColIdx = cols.indexOf('resumeLink');

  console.log('\n--- Live Row Data ---');
  console.log('UUID:           ', row0[idColIdx]);
  console.log('Name:           ', row0[nameColIdx]);
  console.log('Email:          ', row0[emailColIdx]);
  console.log('experienceYears:', JSON.stringify(row0[expYearsColIdx]));
  console.log('receivedAt (raw):', row0[recAtColIdx], `(type: ${typeof row0[recAtColIdx]})`);
  console.log('updatedAt (raw): ', row0[updAtColIdx], `(type: ${typeof row0[updAtColIdx]})`);
  console.log('Reference Type: ', row0[refTypeColIdx]);
  console.log('Referrer Name:  ', JSON.stringify(row0[refNameColIdx]));
  console.log('Referrer Email: ', JSON.stringify(row0[refEmailColIdx]));
  console.log('resumeLink:     ', row0[resumeLinkColIdx]);
  console.log('Experience:\n"""\n' + row0[expColIdx] + '\n"""');

  // Step 3: Check dataBodyRange for receivedAt and updatedAt
  console.log('\n[Step 3] Checking column dataBodyRange format and display text...');
  const recRangeRes = await fetch(`${base}/tables/CandidateTable/columns('receivedAt')/dataBodyRange`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const recRange = await recRangeRes.json();
  console.log('receivedAt range verification:', {
    address: recRange.address,
    value: recRange.values[0][0],
    valueType: recRange.valueTypes[0][0],
    numberFormat: recRange.numberFormat[0][0],
    text: recRange.text[0][0],
  });

  const updRangeRes = await fetch(`${base}/tables/CandidateTable/columns('updatedAt')/dataBodyRange`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const updRange = await updRangeRes.json();
  console.log('updatedAt range verification:', {
    address: updRange.address,
    value: updRange.values[0][0],
    valueType: updRange.valueTypes[0][0],
    numberFormat: updRange.numberFormat[0][0],
    text: updRange.text[0][0],
  });

  const passed =
    typeof recRange.values[0][0] === 'number' &&
    recRange.valueTypes[0][0] === 'Double' &&
    recRange.numberFormat[0][0] === 'dd-mm-yyyy hh:mm:ss AM/PM' &&
    typeof updRange.values[0][0] === 'number' &&
    updRange.valueTypes[0][0] === 'Double' &&
    updRange.numberFormat[0][0] === 'dd-mm-yyyy hh:mm:ss AM/PM';

  console.log('\n=== ALL FORMAT & WORKFLOW CHECKS: ' + (passed ? 'PASSED' : 'FAILED') + ' ===');
}

main().finally(() => prisma.$disconnect());

