import path from 'path';
import 'dotenv/config';
import crypto from 'crypto';
import { PrismaClient } from '@prisma/client';
import { prisma } from '../src/lib/db.js';
import { decrypt } from '../src/lib/crypto.js';
import { refreshAccessToken } from '../src/lib/tokenRefresh.js';

const prismaClient = new PrismaClient();

async function getValidAccessToken(): Promise<string> {
  const conn = await prismaClient.mailboxConnection.findFirst({ orderBy: { createdAt: 'desc' } });
  if (!conn) throw new Error('No mailbox connection found');
  const decryptedAccess = decrypt(conn.accessToken);
  const [, payloadB64] = decryptedAccess.split('.');
  let payload: any = {};
  try { payload = JSON.parse(Buffer.from(payloadB64, 'base64url').toString('utf8')); } catch (_) { payload = {}; }
  const expiresAt = payload.exp ? payload.exp * 1000 : conn.tokenExpiresAt?.getTime();
  const now = Date.now();
  if (!expiresAt || expiresAt - now < 5 * 60 * 1000) {
    await refreshAccessToken(conn.id);
    const refreshed = await prismaClient.mailboxConnection.findUnique({ where: { id: conn.id } });
    if (!refreshed) throw new Error('Failed to reload after refresh');
    return decrypt(refreshed.accessToken);
  }
  return decryptedAccess;
}

function decrypt(payload: string): string {
  const [ivB64, encryptedB64, tagB64] = payload.split('.');
  const iv = Buffer.from(ivB64, 'base64');
  const encrypted = Buffer.from(encryptedB64, 'base64');
  const tag = Buffer.from(tagB64, 'base64');
  const key = Buffer.from(process.env.TOKEN_ENCRYPTION_KEY!, 'base64');
  const decipher = crypto.createDecipheriv('aes-256-gcm', key, iv, { authTagLength: 16 });
  decipher.setAuthTag(tag);
  return Buffer.concat([decipher.update(encrypted), decipher.final()]).toString('utf8');
}

async function fetchJson(url: string, token: string) {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const txt = await res.text();
  let data: any;
  try { data = JSON.parse(txt); } catch (_) { data = txt; }
  return { status: res.status, data };
}

async function postJson(url: string, token: string, body: any) {
  const res = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(body),
  });
  const txt = await res.text();
  let data: any;
  try { data = JSON.parse(txt); } catch (_) { data = txt; }
  return { status: res.status, data };
}

async function main() {
  const token = await getValidAccessToken();
  const wb = await prismaClient.excelWorkbook.findFirst();
  if (!wb) throw new Error('ExcelWorkbook record not found');
  const realId = wb.driveItemId;
  const base = `https://graph.microsoft.com/v1.0/me/drive/items/${realId}/workbook`;

  const wsRes = await fetchJson(`${base}/worksheets`, token);
  if (wsRes.status !== 200) throw new Error(`Failed to list worksheets: ${wsRes.status}`);
  const candidatesWs = (wsRes.data.value || []).find((w: any) => w.name === 'Candidates');
  if (!candidatesWs) throw new Error('Worksheet "Candidates" not found');

  const tblRes = await fetchJson(`${base}/tables`, token);
  if (tblRes.status !== 200) throw new Error(`Failed to list tables: ${tblRes.status}`);
  const candidateTable = (tblRes.data.value || []).find((t: any) => t.name === 'CandidateTable');
  if (!candidateTable) throw new Error('Table "CandidateTable" not found');

  async function getColumnNames() {
    const colsRes = await fetchJson(`${base}/tables/${candidateTable.id}/columns`, token);
    if (colsRes.status !== 200) throw new Error(`Failed to get columns: ${colsRes.status}`);
    return (colsRes.data.value || []).map((c: any) => c.name);
  }

  const beforeHeaders = await getColumnNames();
  console.log('Before headers:', beforeHeaders);

  const required = [
    'Candidate Email', 'Reference Type', 'Referrer Name', 'Referrer Email',
    'Total Experience (Years)', 'Previous/Current Company', 'Job Title / Role',
    'Employment Duration', 'Responsibilities / Work Description',
    'Skills / Technologies', 'Download CV',
  ];
  const missing = required.filter(col => !beforeHeaders.includes(col));
  console.log('Missing columns to add:', missing);

  for (const col of missing) {
    const addRes = await postJson(`${base}/tables/${candidateTable.id}/columns/add`, token, { name: col });
    if (addRes.status !== 200 && addRes.status !== 201) {
      throw new Error(`Failed to add column "${col}": ${addRes.status} ${JSON.stringify(addRes.data)}`);
    }
    console.log(`Added column: ${col} (status ${addRes.status})`);
  }

  const afterHeaders = await getColumnNames();
  console.log('After headers:', afterHeaders);

  const rangeRes = await fetchJson(`${base}/tables/${candidateTable.id}/range`, token);
  if (rangeRes.status !== 200) throw new Error(`Failed to fetch table range: ${rangeRes.status}`);
  console.log('Table range address:', rangeRes.data.address);
  console.log('Table rowCount (including header):', rangeRes.data.rowCount);
  console.log('Table columnCount:', rangeRes.data.columnCount);

  const wsRangeRes = await fetchJson(`${base}/worksheets/${candidatesWs.id}/range`, token);
  if (wsRangeRes.status !== 200) throw new Error(`Failed to fetch worksheet range: ${wsRangeRes.status}`);
  console.log('Worksheet used range address:', wsRangeRes.data.address);
  console.log('Worksheet rowCount:', wsRangeRes.data.rowCount);
  console.log('Worksheet columnCount:', wsRangeRes.data.columnCount);

  const dataRowCount = rangeRes.data.rowCount - 1;
  console.log('Candidate data rows (excluding header):', dataRowCount);

  await prismaClient.$disconnect();
}

main().catch(err => { console.error('Error:', err.message); process.exit(1); });

const prisma = new PrismaClient();

function decrypt(payload: string): string {
  const [ivB64, encryptedB64, tagB64] = payload.split('.');
  const iv = Buffer.from(ivB64, 'base64');
  const encrypted = Buffer.from(encryptedB64, 'base64');
  const tag = Buffer.from(tagB64, 'base64');
  const key = Buffer.from(process.env.TOKEN_ENCRYPTION_KEY!, 'base64');
  const decipher = crypto.createDecipheriv('aes-256-gcm', key, iv, { authTagLength: 16 });
  decipher.setAuthTag(tag);
  return Buffer.concat([decipher.update(encrypted), decipher.final()]).toString('utf8');
}

async function fetchJson(url: string, token: string) {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const txt = await res.text();
  let data: any;
  try { data = JSON.parse(txt); } catch (_) { data = txt; }
  return { status: res.status, data };
}

async function postJson(url: string, token: string, body: any) {
  const res = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(body),
  });
  const txt = await res.text();
  let data: any;
  try { data = JSON.parse(txt); } catch (_) { data = txt; }
  return { status: res.status, data };
}

async function main() {
  // Ensure we have a fresh token
  const token = await getValidAccessToken();

  // Get the workbook DriveItem ID from DB
  const wb = await prisma.excelWorkbook.findFirst();
  if (!wb) throw new Error('ExcelWorkbook record not found');
  const realId = wb.driveItemId;
  const base = `https://graph.microsoft.com/v1.0/me/drive/items/${realId}/workbook`;

  // 1️⃣ Fetch worksheets to locate the main worksheet
  const wsRes = await fetchJson(`${base}/worksheets`, token);
  if (wsRes.status !== 200) throw new Error(`Failed to list worksheets: ${wsRes.status}`);
  const worksheets = wsRes.data.value || [];
  const candidatesWs = worksheets.find((w: any) => w.name === 'Candidates');
  if (!candidatesWs) throw new Error('Worksheet "Candidates" not found');

  // 2️⃣ Fetch tables to locate CandidateTable
  const tblRes = await fetchJson(`${base}/tables`, token);
  if (tblRes.status !== 200) throw new Error(`Failed to list tables: ${tblRes.status}`);
  const tables = tblRes.data.value || [];
  const candidateTable = tables.find((t: any) => t.name === 'CandidateTable');
  if (!candidateTable) throw new Error('Table "CandidateTable" not found');

  // Helper to get current column names
  async function getColumnNames() {
    const colsRes = await fetchJson(`${base}/tables/${candidateTable.id}/columns`, token);
    if (colsRes.status !== 200) throw new Error(`Failed to get columns: ${colsRes.status}`);
    return (colsRes.data.value || []).map((c: any) => c.name);
  }

  const beforeHeaders = await getColumnNames();
  console.log('Before headers:', beforeHeaders);

  // Required columns
  const required = [
    'Candidate Email',
    'Reference Type',
    'Referrer Name',
    'Referrer Email',
    'Total Experience (Years)',
    'Previous/Current Company',
    'Job Title / Role',
    'Employment Duration',
    'Responsibilities / Work Description',
    'Skills / Technologies',
    'Download CV',
  ];

  const missing = required.filter(col => !beforeHeaders.includes(col));
  console.log('Missing columns to add:', missing);

  for (const col of missing) {
    const addRes = await postJson(`${base}/tables/${candidateTable.id}/columns/add`, token, { name: col });
    if (addRes.status !== 200 && addRes.status !== 201) {
      throw new Error(`Failed to add column "${col}": ${addRes.status} ${JSON.stringify(addRes.data)}`);
    }
    console.log(`Added column: ${col} (status ${addRes.status})`);
  }

  const afterHeaders = await getColumnNames();
  console.log('After headers:', afterHeaders);

  // 3️⃣ Verify table range after addition
  const rangeRes = await fetchJson(`${base}/tables/${candidateTable.id}/range`, token);
  if (rangeRes.status !== 200) throw new Error(`Failed to fetch table range: ${rangeRes.status}`);
  const tableRange = rangeRes.data;
  console.log('Table range address:', tableRange.address);
  console.log('Table rowCount (including header):', tableRange.rowCount);
  console.log('Table columnCount:', tableRange.columnCount);

  // 4️⃣ Worksheet used range
  const wsRangeRes = await fetchJson(`${base}/worksheets/${candidatesWs.id}/range`, token);
  if (wsRangeRes.status !== 200) throw new Error(`Failed to fetch worksheet range: ${wsRangeRes.status}`);
  const wsRange = wsRangeRes.data;
  console.log('Worksheet used range address:', wsRange.address);
  console.log('Worksheet rowCount:', wsRange.rowCount);
  console.log('Worksheet columnCount:', wsRange.columnCount);

  // 5️⃣ Row count (excluding header) – can derive from table range
  const dataRowCount = tableRange.rowCount - 1; // minus header row
  console.log('Candidate data rows (excluding header):', dataRowCount);

  await prisma.$disconnect();
}

main().catch(err => {
  console.error('Error:', err.message);
  process.exit(1);
});
