// backend/scripts/verify-recovery-workbook.js
// Read‑only verification of the recovery Excel workbook.
// Uses the same Microsoft Graph token handling as upload-recovery-workbook.js.
// No data is written to the workbook, PostgreSQL, or Outlook.

const fs = require('fs');
const path = require('path');
require('dotenv').config();
const { PrismaClient } = require('@prisma/client');
const crypto = require('crypto');

const prisma = new PrismaClient();

function decrypt(payload) {
  const [ivB64, encryptedB64, tagB64] = payload.split('.');
  if (!ivB64 || !encryptedB64 || !tagB64) throw new Error('Invalid encrypted payload format');
  const iv = Buffer.from(ivB64, 'base64');
  const encrypted = Buffer.from(encryptedB64, 'base64');
  const tag = Buffer.from(tagB64, 'base64');
  const key = Buffer.from(process.env.TOKEN_ENCRYPTION_KEY, 'base64');
  const decipher = crypto.createDecipheriv('aes-256-gcm', key, iv, { authTagLength: 16 });
  decipher.setAuthTag(tag);
  const decrypted = Buffer.concat([decipher.update(encrypted), decipher.final()]);
  return decrypted.toString('utf8');
}

async function getAccessToken() {
  const conn = await prisma.mailboxConnection.findFirst({ orderBy: { createdAt: 'desc' } });
  if (!conn) throw new Error('No mailbox connection found');
  if (new Date() >= conn.tokenExpiresAt) {
    // Refresh token using the same helper used elsewhere
    const { refreshAccessToken } = require('../dist/lib/tokenRefresh');
    await refreshAccessToken(conn.id);
    const refreshed = await prisma.mailboxConnection.findUnique({ where: { id: conn.id } });
    if (!refreshed) throw new Error('Failed to refresh token');
    return decrypt(refreshed.accessToken);
  }
  return decrypt(conn.accessToken);
}

// Helper to perform a GET request with optional retry on 503
async function graphGet(url, token, maxAttempts = 3, delayMs = 2000) {
  let attempt = 0;
  while (attempt < maxAttempts) {
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (res.status !== 503) return { res, attempt };
    attempt++;
    if (attempt < maxAttempts) {
      console.warn(`Received 503 from ${url}, retry ${attempt}/${maxAttempts - 1} after ${delayMs}ms`);
      await new Promise(r => setTimeout(r, delayMs));
      delayMs *= 2; // exponential back‑off
    }
  }
  // final attempt result (still 503)
  const finalRes = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  return { res: finalRes, attempt };
}

async function main() {
  const token = await getAccessToken();
  console.log('✅ Access token acquired');

  // 1. Load ExcelWorkbook record from PostgreSQL
  const workbookRecord = await prisma.excelWorkbook.findFirst();
  if (!workbookRecord) throw new Error('ExcelWorkbook record not found in DB');
  const driveItemId = workbookRecord.driveItemId;

  // Helper to log diagnostics and parse JSON safely
async function logAndParse(res, url) {
  const contentType = res.headers.get('content-type') || '';
  const bodyText = await res.clone().text();
  const snippet = bodyText.slice(0, 200).replace(/[\r\n]+/g, ' ');
  const { hostname, pathname } = new URL(url);
  console.log('\n--- Response diagnostics ---');
  console.log('URL:', `${hostname}${pathname}`);
  console.log('Status:', res.status);
  console.log('Content-Type:', contentType);
  console.log('Body snippet:', snippet);
  return res.json();
}


  const metaUrl = `https://graph.microsoft.com/v1.0/me/drive/items/${driveItemId}`;
  const { res: metaRes } = await graphGet(metaUrl, token);
  if (!metaRes.ok) throw new Error(`Workbook metadata request failed: ${metaRes.status} ${await metaRes.text()}`);
  const meta = await metaRes.json();
  console.log('\n--- Workbook Metadata ---');
  console.log('Filename:', meta.name);
  console.log('DriveItem ID:', meta.id);
  console.log('Web URL:', meta.webUrl);

  const base = `https://graph.microsoft.com/v1.0/me/drive/items/${driveItemId}/workbook`;

  // 3. Verify worksheet & table existence
  const { res: wsRes } = await graphGet(`${base}/worksheets`, token);
  if (!wsRes.ok) throw new Error(`Worksheets request failed: ${wsRes.status} ${await wsRes.text()}`);
  const wsData = await wsRes.json();
  const ws = wsData.value?.find(w => w.name === 'Candidates');
  if (!ws) throw new Error('Worksheet "Candidates" not found');
  console.log('\nWorksheet name:', ws.name);

  const { res: tblRes } = await graphGet(`${base}/tables`, token);
  if (!tblRes.ok) throw new Error(`Tables request failed: ${tblRes.status} ${await tblRes.text()}`);
  const tblData = await tblRes.json();
  const tbl = tblData.value?.find(t => t.name === 'CandidateTable');
  if (!tbl) throw new Error('Table "CandidateTable" not found');
  console.log('Table name:', tbl.name);

  // 4. Fetch column headers (should be 17)
  const { res: colRes } = await graphGet(`${base}/tables/CandidateTable/columns`, token);
  if (!colRes.ok) throw new Error(`Columns request failed: ${colRes.status} ${await colRes.text()}`);
  const colData = await colRes.json();
  const colNames = colData.value?.map(c => c.name) || [];
  console.log('\nColumn headers (count', colNames.length + '):');
  console.log(colNames.join(', '));

  // 5. Fetch all rows
  const { res: rowsRes } = await graphGet(`${base}/tables/CandidateTable/rows`, token);
  if (!rowsRes.ok) throw new Error(`Rows request failed: ${rowsRes.status} ${await rowsRes.text()}`);
  const rowsData = await rowsRes.json();
  const rows = rowsData.value || [];
  console.log('\nTotal rows in table:', rows.length);

  // 6. Duplicate Candidate ID detection
  const idCounts = {};
  rows.forEach(r => {
    const cid = r.values?.[0]?.[0];
    if (cid) idCounts[cid] = (idCounts[cid] || 0) + 1;
  });
  const duplicates = Object.entries(idCounts).filter(([, cnt]) => cnt > 1);
  if (duplicates.length) {
    console.warn('⚠️ Duplicate Candidate IDs found:');
    duplicates.forEach(([id, cnt]) => console.warn(`- ${id}: ${cnt} times`));
  } else {
    console.log('No duplicate Candidate IDs detected.');
  }

  // 7. Compare each row against PostgreSQL record
  const candidates = await prisma.candidate.findMany();
  const candMap = {};
  candidates.forEach(c => (candMap[c.id] = c));

  const expectedHeaders = [
    'id', 'name', 'email', 'phone', 'location', 'appliedRole', 'degree', 'college',
    'graduationYear', 'experienceYears', 'skills', 'confidenceScore', 'reviewStatus',
    'applicationStatus', 'sourceEmailId', 'receivedAt', 'updatedAt'
  ];

  console.log('\n--- Row vs DB comparison ---');
  let mismatched = 0;
  rows.forEach(row => {
    const values = row.values?.[0] || [];
    const rowId = values[0];
    const isSynthetic = values.some(v => typeof v === 'string' && v.includes('[DEV-SYNTHETIC]'));
    const label = isSynthetic ? '[DEV‑SYNTHETIC]' : '';
    const dbCand = candMap[rowId];
    if (!dbCand) {
      console.warn(`Row with Candidate ID ${rowId} ${label} has no matching DB record`);
      mismatched++;
      return;
    }
    // Build expected array from DB using the same ordering as upsertCandidateRow
    const expected = [
      dbCand.id ?? '',
      dbCand.name ?? '',
      dbCand.email ?? '',
      dbCand.phone ?? '',
      '', // location placeholder
      dbCand.appliedRole ?? '',
      dbCand.degree ?? '',
      dbCand.college ?? '',
      dbCand.graduationYear?.toString() ?? '',
      dbCand.experienceYears?.toString() ?? '',
      dbCand.skills?.join(', ') ?? '',
      '', // confidenceScore placeholder
      '', // reviewStatus placeholder
      '', // applicationStatus placeholder
      dbCand.sourceEmailId ?? '',
      dbCand.receivedAt?.toISOString() ?? '',
      dbCand.updatedAt?.toISOString() ?? '',
    ];
    const diffIndices = [];
    expected.forEach((exp, i) => {
      const actual = values[i] ?? '';
      if (String(exp) !== String(actual)) diffIndices.push(i);
    });
    if (diffIndices.length) {
      mismatched++;
      const diffCols = diffIndices.map(i => `${expectedHeaders[i]} (DB="${expected[i]}", Excel="${values[i]}")`).join('; ');
      console.warn(`Mismatch for Candidate ID ${rowId} ${label}: ${diffCols}`);
    } else {
      console.log(`Candidate ID ${rowId} ${label}: ✅ match`);
    }
  });

  console.log('\nComparison summary:');
  console.log('Total rows examined:', rows.length);
  console.log('Mismatched rows:', mismatched);

  // 8. Log HTTP status codes used above
  console.log('\n--- Graph HTTP status codes ---');
  console.log('Workbook metadata:', metaRes.status);
  console.log('Worksheets list:', wsRes.status);
  console.log('Tables list:', tblRes.status);
  console.log('Columns fetch:', colRes.status);
  console.log('Rows fetch:', rowsRes.status);
}

main()
  .catch(err => {
    console.error('❌ Verification error:', err);
    process.exit(1);
  })
  .finally(() => prisma.$disconnect());

