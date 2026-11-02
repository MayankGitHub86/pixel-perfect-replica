// backend/scripts/create-real-workbook.js
// Creates a clean OneDrive workbook containing only the verified genuine candidate.
// No existing workbooks are modified; the new file is uploaded as Candidate_Master_RealData.xlsx.

const fs = require('fs');
const path = require('path');
require('dotenv').config();
const { PrismaClient } = require('@prisma/client');
const crypto = require('crypto');
const ExcelJS = require('exceljs');

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
    const { refreshAccessToken } = require('../dist/lib/tokenRefresh');
    await refreshAccessToken(conn.id);
    const refreshed = await prisma.mailboxConnection.findUnique({ where: { id: conn.id } });
    if (!refreshed) throw new Error('Failed to refresh token');
    return decrypt(refreshed.accessToken);
  }
  return decrypt(conn.accessToken);
}

// Build the exact 17‑column ordering used by the Excel service
const columnHeaders = [
  'id', 'name', 'email', 'phone', 'location', 'appliedRole', 'degree', 'college',
  'graduationYear', 'experienceYears', 'skills', 'confidenceScore', 'reviewStatus',
  'applicationStatus', 'sourceEmailId', 'receivedAt', 'updatedAt',
];

async function buildWorkbook(candidate) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Candidates');

  // Build data row matching the upsert logic (empty strings for placeholders)
  const row = [
    candidate.id ?? '',
    candidate.name ?? '',
    candidate.email ?? '',
    candidate.phone ?? '',
    '', // location placeholder
    candidate.appliedRole ?? '',
    candidate.degree ?? '',
    candidate.college ?? '',
    candidate.graduationYear?.toString() ?? '',
    candidate.experienceYears?.toString() ?? '',
    candidate.skills?.join(', ') ?? '',
    '', // confidenceScore placeholder
    '', // reviewStatus placeholder
    '', // applicationStatus placeholder
    candidate.sourceEmailId ?? '',
    candidate.receivedAt?.toISOString() ?? '',
    candidate.updatedAt?.toISOString() ?? '',
  ];

  // Add table covering headers and rows
  ws.addTable({
    name: 'CandidateTable',
    ref: 'A1',
    headerRow: true,
    totalsRow: false,
    style: { theme: 'TableStyleMedium9' },
    columns: columnHeaders.map(h => ({ name: h })),
    rows: [row],
  });

  return wb;
}

async function main() {
  const token = await getAccessToken();
  console.log('✅ Access token acquired');

  // 1. Fetch the verified genuine candidate by exact ID and verify Outlook provenance
  const candidate = await prisma.candidate.findUnique({
    where: { id: '106945ea-6ae3-4c64-b5bb-9c1148d72985' },
  });
  if (!candidate) throw new Error('Verified genuine candidate not found');
  if (candidate.sourceEmailId !== 'ca1341f7-7987-4cea-a3db-3ce2347425b8') {
    throw new Error('Candidate does not have the expected sourceEmailId – abort');
  }
  console.log('📄 Verified candidate:', candidate.id, candidate.name);


  // 2. Build workbook in memory
  const workbook = await buildWorkbook(candidate);
  const localPath = path.join(__dirname, '../Candidate_Master_RealData.xlsx');
  // Prevent local overwrite
  if (fs.existsSync(localPath)) {
    throw new Error('Local workbook file already exists – abort to avoid overwrite');
  }
  await workbook.xlsx.writeFile(localPath);
  console.log('🗂 Workbook written to', localPath);


  // 3. Upload to OneDrive (no-overwrite: preflight GET + conditional PUT)
  const fileBuffer = fs.readFileSync(localPath);
  const { uploadToOneDrive } = require('./lib/upload-to-onedrive');
  const meta = await uploadToOneDrive({
    token,
    fileBuffer,
    fileName: 'Candidate_Master_RealData.xlsx',
  });
  console.log('📤 Upload successful – metadata:');
  console.log('  DriveItem ID:', meta.id);
  console.log('  Name:', meta.name);
  console.log('  Web URL:', meta.webUrl);
  console.log('  Size (bytes):', meta.size);

  const driveItemId = meta.id;
  const base = `https://graph.microsoft.com/v1.0/me/drive/items/${driveItemId}/workbook`;

  // 4. Verify worksheet and table exist (they should, as they were created in the file)
  const wsRes = await fetch(`${base}/worksheets`, { headers: { Authorization: `Bearer ${token}` } });
  if (!wsRes.ok) throw new Error(`Worksheets request failed: ${wsRes.status}`);
  const wsData = await wsRes.json();
  const ws = wsData.value?.find(w => w.name === 'Candidates');
  console.log('✅ Worksheet found:', ws?.name);

  const tblRes = await fetch(`${base}/tables`, { headers: { Authorization: `Bearer ${token}` } });
  if (!tblRes.ok) throw new Error(`Tables request failed: ${tblRes.status}`);
  const tblData = await tblRes.json();
  const tbl = tblData.value?.find(t => t.name === 'CandidateTable');
  console.log('✅ Table found:', tbl?.name);

  // 5. Verify columns (should be 17)
  const colRes = await fetch(`${base}/tables/CandidateTable/columns`, { headers: { Authorization: `Bearer ${token}` } });
  if (!colRes.ok) throw new Error(`Columns request failed: ${colRes.status}`);
  const colData = await colRes.json();
  const cols = colData.value?.map(c => c.name) || [];
  console.log('🔣 Columns count:', cols.length);
  console.log('Columns:', cols.join(', '));

  // 6. Read back rows
  const rowsRes = await fetch(`${base}/tables/CandidateTable/rows`, { headers: { Authorization: `Bearer ${token}` } });
  if (!rowsRes.ok) throw new Error(`Rows request failed: ${rowsRes.status}`);
  const rowsData = await rowsRes.json();
  const rows = rowsData.value || [];
  console.log('📊 Row count in table:', rows.length);
  if (rows.length !== 1) throw new Error('Unexpected row count – expected exactly 1 row');

  const values = rows[0].values?.[0] || [];
  console.log('🔎 Row values:', values);

  // 7. Compare to DB candidate
  const expected = [
    candidate.id ?? '',
    candidate.name ?? '',
    candidate.email ?? '',
    candidate.phone ?? '',
    '',
    candidate.appliedRole ?? '',
    candidate.degree ?? '',
    candidate.college ?? '',
    candidate.graduationYear?.toString() ?? '',
    candidate.experienceYears?.toString() ?? '',
    candidate.skills?.join(', ') ?? '',
    '',
    '',
    '',
    candidate.sourceEmailId ?? '',
    candidate.receivedAt?.toISOString() ?? '',
    candidate.updatedAt?.toISOString() ?? '',
  ];

  const diffs = [];
  expected.forEach((exp, i) => {
    const act = values[i] ?? '';
    if (String(exp) !== String(act)) diffs.push({ index: i, header: columnHeaders[i], expected: exp, actual: act });
  });
  if (diffs.length) {
    console.error('❌ Mismatches found:');
    diffs.forEach(d => console.error(`  Column ${d.header}: DB="${d.expected}"  Excel="${d.actual}"`));
    throw new Error('Row values do not match DB');
  }
  console.log('✅ Row matches DB exactly');

  // 8. Update ExcelWorkbook DB record atomically
  const { updateWorkbookPointer } = require('./lib/update-workbook-pointer');
  await prisma.$transaction(async (tx) => {
    await updateWorkbookPointer(tx, { driveItemId, webUrl: meta.webUrl });
  });
  console.log('🗄 ExcelWorkbook DB record updated to new workbook');

  // 9. Read back DB record for verification
  const wbRecord = await prisma.excelWorkbook.findFirst();
  console.log('🔎 DB ExcelWorkbook record:', { id: wbRecord.id, driveItemId: wbRecord.driveItemId, webUrl: wbRecord.webUrl });
}

main()
  .catch(err => {
    console.error('❌ Error:', err);
    process.exit(1);
  })
  .finally(() => prisma.$disconnect());

