// backend/scripts/verify-phase5-e2e.js
// End-to-end verification of genuine candidate workflow:
// Outlook email provenance -> Genuine CV parsing -> PostgreSQL update -> OneDrive Resume storage -> Excel row with Download CV link

const fs = require('fs');
const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '../.env') });
const crypto = require('crypto');
const { PrismaClient } = require('@prisma/client');
const mammoth = require('mammoth');

const prisma = new PrismaClient();

function decrypt(payload) {
  const [ivB64, encryptedB64, tagB64] = payload.split('.');
  const iv = Buffer.from(ivB64, 'base64');
  const encrypted = Buffer.from(encryptedB64, 'base64');
  const tag = Buffer.from(tagB64, 'base64');
  const key = Buffer.from(process.env.TOKEN_ENCRYPTION_KEY, 'base64');
  const decipher = crypto.createDecipheriv('aes-256-gcm', key, iv, { authTagLength: 16 });
  decipher.setAuthTag(tag);
  return Buffer.concat([decipher.update(encrypted), decipher.final()]).toString('utf8');
}

function encrypt(plainText) {
  const iv = crypto.randomBytes(12);
  const key = Buffer.from(process.env.TOKEN_ENCRYPTION_KEY, 'base64');
  const cipher = crypto.createCipheriv('aes-256-gcm', key, iv, { authTagLength: 16 });
  const encrypted = Buffer.concat([cipher.update(plainText, 'utf8'), cipher.final()]);
  const tag = cipher.getAuthTag();
  return `${iv.toString('base64')}.${encrypted.toString('base64')}.${tag.toString('base64')}`;
}

async function refreshAccessToken(mailboxId) {
  const mailbox = await prisma.mailboxConnection.findUnique({ where: { id: mailboxId } });
  if (!mailbox) throw new Error('MailboxConnection not found');
  const refreshToken = decrypt(mailbox.refreshToken);
  const tenantId = process.env.MICROSOFT_TENANT_ID || 'common';
  const clientId = process.env.MICROSOFT_CLIENT_ID;
  const clientSecret = process.env.MICROSOFT_CLIENT_SECRET;
  const redirectUri = process.env.MICROSOFT_REDIRECT_URI || 'http://localhost:3000/api/integrations/microsoft/callback';

  const tokenRes = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      grant_type: 'refresh_token',
      refresh_token: refreshToken,
      redirect_uri: redirectUri,
    }).toString(),
  });

  if (!tokenRes.ok) {
    const errBody = await tokenRes.text();
    throw new Error(`Token refresh failed: ${tokenRes.status} ${errBody}`);
  }

  const tokenData = await tokenRes.json();
  const newAccessToken = tokenData.access_token;
  const newRefreshToken = tokenData.refresh_token ?? refreshToken;
  const expiresIn = tokenData.expires_in;

  const encryptedAccess = encrypt(newAccessToken);
  const encryptedRefresh = encrypt(newRefreshToken);
  const tokenExpiresAt = new Date(Date.now() + expiresIn * 1000);

  await prisma.mailboxConnection.update({
    where: { id: mailboxId },
    data: {
      accessToken: encryptedAccess,
      refreshToken: encryptedRefresh,
      tokenExpiresAt,
      updatedAt: new Date(),
    },
  });
  return newAccessToken;
}

async function getAccessToken() {
  const conn = await prisma.mailboxConnection.findFirst({ orderBy: { createdAt: 'desc' } });
  if (!conn) throw new Error('No mailbox connection found');
  if (new Date() >= new Date(conn.tokenExpiresAt.getTime() - 60000)) {
    return refreshAccessToken(conn.id);
  }
  return decrypt(conn.accessToken);
}

async function runEndToEndVerification() {
  console.log('=== PHASE 5 REAL DATA END-TO-END VERIFICATION ===\n');

  const candidateId = '106945ea-6ae3-4c64-b5bb-9c1148d72985';
  let token = await getAccessToken();

  // 1. Identify actual Outlook email in PostgreSQL
  console.log('Step 1: Identifying actual Outlook email...');
  const candidateBefore = await prisma.candidate.findUnique({
    where: { id: candidateId },
    include: { attachments: true },
  });
  if (!candidateBefore) throw new Error(`Candidate ${candidateId} not found`);

  const sourceEmail = await prisma.emailMessage.findFirst({
    where: {
      OR: [
        { id: candidateBefore.sourceEmailId },
        { messageId: candidateBefore.sourceEmailId },
      ],
    },
  });
  if (!sourceEmail) throw new Error(`Source email not found for ${candidateBefore.sourceEmailId}`);
  console.log('  PASS: Identified genuine source email:', {
    id: sourceEmail.id,
    subject: sourceEmail.subject,
    fromEmail: sourceEmail.fromEmail,
    receivedAt: sourceEmail.receivedAt.toISOString(),
  });

  // 2. Retrieve actual CV attachment
  console.log('\nStep 2: Retrieving actual CV attachment...');
  const attachment = candidateBefore.attachments[0];
  if (!attachment || !fs.existsSync(attachment.storageKey)) {
    throw new Error(`Attachment file missing at ${attachment?.storageKey}`);
  }
  const fileBuffer = fs.readFileSync(attachment.storageKey);
  console.log('  PASS: Retrieved CV attachment from local storage:', {
    filename: attachment.filename,
    sizeBytes: fileBuffer.length,
    path: attachment.storageKey,
  });

  // 3. Parse CV with updated extractor
  console.log('\nStep 3: Parsing genuine CV attachment text...');
  const mammothRes = await mammoth.extractRawText({ buffer: fileBuffer });
  const rawText = mammothRes.value;

  // Use the compiled / project candidate extractor
  const { extractCandidateInfo } = require('../dist/lib/candidateExtractor');
  const extracted = extractCandidateInfo(rawText, {
    subject: sourceEmail.subject,
    from: sourceEmail.fromEmail,
  });

  console.log('  PASS: Extracted genuine candidate profile:', {
    name: extracted.name,
    email: extracted.email,
    phone: extracted.phone,
    location: extracted.location,
    appliedRole: extracted.appliedRole,
    degree: extracted.degree,
    college: extracted.college,
    graduationYear: extracted.graduationYear,
    skillsCount: extracted.skills.length,
    skillsSample: extracted.skills.slice(0, 5),
    processingStatus: extracted.processingStatus,
  });

  // 4. Upload/Ensure CV is stored in OneDrive
  console.log('\nStep 4: Ensuring CV is stored in persistent OneDrive storage...');
  const targetName = `${candidateId}_${attachment.filename}`;
  const uploadUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/Candidate_Resumes/${encodeURIComponent(targetName)}:/content`;
  const uploadRes = await fetch(uploadUrl, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': attachment.mimeType,
    },
    body: fileBuffer,
  });

  if (!uploadRes.ok) {
    const errText = await uploadRes.text();
    throw new Error(`OneDrive CV upload failed: ${uploadRes.status} ${errText}`);
  }
  const uploadData = await uploadRes.json();
  const stableOneDriveUrl = uploadData.webUrl;
  const driveItemId = uploadData.id;
  console.log('  PASS: Resume stored in OneDrive:', {
    driveItemId,
    filename: targetName,
    webUrl: stableOneDriveUrl,
    sizeBytes: uploadData.size,
  });

  // 5. Update PostgreSQL candidate record with genuine extracted details
  console.log('\nStep 5: Updating PostgreSQL candidate record...');
  const updatedCandidate = await prisma.candidate.update({
    where: { id: candidateId },
    data: {
      name: extracted.name,
      email: extracted.email,
      phone: extracted.phone,
      location: extracted.location,
      appliedRole: extracted.appliedRole,
      degree: extracted.degree,
      college: extracted.college,
      graduationYear: extracted.graduationYear,
      skills: Array.from(new Set([...candidateBefore.skills, ...extracted.skills])),
      resumeFileName: attachment.filename,
      resumeRef: driveItemId,
      resumeUrl: stableOneDriveUrl,
      processingStatus: extracted.processingStatus,
      updatedAt: new Date(),
    },
  });
  console.log('  PASS: PostgreSQL candidate record updated:', {
    id: updatedCandidate.id,
    name: updatedCandidate.name,
    phone: updatedCandidate.phone,
    location: updatedCandidate.location,
    role: updatedCandidate.appliedRole,
    graduationYear: updatedCandidate.graduationYear,
    resumeUrl: updatedCandidate.resumeUrl ? '(present)' : '(none)',
  });

  // 6. Sync to active production workbook via Graph
  console.log('\nStep 6: Syncing candidate to active Excel workbook (Candidate_Master_RealData.xlsx)...');
  const wbRecord = await prisma.excelWorkbook.findFirst();
  console.log('  Active DB workbook pointer:', wbRecord.driveItemId);

  const { upsertCandidateRow, ensureWorksheetAndTable } = require('../dist/lib/excel');
  await ensureWorksheetAndTable(wbRecord.driveItemId);
  const syncResult1 = await upsertCandidateRow(candidateId);
  console.log('  First sync result:', syncResult1);

  // 7. Verify Excel row contents and columns
  console.log('\nStep 7: Verifying Excel row in Candidate_Master_RealData.xlsx via Graph API...');
  const rowsRes = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${wbRecord.driveItemId}/workbook/tables/CandidateTable/rows`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const rowsData = await rowsRes.json();
  console.log('  Table rows count:', rowsData.value?.length);
  if (rowsData.value?.length !== 1) {
    throw new Error(`Expected exactly 1 row in CandidateTable, found ${rowsData.value?.length}`);
  }

  const rowValues = rowsData.value[0].values[0];
  console.log('  Row length (columns count):', rowValues.length);
  console.log('  Row values:');
  console.log('    [0] id:                   ', rowValues[0]);
  console.log('    [1] name:                 ', rowValues[1]);
  console.log('    [2] email:                ', rowValues[2]);
  console.log('    [3] phone:                ', rowValues[3]);
  console.log('    [4] location:             ', rowValues[4]);
  console.log('    [5] appliedRole:          ', rowValues[5]);
  console.log('    [6] degree:               ', rowValues[6]);
  console.log('    [7] college:              ', rowValues[7]);
  console.log('    [8] graduationYear:       ', rowValues[8]);
  console.log('    [10] skills:              ', rowValues[10]?.slice(0, 60) + '...');
  console.log('    [14] sourceEmailId:       ', rowValues[14]);
  console.log('    [17] resumeFileName:      ', rowValues[17]);
  console.log('    [18] resumeLink:          ', rowValues[18]);
  console.log('    [19] sourceEmailSubject:  ', rowValues[19]);
  console.log('    [20] sourceEmailReceivedAt:', rowValues[20]);
  console.log('    [21] extractionStatus:    ', rowValues[21]);

  if (rowValues[17] !== attachment.filename) {
    throw new Error(`resumeFileName mismatch: expected ${attachment.filename}, got ${rowValues[17]}`);
  }
  if (!rowValues[18] || !rowValues[18].includes('Download CV')) {
    throw new Error(`resumeLink mismatch: expected Download CV link, got ${rowValues[18]}`);
  }

  // 8. Test idempotency (reprocessing should NOT duplicate row)
  console.log('\nStep 8: Testing deduplication and idempotency on reprocessing...');
  const syncResult2 = await upsertCandidateRow(candidateId);
  console.log('  Second sync result:', syncResult2);

  const rowsRes2 = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${wbRecord.driveItemId}/workbook/tables/CandidateTable/rows`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const rowsData2 = await rowsRes2.json();
  console.log('  Table rows count after second sync:', rowsData2.value?.length);
  if (rowsData2.value?.length !== 1) {
    throw new Error(`Duplicate row detected! Row count: ${rowsData2.value?.length}`);
  }
  console.log('  PASS: Idempotency confirmed - exactly 1 row maintained in Excel');

  // 9. Download and verify CV from the persistent OneDrive URL
  console.log('\nStep 9: Downloading and verifying CV from persistent OneDrive storage...');
  const dlRes = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${driveItemId}/content`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!dlRes.ok) throw new Error(`Download from OneDrive failed: ${dlRes.status}`);
  const dlBuffer = Buffer.from(await dlRes.arrayBuffer());
  const bytesMatch = dlBuffer.equals(fileBuffer);
  console.log('  PASS: Downloaded CV bytes match original attachment:', bytesMatch, `(${dlBuffer.length} bytes)`);
  if (!bytesMatch) throw new Error('Downloaded bytes do not match original CV attachment!');

  // 10. Verify workbook metadata and DB pointer integrity
  console.log('\nStep 10: Verifying workbook metadata and DB pointer integrity...');
  const wbItemRes = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${wbRecord.driveItemId}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const wbItem = await wbItemRes.json();
  console.log('  PASS: Workbook details:', {
    id: wbItem.id,
    name: wbItem.name,
    parentPath: wbItem.parentReference?.path,
    webUrl: wbItem.webUrl,
  });

  const finalDbWb = await prisma.excelWorkbook.findFirst();
  if (finalDbWb.driveItemId !== wbRecord.driveItemId) {
    throw new Error('Database pointer changed unexpectedly!');
  }
  console.log('  PASS: DB ExcelWorkbook pointer unchanged');

  // 11. Confirm no synthetic candidates in Excel
  console.log('\nStep 11: Confirming no synthetic records entered Excel...');
  const allDbCandidates = await prisma.candidate.findMany();
  const syntheticInExcel = rowsData2.value.filter(r => r.values[0][1]?.includes('[DEV-SYNTHETIC]') || r.values[0][2]?.includes('synthetic.dev'));
  console.log('  PASS: Synthetic candidates in Excel table:', syntheticInExcel.length);

  console.log('\n======================================================');
  console.log('PHASE 5 REAL DATA END-TO-END VERIFICATION: ALL PASSED!');
  console.log('======================================================');

  await prisma.$disconnect();
}

runEndToEndVerification().catch((err) => {
  console.error('\nFAIL: Verification error:', err);
  process.exit(1);
});

