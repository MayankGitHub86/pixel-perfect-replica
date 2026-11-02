// backend/scripts/check-existing-workbooks.js
// Read‑only script: list OneDrive root and fetch metadata for known workbook IDs.
// Uses the stored access token directly (no refresh) to avoid placeholder client errors.

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

async function main() {
  const conn = await prisma.mailboxConnection.findFirst({ orderBy: { createdAt: 'desc' } });
  if (!conn) throw new Error('No mailbox connection found');
  const token = decrypt(conn.accessToken); // use existing token (may be still valid)
  console.log('✅ Access token (raw) acquired');

  // 1. List root items
  const listRes = await fetch('https://graph.microsoft.com/v1.0/me/drive/root/children', {
    headers: { Authorization: `Bearer ${token}` },
  });
  console.log('Root list HTTP status:', listRes.status);
  if (!listRes.ok) { console.error('Error body:', await listRes.text()); return; }
  const listData = await listRes.json();
  console.log('Root items count:', listData.value?.length ?? 0);
  listData.value?.forEach(item => {
    console.log(`- ${item.name} (id: ${item.id})`);
  });

  // 2. Check metadata for the workbook stored in DB (Recovery_V2)
  const wbRecord = await prisma.excelWorkbook.findFirst();
  if (!wbRecord) { console.error('No ExcelWorkbook record in DB'); return; }
  console.log('\n--- DB ExcelWorkbook record ---');
  console.log('DriveItem ID:', wbRecord.driveItemId);
  console.log('Web URL:', wbRecord.webUrl);

  const metaRes = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${wbRecord.driveItemId}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  console.log('Metadata request status:', metaRes.status);
  if (!metaRes.ok) { console.error('Meta error body:', await metaRes.text()); return; }
  const meta = await metaRes.json();
  console.log('Metadata name:', meta.name);
  console.log('Metadata webUrl:', meta.webUrl);

  // 3. Check if Candidate_Master_RealData.xlsx already exists (search by name)
  const realName = 'Candidate_Master_RealData.xlsx';
  const realItem = listData.value?.find(i => i.name === realName);
  if (realItem) {
    console.log('\n✅ Real workbook already exists with ID:', realItem.id);
    // fetch its rows count
    const rowsRes = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${realItem.id}/workbook/tables/CandidateTable/rows`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    console.log('Rows request status for real workbook:', rowsRes.status);
    if (rowsRes.ok) {
      const rowsData = await rowsRes.json();
      console.log('Row count in real workbook:', rowsData.value?.length ?? 0);
    }
  } else {
    console.log('\n⚠️ Real workbook NOT present in OneDrive root');
  }
}

main()
  .catch(err => console.error('❌ Error:', err))
  .finally(() => prisma.$disconnect());

