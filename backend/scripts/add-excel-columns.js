const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '../.env') });
const crypto = require('crypto');
const { PrismaClient } = require('@prisma/client');
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

async function main() {
  const conn = await prisma.mailboxConnection.findFirst({ orderBy: { createdAt: 'desc' } });
  const token = decrypt(conn.accessToken);
  const wb = await prisma.excelWorkbook.findFirst();
  const realId = wb.driveItemId;
  const base = `https://graph.microsoft.com/v1.0/me/drive/items/${realId}/workbook`;

  const colsRes = await fetch(`${base}/tables/CandidateTable/columns`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  const colsData = await colsRes.json();
  const existingNames = new Set(colsData.value?.map(c => c.name));
  console.log('Existing columns count:', existingNames.size);
  console.log('Existing columns:', Array.from(existingNames));

  const targetColumns = [
    'resumeFileName',
    'resumeLink',
    'sourceEmailSubject',
    'sourceEmailReceivedAt',
    'extractionStatus',
  ];

  for (const colName of targetColumns) {
    if (!existingNames.has(colName)) {
      console.log(`Adding column: ${colName}`);
      const addRes = await fetch(`${base}/tables/CandidateTable/columns/add`, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ name: colName })
      });
      console.log(`Add ${colName} status:`, addRes.status);
    } else {
      console.log(`Column ${colName} already exists`);
    }
  }

  const finalColsRes = await fetch(`${base}/tables/CandidateTable/columns`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  const finalColsData = await finalColsRes.json();
  console.log('Final columns count:', finalColsData.value?.length);
  console.log('Final columns:', finalColsData.value?.map(c => c.name));

  await prisma.$disconnect();
}

main().catch(console.error);

