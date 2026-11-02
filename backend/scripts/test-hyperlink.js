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
  const testId = '5082520FEDE864C4!s87dce66d044a4c47ba66111c3153659b';
  
  const colsRes = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${testId}/workbook/tables/Table1/columns`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  const colsData = await colsRes.json();
  const numCols = colsData.value.length;
  console.log('numCols in Table1:', numCols);
  const rowValues = new Array(numCols).fill('test');
  rowValues[rowValues.length - 1] = '=HYPERLINK("https://example.com/test", "Download CV")';

  const addRes = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${testId}/workbook/tables/Table1/rows/add`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      values: [rowValues]
    })
  });
  console.log('Add row status:', addRes.status);
  const addData = await addRes.json();
  console.log('Add row result:', addData);
  await prisma.$disconnect();
}

main().catch(console.error);
