// list-workbook.ts
// Script to list worksheets and tables in the real workbook via Graph
import path from 'path';
import 'dotenv/config';
import crypto from 'crypto';
import { PrismaClient } from '@prisma/client';

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
  if (!conn) throw new Error('No mailbox connection found');
  const token = decrypt(conn.accessToken);
  const wb = await prisma.excelWorkbook.findFirst();
  if (!wb) throw new Error('ExcelWorkbook record not found');
  const realId = wb.driveItemId;
  const base = `https://graph.microsoft.com/v1.0/me/drive/items/${realId}/workbook`;
  // worksheets
  const wsRes = await fetch(`${base}/worksheets`, { headers: { Authorization: `Bearer ${token}` } });
  const wsData = await wsRes.json();
  console.log('Worksheets:', (wsData.value || []).map(w => w.name));
  // tables
  const tblRes = await fetch(`${base}/tables`, { headers: { Authorization: `Bearer ${token}` } });
  const tblData = await tblRes.json();
  console.log('Tables:', (tblData.value || []).map(t => ({ name: t.name, id: t.id })));
}

main().catch(err => { console.error('Error:', err.message); process.exit(1); });

