import path from 'path';
import 'dotenv/config';
import crypto from 'crypto';
import { PrismaClient } from '@prisma/client';
import { getValidAccessToken } from '../src/lib/graphAuth.ts';

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

async function fetchAndLog(url: string, token: string) {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const body = await res.text();
  let parsed: any;
  try { parsed = JSON.parse(body); } catch (_) { parsed = body; }
  console.log('\n=== GET', url);
  console.log('Status:', res.status);
  console.log('Response body:', parsed);
  return { status: res.status, body: parsed };
}

async function main() {
  // Use the new helper to ensure a fresh token.
  const token = await getValidAccessToken();
  const wb = await prisma.excelWorkbook.findFirst();
  if (!wb) throw new Error('ExcelWorkbook record not found');
  const realId = wb.driveItemId; // should match the ID given
  console.log('Using DriveItem ID:', realId);

  const baseItem = `https://graph.microsoft.com/v1.0/me/drive/items/${realId}`;
  // 1. DriveItem metadata
  await fetchAndLog(baseItem, token);

  // 2. Workbook root
  await fetchAndLog(`${baseItem}/workbook`, token);

  // 3. Worksheets
  await fetchAndLog(`${baseItem}/workbook/worksheets`, token);

  // 4. Tables
  await fetchAndLog(`${baseItem}/workbook/tables`, token);

  // 5. Download content (downloadUrl)
  const metaRes = await fetch(`${baseItem}?select=@microsoft.graph.downloadUrl`, { headers: { Authorization: `Bearer ${token}` } });
  const meta = await metaRes.json();
  console.log('\nDownload URL:', meta['@microsoft.graph.downloadUrl'] || 'none');
  if (meta['@microsoft.graph.downloadUrl']) {
    const dlRes = await fetch(meta['@microsoft.graph.downloadUrl']);
    console.log('Download response status:', dlRes.status);
    const buffer = await dlRes.arrayBuffer();
    const header = Buffer.from(buffer.slice(0, 4)).toString('hex');
    console.log('File signature (first 4 bytes hex):', header);
  }
}

main().catch(err => { console.error('Error:', err.message); process.exit(1); });


