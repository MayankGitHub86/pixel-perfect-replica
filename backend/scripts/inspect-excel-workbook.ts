// inspect-excel-workbook.ts
// Script to inspect the Candidate_Master_RealData.xlsx workbook via Microsoft Graph
// Reports main worksheet, table name, existing headers, used range, and row count.

import path from 'path';
import 'dotenv/config';
import crypto from 'crypto';
import { PrismaClient } from '@prisma/client';

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

async function main() {
  const conn = await prisma.mailboxConnection.findFirst({ orderBy: { createdAt: 'desc' } });
  if (!conn) throw new Error('No mailbox connection found');
  const token = decrypt(conn.accessToken);
  const wbRecord = await prisma.excelWorkbook.findFirst();
  if (!wbRecord) throw new Error('ExcelWorkbook record not found');
  const realId = wbRecord.driveItemId;
  const base = `https://graph.microsoft.com/v1.0/me/drive/items/${realId}/workbook`;

  // Get worksheets
  const wsRes = await fetch(`${base}/worksheets`, { headers: { Authorization: `Bearer ${token}` } });
  const wsData = await wsRes.json();
  const worksheets = wsData.value || [];
  console.log('Worksheets:', worksheets.map((w: any) => w.name));

  // Get tables
  const tablesRes = await fetch(`${base}/tables`, { headers: { Authorization: `Bearer ${token}` } });
  const tablesData = await tablesRes.json();
  const tables = tablesData.value || [];
  console.log('Tables:', tables.map((t: any) => t.name));

  const candidateTable = tables.find((t: any) => t.name === 'CandidateTable');
  if (!candidateTable) {
    console.log('CandidateTable not found');
    return;
  }

  // Get columns
  const colsRes = await fetch(`${base}/tables/${candidateTable.id}/columns`, { headers: { Authorization: `Bearer ${token}` } });
  const colsData = await colsRes.json();
  const columnNames = (colsData.value || []).map((c: any) => c.name);
  console.log('Existing columns:', columnNames);

  // Get used range and row count
  const rangeRes = await fetch(`${base}/tables/${candidateTable.id}/range`, { headers: { Authorization: `Bearer ${token}` } });
  const rangeData = await rangeRes.json();
  console.log('Used range address:', rangeData.address);
  console.log('Row count (including header):', rangeData.rowCount);

  // Sample rows
  const rowsRes = await fetch(`${base}/tables/${candidateTable.id}/rows?$top=5`, { headers: { Authorization: `Bearer ${token}` } });
  const rowsData = await rowsRes.json();
  console.log('Sample rows count:', rowsData.value?.length || 0);
}

main().catch(err => {
  console.error('Error:', err.message);
  process.exit(1);
});

