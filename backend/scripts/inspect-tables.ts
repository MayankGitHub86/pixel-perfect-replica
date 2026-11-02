#!/usr/bin/env node
// backend/scripts/inspect-tables.ts
// Diagnostic: list all tables and their columns in the workbook.

import crypto from 'crypto';
import path from 'path';
import { PrismaClient } from '@prisma/client';
import dotenv from 'dotenv';

dotenv.config({ path: path.resolve(process.cwd(), '.env') });

const prisma = new PrismaClient();

function decryptToken(payload: string): string {
  const [ivB64, encryptedB64, tagB64] = payload.split('.');
  const iv = Buffer.from(ivB64, 'base64');
  const encrypted = Buffer.from(encryptedB64, 'base64');
  const tag = Buffer.from(tagB64, 'base64');
  const key = Buffer.from(process.env.TOKEN_ENCRYPTION_KEY!, 'base64');
  const decipher = crypto.createDecipheriv('aes-256-gcm', key, iv, { authTagLength: 16 });
  decipher.setAuthTag(tag);
  return Buffer.concat([decipher.update(encrypted), decipher.final()]).toString('utf8');
}

async function getValidAccessToken(): Promise<string> {
  const conn = await prisma.mailboxConnection.findFirst({ orderBy: { createdAt: 'desc' } });
  if (!conn) throw new Error('No mailbox connection found');
  const decryptedAccess = decryptToken(conn.accessToken);
  const [, payloadB64] = decryptedAccess.split('.');
  let payload: any = {};
  try { payload = JSON.parse(Buffer.from(payloadB64, 'base64url').toString('utf8')); } catch { payload = {}; }
  const expiresAt = payload.exp ? payload.exp * 1000 : (conn.tokenExpiresAt as any)?.getTime?.();
  const now = Date.now();
  if (!expiresAt || expiresAt - now < 5 * 60 * 1000) {
    const refreshToken = decryptToken(conn.refreshToken);
    const tokenRes = await fetch(
      `https://login.microsoftonline.com/${process.env.MICROSOFT_TENANT_ID}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id: process.env.MICROSOFT_CLIENT_ID!,
          client_secret: process.env.MICROSOFT_CLIENT_SECRET!,
          grant_type: 'refresh_token',
          refresh_token: refreshToken,
          redirect_uri: process.env.MICROSOFT_REDIRECT_URI || 'http://localhost:3000/api/integrations/microsoft/callback',
        }).toString(),
      }
    );
    if (!tokenRes.ok) throw new Error(`Token refresh failed: ${tokenRes.status}`);
    const tokenData: any = await tokenRes.json();
    return tokenData.access_token;
  }
  return decryptedAccess;
}

async function main() {
  const token = await getValidAccessToken();
  console.log('Token acquired');

  const wb = await prisma.excelWorkbook.findFirst();
  if (!wb) throw new Error('Excel workbook record not found');
  const workbookId = wb.driveItemId;
  const base = `https://graph.microsoft.com/v1.0/me/drive/items/${workbookId}/workbook`;

  // List worksheets
  const wsRes = await fetch(`${base}/worksheets`, { headers: { Authorization: `Bearer ${token}` } });
  const wsData: any = await wsRes.json();
  console.log('\n=== Worksheets ===');
  for (const ws of (wsData.value || [])) {
    console.log(`  - ${ws.name} (id: ${ws.id}, visibility: ${ws.visibility})`);
  }

  // List tables
  const tblRes = await fetch(`${base}/tables`, { headers: { Authorization: `Bearer ${token}` } });
  const tblData: any = await tblRes.json();
  console.log('\n=== Tables ===');
  const tables = tblData.value || [];
  for (const t of tables) {
    console.log(`  - ${t.name} (id: ${t.id})`);
    // Get columns
    const colRes = await fetch(`${base}/tables/${t.id}/columns`, { headers: { Authorization: `Bearer ${token}` } });
    const colData: any = await colRes.json();
    const cols = (colData.value || []).map((c: any) => c.name);
    console.log(`    Columns: ${JSON.stringify(cols)}`);
    // Get row count
    const rowRes = await fetch(`${base}/tables/${t.id}/rows`, { headers: { Authorization: `Bearer ${token}` } });
    const rowData: any = await rowRes.json();
    const rows = rowData.value || [];
    console.log(`    Row count: ${rows.length}`);
    // Show first row values
    if (rows.length > 0) {
      console.log(`    First row values: ${JSON.stringify(rows[0].values?.[0])}`);
    }
  }

  // Also check for named ranges or tables within each worksheet
  for (const ws of (wsData.value || [])) {
    const wsTblRes = await fetch(`${base}/worksheets/${ws.id}/tables`, { headers: { Authorization: `Bearer ${token}` } });
    const wsTblData: any = await wsTblRes.json();
    const wsTables = wsTblData.value || [];
    if (wsTables.length > 0) {
      console.log(`\n  Tables in worksheet "${ws.name}":`);
      for (const t of wsTables) {
        console.log(`    - ${t.name}`);
      }
    }
  }

  await prisma.$disconnect();
  console.log('\nDone.');
}

main().catch(err => {
  console.error('Error:', err);
  process.exit(1);
});

