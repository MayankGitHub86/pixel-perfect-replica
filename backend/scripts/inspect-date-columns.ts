import { PrismaClient } from '@prisma/client';
import dotenv from 'dotenv';
import path from 'path';
import crypto from 'crypto';

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
  const wb = await prisma.excelWorkbook.findFirst();
  if (!wb) throw new Error('No workbook found');
  const base = `https://graph.microsoft.com/v1.0/me/drive/items/${wb.driveItemId}/workbook`;

  // Get CandidateTable
  const tblRes = await fetch(`${base}/tables/CandidateTable`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const tbl = await tblRes.json();
  console.log('Table:', { id: tbl.id, name: tbl.name });

  // Get columns
  const colsRes = await fetch(`${base}/tables/CandidateTable/columns`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const cols = await colsRes.json();
  const colNames = (cols.value || []).map((c: any) => ({ id: c.id, name: c.name, index: c.index }));
  console.log('Columns:', colNames);

  // Check receivedAt column
  const receivedAtCol = colNames.find((c: any) => c.name === 'receivedAt');
  const updatedAtCol = colNames.find((c: any) => c.name === 'updatedAt');
  console.log('Target cols:', { receivedAtCol, updatedAtCol });

  if (receivedAtCol) {
    const colRangeRes = await fetch(`${base}/tables/CandidateTable/columns('${receivedAtCol.id}')/dataBodyRange`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    const colRange = await colRangeRes.json();
    console.log('receivedAt dataBodyRange:', {
      address: colRange.address,
      values: colRange.values,
      valueTypes: colRange.valueTypes,
      numberFormat: colRange.numberFormat,
      text: colRange.text,
    });
  }

  // Also check table rows
  const rowsRes = await fetch(`${base}/tables/CandidateTable/rows`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const rows = await rowsRes.json();
  console.log('Rows count:', rows.value?.length);
  if (rows.value?.length) {
    console.log('Row 0 values length:', rows.value[0].values[0].length);
  }
}

main().finally(() => prisma.$disconnect());

