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

  const colsRes = await fetch(`${base}/tables/CandidateTable/columns`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const colsData: any = await colsRes.json();
  const colNames = (colsData.value || []).map((c: any) => c.name);

  const rowsRes = await fetch(`${base}/tables/CandidateTable/rows`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const rowsData: any = await rowsRes.json();
  const rows = rowsData.value || [];

  console.log('Columns count:', colNames.length);
  console.log('Columns:', colNames);
  console.log('Rows count:', rows.length);

  for (let rIdx = 0; rIdx < rows.length; rIdx++) {
    const row = rows[rIdx];
    const vals = row.values[0];
    const rowObj: Record<string, any> = {};
    colNames.forEach((col: string, idx: number) => {
      rowObj[col] = vals[idx];
    });
    console.log(`\n--- ROW ${rIdx} (${rowObj['name']}) ---`);
    console.log(JSON.stringify(rowObj, null, 2));
  }
}

main().finally(() => prisma.$disconnect());

