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

// Convert Date / ISO string to Excel serial date-time number
function toExcelSerialDate(dateInput: Date | string | null | undefined): number | null {
  if (!dateInput) return null;
  const d = typeof dateInput === 'string' ? new Date(dateInput) : dateInput;
  if (isNaN(d.getTime())) return null;

  // Unix epoch: 1970-01-01T00:00:00.000Z
  // Excel epoch: 1899-12-30T00:00:00.000Z (accounting for 1900 leap year bug)
  // Number of days difference = 25569
  const msPerDay = 86400000;
  // Preserve exact timestamp without artificial timezone shifting
  return 25569 + (d.getTime() / msPerDay);
}

async function main() {
  const token = await getValidAccessToken();
  const wb = await prisma.excelWorkbook.findFirst();
  if (!wb) throw new Error('No workbook found');
  const base = `https://graph.microsoft.com/v1.0/me/drive/items/${wb.driveItemId}/workbook`;

  const cand = await prisma.candidate.findUnique({
    where: { id: '106945ea-6ae3-4c64-b5bb-9c1148d72985' },
  });
  if (!cand) throw new Error('Candidate not found');

  const receivedAtSerial = toExcelSerialDate(cand.receivedAt);
  const updatedAtSerial = toExcelSerialDate(cand.updatedAt);
  console.log('Serials:', { receivedAtSerial, updatedAtSerial });

  // Try setting numberFormat on dataBodyRange
  const patchRecRes = await fetch(`${base}/tables/CandidateTable/columns('receivedAt')/dataBodyRange`, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      numberFormat: [['dd-mm-yyyy hh:mm:ss AM/PM']],
      values: [[receivedAtSerial]],
    }),
  });
  console.log('patch receivedAt status:', patchRecRes.status);
  const patchRecData = await patchRecRes.json();
  console.log('patch receivedAt response:', {
    values: patchRecData.values,
    valueTypes: patchRecData.valueTypes,
    numberFormat: patchRecData.numberFormat,
    text: patchRecData.text,
  });

  const patchUpdRes = await fetch(`${base}/tables/CandidateTable/columns('updatedAt')/dataBodyRange`, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      numberFormat: [['dd-mm-yyyy hh:mm:ss AM/PM']],
      values: [[updatedAtSerial]],
    }),
  });
  console.log('patch updatedAt status:', patchUpdRes.status);
  const patchUpdData = await patchUpdRes.json();
  console.log('patch updatedAt response:', {
    values: patchUpdData.values,
    valueTypes: patchUpdData.valueTypes,
    numberFormat: patchUpdData.numberFormat,
    text: patchUpdData.text,
  });

  // Also apply format to the whole column in worksheet Candidates!Q:R so future added rows inherit it
  const patchColFormatRes = await fetch(`${base}/worksheets('Candidates')/range(address='Q:R')/format`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  console.log('format range status:', patchColFormatRes.status);

  // Directly set numberFormat on worksheet range Q:R
  const patchWsRangeRes = await fetch(`${base}/worksheets('Candidates')/range(address='Q:R')`, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      numberFormat: [['dd-mm-yyyy hh:mm:ss AM/PM', 'dd-mm-yyyy hh:mm:ss AM/PM']],
    }),
  });
  console.log('patch worksheet Q:R status:', patchWsRangeRes.status);
}

main().finally(() => prisma.$disconnect());

