const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '../.env') });
const crypto = require('crypto');
const { PrismaClient } = require('@prisma/client');

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

function encrypt(plainText) {
  const iv = crypto.randomBytes(12);
  const key = Buffer.from(process.env.TOKEN_ENCRYPTION_KEY, 'base64');
  const cipher = crypto.createCipheriv('aes-256-gcm', key, iv, { authTagLength: 16 });
  const encrypted = Buffer.concat([cipher.update(plainText, 'utf8'), cipher.final()]);
  const tag = cipher.getAuthTag();
  return `${iv.toString('base64')}.${encrypted.toString('base64')}.${tag.toString('base64')}`;
}

async function refreshAccessToken(mailboxId) {
  const mailbox = await prisma.mailboxConnection.findUnique({ where: { id: mailboxId } });
  if (!mailbox) throw new Error('MailboxConnection not found');
  const refreshToken = decrypt(mailbox.refreshToken);
  const tenantId = process.env.MICROSOFT_TENANT_ID || 'common';
  const clientId = process.env.MICROSOFT_CLIENT_ID;
  const clientSecret = process.env.MICROSOFT_CLIENT_SECRET;
  const redirectUri = process.env.MICROSOFT_REDIRECT_URI || 'http://localhost:3000/api/integrations/microsoft/callback';

  const tokenRes = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      grant_type: 'refresh_token',
      refresh_token: refreshToken,
      redirect_uri: redirectUri,
    }).toString(),
  });

  if (!tokenRes.ok) {
    const errBody = await tokenRes.text();
    throw new Error(`Token refresh failed: ${tokenRes.status} ${errBody}`);
  }

  const tokenData = await tokenRes.json();
  const newAccessToken = tokenData.access_token;
  const newRefreshToken = tokenData.refresh_token ?? refreshToken;
  const expiresIn = tokenData.expires_in;

  const encryptedAccess = encrypt(newAccessToken);
  const encryptedRefresh = encrypt(newRefreshToken);
  const tokenExpiresAt = new Date(Date.now() + expiresIn * 1000);

  await prisma.mailboxConnection.update({
    where: { id: mailboxId },
    data: {
      accessToken: encryptedAccess,
      refreshToken: encryptedRefresh,
      tokenExpiresAt,
      updatedAt: new Date(),
    },
  });
  console.log('Token refreshed successfully. Expires at:', tokenExpiresAt);
  return newAccessToken;
}

async function getAccessToken() {
  const conn = await prisma.mailboxConnection.findFirst({ orderBy: { createdAt: 'desc' } });
  if (!conn) throw new Error('No mailbox connection found');
  if (new Date() >= new Date(conn.tokenExpiresAt.getTime() - 60000)) {
    console.log('Token expired or expiring soon, refreshing...');
    return refreshAccessToken(conn.id);
  }
  return decrypt(conn.accessToken);
}

async function main() {
  let token = await getAccessToken();
  const wb = await prisma.excelWorkbook.findFirst();
  console.log('DB Workbook DriveItemId:', wb.driveItemId);

  const realId = wb.driveItemId;
  let itemRes = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${realId}`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (itemRes.status === 401) {
    console.log('401 received, forcing refresh...');
    const conn = await prisma.mailboxConnection.findFirst({ orderBy: { createdAt: 'desc' } });
    token = await refreshAccessToken(conn.id);
    itemRes = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${realId}`, {
      headers: { Authorization: `Bearer ${token}` }
    });
  }

  console.log('Item Status:', itemRes.status);
  const itemData = await itemRes.json();
  console.log('Graph Item Info:');
  console.log('ID:', itemData.id);
  console.log('Name:', itemData.name);
  console.log('ParentPath:', itemData.parentReference?.path);
  console.log('WebUrl:', itemData.webUrl);

  const tablesRes = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${realId}/workbook/tables`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  console.log('Tables Status:', tablesRes.status);
  const tablesData = await tablesRes.json();
  console.log('Tables:', tablesData.value?.map(t => ({ id: t.id, name: t.name })));

  if (tablesData.value?.length > 0) {
    const tableId = tablesData.value[0].id;
    const colsRes = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${realId}/workbook/tables/${tableId}/columns`, {
      headers: { Authorization: `Bearer ${token}` }
    });
    const colsData = await colsRes.json();
    console.log('Columns:', colsData.value?.map(c => c.name));

    const rowsRes = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${realId}/workbook/tables/${tableId}/rows`, {
      headers: { Authorization: `Bearer ${token}` }
    });
    const rowsData = await rowsRes.json();
    console.log('Rows count:', rowsData.value?.length);
    if (rowsData.value?.length > 0) {
      console.log('Row 1 values:', rowsData.value[0].values);
    }
  }

  await prisma.$disconnect();
}

main().catch(console.error);
