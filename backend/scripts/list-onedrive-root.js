// backend/scripts/list-onedrive-root.js
// Read the stored access token (ignoring expiration) and list OneDrive root items.
// No token refresh – used only for read‑only verification.

const fs = require('fs');
const path = require('path');
require('dotenv').config();
const { PrismaClient } = require('@prisma/client');
const crypto = require('crypto');

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

async function main() {
  const conn = await prisma.mailboxConnection.findFirst({ orderBy: { createdAt: 'desc' } });
  if (!conn) throw new Error('No mailbox connection found');
  const token = decrypt(conn.accessToken); // ignore expiration
  console.log('✅ Access token (raw) acquired');

  const listRes = await fetch('https://graph.microsoft.com/v1.0/me/drive/root/children', {
    headers: { Authorization: `Bearer ${token}` },
  });
  console.log('GET root children status:', listRes.status);
  if (!listRes.ok) {
    const txt = await listRes.text();
    console.error('Error body:', txt);
    return;
  }
  const data = await listRes.json();
  console.log('Root items count:', data.value?.length ?? 0);
  data.value?.forEach(item => {
    console.log(`- ${item.name} (id: ${item.id})`);
  });
}

main()
  .catch(err => console.error('❌ Error:', err))
  .finally(() => prisma.$disconnect());

