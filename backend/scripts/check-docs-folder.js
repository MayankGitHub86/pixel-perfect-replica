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

async function getAccessToken() {
  const conn = await prisma.mailboxConnection.findFirst({ orderBy: { createdAt: 'desc' } });
  return decrypt(conn.accessToken);
}

async function main() {
  const token = await getAccessToken();
  const docsRes = await fetch('https://graph.microsoft.com/v1.0/me/drive/root:/Documents', {
    headers: { Authorization: `Bearer ${token}` }
  });
  console.log('Docs folder status:', docsRes.status);
  const docsData = await docsRes.json();
  console.log('Docs folder id:', docsData.id, docsData.name);
  await prisma.$disconnect();
}

main().catch(console.error);

