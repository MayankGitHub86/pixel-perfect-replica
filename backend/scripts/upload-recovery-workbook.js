const fs = require('fs');
const path = require('path');
const { PrismaClient } = require('@prisma/client');
const crypto = require('crypto');
require('dotenv').config();

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

async function getAccessToken() {
  const conn = await prisma.mailboxConnection.findFirst({ orderBy: { createdAt: 'desc' } });
  if (!conn) throw new Error('No mailbox connection found');
  
  if (new Date() >= conn.tokenExpiresAt) {
    console.log('Token expired. Refreshing token via refresh endpoint...');
    // We can call refreshAccessToken from compiled or ts-node, or do token refresh via fetch
    const { refreshAccessToken } = require('../dist/lib/tokenRefresh');
    await refreshAccessToken(conn.id);
    const refreshed = await prisma.mailboxConnection.findUnique({ where: { id: conn.id } });
    return decrypt(refreshed.accessToken);
  }
  return decrypt(conn.accessToken);
}

async function main() {
  const token = await getAccessToken();
  console.log('Access token acquired.');

  // 1. Check OneDrive root children
  console.log('\n--- Checking OneDrive root children ---');
  const listRes = await fetch('https://graph.microsoft.com/v1.0/me/drive/root/children', {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!listRes.ok) {
    throw new Error(`List root failed: ${listRes.status} ${await listRes.text()}`);
  }
  const listData = await listRes.json();
  const files = listData.value || [];
  console.log(`Found ${files.length} items in OneDrive root:`);
  for (const item of files) {
    console.log(`- ${item.name} (id: ${item.id}, size: ${item.size} bytes)`);
  }

  const existingRecovery = files.find(f => f.name === 'Candidate_Master_Recovery.xlsx');
  let recoveryItem = null;

  if (existingRecovery) {
    console.log('\nFound existing Candidate_Master_Recovery.xlsx, checking validity...');
    // Test if worksheets endpoint works
    const testWs = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${existingRecovery.id}/workbook/worksheets`, {
      headers: { Authorization: `Bearer ${token}` }
    });
    if (testWs.ok) {
      console.log('Existing recovery workbook is valid!');
      recoveryItem = existingRecovery;
    } else {
      console.log('Existing recovery workbook returned error:', testWs.status, await testWs.text());
    }
  }

  if (!recoveryItem) {
    console.log('\nUploading Candidate_Master_Recovery.xlsx to OneDrive root...');
    const localFile = path.join(__dirname, '../Candidate_Master_Recovery.xlsx');
    const buffer = fs.readFileSync(localFile);
    console.log('Local file size to upload:', buffer.length, 'bytes');

    const uploadRes = await fetch('https://graph.microsoft.com/v1.0/me/drive/root:/Candidate_Master_Recovery.xlsx:/content', {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Length': buffer.length.toString(),
      },
      body: buffer
    });

    if (!uploadRes.ok) {
      throw new Error(`Upload failed: ${uploadRes.status} ${await uploadRes.text()}`);
    }

    recoveryItem = await uploadRes.json();
    console.log('Upload successful! HTTP 200/201');
  }

  console.log('\n=== Actual Recovery DriveItem Metadata ===');
  console.log('Actual DriveItem ID:', recoveryItem.id);
  console.log('Actual Filename:', recoveryItem.name);
  console.log('Actual webUrl:', recoveryItem.webUrl);
  console.log('Actual Parent Location:', recoveryItem.parentReference ? JSON.stringify(recoveryItem.parentReference) : 'root');
  console.log('Actual File Size:', recoveryItem.size, 'bytes');

  // Phase 5 Graph Verification
  console.log('\n=== Graph Verification: Worksheets ===');
  const wsRes = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${recoveryItem.id}/workbook/worksheets`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  console.log('Worksheets HTTP Status:', wsRes.status);
  const wsData = await wsRes.json();
  console.log('Worksheets:', wsData.value?.map(w => w.name));

  console.log('\n=== Graph Verification: Tables ===');
  const tRes = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${recoveryItem.id}/workbook/tables`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  console.log('Tables HTTP Status:', tRes.status);
  const tData = await tRes.json();
  console.log('Tables:', tData.value?.map(t => ({ name: t.name, showHeaders: t.showHeaders })));

  console.log('\n=== Graph Verification: Columns ===');
  const colRes = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${recoveryItem.id}/workbook/tables/CandidateTable/columns`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  console.log('Columns HTTP Status:', colRes.status);
  const colData = await colRes.json();
  const colNames = colData.value?.map(c => c.name) || [];
  console.log(`Columns count: ${colNames.length}`);
  console.log('Columns:', colNames);

  // Write output json for verification
  fs.writeFileSync(path.join(__dirname, 'recovery-meta.json'), JSON.stringify(recoveryItem, null, 2));
  console.log('\nRecovery metadata saved to recovery-meta.json');
}

main().catch(err => {
  console.error('Upload & verify error:', err);
  process.exit(1);
}).finally(() => prisma.$disconnect());

