const fs = require('fs');
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

async function main() {
  const conn = await prisma.mailboxConnection.findFirst({ orderBy: { createdAt: 'desc' } });
  const token = decrypt(conn.accessToken);
  const localPath = "C:\\Users\\ASUS\\Documents\\lcc outlook project\\pixel-perfect-replica\\backend\\storage\\attachments\\ca1341f7-7987-4cea-a3db-3ce2347425b8\\cv-att-1_Chirayu_Sharma_Resume_TridhyaTech_BA.docx";
  const fileBuffer = fs.readFileSync(localPath);
  const candidateId = "106945ea-6ae3-4c64-b5bb-9c1148d72985";
  const targetName = `${candidateId}_Chirayu_Sharma_Resume_TridhyaTech_BA.docx`;

  const uploadUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/Candidate_Resumes/${encodeURIComponent(targetName)}:/content`;
  console.log('Uploading to:', uploadUrl);
  const res = await fetch(uploadUrl, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    },
    body: fileBuffer
  });
  console.log('Upload status:', res.status);
  const data = await res.json();
  console.log('Uploaded Item ID:', data.id);
  console.log('Uploaded Item Name:', data.name);
  console.log('Uploaded Item WebUrl:', data.webUrl);
  console.log('Uploaded Item Size:', data.size);
  await prisma.$disconnect();
}

main().catch(console.error);

