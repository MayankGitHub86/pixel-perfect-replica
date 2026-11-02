const fs = require('fs');
const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '../.env') });
const { PrismaClient } = require('@prisma/client');
const prisma = new PrismaClient();

async function backup() {
  const data = {
    candidates: await prisma.candidate.findMany(),
    emailMessages: await prisma.emailMessage.findMany(),
    attachments: await prisma.attachment.findMany(),
    applications: await prisma.application.findMany(),
    excelWorkbooks: await prisma.excelWorkbook.findMany(),
    excelSyncRuns: await prisma.excelSyncRun.findMany(),
    processingJobs: await prisma.processingJob.findMany(),
    mailboxConnections: await prisma.mailboxConnection.findMany(),
  };
  const backupPath = path.join(__dirname, '../storage/db_backup_' + Date.now() + '.json');
  fs.mkdirSync(path.dirname(backupPath), { recursive: true });
  fs.writeFileSync(backupPath, JSON.stringify(data, null, 2));
  console.log('Database backup created at:', backupPath);
  await prisma.$disconnect();
}

backup().catch(console.error);

