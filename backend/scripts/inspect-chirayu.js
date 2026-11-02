require('dotenv').config();
const { PrismaClient } = require('@prisma/client');
const prisma = new PrismaClient();

async function main() {
  const candidate = await prisma.candidate.findUnique({
    where: { id: '106945ea-6ae3-4c64-b5bb-9c1148d72985' },
    include: {
      attachments: true,
      applications: true,
      excelSyncRuns: true,
    }
  });
  console.log('CANDIDATE:', JSON.stringify(candidate, null, 2));

  if (candidate && candidate.sourceEmailId) {
    const email = await prisma.emailMessage.findFirst({
      where: {
        OR: [
          { id: candidate.sourceEmailId },
          { messageId: candidate.sourceEmailId },
        ]
      },
      include: {
        attachments: true,
        processingJob: true,
      }
    });
    console.log('EMAIL:', JSON.stringify(email, null, 2));
  }
  await prisma.$disconnect();
}

main().catch(console.error);

