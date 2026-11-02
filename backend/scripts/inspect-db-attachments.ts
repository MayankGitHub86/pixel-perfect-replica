import { PrismaClient } from '@prisma/client';
const prisma = new PrismaClient();

async function main() {
  const atts = await prisma.attachment.findMany();
  console.log('Attachments in DB:');
  for (const a of atts) {
    console.log({
      id: a.id,
      emailId: a.emailId,
      filename: a.filename,
      candidateId: a.candidateId,
      storageKey: a.storageKey,
      status: a.status
    });
  }
}

main().finally(() => prisma.$disconnect());

