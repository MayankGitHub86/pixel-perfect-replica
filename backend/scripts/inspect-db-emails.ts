import { PrismaClient } from '@prisma/client';
const prisma = new PrismaClient();

async function main() {
  const emails = await prisma.emailMessage.findMany({
    include: { attachments: true }
  });
  console.log('Total emails in DB:', emails.length);
  for (const e of emails) {
    console.log({
      id: e.id,
      messageId: e.messageId,
      subject: e.subject,
      fromEmail: e.fromEmail,
      candidateId: e.candidateId,
      history: e.history,
    });
  }
}

main().finally(() => prisma.$disconnect());

