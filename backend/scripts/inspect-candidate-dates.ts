import { PrismaClient } from '@prisma/client';
import dotenv from 'dotenv';
import path from 'path';

dotenv.config({ path: path.resolve(process.cwd(), '.env') });
const prisma = new PrismaClient();

async function main() {
  const c = await prisma.candidate.findUnique({
    where: { id: '106945ea-6ae3-4c64-b5bb-9c1148d72985' },
  });
  console.log('Candidate Dates:', {
    receivedAt: c?.receivedAt,
    receivedAtISO: c?.receivedAt?.toISOString(),
    updatedAt: c?.updatedAt,
    updatedAtISO: c?.updatedAt?.toISOString(),
  });
  if (c?.sourceEmailId) {
    const em = await prisma.emailMessage.findFirst({
      where: { OR: [{ id: c.sourceEmailId }, { messageId: c.sourceEmailId }] },
    });
    console.log('Email Message Dates:', {
      receivedAt: em?.receivedAt,
      receivedAtISO: em?.receivedAt?.toISOString(),
      createdAt: em?.createdAt,
      createdAtISO: em?.createdAt?.toISOString(),
    });
  }
}

main().finally(() => prisma.$disconnect());

