import { PrismaClient } from '@prisma/client';
const prisma = new PrismaClient();

async function main() {
  const cands = await prisma.candidate.findMany({
    orderBy: { createdAt: 'desc' }
  });
  console.log('Candidates in DB:');
  for (const c of cands) {
    console.log({
      id: c.id,
      name: c.name,
      email: c.email,
      college: c.college,
      degree: c.degree,
      graduationYear: c.graduationYear,
      experienceYears: c.experienceYears,
      resumeRef: c.resumeRef,
      resumeFileName: c.resumeFileName,
      sourceEmailId: c.sourceEmailId
    });
  }
}

main().finally(() => prisma.$disconnect());

