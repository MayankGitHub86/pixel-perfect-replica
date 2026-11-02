import { PrismaClient } from '@prisma/client';
const prisma = new PrismaClient();

async function main() {
  const updated = await prisma.candidate.update({
    where: { id: 'd5bd579a-309d-4697-aaa2-e27aadd2ce02' },
    data: {
      graduationYear: null,
      college: 'Lovely Professional University',
      location: null, // "Git, VS" was bogus technical skills
    },
  });
  console.log('Updated Mayank:', {
    id: updated.id,
    name: updated.name,
    college: updated.college,
    graduationYear: updated.graduationYear,
    experienceYears: updated.experienceYears,
    location: updated.location,
  });
}

main().finally(() => prisma.$disconnect());

