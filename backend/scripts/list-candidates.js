const { PrismaClient } = require('@prisma/client');
(async () => {
  const prisma = new PrismaClient();
  const candidates = await prisma.candidate.findMany({ orderBy: { createdAt: 'asc' } });
  console.log('Candidates:');
  candidates.forEach(c => {
    console.log(`- ID: ${c.id}\n  Name: ${c.name}\n  Email: ${c.email}\n  SourceEmailId: ${c.sourceEmailId}\n  ApplicationId: ${c.applicationId}\n`);
  });
  await prisma.$disconnect();
})();

