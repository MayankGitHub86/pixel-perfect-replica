import { upsertCandidateRow } from '../src/lib/excel';
import { prisma } from '../src/lib/db';

async function main() {
  const mayankId = 'd5bd579a-309d-4697-aaa2-e27aadd2ce02';
  console.log('Syncing Mayank to Excel with updated code...');
  const res = await upsertCandidateRow(mayankId);
  console.log('Sync result:', res);
}

main().finally(() => prisma.$disconnect());

