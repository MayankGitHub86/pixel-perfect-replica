import fs from 'fs';
import path from 'path';
import { prisma } from '../src/lib/db';
import { uploadResumeToOneDrive } from '../src/lib/attachmentService';
import { upsertCandidateRow } from '../src/lib/excel';

async function main() {
  const mayankId = 'd5bd579a-309d-4697-aaa2-e27aadd2ce02';
  const candidate = await prisma.candidate.findUnique({
    where: { id: mayankId },
    include: { attachments: true },
  });

  if (!candidate) throw new Error('Mayank not found');

  const att = candidate.attachments[0];
  if (!att || !att.storageKey) throw new Error('Attachment storageKey not found');

  // Sync to Excel so Mayank gets his own unique HYPERLINK!
  const synced = await upsertCandidateRow(candidate.id);
  console.log('Successfully synced Mayank to Excel with his own OneDrive resume link! Result:', synced);
}

main().finally(() => prisma.$disconnect());
