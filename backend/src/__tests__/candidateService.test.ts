// backend/src/__tests__/candidateService.test.ts
import { describe, it, expect, beforeEach, afterAll } from 'vitest';
import { prisma } from '../lib/db';
import { ingestCandidate } from '../lib/candidateService';
import type { ExtractedCandidate } from '../lib/candidateExtractor';

describe('Candidate Ingestion & Deduplication Service', () => {
  const testEmail = 'dedup.test@example.com';

  beforeEach(async () => {
    // Clean up test records
    await prisma.auditLog.deleteMany({ where: { payload: { path: ['email'], equals: testEmail } } });
    await prisma.application.deleteMany({ where: { sourceEmail: testEmail } });
    await prisma.candidate.deleteMany({ where: { email: testEmail } });
  });

  afterAll(async () => {
    await prisma.auditLog.deleteMany({ where: { payload: { path: ['email'], equals: testEmail } } });
    await prisma.application.deleteMany({ where: { sourceEmail: testEmail } });
    await prisma.candidate.deleteMany({ where: { email: testEmail } });
    await prisma.$disconnect();
  });

  it('creates new candidate and application on first ingestion', async () => {
    const extracted: ExtractedCandidate = {
      name: 'Alex Mercer',
      email: testEmail,
      phone: '123-456-7890',
      location: null,
      skills: ['TypeScript', 'Node.js'],
      degree: 'B.Tech',
      college: 'Test College',
      graduationYear: 2021,
      experienceYears: 3,
      projects: [],
      certifications: [],
      appliedRole: 'Backend Engineer',
      confidence: 0.9,
      reviewFields: [],
      processingStatus: 'processed',
    };

    const res = await ingestCandidate({
      extracted,
      sourceEmailAddress: testEmail,
    });

    expect(res.isDuplicate).toBe(false);
    expect(res.candidateId).toBeDefined();
    expect(res.applicationId).toBeDefined();

    // Verify candidate in DB
    const candidate = await prisma.candidate.findUnique({ where: { id: res.candidateId } });
    expect(candidate).not.toBeNull();
    expect(candidate?.name).toBe('Alex Mercer');
    expect(candidate?.email).toBe(testEmail);
    expect(candidate?.skills).toEqual(['TypeScript', 'Node.js']);
  });

  it('deduplicates candidate and updates existing record on second ingestion', async () => {
    const firstExtracted: ExtractedCandidate = {
      name: 'Alex Mercer',
      email: testEmail,
      phone: '123-456-7890',
      location: null,
      skills: ['TypeScript'],
      degree: 'B.Tech',
      college: null,
      graduationYear: null,
      experienceYears: 3,
      projects: [],
      certifications: [],
      appliedRole: 'Backend Engineer',
      confidence: 0.8,
      reviewFields: [],
      processingStatus: 'processed',
    };

    const first = await ingestCandidate({ extracted: firstExtracted, sourceEmailAddress: testEmail });
    expect(first.isDuplicate).toBe(false);

    // Second application from same candidate with new skill and college info
    const secondExtracted: ExtractedCandidate = {
      name: 'Alex Mercer',
      email: testEmail,
      phone: '123-456-7890',
      location: null,
      skills: ['TypeScript', 'GraphQL'],
      degree: 'B.Tech',
      college: 'Updated University',
      graduationYear: 2021,
      experienceYears: 3.5,
      projects: [],
      certifications: [],
      appliedRole: 'Full Stack Engineer',
      confidence: 0.9,
      reviewFields: [],
      processingStatus: 'processed',
    };

    const second = await ingestCandidate({ extracted: secondExtracted, sourceEmailAddress: testEmail });
    expect(second.isDuplicate).toBe(true);
    expect(second.candidateId).toBe(first.candidateId);
    expect(second.applicationId).not.toBe(first.applicationId); // New application created!

    // Verify candidate was merged without duplicate candidate record
    const totalCandidates = await prisma.candidate.count({ where: { email: testEmail } });
    expect(totalCandidates).toBe(1);

    const merged = await prisma.candidate.findUnique({ where: { id: first.candidateId } });
    expect(merged?.skills).toContain('TypeScript');
    expect(merged?.skills).toContain('GraphQL');
    expect(merged?.college).toBe('Updated University');

    // Verify both applications exist
    const apps = await prisma.application.findMany({ where: { candidateId: first.candidateId } });
    expect(apps.length).toBe(2);
  });
});

