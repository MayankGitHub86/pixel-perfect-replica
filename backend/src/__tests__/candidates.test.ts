import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import request from 'supertest';
import { createApp } from '../app';
import { prisma } from '../lib/db';

const app = createApp();

describe('Candidate & Application APIs', () => {
  let testCandidateId: string;
  let testApplicationId: string;
  let testJobId: string;

  beforeAll(async () => {
    // Clean any previous test data
    await prisma.auditLog.deleteMany({ where: { actor: 'test-runner' } });
    await prisma.application.deleteMany({ where: { sourceEmail: 'test.cand@example.com' } });
    await prisma.processingJob.deleteMany({ where: { related: 'Test Job' } });
    await prisma.emailMessage.deleteMany({ where: { fromEmail: 'test.cand@example.com' } });
    await prisma.candidate.deleteMany({ where: { email: 'test.cand@example.com' } });

    // Create a real test candidate
    const candidate = await prisma.candidate.create({
      data: {
        name: 'Test Candidate',
        email: 'test.cand@example.com',
        phone: '+1 555-0100',
        college: 'Test University',
        degree: 'B.S.',
        graduationYear: 2022,
        skills: ['TypeScript', 'Node.js'],
        appliedRole: 'Software Engineer',
        processingStatus: 'needs_review',
        interviewStatus: 'not_scheduled',
        reviewFields: ['college', 'graduationYear'],
      },
    });
    testCandidateId = candidate.id;

    const application = await prisma.application.create({
      data: {
        candidateId: candidate.id,
        role: 'Software Engineer',
        sourceEmail: 'test.cand@example.com',
        status: 'new',
      },
    });
    testApplicationId = application.id;

    const job = await prisma.processingJob.create({
      data: {
        type: 'field_extraction',
        related: 'Test Job',
        status: 'failed',
        attempts: 1,
        lastError: 'Parse timeout',
      },
    });
    testJobId = job.id;
  });

  afterAll(async () => {
    // Cleanup
    await prisma.auditLog.deleteMany({ where: { actor: 'test-runner' } });
    await prisma.application.deleteMany({ where: { id: testApplicationId } });
    await prisma.processingJob.deleteMany({ where: { id: testJobId } });
    await prisma.candidate.deleteMany({ where: { id: testCandidateId } });
    await prisma.$disconnect();
  });

  describe('GET /api/candidates', () => {
    it('should return paginated candidates list with correct shape', async () => {
      const res = await request(app).get('/api/candidates?page=1&pageSize=10');

      expect(res.status).toBe(200);
      expect(res.body.page).toBe(1);
      expect(res.body.pageSize).toBe(10);
      expect(typeof res.body.total).toBe('number');
      expect(Array.isArray(res.body.items)).toBe(true);

      const found = res.body.items.find((c: any) => c.id === testCandidateId);
      expect(found).toBeDefined();
      expect(found.name).toBe('Test Candidate');
      expect(found.email).toBe('test.cand@example.com');
      expect(found.processingStatus).toBe('needs_review');
    });

    it('should filter candidates by appliedRole', async () => {
      const res = await request(app).get('/api/candidates?appliedRole=Software%20Engineer');

      expect(res.status).toBe(200);
      expect(res.body.items.length).toBeGreaterThanOrEqual(1);
      expect(res.body.items.every((c: any) => c.appliedRole === 'Software Engineer')).toBe(true);
    });

    it('should filter candidates by search term', async () => {
      const res = await request(app).get('/api/candidates?search=Test%20Candidate');

      expect(res.status).toBe(200);
      expect(res.body.items.some((c: any) => c.id === testCandidateId)).toBe(true);
    });
  });

  describe('GET /api/candidates/:id', () => {
    it('should return candidate detail with applications and audit', async () => {
      const res = await request(app).get(`/api/candidates/${testCandidateId}`);

      expect(res.status).toBe(200);
      expect(res.body.id).toBe(testCandidateId);
      expect(res.body.name).toBe('Test Candidate');
      expect(Array.isArray(res.body.applications)).toBe(true);
      expect(res.body.applications[0].id).toBe(testApplicationId);
      expect(Array.isArray(res.body.audit)).toBe(true);
    });

    it('should return 404 when candidate does not exist', async () => {
      const res = await request(app).get('/api/candidates/non-existent-uuid-12345');

      expect(res.status).toBe(404);
      expect(res.body.error.code).toBe('NOT_FOUND');
    });
  });

  describe('PATCH /api/candidates/:id', () => {
    it('should update candidate fields and resolve reviewFields', async () => {
      const res = await request(app)
        .patch(`/api/candidates/${testCandidateId}`)
        .send({
          college: 'MIT Updated',
          graduationYear: 2023,
          interviewStatus: 'scheduled',
        });

      expect(res.status).toBe(200);
      expect(res.body.college).toBe('MIT Updated');
      expect(res.body.graduationYear).toBe(2023);
      expect(res.body.interviewStatus).toBe('scheduled');
      // Both review fields were resolved, so reviewFields should now be empty and status promoted to processed
      expect(res.body.reviewFields).toEqual([]);
      expect(res.body.processingStatus).toBe('processed');

      // Verify audit log was written
      const audit = await prisma.auditLog.findFirst({
        where: { related: testCandidateId },
        orderBy: { createdAt: 'desc' },
      });
      expect(audit).toBeDefined();
    });

    it('should reject invalid patch body with 400', async () => {
      const res = await request(app)
        .patch(`/api/candidates/${testCandidateId}`)
        .send({
          interviewStatus: 'invalid_status_value',
        });

      expect(res.status).toBe(400);
      expect(res.body.error.code).toBe('VALIDATION_ERROR');
    });
  });

  describe('GET /api/applications', () => {
    it('should list applications with candidate name', async () => {
      const res = await request(app).get('/api/applications');

      expect(res.status).toBe(200);
      expect(Array.isArray(res.body.items)).toBe(true);
      const appItem = res.body.items.find((a: any) => a.id === testApplicationId);
      expect(appItem).toBeDefined();
      expect(appItem.role).toBe('Software Engineer');
      expect(appItem.candidateName).toBe('Test Candidate');
    });
  });

  describe('PATCH /api/applications/:id/status', () => {
    it('should update application status and write audit log', async () => {
      const res = await request(app)
        .patch(`/api/applications/${testApplicationId}/status`)
        .send({ status: 'interview' });

      expect(res.status).toBe(200);
      expect(res.body.status).toBe('interview');

      // Verify persistence in DB
      const dbApp = await prisma.application.findUnique({ where: { id: testApplicationId } });
      expect(dbApp?.status).toBe('interview');
    });

    it('should reject invalid application status with 400', async () => {
      const res = await request(app)
        .patch(`/api/applications/${testApplicationId}/status`)
        .send({ status: 'not_a_real_status' });

      expect(res.status).toBe(400);
      expect(res.body.error.code).toBe('VALIDATION_ERROR');
    });
  });

  describe('POST /api/jobs/:id/retry', () => {
    it('should retry a failed job and increment attempts', async () => {
      const res = await request(app).post(`/api/jobs/${testJobId}/retry`);

      expect(res.status).toBe(200);
      expect(res.body.status).toBe('queued');
      expect(res.body.attempts).toBe(2);
      expect(res.body.lastError).toBeNull();
    });
  });

  describe('Integration Status Truthfulness', () => {
    it('GET /api/excel-sync/status should return genuine state', async () => {
      const res = await request(app).get('/api/excel-sync/status');

      expect(res.status).toBe(200);
      expect(['not_connected', 'connected']).toContain(res.body.state);
      expect(typeof res.body.pendingRows).toBe('number');
      expect(Array.isArray(res.body.history)).toBe(true);
    });

    it('GET /api/integrations/microsoft/status should return genuine state', async () => {
      const res = await request(app).get('/api/integrations/microsoft/status');

      expect(res.status).toBe(200);
      expect(['not_connected', 'connected', 'expired']).toContain(res.body.state);
    });
  });
});

