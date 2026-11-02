// backend/src/lib/queue.ts
// BullMQ processing queue and worker for resume ingestion and candidate extraction.

import { Queue, Worker, Job } from 'bullmq';
import { redis } from './redis';
import { prisma } from './db';
import { logger } from '../logger';
import { parseResumeBuffer } from './resumeParser';
import { extractCandidateInfo } from './candidateExtractor';
import { ingestCandidate } from './candidateService';
import { readAttachmentFile } from './attachmentService';
// import { excelSyncQueue } from './queue'; // placeholder removed – defined later in this file

export interface EmailProcessingJobData {
  emailId: string;
  requestId?: string;
}

export const EMAIL_PROCESSING_QUEUE_NAME = 'email-processing';

export const emailProcessingQueue = new Queue<EmailProcessingJobData>(
  EMAIL_PROCESSING_QUEUE_NAME,
  {
    connection: redis as any,
    defaultJobOptions: {
      attempts: 3,
      backoff: { type: 'exponential', delay: 2000 },
      removeOnComplete: true,
      removeOnFail: false,
    },
  }
);

export async function processEmailJob(jobData: EmailProcessingJobData): Promise<void> {
  const { emailId, requestId = 'system' } = jobData;

  const email = await prisma.emailMessage.findUnique({
    where: { id: emailId },
    include: { attachments: true, processingJob: true },
  });
  if (!email) throw new Error(`Email ${emailId} not found.`);

  await prisma.emailMessage.update({ where: { id: emailId }, data: { status: 'processing' } });
  if (email.processingJob) {
    await prisma.processingJob.update({ where: { id: email.processingJob.id }, data: { status: 'processing', startedAt: new Date() } });
  }

  try {
    let resumeText = '';
    let usedAttachmentId: string | undefined;
    let usedAttachmentFilename: string | undefined;
    let usedStorageKey: string | undefined;
    let resumeUrl: string | undefined;

    const validAttachment = email.attachments.find(
      (a) => (a.status === 'new' || a.status === 'done') && a.storageKey && a.storageKey !== 'rejected' && a.storageKey !== 'pending_download'
    );
    if (validAttachment) {
      try {
        const fileBuffer = await readAttachmentFile(validAttachment.storageKey);
        const parsed = await parseResumeBuffer(fileBuffer, validAttachment.filename);
        resumeText = parsed.text;
        usedAttachmentId = validAttachment.id;
        usedAttachmentFilename = validAttachment.filename;
        usedStorageKey = validAttachment.storageKey;

        try {
          const { uploadResumeToOneDrive } = await import('./attachmentService');
          const odResult = await uploadResumeToOneDrive(
            email.candidateId || email.id,
            validAttachment.filename,
            fileBuffer,
            validAttachment.mimeType
          );
          resumeUrl = odResult.webUrl;
        } catch (odErr: any) {
          logger.warn({ attachmentId: validAttachment.id, error: odErr.message }, 'Failed to upload attachment to OneDrive');
        }
      } catch (attErr: any) {
        logger.warn({ attachmentId: validAttachment.id, error: attErr.message }, 'Failed to read attachment buffer');
      }
    }

    const extracted = extractCandidateInfo(resumeText, { subject: email.subject, from: email.fromEmail });
    const ingestResult = await ingestCandidate({
      extracted,
      sourceEmailId: email.id,
      sourceEmailAddress: email.fromEmail,
      attachmentId: usedAttachmentId,
      attachmentFilename: usedAttachmentFilename,
      storageKey: usedStorageKey,
      resumeUrl,
      requestId,
    });

    // Enqueue Excel sync for this candidate automatically
    await excelSyncQueue.add('sync', { candidateId: ingestResult.candidateId });

    const history = Array.isArray(email.history) ? [...(email.history as any[])] : [];
    history.push({ at: new Date().toISOString(), step: 'field_extraction', status: 'succeeded', message: `Candidate extracted (Status: ${ingestResult.processingStatus})` });
    await prisma.emailMessage.update({
      where: { id: emailId },
      data: { status: ingestResult.processingStatus, candidateId: ingestResult.candidateId, applicationId: ingestResult.applicationId, processedAt: new Date(), history },
    });

    if (email.processingJob) {
      await prisma.processingJob.update({ where: { id: email.processingJob.id }, data: { status: 'succeeded', completedAt: new Date() } });
    }
    logger.info({ emailId, candidateId: ingestResult.candidateId }, 'Email processing job completed');
  } catch (err: any) {
    logger.error({ emailId, error: err.message }, 'Email processing job failed');
    await prisma.emailMessage.update({ where: { id: emailId }, data: { status: 'failed', error: err.message } });
    if (email.processingJob) {
      await prisma.processingJob.update({ where: { id: email.processingJob.id }, data: { status: 'failed', lastError: err.message, completedAt: new Date() } });
    }
    throw err;
  }
}

export const emailWorker = new Worker<EmailProcessingJobData>(
  EMAIL_PROCESSING_QUEUE_NAME,
  async (job: Job<EmailProcessingJobData>) => {
    await processEmailJob(job.data);
  },
  { connection: redis as any, concurrency: 3 }
);
emailWorker.on('failed', (job, err) => logger.error({ jobId: job?.id, error: err.message }, 'BullMQ email worker job failed'));

// ---- Excel sync queue and worker ----
export const EXCEL_SYNC_QUEUE_NAME = 'excel-sync';
export const excelSyncQueue = new Queue<{ candidateId: string }>(EXCEL_SYNC_QUEUE_NAME, { connection: redis as any });

export const excelSyncWorker = new Worker<{ candidateId: string }, void>(
  EXCEL_SYNC_QUEUE_NAME,
  async (job) => {
    const { candidateId } = job.data;
    const { upsertCandidateRow } = await import('./excel');
    await upsertCandidateRow(candidateId);
  },
  { connection: redis as any, concurrency: 3 }
);
excelSyncWorker.on('failed', (job, err) => logger.error({ jobId: job?.id, error: err.message }, 'Excel sync job failed'));
