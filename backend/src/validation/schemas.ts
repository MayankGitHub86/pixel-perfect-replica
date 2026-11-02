import { z } from 'zod';

export const paginationQuerySchema = z.object({
  page: z.coerce.number().int().positive().default(1),
  pageSize: z.coerce.number().int().positive().max(1000).default(10),
  search: z.string().optional(),
  sortBy: z.string().optional(),
  sortDir: z.enum(['asc', 'desc']).default('desc'),
});

export const candidateListQuerySchema = paginationQuerySchema.extend({
  since: z.coerce.number().optional(),
  appliedRole: z.string().optional(),
  degree: z.string().optional(),
  processingStatus: z.enum(['new', 'processing', 'processed', 'needs_review', 'failed']).optional(),
  interviewStatus: z.enum(['not_scheduled', 'scheduled', 'completed', 'no_show']).optional(),
});

export const candidatePatchSchema = z.object({
  name: z.string().min(1).max(200).nullable().optional(),
  phone: z.string().max(50).nullable().optional(),
  college: z.string().max(200).nullable().optional(),
  degree: z.string().max(100).nullable().optional(),
  graduationYear: z.coerce.number().int().min(1950).max(2050).nullable().optional(),
  appliedRole: z.string().max(100).nullable().optional(),
  interviewStatus: z.enum(['not_scheduled', 'scheduled', 'completed', 'no_show']).optional(),
  skills: z.array(z.string()).optional(),
  experienceYears: z.coerce.number().min(0).max(100).nullable().optional(),
});

export const applicationListQuerySchema = paginationQuerySchema.extend({
  status: z.enum(['new', 'under_review', 'shortlisted', 'interview', 'rejected', 'hired', 'all']).optional(),
  candidateId: z.string().uuid().optional(),
});

export const applicationStatusPatchSchema = z.object({
  status: z.enum(['new', 'under_review', 'shortlisted', 'interview', 'rejected', 'hired']),
});

export const emailListQuerySchema = paginationQuerySchema.extend({
  status: z.enum(['new', 'processing', 'processed', 'needs_review', 'failed']).optional(),
});

export const jobListQuerySchema = paginationQuerySchema.extend({
  status: z.enum(['queued', 'processing', 'succeeded', 'retrying', 'failed', 'cancelled']).optional(),
  type: z.enum(['email_ingest', 'attachment_parse', 'ocr', 'field_extraction', 'excel_sync']).optional(),
});

export const auditListQuerySchema = paginationQuerySchema.extend({
  result: z.enum(['success', 'failure', 'info']).optional(),
});

