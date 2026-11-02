// backend/src/lib/candidateService.ts
// Candidate creation, deduplication, and application linking with audit logging.

import { prisma } from './db';
import { logger } from '../logger';
import type { ExtractedCandidate } from './candidateExtractor';

export interface IngestCandidateParams {
  extracted: ExtractedCandidate;
  sourceEmailId?: string;
  sourceEmailAddress?: string;
  attachmentId?: string;
  attachmentFilename?: string;
  storageKey?: string;
  resumeUrl?: string;
  requestId?: string;
}

export interface IngestCandidateResult {
  candidateId: string;
  applicationId: string;
  isDuplicate: boolean;
  processingStatus: string;
}

/**
 * Ingest candidate and application with deduplication on normalized email.
 */
export async function ingestCandidate(params: IngestCandidateParams): Promise<IngestCandidateResult> {
  const {
    extracted,
    sourceEmailId,
    sourceEmailAddress = '',
    attachmentId,
    attachmentFilename,
    storageKey,
    resumeUrl,
    requestId = 'system',
  } = params;

  const normalizedEmail = extracted.email ? extracted.email.toLowerCase().trim() : null;

  // 1. Check for existing candidate if email is present
  let existingCandidate = normalizedEmail
    ? await prisma.candidate.findUnique({
        where: { email: normalizedEmail },
      })
    : null;

  let candidateId: string;
  let isDuplicate = false;

  if (existingCandidate) {
    // Duplicate detected — merge / update existing candidate record
    isDuplicate = true;
    candidateId = existingCandidate.id;

    // Merge skills without duplicating
    const mergedSkills = Array.from(new Set([...existingCandidate.skills, ...extracted.skills]));

    await prisma.candidate.update({
      where: { id: candidateId },
      data: {
        skills: mergedSkills,
        // Update fields if missing on existing record
        ...(existingCandidate.name ? {} : { name: extracted.name }),
        ...(existingCandidate.phone ? {} : { phone: extracted.phone }),
        ...(existingCandidate.location ? {} : { location: extracted.location }),
        ...(existingCandidate.college ? {} : { college: extracted.college }),
        ...(existingCandidate.degree ? {} : { degree: extracted.degree }),
        ...(existingCandidate.graduationYear != null ? {} : (extracted.graduationYear != null ? { graduationYear: extracted.graduationYear } : {})),
        ...(existingCandidate.experienceYears != null ? {} : (extracted.experienceYears != null ? { experienceYears: extracted.experienceYears } : {})),
        // Update resume reference if new attachment provided
        ...(storageKey ? { resumeRef: storageKey } : {}),
        ...(attachmentFilename ? { resumeFileName: attachmentFilename } : {}),
        ...(resumeUrl ? { resumeUrl } : {}),
        updatedAt: new Date(),
      },
    });

    // Create Audit Log for merge
    await prisma.auditLog.create({
      data: {
        actor: 'system',
        action: 'candidate_merged',
        related: candidateId,
        entity: 'Candidate',
        entityId: candidateId,
        result: 'success',
        requestId,
        payload: {
          email: normalizedEmail,
          sourceEmailId,
          reason: 'Duplicate candidate matched on email address',
        },
      },
    });

    logger.info({ candidateId, email: normalizedEmail }, 'Existing candidate deduplicated and updated');
  } else {
    // New candidate — create fresh record
    const newCandidate = await prisma.candidate.create({
      data: {
        name: extracted.name,
        email: normalizedEmail,
        phone: extracted.phone,
        location: extracted.location || null,
        college: extracted.college,
        degree: extracted.degree,
        graduationYear: extracted.graduationYear,
        skills: extracted.skills,
        experienceYears: extracted.experienceYears,
        appliedRole: extracted.appliedRole || 'General Application',
        processingStatus: extracted.processingStatus,
        resumeRef: storageKey || null,
        resumeFileName: attachmentFilename || null,
        resumeUrl: resumeUrl || null,
        sourceEmailId: sourceEmailId || null,
        reviewFields: extracted.reviewFields,
      },
    });

    candidateId = newCandidate.id;

    // Create Audit Log for creation
    await prisma.auditLog.create({
      data: {
        actor: 'system',
        action: 'candidate_created',
        related: candidateId,
        entity: 'Candidate',
        entityId: candidateId,
        result: 'success',
        requestId,
        payload: {
          email: normalizedEmail,
          appliedRole: extracted.appliedRole,
          confidence: extracted.confidence,
          processingStatus: extracted.processingStatus,
        },
      },
    });

    logger.info({ candidateId, email: normalizedEmail }, 'New candidate created');
  }

  // 2. Create Application record linked to Candidate
  const application = await prisma.application.create({
    data: {
      candidateId,
      role: extracted.appliedRole || 'General Application',
      sourceEmail: sourceEmailAddress || (normalizedEmail || ''),
      status: 'new',
    },
  });

  // 3. Link Attachment to Candidate if attachmentId is provided
  if (attachmentId) {
    await prisma.attachment.update({
      where: { id: attachmentId },
      data: {
        candidateId,
        status: 'done',
        processedAt: new Date(),
      },
    }).catch((err) => {
      logger.warn({ attachmentId, err: err.message }, 'Failed to link attachment to candidate');
    });
  }

  return {
    candidateId,
    applicationId: application.id,
    isDuplicate,
    processingStatus: extracted.processingStatus,
  };
}

