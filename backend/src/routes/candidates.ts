import { Router, Request, Response, NextFunction } from 'express';
import { prisma } from '../lib/db';
import { candidateListQuerySchema, candidatePatchSchema } from '../validation/schemas';
import { Prisma } from '@prisma/client';
import path from 'path';
import { ATTACHMENTS_STORAGE_DIR } from '../lib/attachmentService';

export const candidateRouter = Router();

/**
 * Resolve a stored resumeRef / storageKey to an absolute path, or null if it cannot be resolved.
 * Accepts both relative keys ("emailId/key_file.pdf") and legacy absolute paths.
 * Never exposes the resolved path outside this module.
 */
function resolveLocalAttachmentPath(key: string): string | null {
  try {
    const resolved = path.isAbsolute(key)
      ? path.resolve(key)
      : path.resolve(ATTACHMENTS_STORAGE_DIR, key);
    if (!resolved.startsWith(ATTACHMENTS_STORAGE_DIR)) return null; // sandbox violation
    return resolved;
  } catch {
    return null;
  }
}

/**
 * Safe boolean — does a candidate have a downloadable resume?
 * Used to indicate to the frontend whether a Download button should appear,
 * without ever leaking raw file paths over the API.
 */
function hasResume(resumeRef: string | null): boolean {
  if (!resumeRef) return false;
  // OneDrive DriveItem IDs look like "XXXXXXXX!sYYY" — relative keys contain "/"
  return true; // presence of any non-null ref is sufficient; download endpoint does the real check
}


candidateRouter.get('/', async (req: Request, res: Response, next: NextFunction) => {
  try {
    const q = candidateListQuerySchema.parse(req.query);
    const { page, pageSize, search, sortBy, sortDir, since, appliedRole, degree, processingStatus, interviewStatus } = q;

    const where: Prisma.CandidateWhereInput = {};

    if (search) {
      where.OR = [
        { name: { contains: search, mode: 'insensitive' } },
        { email: { contains: search, mode: 'insensitive' } },
        { appliedRole: { contains: search, mode: 'insensitive' } },
        { college: { contains: search, mode: 'insensitive' } },
        { degree: { contains: search, mode: 'insensitive' } },
        { skills: { has: search } },
      ];
    }

    if (appliedRole && appliedRole !== 'all') {
      where.appliedRole = appliedRole;
    }

    if (degree && degree !== 'all') {
      where.degree = degree;
    }

    if (processingStatus && (processingStatus as string) !== 'all') {
      where.processingStatus = processingStatus;
    }

    if (interviewStatus && (interviewStatus as string) !== 'all') {
      where.interviewStatus = interviewStatus;
    }

    if (since) {
      const minDate = new Date(Date.now() - since * 86_400_000);
      where.receivedAt = { gte: minDate };
    }

    const orderBy: Prisma.CandidateOrderByWithRelationInput = {};
    if (sortBy && ['name', 'receivedAt', 'updatedAt', 'appliedRole', 'experienceYears'].includes(sortBy)) {
      orderBy[sortBy as keyof Prisma.CandidateOrderByWithRelationInput] = sortDir;
    } else {
      orderBy.receivedAt = 'desc';
    }

    const [total, candidates] = await Promise.all([
      prisma.candidate.count({ where }),
      prisma.candidate.findMany({
        where,
        orderBy,
        skip: (page - 1) * pageSize,
        take: pageSize,
      }),
    ]);

    const items = candidates.map((c) => ({
      id: c.id,
      name: c.name,
      email: c.email,
      phone: c.phone,
      college: c.college,
      degree: c.degree,
      graduationYear: c.graduationYear,
      skills: c.skills,
      experienceYears: c.experienceYears,
      appliedRole: c.appliedRole,
      receivedAt: c.receivedAt.toISOString(),
      processingStatus: c.processingStatus as any,
      interviewStatus: c.interviewStatus as any,
      updatedAt: c.updatedAt.toISOString(),
      // Never expose raw file-system paths. Provide a boolean so the UI can
      // show the Download button, and a safe identifier for the download URL.
      resumeRef: hasResume(c.resumeRef) ? c.id : null,
      resumeFileName: c.resumeFileName,
      sourceEmailId: c.sourceEmailId,
      reviewFields: c.reviewFields,
      isSample: false,
    }));

    res.json({
      items,
      total,
      page,
      pageSize,
    });
  } catch (err) {
    next(err);
  }
});

// GET /api/candidates/:id
candidateRouter.get('/:id', async (req: Request, res: Response, next: NextFunction) => {
  try {
    const { id } = req.params;

    const candidate = await prisma.candidate.findUnique({
      where: { id },
      include: {
        applications: { orderBy: { submissionDate: 'desc' } },
      },
    });

    if (!candidate) {
      res.status(404).json({
        error: {
          code: 'NOT_FOUND',
          message: `Candidate ${id} was not found.`,
          requestId: req.id,
        },
      });
      return;
    }

    // Fetch related source email if known
    let sourceEmail: { from: string; subject: string; receivedAt: string; messageId: string; bodyText?: string | null } | null = null;
    let processingHistory: { at: string; step: string; status: any; message?: string }[] = [];

    if (candidate.sourceEmailId) {
      const email = await prisma.emailMessage.findFirst({
        where: {
          OR: [
            { id: candidate.sourceEmailId },
            { messageId: candidate.sourceEmailId },
          ],
        },
      });

      if (email) {
        const { getCleanEmailBody } = await import('../lib/emailService');
        const bodyText = await getCleanEmailBody(email.id);

        sourceEmail = {
          from: email.fromEmail,
          subject: email.subject,
          receivedAt: email.receivedAt.toISOString(),
          messageId: email.messageId,
          bodyText,
        };
        if (Array.isArray(email.history)) {
          processingHistory = email.history as any;
        }
      }
    }

    // Fetch audit entries related to this candidate
    const auditRecords = await prisma.auditLog.findMany({
      where: { related: id },
      orderBy: { createdAt: 'desc' },
      take: 20,
    });

    const audit = auditRecords.map((a) => ({
      id: a.id,
      at: a.createdAt.toISOString(),
      actor: a.actor,
      action: a.action,
      related: a.related,
      result: a.result as any,
      requestId: a.requestId,
      isSample: false,
    }));

    const formattedApplications = candidate.applications.map((a) => ({
      id: a.id,
      candidateId: a.candidateId,
      candidateName: candidate.name,
      role: a.role,
      sourceEmail: a.sourceEmail,
      receivedAt: a.submissionDate.toISOString(),
      status: a.status as any,
      recruiter: a.recruiter,
      updatedAt: a.updatedAt.toISOString(),
      isSample: false,
    }));

    res.json({
      id: candidate.id,
      name: candidate.name,
      email: candidate.email,
      phone: candidate.phone,
      college: candidate.college,
      degree: candidate.degree,
      graduationYear: candidate.graduationYear,
      skills: candidate.skills,
      experienceYears: candidate.experienceYears,
      appliedRole: candidate.appliedRole,
      receivedAt: candidate.receivedAt.toISOString(),
      processingStatus: candidate.processingStatus as any,
      interviewStatus: candidate.interviewStatus as any,
      updatedAt: candidate.updatedAt.toISOString(),
      // Expose a safe non-path resumeRef (just the candidate id as token) so the
      // frontend can render the Download button. Raw file paths are never sent.
      resumeRef: hasResume(candidate.resumeRef) ? candidate.id : null,
      resumeFileName: candidate.resumeFileName,
      sourceEmailId: candidate.sourceEmailId,
      reviewFields: candidate.reviewFields,
      sourceEmail,
      applications: formattedApplications,
      processingHistory,
      audit,
      isSample: false,
    });
  } catch (err) {
    next(err);
  }
});

// PATCH /api/candidates/:id
candidateRouter.patch('/:id', async (req: Request, res: Response, next: NextFunction) => {
  try {
    const { id } = req.params;
    const patch = candidatePatchSchema.parse(req.body);

    const existing = await prisma.candidate.findUnique({ where: { id } });
    if (!existing) {
      res.status(404).json({
        error: {
          code: 'NOT_FOUND',
          message: 'Candidate not found',
          requestId: req.id,
        },
      });
      return;
    }

    // Filter reviewFields to remove any field that is now resolved
    const patchKeys = Object.keys(patch);
    const updatedReviewFields = existing.reviewFields.filter(
      (f) => !patchKeys.includes(f) || (patch as any)[f] == null
    );

    const updated = await prisma.candidate.update({
      where: { id },
      data: {
        ...patch,
        reviewFields: updatedReviewFields,
        // If all review fields are resolved and processing was needs_review, promote to processed
        processingStatus:
          updatedReviewFields.length === 0 && existing.processingStatus === 'needs_review'
            ? 'processed'
            : existing.processingStatus,
      },
    });

    // Record audit log
    await prisma.auditLog.create({
      data: {
        actor: 'recruiter',
        action: 'Candidate details updated',
        related: id,
        result: 'success',
        requestId: String(req.id || ''),
        entity: 'Candidate',
        entityId: id,
        payload: patch as any,
      },
    });

    res.json({
      id: updated.id,
      name: updated.name,
      email: updated.email,
      phone: updated.phone,
      college: updated.college,
      degree: updated.degree,
      graduationYear: updated.graduationYear,
      skills: updated.skills,
      experienceYears: updated.experienceYears,
      appliedRole: updated.appliedRole,
      receivedAt: updated.receivedAt.toISOString(),
      processingStatus: updated.processingStatus as any,
      interviewStatus: updated.interviewStatus as any,
      updatedAt: updated.updatedAt.toISOString(),
      resumeRef: hasResume(updated.resumeRef) ? updated.id : null,
      resumeFileName: updated.resumeFileName,
      sourceEmailId: updated.sourceEmailId,
      reviewFields: updated.reviewFields,
      isSample: false,
    });
  } catch (err) {
    next(err);
  }
});

// GET /api/candidates/:id/resume - Direct CV download/view endpoint
candidateRouter.get('/:id/resume', async (req: Request, res: Response, next: NextFunction) => {
  try {
    const { id } = req.params;
    const candidate = await prisma.candidate.findUnique({
      where: { id },
      include: { attachments: true },
    });

    if (!candidate) {
      res.status(404).json({ error: { code: 'NOT_FOUND', message: 'Candidate not found' } });
      return;
    }

    const filename = candidate.resumeFileName || 'resume.docx';
    const fsModule = await import('fs');

    // 1. If candidate has resumeRef — try local path first, then OneDrive DriveItem ID
    if (candidate.resumeRef) {
      // Attempt to resolve as a local storage key (relative or absolute).
      // resolveLocalAttachmentPath validates against ATTACHMENTS_STORAGE_DIR sandbox.
      const localPath = resolveLocalAttachmentPath(candidate.resumeRef);
      if (localPath && fsModule.existsSync(localPath)) {
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.sendFile(localPath);
        return;
      }

      // Otherwise treat resumeRef as an OneDrive DriveItem ID and download via Microsoft Graph
      try {
        const { downloadResumeFromOneDrive } = await import('../lib/attachmentService');
        const downloaded = await downloadResumeFromOneDrive(candidate.resumeRef);
        if (downloaded && downloaded.buffer && downloaded.buffer.length > 0) {
          res.setHeader('Content-Type', downloaded.contentType || 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
          res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
          res.send(downloaded.buffer);
          return;
        }
      } catch (_graphErr: any) {
        // Fall through to local attachment fallback
      }
    }

    // 2. Check candidate's linked local attachment as fallback (only attachments owned by this candidate)
    const attachment = candidate.attachments?.find((a) => a.candidateId === candidate.id && a.storageKey);
    if (attachment && attachment.storageKey) {
      const localPath = resolveLocalAttachmentPath(attachment.storageKey);
      if (localPath && fsModule.existsSync(localPath)) {
        res.setHeader('Content-Disposition', `attachment; filename="${attachment.filename || filename}"`);
        res.sendFile(localPath);
        return;
      }
    }

    // 3. Fallback redirect if candidate has a valid resumeUrl (public OneDrive link)
    if (candidate.resumeUrl) {
      res.redirect(candidate.resumeUrl);
      return;
    }

    res.status(404).json({ error: { code: 'NOT_FOUND', message: 'Candidate resume file not available from OneDrive or local storage' } });
  } catch (err) {
    next(err);
  }
});
