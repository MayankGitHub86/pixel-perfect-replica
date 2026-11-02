import { Router, Request, Response, NextFunction } from 'express';
import { prisma } from '../lib/db';
import { applicationListQuerySchema, applicationStatusPatchSchema } from '../validation/schemas';
import { Prisma } from '@prisma/client';

export const applicationRouter = Router();

// GET /api/applications
applicationRouter.get('/', async (req: Request, res: Response, next: NextFunction) => {
  try {
    const q = applicationListQuerySchema.parse(req.query);
    const { page, pageSize, search, status, sortBy, sortDir, candidateId } = q;

    const where: Prisma.ApplicationWhereInput = {};

    if (candidateId) {
      where.candidateId = candidateId;
    }

    if (status && status !== 'all') {
      where.status = status;
    }

    if (search) {
      where.OR = [
        { role: { contains: search, mode: 'insensitive' } },
        { sourceEmail: { contains: search, mode: 'insensitive' } },
        { candidate: { name: { contains: search, mode: 'insensitive' } } },
      ];
    }

    const orderBy: Prisma.ApplicationOrderByWithRelationInput = {};
    if (sortBy && ['role', 'status', 'submissionDate', 'createdAt', 'updatedAt'].includes(sortBy)) {
      orderBy[sortBy as keyof Prisma.ApplicationOrderByWithRelationInput] = sortDir;
    } else {
      orderBy.submissionDate = 'desc';
    }

    const [total, applications] = await Promise.all([
      prisma.application.count({ where }),
      prisma.application.findMany({
        where,
        orderBy,
        skip: (page - 1) * pageSize,
        take: pageSize,
        include: {
          candidate: { select: { name: true } },
        },
      }),
    ]);

    const items = applications.map((a) => ({
      id: a.id,
      candidateId: a.candidateId,
      candidateName: a.candidate.name,
      role: a.role,
      sourceEmail: a.sourceEmail,
      receivedAt: a.submissionDate.toISOString(),
      status: a.status as any,
      recruiter: a.recruiter,
      updatedAt: a.updatedAt.toISOString(),
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

// PATCH /api/applications/:id/status
applicationRouter.patch('/:id/status', async (req: Request, res: Response, next: NextFunction) => {
  try {
    const { id } = req.params;
    const { status } = applicationStatusPatchSchema.parse(req.body);

    const existing = await prisma.application.findUnique({
      where: { id },
      include: { candidate: { select: { name: true } } },
    });

    if (!existing) {
      res.status(404).json({
        error: {
          code: 'NOT_FOUND',
          message: 'Application not found',
          requestId: req.id,
        },
      });
      return;
    }

    const updated = await prisma.application.update({
      where: { id },
      data: { status },
      include: { candidate: { select: { name: true } } },
    });

    // Record audit log
    await prisma.auditLog.create({
      data: {
        actor: 'recruiter',
        action: `Application status changed to ${status}`,
        related: existing.candidateId,
        result: 'success',
        requestId: String(req.id || ''),
        entity: 'Application',
        entityId: id,
        payload: { previousStatus: existing.status, newStatus: status },
      },
    });

    res.json({
      id: updated.id,
      candidateId: updated.candidateId,
      candidateName: updated.candidate.name,
      role: updated.role,
      sourceEmail: updated.sourceEmail,
      receivedAt: updated.submissionDate.toISOString(),
      status: updated.status as any,
      recruiter: updated.recruiter,
      updatedAt: updated.updatedAt.toISOString(),
      isSample: false,
    });
  } catch (err) {
    next(err);
  }
});
