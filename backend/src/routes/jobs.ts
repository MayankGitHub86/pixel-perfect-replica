import { Router, Request, Response, NextFunction } from 'express';
import { prisma } from '../lib/db';
import { jobListQuerySchema } from '../validation/schemas';
import { Prisma } from '@prisma/client';

export const jobRouter = Router();

// GET /api/jobs
jobRouter.get('/', async (req: Request, res: Response, next: NextFunction) => {
  try {
    const q = jobListQuerySchema.parse(req.query);
    const { page, pageSize, search, status, type, sortBy, sortDir } = q;

    const where: Prisma.ProcessingJobWhereInput = {};

    if (status && (status as string) !== 'all') {
      where.status = status;
    }

    if (type && (type as string) !== 'all') {
      where.type = type;
    }

    if (search) {
      where.OR = [
        { related: { contains: search, mode: 'insensitive' } },
        { id: { contains: search, mode: 'insensitive' } },
        { type: { contains: search, mode: 'insensitive' } },
      ];
    }

    const orderBy: Prisma.ProcessingJobOrderByWithRelationInput = {};
    if (sortBy && ['type', 'status', 'attempts', 'startedAt', 'createdAt'].includes(sortBy)) {
      orderBy[sortBy as keyof Prisma.ProcessingJobOrderByWithRelationInput] = sortDir;
    } else {
      orderBy.createdAt = 'desc';
    }

    const [total, jobs] = await Promise.all([
      prisma.processingJob.count({ where }),
      prisma.processingJob.findMany({
        where,
        orderBy,
        skip: (page - 1) * pageSize,
        take: pageSize,
      }),
    ]);

    const items = jobs.map((j) => ({
      id: j.id,
      type: j.type as any,
      related: j.related,
      status: j.status as any,
      attempts: j.attempts,
      startedAt: j.startedAt?.toISOString() ?? null,
      completedAt: j.completedAt?.toISOString() ?? null,
      lastError: j.lastError,
      nextRetryAt: j.nextRetryAt?.toISOString() ?? null,
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

// POST /api/jobs/:id/retry
jobRouter.post('/:id/retry', async (req: Request, res: Response, next: NextFunction) => {
  try {
    const { id } = req.params;

    const existing = await prisma.processingJob.findUnique({ where: { id } });
    if (!existing) {
      res.status(404).json({
        error: {
          code: 'NOT_FOUND',
          message: `Job ${id} was not found`,
          requestId: req.id,
        },
      });
      return;
    }

    const updated = await prisma.processingJob.update({
      where: { id },
      data: {
        status: 'queued',
        attempts: existing.attempts + 1,
        lastError: null,
      },
    });

    await prisma.auditLog.create({
      data: {
        actor: 'recruiter',
        action: `Job retry triggered for ${updated.type}`,
        related: updated.related,
        result: 'info',
        requestId: String(req.id || ''),
        entity: 'ProcessingJob',
        entityId: id,
        payload: { jobType: updated.type, attempt: updated.attempts },
      },
    });

    res.json({
      id: updated.id,
      type: updated.type as any,
      related: updated.related,
      status: updated.status as any,
      attempts: updated.attempts,
      startedAt: updated.startedAt?.toISOString() ?? null,
      completedAt: updated.completedAt?.toISOString() ?? null,
      lastError: updated.lastError,
      nextRetryAt: updated.nextRetryAt?.toISOString() ?? null,
      isSample: false,
    });
  } catch (err) {
    next(err);
  }
});
