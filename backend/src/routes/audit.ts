import { Router, Request, Response, NextFunction } from 'express';
import { prisma } from '../lib/db';
import { auditListQuerySchema } from '../validation/schemas';
import { Prisma } from '@prisma/client';

export const auditRouter = Router();

// GET /api/audit-logs
auditRouter.get('/', async (req: Request, res: Response, next: NextFunction) => {
  try {
    const q = auditListQuerySchema.parse(req.query);
    const { page, pageSize, search, result, sortBy, sortDir } = q;

    const where: Prisma.AuditLogWhereInput = {};

    if (result && (result as string) !== 'all') {
      where.result = result;
    }

    if (search) {
      where.OR = [
        { actor: { contains: search, mode: 'insensitive' } },
        { action: { contains: search, mode: 'insensitive' } },
        { related: { contains: search, mode: 'insensitive' } },
        { requestId: { contains: search, mode: 'insensitive' } },
      ];
    }

    const orderBy: Prisma.AuditLogOrderByWithRelationInput = {};
    if (sortBy && ['actor', 'action', 'result', 'createdAt'].includes(sortBy)) {
      orderBy[sortBy as keyof Prisma.AuditLogOrderByWithRelationInput] = sortDir;
    } else {
      orderBy.createdAt = 'desc';
    }

    const [total, logs] = await Promise.all([
      prisma.auditLog.count({ where }),
      prisma.auditLog.findMany({
        where,
        orderBy,
        skip: (page - 1) * pageSize,
        take: pageSize,
      }),
    ]);

    const items = logs.map((l) => ({
      id: l.id,
      at: l.createdAt.toISOString(),
      actor: l.actor,
      action: l.action,
      related: l.related,
      result: l.result as any,
      requestId: l.requestId,
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
