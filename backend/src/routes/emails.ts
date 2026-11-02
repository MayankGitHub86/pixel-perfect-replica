import { Router, Request, Response, NextFunction } from 'express';
import { prisma } from '../lib/db';
import { emailListQuerySchema } from '../validation/schemas';
import { Prisma } from '@prisma/client';

export const emailRouter = Router();

// GET /api/emails
emailRouter.get('/', async (req: Request, res: Response, next: NextFunction) => {
  try {
    const q = emailListQuerySchema.parse(req.query);
    const { page, pageSize, search, status, sortBy, sortDir } = q;

    const where: Prisma.EmailMessageWhereInput = {};

    if (status && (status as string) !== 'all') {
      where.status = status;
    }

    if (search) {
      where.OR = [
        { fromEmail: { contains: search, mode: 'insensitive' } },
        { subject: { contains: search, mode: 'insensitive' } },
      ];
    }

    const orderBy: Prisma.EmailMessageOrderByWithRelationInput = {};
    if (sortBy && ['subject', 'fromEmail', 'receivedAt', 'status'].includes(sortBy)) {
      orderBy[sortBy as keyof Prisma.EmailMessageOrderByWithRelationInput] = sortDir;
    } else {
      orderBy.receivedAt = 'desc';
    }

    const [total, emails] = await Promise.all([
      prisma.emailMessage.count({ where }),
      prisma.emailMessage.findMany({
        where,
        orderBy,
        skip: (page - 1) * pageSize,
        take: pageSize,
        include: {
          attachments: true,
        },
      }),
    ]);

    const items = emails.map((e) => ({
      id: e.id,
      from: e.fromEmail,
      subject: e.subject,
      receivedAt: e.receivedAt.toISOString(),
      attachments: e.attachments.map((a) => ({
        id: a.id,
        fileName: a.filename,
        mimeType: a.mimeType,
        sizeBytes: a.sizeBytes,
      })),
      status: e.status as any,
      candidateId: e.candidateId,
      applicationId: e.applicationId,
      error: e.error,
      history: Array.isArray(e.history) ? (e.history as any) : [],
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

// GET /api/emails/attachments/list
emailRouter.get('/attachments/list', async (req: Request, res: Response, next: NextFunction) => {
  try {
    const attachments = await prisma.attachment.findMany({
      orderBy: { createdAt: 'desc' },
      take: 50,
      include: {
        email: {
          select: { subject: true, fromEmail: true },
        },
      },
    });

    res.json({
      items: attachments.map((a) => ({
        id: a.id,
        emailId: a.emailId,
        fileName: a.filename,
        mimeType: a.mimeType,
        sizeBytes: a.sizeBytes,
        status: a.status,
        emailSubject: a.email?.subject,
        emailFrom: a.email?.fromEmail,
        createdAt: a.createdAt.toISOString(),
      })),
      total: attachments.length,
    });
  } catch (err) {
    next(err);
  }
});

// GET /api/emails/:id
emailRouter.get('/:id', async (req: Request, res: Response, next: NextFunction) => {
  try {
    const { id } = req.params;
    const email = await prisma.emailMessage.findUnique({
      where: { id },
      include: {
        attachments: true,
        processingJob: true,
      },
    });

    if (!email) {
      res.status(404).json({
        error: {
          code: 'NOT_FOUND',
          message: `Email ${id} was not found`,
          requestId: req.id,
        },
      });
      return;
    }

    res.json({
      id: email.id,
      from: email.fromEmail,
      subject: email.subject,
      receivedAt: email.receivedAt.toISOString(),
      status: email.status as any,
      candidateId: email.candidateId,
      applicationId: email.applicationId,
      error: email.error,
      history: Array.isArray(email.history) ? email.history : [],
      attachments: email.attachments.map((a) => ({
        id: a.id,
        fileName: a.filename,
        mimeType: a.mimeType,
        sizeBytes: a.sizeBytes,
        status: a.status,
        storageKey: a.storageKey,
      })),
      isSample: false,
    });
  } catch (err) {
    next(err);
  }
});
