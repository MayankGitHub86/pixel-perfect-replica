import { Request, Response, NextFunction } from 'express';
import { logger } from '../logger';
import { config } from '../config';
import { ZodError } from 'zod';

export interface AppError extends Error {
  statusCode?: number;
  code?: string;
  details?: unknown;
}

export function notFoundHandler(req: Request, res: Response): void {
  res.status(404).json({
    error: {
      code: 'NOT_FOUND',
      message: `Cannot ${req.method} ${req.originalUrl}`,
      requestId: req.id,
    },
  });
}

export function errorHandler(
  err: AppError | ZodError,
  req: Request,
  res: Response,
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  next: NextFunction
): void {
  const requestId = req.id;

  if (err instanceof ZodError) {
    logger.warn({ requestId, issues: err.issues }, 'Validation failed');
    res.status(400).json({
      error: {
        code: 'VALIDATION_ERROR',
        message: 'Request payload validation failed',
        details: err.issues.map((i) => ({ path: i.path.join('.'), message: i.message })),
        requestId,
      },
    });
    return;
  }

  const statusCode = err.statusCode || 500;
  const code = err.code || (statusCode === 500 ? 'INTERNAL_SERVER_ERROR' : 'BAD_REQUEST');
  const message = statusCode === 500 && config.NODE_ENV === 'production' 
    ? 'An unexpected error occurred' 
    : err.message || 'Unknown error';

  logger.error(
    {
      requestId,
      statusCode,
      code,
      message: err.message,
      stack: config.NODE_ENV === 'development' ? err.stack : undefined,
    },
    'Request error'
  );

  res.status(statusCode).json({
    error: {
      code,
      message,
      requestId,
    },
  });
}

