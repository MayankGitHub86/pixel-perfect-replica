import express from 'express';
import cors from 'cors';
import pinoHttp from 'pino-http';
import cookieParser from 'cookie-parser';

import { config } from './config';
import { logger } from './logger';
import { requestIdMiddleware } from './middleware/requestId';
import { errorHandler, notFoundHandler } from './middleware/error';
import { healthHandler, readyHandler } from './routes/health';
import { candidateRouter } from './routes/candidates';
import { applicationRouter } from './routes/applications';
import { emailRouter } from './routes/emails';
import { jobRouter } from './routes/jobs';
import { excelRouter } from './routes/excel';
import { integrationRouter } from './routes/integrations';
import { auditRouter } from './routes/audit';

export function createApp(): express.Application {
  const app = express();

  app.use(cookieParser());


  // CORS configuration for frontend
  const allowedOrigins = config.CORS_ORIGIN.split(',').map((o) => o.trim());
  app.use(
    cors({
      origin: (origin, callback) => {
        if (!origin || allowedOrigins.includes(origin) || allowedOrigins.includes('*')) {
          callback(null, true);
        } else {
          callback(new Error(`CORS origin ${origin} not permitted`));
        }
      },
      credentials: true,
    })
  );

  app.use(express.json({ limit: '10mb' }));
  app.use(requestIdMiddleware);

  if (config.NODE_ENV !== 'test') {
    app.use(
      pinoHttp({
        logger: logger as any,
        genReqId: (req) => req.id || req.headers['x-request-id']?.toString() || 'anonymous',
        customLogLevel: (req, res, err) => {
          if (res.statusCode >= 500 || err) return 'error';
          if (res.statusCode >= 400) return 'warn';
          return 'info';
        },
      })
    );
  }

  // Health and Readiness
  app.get('/api/health', healthHandler);
  app.get('/api/ready', readyHandler);

  // Business Domain APIs
  app.use('/api/candidates', candidateRouter);
  app.use('/api/applications', applicationRouter);
  app.use('/api/emails', emailRouter);
  app.use('/api/jobs', jobRouter);
  app.use('/api/excel-sync', excelRouter);
  app.use('/api/integrations', integrationRouter);
  app.use('/api/audit-logs', auditRouter);

  // 404 and Centralized Error Handling
  app.use(notFoundHandler);
  app.use(errorHandler);

  return app;
}

export const app = createApp();
export default app;
