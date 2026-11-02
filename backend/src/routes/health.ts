import { Request, Response } from 'express';
import { checkDatabaseConnection } from '../lib/db';
import { checkRedisConnection } from '../lib/redis';

export const healthHandler = (req: Request, res: Response): void => {
  res.status(200).json({
    status: 'ok',
    timestamp: new Date().toISOString(),
    uptime: process.uptime(),
    requestId: req.id,
  });
};

export const readyHandler = async (req: Request, res: Response): Promise<void> => {
  const [dbHealthy, redisHealthy] = await Promise.all([
    checkDatabaseConnection(),
    checkRedisConnection(),
  ]);

  const allReady = dbHealthy && redisHealthy;
  const status = allReady ? 'ready' : 'unavailable';
  const statusCode = allReady ? 200 : 503;

  res.status(statusCode).json({
    status,
    timestamp: new Date().toISOString(),
    requestId: req.id,
    checks: {
      database: dbHealthy ? 'up' : 'down',
      redis: redisHealthy ? 'up' : 'down',
    },
  });
};
