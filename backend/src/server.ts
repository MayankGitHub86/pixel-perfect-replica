import { app } from './app';
import { config } from './config';
import { logger } from './logger';
import { prisma } from './lib/db';
import { redis } from './lib/redis';
import './lib/queue';

// Periodic Outlook inbox sync scheduler (every 60 seconds)
const SYNC_INTERVAL_MS = 60 * 1000;
const syncTimer = setInterval(async () => {
  try {
    const { syncOutlookInbox } = await import('./lib/outlookSync');
    const res = await syncOutlookInbox({ batchSize: 10, requestId: 'scheduler' });
    if (res.newIngested > 0) {
      logger.info({ newIngested: res.newIngested }, 'Scheduled Outlook sync ingested new messages');
    }
  } catch (err: any) {
    logger.debug({ err: err.message }, 'Scheduled Outlook sync poll');
  }
}, SYNC_INTERVAL_MS);

const server = app.listen(config.PORT, () => {
  logger.info(
    {
      port: config.PORT,
      env: config.NODE_ENV,
      cors: config.CORS_ORIGIN,
      pollingIntervalMs: SYNC_INTERVAL_MS,
    },
    'Recruitment backend service started successfully with background workers and scheduler'
  );
});

let isShuttingDown = false;

async function handleShutdown(signal: string): Promise<void> {
  if (isShuttingDown) return;
  isShuttingDown = true;

  if (syncTimer) clearInterval(syncTimer);
  logger.info({ signal }, 'Graceful shutdown initiated');

  server.close(async () => {
    logger.info('HTTP server closed');

    try {
      await prisma.$disconnect();
      logger.info('Prisma disconnected');
    } catch (err) {
      logger.error({ err }, 'Error disconnecting Prisma');
    }

    try {
      await redis.quit();
      logger.info('Redis connection closed');
    } catch (err) {
      logger.error({ err }, 'Error closing Redis connection');
    }

    logger.info('Shutdown complete');
    process.exit(0);
  });

  // Force exit after 10 seconds if graceful shutdown times out
  setTimeout(() => {
    logger.error('Graceful shutdown timed out, force exiting');
    process.exit(1);
  }, 10000).unref();
}

process.on('SIGTERM', () => handleShutdown('SIGTERM'));
process.on('SIGINT', () => handleShutdown('SIGINT'));

export default server;
