import { PrismaClient } from '@prisma/client';
import { logger } from '../logger';

export const prisma = new PrismaClient({
  log: process.env.NODE_ENV === 'development' ? ['query', 'error', 'warn'] : ['error'],
});

export async function checkDatabaseConnection(timeoutMs = 2000): Promise<boolean> {
  try {
    const checkPromise = prisma.$queryRaw`SELECT 1`;
    const timeoutPromise = new Promise((_, reject) =>
      setTimeout(() => reject(new Error('Database check timed out')), timeoutMs)
    );
    await Promise.race([checkPromise, timeoutPromise]);
    return true;
  } catch (error) {
    logger.warn({ error }, 'Database connection check failed or timed out');
    return false;
  }
}
