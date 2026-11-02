import IORedis from 'ioredis';
import { config } from '../config';
import { logger } from '../logger';

export const redis = new IORedis(config.REDIS_URL, {
  maxRetriesPerRequest: null,
  lazyConnect: true,
  enableReadyCheck: false,
});

redis.on('error', (err) => {
  if (config.NODE_ENV !== 'test') {
    logger.warn({ err: err.message }, 'Redis connection warning');
  }
});

export async function checkRedisConnection(timeoutMs = 2000): Promise<boolean> {
  let probeClient: IORedis | null = null;
  try {
    if (redis.status === 'ready') {
      const pong = await Promise.race([
        redis.ping(),
        new Promise<string>((_, reject) =>
          setTimeout(() => reject(new Error('Redis ping timeout')), timeoutMs)
        ),
      ]);
      return pong === 'PONG';
    }

    probeClient = new IORedis(config.REDIS_URL, {
      connectTimeout: timeoutMs,
      commandTimeout: timeoutMs,
      enableOfflineQueue: false,
      maxRetriesPerRequest: 1,
      retryStrategy: () => null,
      lazyConnect: true,
    });

    await probeClient.connect();
    const pong = await probeClient.ping();
    await probeClient.quit().catch(() => {});
    return pong === 'PONG';
  } catch (error) {
    if (probeClient) {
      probeClient.disconnect();
    }
    logger.warn({ error }, 'Redis readiness check failed or timed out');
    return false;
  }
}
