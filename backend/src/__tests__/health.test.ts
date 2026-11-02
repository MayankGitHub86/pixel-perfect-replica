import { describe, it, expect, vi, beforeEach } from 'vitest';
import request from 'supertest';
import { createApp } from '../app';
import * as dbModule from '../lib/db';
import * as redisModule from '../lib/redis';

describe('Health and Readiness Endpoints', () => {
  const app = createApp();

  beforeEach(() => {
    vi.restoreAllMocks();
  });

  describe('GET /api/health', () => {
    it('should return 200 with status ok and valid metadata', async () => {
      const res = await request(app).get('/api/health');

      expect(res.status).toBe(200);
      expect(res.body.status).toBe('ok');
      expect(res.body.timestamp).toBeDefined();
      expect(res.body.uptime).toBeGreaterThanOrEqual(0);
      expect(res.headers['x-request-id']).toBeDefined();
    });

    it('should respect incoming X-Request-Id header', async () => {
      const customId = 'custom-request-id-12345';
      const res = await request(app)
        .get('/api/health')
        .set('X-Request-Id', customId);

      expect(res.status).toBe(200);
      expect(res.headers['x-request-id']).toBe(customId);
      expect(res.body.requestId).toBe(customId);
    });
  });

  describe('GET /api/ready', () => {
    it('should return 200 and ready status when both DB and Redis are up', async () => {
      vi.spyOn(dbModule, 'checkDatabaseConnection').mockResolvedValue(true);
      vi.spyOn(redisModule, 'checkRedisConnection').mockResolvedValue(true);

      const res = await request(app).get('/api/ready');

      expect(res.status).toBe(200);
      expect(res.body.status).toBe('ready');
      expect(res.body.checks).toEqual({
        database: 'up',
        redis: 'up',
      });
    });

    it('should return 503 when database is unavailable', async () => {
      vi.spyOn(dbModule, 'checkDatabaseConnection').mockResolvedValue(false);
      vi.spyOn(redisModule, 'checkRedisConnection').mockResolvedValue(true);

      const res = await request(app).get('/api/ready');

      expect(res.status).toBe(503);
      expect(res.body.status).toBe('unavailable');
      expect(res.body.checks).toEqual({
        database: 'down',
        redis: 'up',
      });
    });

    it('should return 503 when Redis is unavailable', async () => {
      vi.spyOn(dbModule, 'checkDatabaseConnection').mockResolvedValue(true);
      vi.spyOn(redisModule, 'checkRedisConnection').mockResolvedValue(false);

      const res = await request(app).get('/api/ready');

      expect(res.status).toBe(503);
      expect(res.body.status).toBe('unavailable');
      expect(res.body.checks).toEqual({
        database: 'up',
        redis: 'down',
      });
    });

    it('should return 503 when both database and Redis are unavailable', async () => {
      vi.spyOn(dbModule, 'checkDatabaseConnection').mockResolvedValue(false);
      vi.spyOn(redisModule, 'checkRedisConnection').mockResolvedValue(false);

      const res = await request(app).get('/api/ready');

      expect(res.status).toBe(503);
      expect(res.body.status).toBe('unavailable');
      expect(res.body.checks).toEqual({
        database: 'down',
        redis: 'down',
      });
    });
  });
});

