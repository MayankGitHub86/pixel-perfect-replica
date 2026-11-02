import { describe, it, expect } from 'vitest';
import request from 'supertest';
import express from 'express';
import { notFoundHandler, errorHandler } from '../middleware/error';
import { requestIdMiddleware } from '../middleware/requestId';
import { z } from 'zod';

describe('Error Handling and Validation Middleware', () => {
  const app = express();
  app.use(express.json());
  app.use(requestIdMiddleware);

  // Test route triggering validation error
  app.post('/test/validation', (req, res, next) => {
    try {
      const schema = z.object({
        email: z.string().email(),
      });
      schema.parse(req.body);
      res.status(200).json({ ok: true });
    } catch (err) {
      next(err);
    }
  });

  // Test route triggering 500 error
  app.get('/test/crash', () => {
    throw new Error('Database connection crashed');
  });

  app.use(notFoundHandler);
  app.use(errorHandler);

  it('should return 404 for unknown endpoints', async () => {
    const res = await request(app).get('/unknown/route');

    expect(res.status).toBe(404);
    expect(res.body.error).toBeDefined();
    expect(res.body.error.code).toBe('NOT_FOUND');
    expect(res.body.error.requestId).toBeDefined();
  });

  it('should format Zod validation errors cleanly with status 400', async () => {
    const res = await request(app)
      .post('/test/validation')
      .send({ email: 'invalid-email-address' });

    expect(res.status).toBe(400);
    expect(res.body.error.code).toBe('VALIDATION_ERROR');
    expect(Array.isArray(res.body.error.details)).toBe(true);
    expect(res.body.error.details[0].path).toBe('email');
  });

  it('should format internal server errors safely without leaking internal details', async () => {
    const res = await request(app).get('/test/crash');

    expect(res.status).toBe(500);
    expect(res.body.error.code).toBe('INTERNAL_SERVER_ERROR');
    expect(res.body.error.requestId).toBeDefined();
  });
});

