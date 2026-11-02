// backend/src/middleware/auth.ts
// Simple security middleware for integration endpoints.
// In development, allow all localhost requests.
// In other environments, block access (return 403) unless explicitly enabled via env.

import { Request, Response, NextFunction } from 'express';
import { config } from '../config';

export function integrationAuthMiddleware(req: Request, res: Response, next: NextFunction) {
  // Allow in development and test environments.
  if (config.NODE_ENV === 'development' || config.NODE_ENV === 'test') {
    return next();
  }
  // Production / other envs: block integration routes.
  res.status(403).json({
    error: {
      code: 'FORBIDDEN',
      message: 'Integration endpoints are disabled in this environment.',
    },
  });
}

