import pino from 'pino';
import { config } from './config';

export const logger = pino({
  level: config.NODE_ENV === 'test' ? 'silent' : config.NODE_ENV === 'production' ? 'info' : 'debug',
  base: {
    service: 'recruitment-backend',
    env: config.NODE_ENV,
  },
  timestamp: pino.stdTimeFunctions.isoTime,
});
