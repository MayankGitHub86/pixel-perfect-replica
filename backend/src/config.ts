import { z } from 'zod';
import dotenv from 'dotenv';
import path from 'path';

// Load .env
dotenv.config({ path: path.resolve(process.cwd(), '.env') });

const envSchema = z.object({
  PORT: z.coerce.number().default(3000),
  NODE_ENV: z.enum(['development', 'production', 'test']).default('development'),
  DATABASE_URL: z.string().url(),
  REDIS_URL: z.string().url(),
  CORS_ORIGIN: z.string().default('http://localhost:5173,http://localhost:8080'),

  // Public backend URL used in Excel hyperlinks and external references.
  // In production, set to the actual deployed HTTPS domain (e.g. https://api.recruitflow.example.com).
  // In development, defaults to http://localhost:3000.
  // If empty in production, Excel resume hyperlinks are left blank to prevent localhost URLs in production workbooks.
  BACKEND_PUBLIC_URL: z.string().default(''),

  // Microsoft OAuth (required)
  MICROSOFT_CLIENT_ID: z.string(),
  MICROSOFT_CLIENT_SECRET: z.string(),
  MICROSOFT_TENANT_ID: z.string(),
  MICROSOFT_REDIRECT_URI: z.string().url().optional().default('http://localhost:3000/api/integrations/microsoft/callback'),

  // Excel Integration (Phase 5 placeholders)
  EXCEL_WORKBOOK_ID: z.string().optional().default('placeholder_workbook_id'),
  EXCEL_TABLE_NAME: z.string().default('CandidateMaster'),
});

export const config = envSchema.parse(process.env);
export type Config = z.infer<typeof envSchema>;
