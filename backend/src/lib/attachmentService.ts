// backend/src/lib/attachmentService.ts
// Secure attachment detection, validation, and storage service.

import path from 'path';
import fs from 'fs';
import { logger } from '../logger';

export interface AttachmentValidationResult {
  valid: boolean;
  mimeType: string;
  extension: string;
  sizeBytes: number;
  reason?: string;
}

export const ALLOWED_EXTENSIONS = new Set(['.pdf', '.docx', '.doc']);
export const ALLOWED_MIME_TYPES = new Set([
  'application/pdf',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/msword',
  'application/octet-stream', // often returned by Microsoft Graph for attachments
]);

export const FORBIDDEN_EXTENSIONS = new Set([
  '.exe', '.bat', '.cmd', '.sh', '.bin', '.js', '.vbs', '.msi',
  '.zip', '.rar', '.7z', '.tar', '.gz', '.iso', '.dmg',
]);

export const MAX_ATTACHMENT_SIZE_BYTES = 10 * 1024 * 1024; // 10MB
export const ATTACHMENTS_STORAGE_DIR = path.resolve(__dirname, '../../storage/attachments');

/**
 * Validate an attachment's filename, MIME type, and byte size.
 */
export function validateAttachment(
  filename: string,
  mimeType: string,
  sizeBytes: number
): AttachmentValidationResult {
  const ext = path.extname(filename || '').toLowerCase();

  // 1. Check for dangerous / executable extensions
  if (FORBIDDEN_EXTENSIONS.has(ext)) {
    return {
      valid: false,
      mimeType,
      extension: ext,
      sizeBytes,
      reason: `Executable or archive file extension '${ext}' is rejected for security.`,
    };
  }

  // 2. Check for supported resume formats
  if (!ALLOWED_EXTENSIONS.has(ext)) {
    return {
      valid: false,
      mimeType,
      extension: ext,
      sizeBytes,
      reason: `Unsupported file extension '${ext}'. Only PDF and DOCX documents are accepted.`,
    };
  }

  // 3. Check MIME type
  const normalizedMime = mimeType ? mimeType.toLowerCase().split(';')[0].trim() : 'application/octet-stream';
  if (!ALLOWED_MIME_TYPES.has(normalizedMime)) {
    return {
      valid: false,
      mimeType: normalizedMime,
      extension: ext,
      sizeBytes,
      reason: `Unsupported MIME type '${normalizedMime}'.`,
    };
  }

  // 4. Check file size
  if (sizeBytes <= 0) {
    return {
      valid: false,
      mimeType: normalizedMime,
      extension: ext,
      sizeBytes,
      reason: 'Attachment is empty (0 bytes).',
    };
  }

  if (sizeBytes > MAX_ATTACHMENT_SIZE_BYTES) {
    return {
      valid: false,
      mimeType: normalizedMime,
      extension: ext,
      sizeBytes,
      reason: `Attachment exceeds maximum size of 10MB (${(sizeBytes / (1024 * 1024)).toFixed(1)}MB).`,
    };
  }

  return {
    valid: true,
    mimeType: normalizedMime,
    extension: ext,
    sizeBytes,
  };
}

/**
 * Sanitize filename to prevent directory traversal or invalid characters.
 */
export function sanitizeFilename(filename: string): string {
  const base = path.basename(filename);
  return base.replace(/[^a-zA-Z0-9._-]/g, '_');
}

/**
 * Generate secure local storage path for an attachment.
 */
export function getStoragePath(emailId: string, attachmentId: string, filename: string): string {
  const safeFilename = sanitizeFilename(filename);
  const targetDir = path.join(ATTACHMENTS_STORAGE_DIR, emailId);
  return path.join(targetDir, `${attachmentId}_${safeFilename}`);
}

/**
 * Save attachment buffer to secure backend-controlled storage.
 * Returns a relative storage key (emailId/filename) instead of an absolute path,
 * so that the key is safe to store in the database and never exposes local file paths.
 */
export async function saveAttachmentFile(
  emailId: string,
  attachmentId: string,
  filename: string,
  data: Buffer
): Promise<string> {
  const filePath = getStoragePath(emailId, attachmentId, filename);
  const dir = path.dirname(filePath);

  await fs.promises.mkdir(dir, { recursive: true });
  await fs.promises.writeFile(filePath, data);

  logger.info({ emailId, attachmentId, sizeBytes: data.length }, 'Attachment saved to secure storage');
  // Return a stable relative key: "emailId/attachmentId_filename" — never the absolute path.
  const safeFilename = sanitizeFilename(filename);
  return `${emailId}/${attachmentId}_${safeFilename}`;
}

/**
 * Read attachment file safely from storage.
 * Accepts either:
 *  - a relative key ("emailId/attachmentId_filename") stored in newer records, or
 *  - a legacy absolute path (still accepted so existing records keep working).
 * Always validates the resolved path stays inside ATTACHMENTS_STORAGE_DIR.
 */
export async function readAttachmentFile(storageKeyOrPath: string): Promise<Buffer> {
  // Resolve relative keys against the canonical storage directory
  const resolved = path.isAbsolute(storageKeyOrPath)
    ? path.resolve(storageKeyOrPath)
    : path.resolve(ATTACHMENTS_STORAGE_DIR, storageKeyOrPath);
  // Security: ensure the resolved path is inside ATTACHMENTS_STORAGE_DIR
  if (!resolved.startsWith(ATTACHMENTS_STORAGE_DIR)) {
    throw new Error('Access denied: attachment path outside storage directory.');
  }
  return fs.promises.readFile(resolved);
}

export interface OneDriveResumeResult {
  driveItemId: string;
  webUrl: string;
}

/**
 * Upload a candidate resume to persistent OneDrive storage with a deterministic name.
 */
export async function uploadResumeToOneDrive(
  candidateId: string,
  filename: string,
  buffer: Buffer,
  mimeType: string = 'application/octet-stream'
): Promise<OneDriveResumeResult> {
  const { prisma } = await import('./db');
  const { decrypt } = await import('./crypto');
  const { refreshAccessToken } = await import('./tokenRefresh');

  const conn = await prisma.mailboxConnection.findFirst({ orderBy: { createdAt: 'desc' } });
  if (!conn) throw new Error('No mailbox connection found');

  let token: string;
  if (new Date() >= new Date(conn.tokenExpiresAt.getTime() - 60000)) {
    await refreshAccessToken(conn.id);
    const refreshed = await prisma.mailboxConnection.findUnique({ where: { id: conn.id } });
    if (!refreshed) throw new Error('Failed to refresh token');
    token = decrypt(refreshed.accessToken);
  } else {
    token = decrypt(conn.accessToken);
  }

  const safeFilename = sanitizeFilename(filename);
  const targetName = `${candidateId}_${safeFilename}`;
  const uploadUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/Candidate_Resumes/${encodeURIComponent(targetName)}:/content`;

  const res = await fetch(uploadUrl, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': mimeType,
    },
    body: buffer,
  });

  if (!res.ok) {
    const errText = await res.text();
    throw new Error(`Failed to upload resume to OneDrive: ${res.status} ${errText}`);
  }

  const data = (await res.json()) as { id: string; webUrl: string };
  logger.info({ candidateId, driveItemId: data.id }, 'Uploaded resume to OneDrive');
  return {
    driveItemId: data.id,
    webUrl: data.webUrl,
  };
}

/**
 * Download a candidate resume from OneDrive via Microsoft Graph using its driveItemId.
 */
export async function downloadResumeFromOneDrive(
  driveItemId: string
): Promise<{ buffer: Buffer; contentType: string } | null> {
  const { prisma } = await import('./db');
  const { decrypt } = await import('./crypto');
  const { refreshAccessToken } = await import('./tokenRefresh');

  const conn = await prisma.mailboxConnection.findFirst({ orderBy: { createdAt: 'desc' } });
  if (!conn) throw new Error('No mailbox connection found');

  let token: string;
  if (new Date() >= new Date(conn.tokenExpiresAt.getTime() - 60000)) {
    await refreshAccessToken(conn.id);
    const refreshed = await prisma.mailboxConnection.findUnique({ where: { id: conn.id } });
    if (!refreshed) throw new Error('Failed to refresh token');
    token = decrypt(refreshed.accessToken);
  } else {
    token = decrypt(conn.accessToken);
  }

  try {
    const res = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${encodeURIComponent(driveItemId)}/content`, {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });

    if (!res.ok) {
      const errBody = await res.text().catch(() => '');
      logger.warn({ driveItemId, status: res.status, errBody }, 'Failed to download resume from Microsoft Graph');
      return null;
    }

    const arrayBuf = await res.arrayBuffer();
    const contentType = res.headers.get('content-type') || 'application/octet-stream';
    return {
      buffer: Buffer.from(arrayBuf),
      contentType,
    };
  } catch (err: any) {
    logger.warn({ driveItemId, err: err.message }, 'Exception during Microsoft Graph resume download');
    return null;
  }
}


