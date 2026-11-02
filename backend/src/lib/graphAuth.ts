// backend/src/lib/graphAuth.ts
// Centralised helper to obtain a valid Microsoft Graph access token.
// It checks the stored expiry, refreshes if necessary, and returns the
// decrypted access token ready for use in Authorization headers.

import { prisma } from './db';
import { decrypt } from './crypto';
import { refreshAccessToken } from './tokenRefresh';

/**
 * Returns a fresh access token for the most recent mailbox connection.
 * If the stored token is close to expiry (within 5 minutes) it will be
 * refreshed via the refresh token flow. The refreshed token (and possibly
 * refreshed refresh token) are persisted back to the DB.
 */
export async function getValidAccessToken(): Promise<string> {
  // Get the latest mailbox connection record.
  const conn = await prisma.mailboxConnection.findFirst({ orderBy: { createdAt: 'desc' } });
  if (!conn) throw new Error('No mailbox connection found');

  // Decrypt the stored access token just for expiry inspection.
  // The token is a JWT; its payload is base64url‑encoded.
  const decryptedAccess = decrypt(conn.accessToken);
  // Parse JWT payload without verifying signature (only to read exp claim).
  const [, payloadB64] = decryptedAccess.split('.');
  let payload: any = {};
  try {
    const json = Buffer.from(payloadB64, 'base64url').toString('utf8');
    payload = JSON.parse(json);
  } catch (_) {
    // If parsing fails we assume the token is malformed and force a refresh.
    payload = {};
  }

  const expiresAt = payload.exp ? payload.exp * 1000 : conn.tokenExpiresAt?.getTime();
  const now = Date.now();
  const fiveMinutes = 5 * 60 * 1000;

  // If we have no expiry info or the token will expire within 5 min, refresh.
  if (!expiresAt || expiresAt - now < fiveMinutes) {
    // Use the stored mailbox ID to refresh; the function updates DB.
    await refreshAccessToken(conn.id);
    // After refresh, fetch the updated record and decrypt the new access token.
    const refreshed = await prisma.mailboxConnection.findUnique({ where: { id: conn.id } });
    if (!refreshed) throw new Error('Failed to reload mailbox connection after refresh');
    return decrypt(refreshed.accessToken);
  }

  // Token is still valid.
  return decryptedAccess;
}

