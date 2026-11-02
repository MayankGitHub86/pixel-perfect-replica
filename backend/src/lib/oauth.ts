// backend/src/lib/oauth.ts
// PKCE and OAuth state helpers for Microsoft Graph integration.

import crypto from 'crypto';
import { prisma } from '../lib/db';
import { config } from '../config';
import { addMinutes, isAfter } from 'date-fns';

/** Generate a random URL‑safe string for PKCE code verifier */
export function generateCodeVerifier(): string {
  const random = crypto.randomBytes(32);
  return base64UrlEncode(random);
}

/** Derive the code challenge (SHA‑256) from a verifier */
export function generateCodeChallenge(verifier: string): string {
  const sha256 = crypto.createHash('sha256').update(verifier).digest();
  return base64UrlEncode(sha256);
}

function base64UrlEncode(buf: Buffer): string {
  return buf.toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
}

/** Create a one‑time OAuth state record.
 * Returns the generated state string (used as the `state` query param).
 */
export async function createOAuthState(): Promise<{ state: string; verifier: string }> {
  const verifier = generateCodeVerifier();
  const state = crypto.randomUUID();
  const expiresAt = addMinutes(new Date(), 10); // 10‑minute validity

  await prisma.oAuthState.create({
    data: {
      id: crypto.randomUUID(),
      state,
      codeVerifier: verifier,
      expiresAt,
    },
  });
  return { state, verifier };
}

/** Validate and consume a stored OAuth state.
 * Throws if not found, expired, or mismatched.
 */
export async function consumeOAuthState(state: string): Promise<string> {
  const record = await prisma.oAuthState.findUnique({ where: { state } });
  if (!record) {
    throw new Error('Invalid OAuth state');
  }
  if (isAfter(new Date(), record.expiresAt)) {
    // Clean up expired record
    await prisma.oAuthState.delete({ where: { state } });
    throw new Error('OAuth state expired');
  }
  // Delete after successful consumption to enforce single‑use
  await prisma.oAuthState.delete({ where: { state } });
  return record.codeVerifier;
}

