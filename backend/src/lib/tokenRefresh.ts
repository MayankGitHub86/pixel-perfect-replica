// backend/src/lib/tokenRefresh.ts
// Secure Microsoft OAuth access-token refresh helper.

import { prisma } from './db';
import { decrypt, encrypt } from './crypto';
import { config } from '../config';

/** Shape of a successful Microsoft token endpoint response. */
interface MicrosoftTokenResponse {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
}

/**
 * Refresh the access token for a mailbox connection.
 * @param mailboxId The UUID of the MailboxConnection record.
 * @throws Error if the mailbox is not found or the token endpoint returns a failure.
 */
export async function refreshAccessToken(mailboxId: string): Promise<void> {
  const mailbox = await prisma.mailboxConnection.findUnique({ where: { id: mailboxId } });
  if (!mailbox) {
    throw new Error('MailboxConnection not found');
  }

  // Decrypt stored refresh token (never log the value).
  const refreshToken = decrypt(mailbox.refreshToken);

  const tokenRes = await fetch(
    `https://login.microsoftonline.com/${config.MICROSOFT_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: config.MICROSOFT_CLIENT_ID,
        client_secret: config.MICROSOFT_CLIENT_SECRET,
        grant_type: 'refresh_token',
        refresh_token: refreshToken,
        redirect_uri: config.MICROSOFT_REDIRECT_URI,
      }).toString(),
    }
  );

  if (!tokenRes.ok) {
    const errBody = await tokenRes.text();
    throw new Error(`Token refresh failed: ${tokenRes.status} ${errBody}`);
  }

  const tokenData = await tokenRes.json() as MicrosoftTokenResponse;
  const newAccessToken = tokenData.access_token;
  const newRefreshToken = tokenData.refresh_token ?? refreshToken;
  const expiresIn = tokenData.expires_in; // seconds

  const encryptedAccess = encrypt(newAccessToken);
  const encryptedRefresh = encrypt(newRefreshToken);
  const tokenExpiresAt = new Date(Date.now() + expiresIn * 1000);

  // Persist rotated tokens.
  await prisma.mailboxConnection.update({
    where: { id: mailboxId },
    data: {
      accessToken: encryptedAccess,
      refreshToken: encryptedRefresh,
      tokenExpiresAt,
      updatedAt: new Date(),
    },
  });
}

