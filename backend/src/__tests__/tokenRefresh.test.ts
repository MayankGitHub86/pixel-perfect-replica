// backend/src/__tests__/tokenRefresh.test.ts
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { refreshAccessToken } from '../lib/tokenRefresh';
import { prisma } from '../lib/db';
import { encrypt, decrypt } from '../lib/crypto';
import { config } from '../config';

// Mock global fetch
const fetchMock = vi.fn();
global.fetch = fetchMock as any;

vi.mock('../lib/db', async () => {
  const actual = await vi.importActual('../lib/db');
  return {
    ...actual,
    prisma: {
      mailboxConnection: {
        findUnique: vi.fn(),
        update: vi.fn(),
      },
    },
  };
});

// Use real crypto functions for encryption (they rely on env var)
process.env.TOKEN_ENCRYPTION_KEY = Buffer.from('a'.repeat(32)).toString('base64');

const SAMPLE_MAILBOX = {
  id: 'mailbox-1',
  refreshToken: encrypt('old-refresh-token'),
  accessToken: encrypt('old-access-token'),
  tokenExpiresAt: new Date(0),
  updatedAt: new Date(0),
  emailAddress: 'user@example.com',
  userId: 'dev',
  createdAt: new Date(0),
};

describe('refreshAccessToken', () => {
  beforeEach(() => {
    vi.resetAllMocks();
  });

  it('refreshes tokens successfully', async () => {
    (prisma.mailboxConnection.findUnique as any).mockResolvedValue(SAMPLE_MAILBOX);
    fetchMock.mockResolvedValue({
      ok: true,
      json: async () => ({
        access_token: 'new-access',
        refresh_token: 'new-refresh',
        expires_in: 3600,
      }),
    });
    (prisma.mailboxConnection.update as any).mockResolvedValue({});

    await refreshAccessToken(SAMPLE_MAILBOX.id);

    expect(prisma.mailboxConnection.findUnique).toHaveBeenCalledWith({ where: { id: SAMPLE_MAILBOX.id } });
    expect(fetchMock).toHaveBeenCalled();
    expect(prisma.mailboxConnection.update).toHaveBeenCalledWith({
      where: { id: SAMPLE_MAILBOX.id },
      data: expect.objectContaining({
        accessToken: expect.any(String),
        refreshToken: expect.any(String),
        tokenExpiresAt: expect.any(Date),
      }),
    });
  });

  it('throws when mailbox not found', async () => {
    (prisma.mailboxConnection.findUnique as any).mockResolvedValue(null);
    await expect(refreshAccessToken('unknown')).rejects.toThrow('MailboxConnection not found');
  });

  it('throws on token endpoint failure', async () => {
    (prisma.mailboxConnection.findUnique as any).mockResolvedValue(SAMPLE_MAILBOX);
    fetchMock.mockResolvedValue({ ok: false, status: 400, text: async () => 'bad request' });
    await expect(refreshAccessToken(SAMPLE_MAILBOX.id)).rejects.toThrow('Token refresh failed');
  });
});

