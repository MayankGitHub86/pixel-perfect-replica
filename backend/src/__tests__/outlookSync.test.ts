// backend/src/__tests__/outlookSync.test.ts
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { syncOutlookInbox } from '../lib/outlookSync';
import { prisma } from '../lib/db';
import * as cryptoLib from '../lib/crypto';

describe('Outlook Email Synchronization Service', () => {
  const originalFetch = global.fetch;

  beforeEach(() => {
    vi.restoreAllMocks();
  });

  afterEach(() => {
    global.fetch = originalFetch;
  });

  it('synchronizes incoming messages with duplicate prevention and attachment detection', async () => {
    // Mock MailboxConnection
    vi.spyOn(prisma.mailboxConnection, 'findFirst').mockResolvedValue({
      id: 'mock-mailbox-1',
      userId: 'mock-user-1',
      emailAddress: 'test@outlook.com',
      accessToken: 'mock-encrypted-token',
      refreshToken: 'mock-refresh-token',
      tokenExpiresAt: new Date(Date.now() + 3600 * 1000),
      createdAt: new Date(),
      updatedAt: new Date(),
    } as any);

    vi.spyOn(prisma.mailboxConnection, 'findUnique').mockResolvedValue({
      id: 'mock-mailbox-1',
      tokenExpiresAt: new Date(Date.now() + 3600 * 1000),
      accessToken: 'mock-encrypted-token',
    } as any);

    vi.spyOn(cryptoLib, 'decrypt').mockReturnValue('mock-decrypted-token');

    // Message 1 is new, Message 2 is already stored (duplicate)
    const mockMessages = [
      {
        id: 'msg-new-001',
        subject: 'Job Application: Full Stack Engineer',
        receivedDateTime: '2026-10-30T10:00:00Z',
        hasAttachments: true,
        bodyPreview: 'Please find attached my resume.',
        from: {
          emailAddress: {
            name: 'Candidate One',
            address: 'cand1@example.com',
          },
        },
      },
      {
        id: 'msg-existing-002',
        subject: 'Already Processed Email',
        receivedDateTime: '2026-10-29T10:00:00Z',
        hasAttachments: false,
        from: {
          emailAddress: {
            name: 'Candidate Two',
            address: 'cand2@example.com',
          },
        },
      },
    ];

    // Mock existing check: msg-new-001 -> null, msg-existing-002 -> found
    vi.spyOn(prisma.emailMessage, 'findUnique').mockImplementation((({ where }: any) => {
      if (where.messageId === 'msg-existing-002') {
        return Promise.resolve({ id: 'email-db-002', messageId: 'msg-existing-002' });
      }
      return Promise.resolve(null);
    }) as any);

    vi.spyOn(prisma.emailMessage, 'create').mockResolvedValue({
      id: 'email-db-001',
      messageId: 'msg-new-001',
    } as any);

    vi.spyOn(prisma.processingJob, 'create').mockResolvedValue({
      id: 'job-001',
    } as any);

    vi.spyOn(prisma.attachment, 'create').mockResolvedValue({
      id: 'att-001',
    } as any);

    vi.spyOn(prisma.auditLog, 'create').mockResolvedValue({
      id: 'audit-001',
    } as any);

    // Mock Graph API fetch
    global.fetch = vi.fn().mockImplementation(async (url: string) => {
      if (url.includes('/messages/msg-new-001/attachments')) {
        return {
          ok: true,
          json: async () => ({
            value: [
              {
                id: 'graph-att-1',
                name: 'Resume.pdf',
                contentType: 'application/pdf',
                size: 1024 * 50,
                isInline: false,
              },
            ],
          }),
        };
      }
      if (url.includes('/messages?')) {
        return {
          ok: true,
          json: async () => ({
            value: mockMessages,
          }),
        };
      }
      return { ok: false, status: 404, text: async () => 'Not found' };
    }) as any;

    const result = await syncOutlookInbox({ batchSize: 5 });

    expect(result.totalScanned).toBe(2);
    expect(result.newIngested).toBe(1);
    expect(result.duplicatesSkipped).toBe(1);
    expect(result.attachmentsDetected).toBe(1);
    expect(result.errors).toHaveLength(0);

    expect(prisma.emailMessage.create).toHaveBeenCalledTimes(1);
    expect(prisma.processingJob.create).toHaveBeenCalledTimes(1);
    expect(prisma.attachment.create).toHaveBeenCalledTimes(1);
  });

  it('handles Graph API error when mailbox cannot be queried', async () => {
    vi.spyOn(prisma.mailboxConnection, 'findFirst').mockResolvedValue({
      id: 'mock-mailbox-1',
      emailAddress: 'test@outlook.com',
      accessToken: 'mock-encrypted-token',
      tokenExpiresAt: new Date(Date.now() + 3600 * 1000),
    } as any);

    vi.spyOn(prisma.mailboxConnection, 'findUnique').mockResolvedValue({
      id: 'mock-mailbox-1',
      tokenExpiresAt: new Date(Date.now() + 3600 * 1000),
      accessToken: 'mock-encrypted-token',
    } as any);

    vi.spyOn(cryptoLib, 'decrypt').mockReturnValue('mock-decrypted-token');

    global.fetch = vi.fn().mockResolvedValue({
      ok: false,
      status: 401,
      text: async () => 'Unauthorized token',
    } as any);

    await expect(syncOutlookInbox()).rejects.toThrow(/Graph API returned 401/);
  });
});
