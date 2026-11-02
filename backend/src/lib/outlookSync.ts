// backend/src/lib/outlookSync.ts
// Microsoft Graph Outlook email synchronization service with duplicate prevention and attachment detection.

import { prisma } from './db';
import { decrypt } from './crypto';
import { refreshAccessToken } from './tokenRefresh';
import { validateAttachment } from './attachmentService';
import { logger } from '../logger';

export interface SyncOptions {
  batchSize?: number;
  folderId?: string; // defaults to 'inbox'
  requestId?: string;
}

export interface SyncResult {
  totalScanned: number;
  newIngested: number;
  duplicatesSkipped: number;
  attachmentsDetected: number;
  errors: string[];
}

export interface GraphMessageItem {
  id: string;
  subject: string;
  receivedDateTime: string;
  hasAttachments: boolean;
  bodyPreview?: string;
  from?: {
    emailAddress?: {
      name?: string;
      address?: string;
    };
  };
}

export interface GraphAttachmentItem {
  id: string;
  name: string;
  contentType: string;
  size: number;
  isInline: boolean;
}

/**
 * Get valid decrypted Graph access token for the mailbox connection.
 */
export async function getValidAccessToken(mailboxId: string): Promise<string> {
  let mailbox = await prisma.mailboxConnection.findUnique({ where: { id: mailboxId } });
  if (!mailbox) {
    throw new Error(`MailboxConnection ${mailboxId} not found.`);
  }

  // Refresh if token is expired or expiring within 5 minutes
  const fiveMinutesFromNow = new Date(Date.now() + 5 * 60 * 1000);
  if (mailbox.tokenExpiresAt < fiveMinutesFromNow) {
    logger.info({ mailboxId }, 'Token near expiration, refreshing before sync...');
    await refreshAccessToken(mailboxId);
    mailbox = await prisma.mailboxConnection.findUnique({ where: { id: mailboxId } });
    if (!mailbox) {
      throw new Error(`MailboxConnection ${mailboxId} disappeared after refresh.`);
    }
  }

  return decrypt(mailbox.accessToken);
}

/**
 * Manually synchronize emails from Microsoft Graph inbox into PostgreSQL.
 */
export async function syncOutlookInbox(options: SyncOptions = {}): Promise<SyncResult> {
  const { batchSize = 5, folderId = 'inbox', requestId = 'system' } = options;

  // 1. Fetch connected mailbox
  const mailbox = await prisma.mailboxConnection.findFirst({
    orderBy: { updatedAt: 'desc' },
  });

  if (!mailbox) {
    throw new Error('No connected Outlook mailbox found. Please connect via OAuth first.');
  }

  // 2. Obtain decrypted Graph access token
  const accessToken = await getValidAccessToken(mailbox.id);
  const headers = {
    Authorization: `Bearer ${accessToken}`,
    Accept: 'application/json',
  };

  // 3. Query Microsoft Graph inbox messages
  const graphUrl = `https://graph.microsoft.com/v1.0/me/mailFolders/${encodeURIComponent(folderId)}/messages?$top=${batchSize}&$select=id,subject,from,receivedDateTime,hasAttachments,bodyPreview&$orderby=receivedDateTime desc`;
  
  const res = await fetch(graphUrl, { headers });
  if (!res.ok) {
    const errorText = await res.text();
    throw new Error(`Graph API returned ${res.status}: ${errorText}`);
  }

  const data = await res.json() as { value?: GraphMessageItem[] };
  const messages: GraphMessageItem[] = data.value || [];

  let newIngested = 0;
  let duplicatesSkipped = 0;
  let attachmentsDetected = 0;
  const errors: string[] = [];

  // 4. Process each message with idempotency
  for (const msg of messages) {
    try {
      // Check if message is already stored in PostgreSQL
      const existing = await prisma.emailMessage.findUnique({
        where: { messageId: msg.id },
      });

      if (existing) {
        duplicatesSkipped++;
        continue;
      }

      const senderEmail = msg.from?.emailAddress?.address?.toLowerCase().trim() || 'unknown@unknown.com';
      const subject = msg.subject || '(No Subject)';
      const receivedAt = new Date(msg.receivedDateTime);

      // Ingest EmailMessage record
      const emailRecord = await prisma.emailMessage.create({
        data: {
          mailboxId: mailbox.id,
          messageId: msg.id,
          subject,
          fromEmail: senderEmail,
          receivedAt,
          status: 'new',
          history: [
            {
              at: new Date().toISOString(),
              step: 'email_ingest',
              status: 'succeeded',
              message: 'Synchronized from Outlook inbox',
            },
          ],
        },
      });

      newIngested++;

      // Create linked ProcessingJob
      await prisma.processingJob.create({
        data: {
          type: 'email_ingest',
          related: subject,
          status: 'queued',
          emailId: emailRecord.id,
        },
      });

      // 5. Detect and record attachments if message indicates attachments
      if (msg.hasAttachments) {
        try {
          const attachUrl = `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(msg.id)}/attachments?$select=id,name,contentType,size,isInline`;
          const attachRes = await fetch(attachUrl, { headers });

          if (attachRes.ok) {
            const attachData = await attachRes.json() as { value?: GraphAttachmentItem[] };
            const attachments = attachData.value || [];

            for (const att of attachments) {
              if (att.isInline) continue; // Skip inline images/signatures

              const validation = validateAttachment(att.name, att.contentType, att.size);

              let storageKey = validation.valid ? 'pending_download' : 'rejected';
              if (validation.valid) {
                try {
                  const singleAttachUrl = `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(msg.id)}/attachments/${encodeURIComponent(att.id)}`;
                  const singleRes = await fetch(singleAttachUrl, { headers });
                  if (singleRes.ok) {
                    const singleData = (await singleRes.json()) as any;
                    if (singleData.contentBytes) {
                      const { saveAttachmentFile } = await import('./attachmentService');
                      const buf = Buffer.from(singleData.contentBytes, 'base64');
                      storageKey = await saveAttachmentFile(emailRecord.id, att.id, att.name, buf);
                    }
                  }
                } catch (dlErr: any) {
                  logger.debug({ attachmentId: att.id, err: dlErr.message }, 'Attachment bytes download skipped');
                }
              }

              await prisma.attachment.create({
                data: {
                  emailId: emailRecord.id,
                  filename: att.name,
                  mimeType: validation.mimeType,
                  sizeBytes: att.size,
                  storageKey,
                  status: validation.valid ? 'new' : 'error',
                },
              });

              attachmentsDetected++;
            }
          } else {
            const attachErr = await attachRes.text();
            logger.warn({ messageId: msg.id, attachErr }, 'Failed to fetch message attachments metadata');
          }
        } catch (attErr: any) {
          logger.warn({ messageId: msg.id, error: attErr.message }, 'Error fetching attachment metadata');
        }
      }

      // 6. Enqueue email processing job into BullMQ queue
      try {
        const { emailProcessingQueue } = await import('./queue');
        await emailProcessingQueue.add('process-email', { emailId: emailRecord.id, requestId });
      } catch (qErr: any) {
        logger.debug({ emailId: emailRecord.id, error: qErr.message }, 'BullMQ email queue add skipped');
      }
    } catch (msgErr: any) {
      errors.push(`Failed to process message ${msg.id}: ${msgErr.message}`);
      logger.error({ messageId: msg.id, err: msgErr.message }, 'Error ingesting email message');
    }
  }

  // 6. Record Audit Log for the sync execution
  await prisma.auditLog.create({
    data: {
      actor: 'system',
      action: 'outlook_inbox_sync',
      related: mailbox.emailAddress,
      entity: 'MailboxConnection',
      entityId: mailbox.id,
      result: errors.length === 0 ? 'success' : 'partial',
      requestId,
      payload: {
        totalScanned: messages.length,
        newIngested,
        duplicatesSkipped,
        attachmentsDetected,
        errors,
      },
    },
  });

  return {
    totalScanned: messages.length,
    newIngested,
    duplicatesSkipped,
    attachmentsDetected,
    errors,
  };
}

