// backend/src/lib/emailService.ts
// Secure email body parsing, sanitization, and retrieval.

import { prisma } from './db';
import { decrypt } from './crypto';
import { refreshAccessToken } from './tokenRefresh';
import { logger } from '../logger';

/**
 * Sanitize HTML or plain text into clean, safe, readable text.
 * Strips script, style, and tag elements, decodes entities, and preserves line breaks.
 */
export function sanitizeAndExtractText(htmlOrText: string): string {
  if (!htmlOrText) return '';
  if (!htmlOrText.includes('<') && !htmlOrText.includes('>')) {
    return htmlOrText.trim();
  }

  let text = htmlOrText;
  text = text.replace(/<style[\s\S]*?<\/style>/gi, '');
  text = text.replace(/<head[\s\S]*?<\/head>/gi, '');
  text = text.replace(/<script[\s\S]*?<\/script>/gi, '');
  text = text.replace(/<br\s*[\/]?>/gi, '\n');
  text = text.replace(/<p[^>]*>/gi, '\n');
  text = text.replace(/<\/p>/gi, '\n\n');
  text = text.replace(/<div[^>]*>/gi, '\n');
  text = text.replace(/<\/div>/gi, '\n');
  text = text.replace(/<li[^>]*>/gi, '\n• ');
  text = text.replace(/<tr[^>]*>/gi, '\n');
  text = text.replace(/<[^>]+>/g, '');
  text = text.replace(/&nbsp;/gi, ' ')
             .replace(/&amp;/gi, '&')
             .replace(/&lt;/gi, '<')
             .replace(/&gt;/gi, '>')
             .replace(/&quot;/gi, '"')
             .replace(/&#39;/gi, "'");
  text = text.split('\n').map(l => l.trim()).filter(Boolean).join('\n');
  return text;
}

/**
 * Fetch and sanitize email body for a given messageId or email record.
 * Uses cached history if available, otherwise queries Microsoft Graph.
 */
export async function getCleanEmailBody(emailRecordIdOrMessageId: string): Promise<string | null> {
  try {
    const email = await prisma.emailMessage.findFirst({
      where: {
        OR: [
          { id: emailRecordIdOrMessageId },
          { messageId: emailRecordIdOrMessageId },
        ],
      },
    });

    if (!email) return null;

    // 1. Check if clean body is already cached in history JSON
    if (Array.isArray(email.history)) {
      const cached = (email.history as any[]).find((h) => h.step === 'email_body' && typeof h.bodyText === 'string');
      if (cached) {
        return cached.bodyText;
      }
    }

    // 2. If it's a synthetic message or doesn't have an Outlook Graph messageId, return null or fallback
    if (!email.messageId || !email.messageId.startsWith('AQMk')) {
      return null;
    }

    // 3. Query Microsoft Graph for message body
    const conn = await prisma.mailboxConnection.findFirst({ orderBy: { createdAt: 'desc' } });
    if (!conn) return null;

    let token = decrypt(conn.accessToken);
    const expiresAt = conn.tokenExpiresAt.getTime();
    if (Date.now() >= expiresAt - 60000) {
      await refreshAccessToken(conn.id);
      const refreshed = await prisma.mailboxConnection.findUnique({ where: { id: conn.id } });
      if (refreshed) token = decrypt(refreshed.accessToken);
    }

    const res = await fetch(
      `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(email.messageId)}?$select=id,subject,bodyPreview,body`,
      { headers: { Authorization: `Bearer ${token}` } }
    );

    if (!res.ok) {
      logger.warn({ messageId: email.messageId, status: res.status }, 'Failed to fetch message body from Graph');
      return null;
    }

    const data: any = await res.json();
    let bodyText = '';
    if (data.body?.content) {
      bodyText = sanitizeAndExtractText(data.body.content);
    }
    if (!bodyText && data.bodyPreview) {
      bodyText = data.bodyPreview.trim();
    }

    // Cache clean text into history
    const existingHistory = Array.isArray(email.history) ? [...(email.history as any[])] : [];
    existingHistory.push({
      at: new Date().toISOString(),
      step: 'email_body',
      status: 'succeeded',
      bodyText,
    });

    await prisma.emailMessage.update({
      where: { id: email.id },
      data: { history: existingHistory },
    }).catch(() => {});

    return bodyText || null;
  } catch (err: any) {
    logger.warn({ error: err.message }, 'Failed to get clean email body');
    return null;
  }
}

