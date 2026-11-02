import crypto from 'crypto';
import { Router, Request, Response, NextFunction } from 'express';
import { prisma } from '../lib/db';
import { config } from '../config';
import { integrationAuthMiddleware } from '../middleware/auth';
import { consumeOAuthState } from '../lib/oauth';
import { encrypt } from '../lib/crypto';
import { syncOutlookInbox } from '../lib/outlookSync';
import { logger } from '../logger';

export const integrationRouter = Router();

integrationRouter.use(integrationAuthMiddleware);

function timingSafeEqualStr(a: string, b: string): boolean {
  const bufA = Buffer.from(a, 'utf8');
  const bufB = Buffer.from(b, 'utf8');
  if (bufA.length !== bufB.length) {
    return false;
  }
  return crypto.timingSafeEqual(bufA, bufB);
}

// GET /api/integrations/microsoft/status
integrationRouter.get('/microsoft/status', async (req: Request, res: Response, next: NextFunction) => {
  try {
    const connection = await prisma.mailboxConnection.findFirst({
      orderBy: { createdAt: 'desc' },
    });

    if (!connection) {
      res.json({
        state: 'not_connected',
        accountEmail: null,
        mailboxPermission: 'missing',
        lastMailboxCheck: null,
        subscriptionStatus: 'none',
        subscriptionExpiresAt: null,
        lastReconciliation: null,
      });
      return;
    }

    const isExpired = connection.tokenExpiresAt < new Date();
    res.json({
      state: isExpired ? 'expired' : 'connected',
      accountEmail: connection.emailAddress,
      mailboxPermission: 'granted',
      lastMailboxCheck: connection.updatedAt.toISOString(),
      subscriptionStatus: 'none',
      subscriptionExpiresAt: null,
      lastReconciliation: null,
    });
  } catch (err) {
    next(err);
  }
});

// POST /api/integrations/microsoft/sync
integrationRouter.post('/microsoft/sync', async (req: Request, res: Response, next: NextFunction) => {
  try {
    const rawBatch = req.body?.batchSize;
    const batchSize = typeof rawBatch === 'number' && !isNaN(rawBatch)
      ? Math.min(Math.max(rawBatch, 1), 20)
      : 5;

    const result = await syncOutlookInbox({
      batchSize,
      requestId: req.id?.toString(),
    });

    res.json({
      success: true,
      ...result,
    });
  } catch (err) {
    next(err);
  }
});

// GET /api/integrations/microsoft/connect
// GET /api/integrations/microsoft/connect – PKCE OAuth initiation
integrationRouter.get('/microsoft/connect', async (req: Request, res: Response, next: NextFunction) => {
  try {
    // Generate state and PKCE verifier/challenge
    const { state, verifier } = await import('../lib/oauth').then(m => m.createOAuthState());
    const challenge = await import('../lib/oauth').then(m => m.generateCodeChallenge(verifier));

    // Store state in a short‑lived HttpOnly cookie (10 min)
    res.cookie('oauth_state', state, {
      httpOnly: true,
      sameSite: 'lax',
      maxAge: 10 * 60 * 1000,
      secure: config.NODE_ENV !== 'development',
    });

    const params = new URLSearchParams({
      client_id: config.MICROSOFT_CLIENT_ID,
      response_type: 'code',
      redirect_uri: config.MICROSOFT_REDIRECT_URI,
      response_mode: 'query',
      scope: 'openid profile offline_access User.Read Mail.Read Files.ReadWrite',
      state,
      code_challenge: challenge,
      code_challenge_method: 'S256',
    });

    const authUrl = `https://login.microsoftonline.com/${config.MICROSOFT_TENANT_ID}/oauth2/v2.0/authorize?${params.toString()}`;
    res.redirect(authUrl);
  } catch (err) {
    next(err);
  }
});

// GET /api/integrations/microsoft/callback – OAuth authorization code callback
integrationRouter.get('/microsoft/callback', async (req: Request, res: Response, next: NextFunction) => {
  const clearStateCookie = () => {
    res.clearCookie('oauth_state', {
      httpOnly: true,
      sameSite: 'lax',
      secure: config.NODE_ENV !== 'development',
    });
  };

  try {
    // 1. Handle Microsoft OAuth error responses safely
    const { error, error_description: errorDescription, code, state } = req.query;

    if (error) {
      clearStateCookie();
      logger.warn({ error: String(error), requestId: req.id }, 'Microsoft OAuth error returned by provider');
      res.status(400).json({
        error: {
          code: 'OAUTH_ERROR',
          message: error === 'access_denied'
            ? 'Microsoft authorization was cancelled or denied'
            : 'Microsoft authentication failed',
          requestId: req.id,
        },
      });
      return;
    }

    // 2. Validate authorization code and state query parameters
    if (!code || typeof code !== 'string') {
      clearStateCookie();
      res.status(400).json({
        error: {
          code: 'MISSING_CODE',
          message: 'Missing or invalid authorization code',
          requestId: req.id,
        },
      });
      return;
    }

    if (!state || typeof state !== 'string') {
      clearStateCookie();
      res.status(400).json({
        error: {
          code: 'MISSING_STATE',
          message: 'Missing or invalid state parameter',
          requestId: req.id,
        },
      });
      return;
    }

    // 3. Validate oauth_state cookie
    const cookieState = req.cookies?.oauth_state;
    if (!cookieState || typeof cookieState !== 'string') {
      clearStateCookie();
      res.status(400).json({
        error: {
          code: 'MISSING_COOKIE_STATE',
          message: 'Missing OAuth state cookie',
          requestId: req.id,
        },
      });
      return;
    }

    // 4. Compare cookie state with query state using timing-safe comparison
    if (!timingSafeEqualStr(cookieState, state)) {
      clearStateCookie();
      res.status(400).json({
        error: {
          code: 'STATE_MISMATCH',
          message: 'OAuth state cookie does not match state parameter',
          requestId: req.id,
        },
      });
      return;
    }

    // 5. Consume the OAuthState database record (enforcing single-use, validity, and expiry)
    let codeVerifier: string;
    try {
      codeVerifier = await consumeOAuthState(state);
    } catch (err: any) {
      clearStateCookie();
      const isExpired = err?.message === 'OAuth state expired';
      res.status(400).json({
        error: {
          code: isExpired ? 'STATE_EXPIRED' : 'INVALID_STATE',
          message: isExpired ? 'OAuth state has expired' : 'Invalid or reused OAuth state',
          requestId: req.id,
        },
      });
      return;
    }

    // 6. Exchange authorization code using the saved PKCE verifier
    const tokenUrl = `https://login.microsoftonline.com/${config.MICROSOFT_TENANT_ID}/oauth2/v2.0/token`;
    const tokenParams = new URLSearchParams({
      client_id: config.MICROSOFT_CLIENT_ID,
      client_secret: config.MICROSOFT_CLIENT_SECRET,
      grant_type: 'authorization_code',
      code,
      redirect_uri: config.MICROSOFT_REDIRECT_URI,
      code_verifier: codeVerifier,
    });

    let tokenRes: globalThis.Response;
    try {
      tokenRes = await fetch(tokenUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: tokenParams.toString(),
      });
    } catch (netErr: any) {
      clearStateCookie();
      logger.error({ error: netErr?.message, requestId: req.id }, 'Network error during Microsoft token exchange');
      res.status(502).json({
        error: {
          code: 'NETWORK_ERROR',
          message: 'Failed to connect to Microsoft token service',
          requestId: req.id,
        },
      });
      return;
    }

    if (!tokenRes.ok) {
      clearStateCookie();
      logger.error({ status: tokenRes.status, requestId: req.id }, 'Microsoft token exchange returned non-OK status');
      res.status(502).json({
        error: {
          code: 'TOKEN_EXCHANGE_FAILED',
          message: 'Failed to exchange authorization code for tokens',
          requestId: req.id,
        },
      });
      return;
    }

    interface MicrosoftTokenResponse {
      access_token: string;
      refresh_token?: string;
      expires_in: number;
    }

    const tokenData = (await tokenRes.json()) as MicrosoftTokenResponse;
    if (!tokenData.access_token || typeof tokenData.access_token !== 'string') {
      clearStateCookie();
      res.status(502).json({
        error: {
          code: 'INVALID_TOKEN_RESPONSE',
          message: 'Microsoft token response missing access token',
          requestId: req.id,
        },
      });
      return;
    }

    // 7. Call Microsoft Graph GET /v1.0/me to obtain and validate identity
    let graphRes: globalThis.Response;
    try {
      graphRes = await fetch('https://graph.microsoft.com/v1.0/me', {
        headers: {
          Authorization: `Bearer ${tokenData.access_token}`,
        },
      });
    } catch (netErr: any) {
      clearStateCookie();
      logger.error({ error: netErr?.message, requestId: req.id }, 'Network error calling Microsoft Graph /me');
      res.status(502).json({
        error: {
          code: 'NETWORK_ERROR',
          message: 'Failed to connect to Microsoft Graph',
          requestId: req.id,
        },
      });
      return;
    }

    if (!graphRes.ok) {
      clearStateCookie();
      logger.error({ status: graphRes.status, requestId: req.id }, 'Microsoft Graph /me returned non-OK status');
      res.status(502).json({
        error: {
          code: 'GRAPH_API_FAILED',
          message: 'Failed to retrieve user profile from Microsoft Graph',
          requestId: req.id,
        },
      });
      return;
    }

    interface MicrosoftProfileResponse {
      id?: string;
      mail?: string | null;
      userPrincipalName?: string | null;
    }

    const profile = (await graphRes.json()) as MicrosoftProfileResponse;
    const accountEmail = profile.mail || profile.userPrincipalName;

    if (!accountEmail || typeof accountEmail !== 'string') {
      clearStateCookie();
      logger.error({ requestId: req.id }, 'Microsoft Graph profile missing mail and userPrincipalName');
      res.status(502).json({
        error: {
          code: 'INVALID_PROFILE',
          message: 'Could not determine account email from Microsoft profile',
          requestId: req.id,
        },
      });
      return;
    }

    const normalizedEmail = accountEmail.toLowerCase().trim();
    const userId = profile.id || normalizedEmail;

    // 8. Encrypt tokens before persistence
    const encryptedAccessToken = encrypt(tokenData.access_token);
    const encryptedRefreshToken = tokenData.refresh_token ? encrypt(tokenData.refresh_token) : undefined;
    const tokenExpiresAt = new Date(Date.now() + (tokenData.expires_in || 3600) * 1000);

    // 9. Upsert MailboxConnection record
    const existing = await prisma.mailboxConnection.findUnique({
      where: { emailAddress: normalizedEmail },
    });

    const finalRefreshToken = encryptedRefreshToken || (existing ? existing.refreshToken : encrypt(''));

    const connection = await prisma.mailboxConnection.upsert({
      where: { emailAddress: normalizedEmail },
      create: {
        userId,
        emailAddress: normalizedEmail,
        accessToken: encryptedAccessToken,
        refreshToken: finalRefreshToken,
        tokenExpiresAt,
      },
      update: {
        userId,
        accessToken: encryptedAccessToken,
        ...(encryptedRefreshToken ? { refreshToken: encryptedRefreshToken } : {}),
        tokenExpiresAt,
        updatedAt: new Date(),
      },
    });

    // 10. Clear OAuth state cookie on success
    clearStateCookie();
    logger.info({ email: connection.emailAddress, requestId: req.id }, 'Microsoft mailbox connected successfully');

    // 11. Return response or redirect to frontend settings
    const frontendUrl = config.CORS_ORIGIN.split(',')[0].trim() || 'http://localhost:5173';
    if (req.headers.accept?.includes('application/json')) {
      res.json({
        success: true,
        emailAddress: connection.emailAddress,
        state: 'connected',
      });
      return;
    }

    res.redirect(`${frontendUrl}/settings?microsoft=connected`);
  } catch (err) {
    clearStateCookie();
    next(err);
  }
});

