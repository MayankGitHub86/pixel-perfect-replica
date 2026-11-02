// backend/src/__tests__/oauthCallback.test.ts
// Focused automated tests for GET /api/integrations/microsoft/callback
// Mocked Microsoft endpoints — no real credentials used.

import { describe, it, expect, vi, beforeEach, afterAll } from 'vitest';
import request from 'supertest';
import { createApp } from '../app';
import { prisma } from '../lib/db';
import { createOAuthState } from '../lib/oauth';
import { decrypt } from '../lib/crypto';

const app = createApp();

const TEST_EMAIL = 'ms-oauth-test@example.com';

function hasClearedOAuthCookie(setCookie: string | string[] | undefined): boolean {
  if (!setCookie) return false;
  const list = Array.isArray(setCookie) ? setCookie : [setCookie];
  return list.some((c) => c.includes('oauth_state=;') || c.includes('Max-Age=0'));
}

// Mock global fetch for Microsoft endpoints
const fetchMock = vi.fn();
global.fetch = fetchMock as any;

describe('GET /api/integrations/microsoft/callback', () => {
  beforeEach(async () => {
    vi.clearAllMocks();
    // Clean up test data
    await prisma.mailboxConnection.deleteMany({
      where: { emailAddress: TEST_EMAIL },
    });
  });

  afterAll(async () => {
    await prisma.mailboxConnection.deleteMany({
      where: { emailAddress: TEST_EMAIL },
    });
    await prisma.oAuthState.deleteMany({
      where: { state: { startsWith: 'test-' } },
    });
  });

  it('successfully exchanges code, fetches profile, encrypts tokens, and connects mailbox (JSON response)', async () => {
    const { state, verifier } = await createOAuthState();

    // Mock token exchange
    fetchMock.mockResolvedValueOnce({
      ok: true,
      json: async () => ({
        access_token: 'mock-access-token-12345',
        refresh_token: 'mock-refresh-token-67890',
        expires_in: 3600,
      }),
    });

    // Mock Graph /v1.0/me
    fetchMock.mockResolvedValueOnce({
      ok: true,
      json: async () => ({
        id: 'ms-graph-user-id-001',
        mail: TEST_EMAIL,
        displayName: 'Test MS User',
      }),
    });

    const res = await request(app)
      .get(`/api/integrations/microsoft/callback?code=mock-auth-code&state=${state}`)
      .set('Cookie', `oauth_state=${state}`)
      .set('Accept', 'application/json');

    expect(res.status).toBe(200);
    expect(res.body.success).toBe(true);
    expect(res.body.emailAddress).toBe(TEST_EMAIL);
    expect(res.body.state).toBe('connected');

    // Never return tokens in API response
    expect(res.body.accessToken).toBeUndefined();
    expect(res.body.refreshToken).toBeUndefined();
    expect(res.body.codeVerifier).toBeUndefined();

    // Cookie must be cleared
    expect(hasClearedOAuthCookie(res.headers['set-cookie'])).toBe(true);

    // Verify token endpoint was called with correct parameters (including PKCE verifier)
    expect(fetchMock).toHaveBeenCalledTimes(2);
    const tokenCall = fetchMock.mock.calls[0];
    const tokenBody = new URLSearchParams(tokenCall[1].body);
    expect(tokenBody.get('grant_type')).toBe('authorization_code');
    expect(tokenBody.get('code')).toBe('mock-auth-code');
    expect(tokenBody.get('code_verifier')).toBe(verifier);

    // Verify Graph call was made with Bearer token
    const graphCall = fetchMock.mock.calls[1];
    expect(graphCall[0]).toBe('https://graph.microsoft.com/v1.0/me');
    expect(graphCall[1].headers.Authorization).toBe('Bearer mock-access-token-12345');

    // Verify state record was consumed (single-use)
    const consumedState = await prisma.oAuthState.findUnique({ where: { state } });
    expect(consumedState).toBeNull();

    // Verify MailboxConnection record and token encryption in DB
    const saved = await prisma.mailboxConnection.findUnique({
      where: { emailAddress: TEST_EMAIL },
    });
    expect(saved).not.toBeNull();
    expect(saved!.emailAddress).toBe(TEST_EMAIL);
    // Tokens must be encrypted, not plaintext
    expect(saved!.accessToken).not.toBe('mock-access-token-12345');
    expect(saved!.refreshToken).not.toBe('mock-refresh-token-67890');
    // Tokens must decrypt to original values
    expect(decrypt(saved!.accessToken)).toBe('mock-access-token-12345');
    expect(decrypt(saved!.refreshToken)).toBe('mock-refresh-token-67890');
    expect(saved!.tokenExpiresAt.getTime()).toBeGreaterThan(Date.now());
  });

  it('redirects to frontend settings on successful browser navigation', async () => {
    const { state } = await createOAuthState();

    fetchMock.mockResolvedValueOnce({
      ok: true,
      json: async () => ({
        access_token: 'mock-access-token-browser',
        refresh_token: 'mock-refresh-token-browser',
        expires_in: 3600,
      }),
    });

    fetchMock.mockResolvedValueOnce({
      ok: true,
      json: async () => ({
        id: 'ms-graph-user-id-002',
        userPrincipalName: TEST_EMAIL,
      }),
    });

    const res = await request(app)
      .get(`/api/integrations/microsoft/callback?code=mock-auth-code&state=${state}`)
      .set('Cookie', `oauth_state=${state}`);

    expect(res.status).toBe(302);
    expect(res.headers.location).toContain('/settings?microsoft=connected');
  });

  it('rejects missing authorization code with 400', async () => {
    const { state } = await createOAuthState();

    const res = await request(app)
      .get(`/api/integrations/microsoft/callback?state=${state}`)
      .set('Cookie', `oauth_state=${state}`);

    expect(res.status).toBe(400);
    expect(res.body.error.code).toBe('MISSING_CODE');
    expect(hasClearedOAuthCookie(res.headers['set-cookie'])).toBe(true);
  });

  it('rejects missing state query parameter with 400', async () => {
    const res = await request(app)
      .get('/api/integrations/microsoft/callback?code=auth-code-123')
      .set('Cookie', 'oauth_state=some-state');

    expect(res.status).toBe(400);
    expect(res.body.error.code).toBe('MISSING_STATE');
  });

  it('rejects missing oauth_state cookie with 400', async () => {
    const res = await request(app)
      .get('/api/integrations/microsoft/callback?code=auth-code-123&state=some-state');

    expect(res.status).toBe(400);
    expect(res.body.error.code).toBe('MISSING_COOKIE_STATE');
  });

  it('rejects state mismatch between cookie and query with 400', async () => {
    const res = await request(app)
      .get('/api/integrations/microsoft/callback?code=auth-code-123&state=state-from-query')
      .set('Cookie', 'oauth_state=different-state-in-cookie');

    expect(res.status).toBe(400);
    expect(res.body.error.code).toBe('STATE_MISMATCH');
    expect(hasClearedOAuthCookie(res.headers['set-cookie'])).toBe(true);
  });

  it('rejects invalid or unknown state with 400', async () => {
    const nonExistentState = 'non-existent-state-12345';

    const res = await request(app)
      .get(`/api/integrations/microsoft/callback?code=mock-code&state=${nonExistentState}`)
      .set('Cookie', `oauth_state=${nonExistentState}`);

    expect(res.status).toBe(400);
    expect(res.body.error.code).toBe('INVALID_STATE');
  });

  it('rejects expired state with 400 and cleans it up from the database', async () => {
    const expiredState = 'test-expired-state-' + Date.now();
    await prisma.oAuthState.create({
      data: {
        id: crypto.randomUUID(),
        state: expiredState,
        codeVerifier: 'expired-verifier',
        expiresAt: new Date(Date.now() - 60000), // 1 minute in the past
      },
    });

    const res = await request(app)
      .get(`/api/integrations/microsoft/callback?code=mock-code&state=${expiredState}`)
      .set('Cookie', `oauth_state=${expiredState}`);

    expect(res.status).toBe(400);
    expect(res.body.error.code).toBe('STATE_EXPIRED');

    // State record must have been deleted
    const checkDb = await prisma.oAuthState.findUnique({ where: { state: expiredState } });
    expect(checkDb).toBeNull();
  });

  it('rejects reused state on subsequent callback attempt', async () => {
    const { state } = await createOAuthState();

    fetchMock.mockResolvedValueOnce({
      ok: true,
      json: async () => ({
        access_token: 'mock-access-token',
        refresh_token: 'mock-refresh-token',
        expires_in: 3600,
      }),
    });

    fetchMock.mockResolvedValueOnce({
      ok: true,
      json: async () => ({
        id: 'ms-user-id',
        mail: TEST_EMAIL,
      }),
    });

    // First use: succeeds
    const firstRes = await request(app)
      .get(`/api/integrations/microsoft/callback?code=mock-code&state=${state}`)
      .set('Cookie', `oauth_state=${state}`)
      .set('Accept', 'application/json');
    expect(firstRes.status).toBe(200);

    // Second use with the same state: rejected as invalid state (single-use)
    const secondRes = await request(app)
      .get(`/api/integrations/microsoft/callback?code=mock-code&state=${state}`)
      .set('Cookie', `oauth_state=${state}`)
      .set('Accept', 'application/json');
    expect(secondRes.status).toBe(400);
    expect(secondRes.body.error.code).toBe('INVALID_STATE');
  });

  it('handles Microsoft user cancellation / error callback safely', async () => {
    const res = await request(app)
      .get('/api/integrations/microsoft/callback?error=access_denied&error_description=User+declined+consent')
      .set('Cookie', 'oauth_state=some-state');

    expect(res.status).toBe(400);
    expect(res.body.error.code).toBe('OAUTH_ERROR');
    expect(res.body.error.message).toContain('cancelled or denied');
    expect(hasClearedOAuthCookie(res.headers['set-cookie'])).toBe(true);
  });

  it('handles token exchange failure without exposing provider response body', async () => {
    const { state } = await createOAuthState();

    fetchMock.mockResolvedValueOnce({
      ok: false,
      status: 400,
      text: async () => 'sensitive_internal_microsoft_error_details',
    });

    const res = await request(app)
      .get(`/api/integrations/microsoft/callback?code=bad-code&state=${state}`)
      .set('Cookie', `oauth_state=${state}`);

    expect(res.status).toBe(502);
    expect(res.body.error.code).toBe('TOKEN_EXCHANGE_FAILED');
    // Sensitive provider details must not leak
    expect(JSON.stringify(res.body)).not.toContain('sensitive_internal_microsoft_error_details');
    expect(hasClearedOAuthCookie(res.headers['set-cookie'])).toBe(true);
  });

  it('handles Graph API profile failure safely with 502', async () => {
    const { state } = await createOAuthState();

    fetchMock.mockResolvedValueOnce({
      ok: true,
      json: async () => ({
        access_token: 'mock-access-token',
        expires_in: 3600,
      }),
    });

    fetchMock.mockResolvedValueOnce({
      ok: false,
      status: 401,
      text: async () => 'unauthorized_graph_call',
    });

    const res = await request(app)
      .get(`/api/integrations/microsoft/callback?code=mock-code&state=${state}`)
      .set('Cookie', `oauth_state=${state}`);

    expect(res.status).toBe(502);
    expect(res.body.error.code).toBe('GRAPH_API_FAILED');
    expect(hasClearedOAuthCookie(res.headers['set-cookie'])).toBe(true);
  });

  it('handles network error during token exchange with 502', async () => {
    const { state } = await createOAuthState();

    fetchMock.mockRejectedValueOnce(new Error('ECONNRESET'));

    const res = await request(app)
      .get(`/api/integrations/microsoft/callback?code=mock-code&state=${state}`)
      .set('Cookie', `oauth_state=${state}`);

    expect(res.status).toBe(502);
    expect(res.body.error.code).toBe('NETWORK_ERROR');
  });
});
