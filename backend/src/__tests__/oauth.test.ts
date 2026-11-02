// backend/src/__tests__/oauth.test.ts
// Focused tests for Microsoft OAuth PKCE state management,
// token encryption, and refresh rotation.
// All Microsoft network calls are mocked — no real credentials used.

import { describe, it, expect, vi, beforeEach } from 'vitest';

// ─── Encryption key for tests ──────────────────────────────────────────────
process.env.TOKEN_ENCRYPTION_KEY = Buffer.from('b'.repeat(32)).toString('base64');

// ─── Mock prisma ─────────────────────────────────────────────────────────────
// vi.mock is hoisted to the top of the file by Vitest, so factory functions
// must not reference variables declared in module scope. Use vi.hoisted() for
// any values that must be shared between the factory and the test body.
const { mockOAuthState, mockMailboxConnection } = vi.hoisted(() => ({
  mockOAuthState: {
    findUnique: vi.fn(),
    create: vi.fn(),
    delete: vi.fn(),
  },
  mockMailboxConnection: {
    findUnique: vi.fn(),
    update: vi.fn(),
  },
}));

vi.mock('../lib/db', () => ({
  prisma: {
    oAuthState: mockOAuthState,
    mailboxConnection: mockMailboxConnection,
  },
}));

// ─── Import after mocks ───────────────────────────────────────────────────────
import { consumeOAuthState, createOAuthState, generateCodeVerifier, generateCodeChallenge } from '../lib/oauth';
import { refreshAccessToken } from '../lib/tokenRefresh';
import { encrypt, decrypt } from '../lib/crypto';

// ─── Helpers ─────────────────────────────────────────────────────────────────
const FUTURE = new Date(Date.now() + 10 * 60 * 1000);
const PAST   = new Date(Date.now() - 10 * 60 * 1000);

const SAMPLE_STATE = {
  id: 'state-uuid-1',
  state: 'random-state-value',
  codeVerifier: 'verifier-abc',
  expiresAt: FUTURE,
  createdAt: new Date(),
};

// Build SAMPLE_MAILBOX lazily so encrypt() runs after TOKEN_ENCRYPTION_KEY is set
function makeSampleMailbox() {
  return {
    id: 'mailbox-1',
    refreshToken: encrypt('stored-refresh-token'),
    accessToken:  encrypt('stored-access-token'),
    tokenExpiresAt: PAST,
    updatedAt: PAST,
    emailAddress: 'user@example.com',
    userId: 'dev',
    createdAt: PAST,
  };
}

// Mock global fetch for token-refresh tests
const fetchMock = vi.fn();
global.fetch = fetchMock as any;

// ─── PKCE helpers ─────────────────────────────────────────────────────────────
describe('PKCE helpers', () => {
  it('generateCodeVerifier produces a URL-safe base64 string of sufficient length', () => {
    const v = generateCodeVerifier();
    expect(v.length).toBeGreaterThanOrEqual(43); // RFC 7636 min 43 chars
    expect(v).toMatch(/^[A-Za-z0-9\-_]+$/);       // URL-safe only — no + / =
  });

  it('generateCodeChallenge produces a deterministic S256 challenge', () => {
    const verifier = 'dBjftJeZ4CVP-mB92K27uhbUJU1p1r_wW1gFWFOEjXk';
    const challenge = generateCodeChallenge(verifier);
    expect(challenge).toMatch(/^[A-Za-z0-9\-_]+$/);
    expect(challenge.length).toBeGreaterThan(0);
    // Must be deterministic
    expect(generateCodeChallenge(verifier)).toBe(challenge);
  });

  it('verifier and challenge are distinct values', () => {
    const v = generateCodeVerifier();
    const c = generateCodeChallenge(v);
    expect(v).not.toBe(c);
  });
});

// ─── Token encryption / decryption ───────────────────────────────────────────
describe('Token encryption / decryption (AES-256-GCM)', () => {
  it('round-trips a plaintext token', () => {
    const plain = 'eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.access_token_example';
    const cipher = encrypt(plain);
    expect(cipher).not.toBe(plain);
    expect(decrypt(cipher)).toBe(plain);
  });

  it('produces distinct ciphertext on each call (random IV)', () => {
    const plain = 'same-plaintext';
    const c1 = encrypt(plain);
    const c2 = encrypt(plain);
    expect(c1).not.toBe(c2);
    expect(decrypt(c1)).toBe(plain);
    expect(decrypt(c2)).toBe(plain);
  });

  it('throws on tampered ciphertext (GCM auth tag mismatch)', () => {
    const cipher = encrypt('sensitive-value');
    const parts = cipher.split('.');
    parts[1] = Buffer.from('corrupted').toString('base64');
    expect(() => decrypt(parts.join('.'))).toThrow();
  });

  it('throws on malformed payload (missing segments)', () => {
    expect(() => decrypt('onlyone')).toThrow('Invalid encrypted payload format');
  });
});

// ─── consumeOAuthState ───────────────────────────────────────────────────────
describe('consumeOAuthState', () => {
  beforeEach(() => {
    vi.resetAllMocks();
  });

  it('returns the codeVerifier for a valid, unexpired state', async () => {
    mockOAuthState.findUnique.mockResolvedValue(SAMPLE_STATE);
    mockOAuthState.delete.mockResolvedValue({});

    const verifier = await consumeOAuthState(SAMPLE_STATE.state);

    expect(verifier).toBe(SAMPLE_STATE.codeVerifier);
    // Single-use: must be deleted immediately after successful consumption
    expect(mockOAuthState.delete).toHaveBeenCalledWith({ where: { state: SAMPLE_STATE.state } });
  });

  it('throws "Invalid OAuth state" for unknown / forged state', async () => {
    mockOAuthState.findUnique.mockResolvedValue(null);

    await expect(consumeOAuthState('forged-state')).rejects.toThrow('Invalid OAuth state');
    expect(mockOAuthState.delete).not.toHaveBeenCalled();
  });

  it('throws "OAuth state expired" and cleans up the expired record', async () => {
    const expired = { ...SAMPLE_STATE, expiresAt: PAST };
    mockOAuthState.findUnique.mockResolvedValue(expired);
    mockOAuthState.delete.mockResolvedValue({});

    await expect(consumeOAuthState(expired.state)).rejects.toThrow('OAuth state expired');
    expect(mockOAuthState.delete).toHaveBeenCalledWith({ where: { state: expired.state } });
  });

  it('prevents reuse — second consumption of the same state fails', async () => {
    mockOAuthState.findUnique
      .mockResolvedValueOnce(SAMPLE_STATE) // first call — found
      .mockResolvedValueOnce(null);         // second call — already deleted
    mockOAuthState.delete.mockResolvedValue({});

    await consumeOAuthState(SAMPLE_STATE.state);
    await expect(consumeOAuthState(SAMPLE_STATE.state)).rejects.toThrow('Invalid OAuth state');
  });
});

// ─── createOAuthState ─────────────────────────────────────────────────────────
describe('createOAuthState', () => {
  beforeEach(() => {
    vi.resetAllMocks();
  });

  it('persists a state record and returns state + verifier', async () => {
    mockOAuthState.create.mockResolvedValue({});

    const { state, verifier } = await createOAuthState();

    expect(typeof state).toBe('string');
    expect(state.length).toBeGreaterThan(0);
    expect(typeof verifier).toBe('string');
    expect(verifier.length).toBeGreaterThan(0);

    expect(mockOAuthState.create).toHaveBeenCalledOnce();
    const createArg = mockOAuthState.create.mock.calls[0][0].data;
    expect(createArg.state).toBe(state);
    expect(createArg.codeVerifier).toBe(verifier);
    expect(createArg.expiresAt).toBeInstanceOf(Date);
    expect(createArg.expiresAt.getTime()).toBeGreaterThan(Date.now());
  });

  it('state and verifier are different values', async () => {
    mockOAuthState.create.mockResolvedValue({});
    const { state, verifier } = await createOAuthState();
    expect(state).not.toBe(verifier);
  });
});

// ─── refreshAccessToken — token rotation ─────────────────────────────────────
describe('refreshAccessToken — token rotation', () => {
  beforeEach(() => {
    vi.resetAllMocks();
  });

  it('rotates both access and refresh tokens when Microsoft returns a new refresh token', async () => {
    const mailbox = makeSampleMailbox();
    mockMailboxConnection.findUnique.mockResolvedValue(mailbox);
    fetchMock.mockResolvedValue({
      ok: true,
      json: async () => ({ access_token: 'new-access-token', refresh_token: 'new-refresh-token', expires_in: 3600 }),
    });
    mockMailboxConnection.update.mockResolvedValue({});

    await refreshAccessToken(mailbox.id);

    const { accessToken, refreshToken, tokenExpiresAt } = mockMailboxConnection.update.mock.calls[0][0].data;
    // Must be encrypted, not plaintext
    expect(accessToken).not.toBe('new-access-token');
    expect(refreshToken).not.toBe('new-refresh-token');
    // Must decrypt to correct values
    expect(decrypt(accessToken)).toBe('new-access-token');
    expect(decrypt(refreshToken)).toBe('new-refresh-token');
    // Expiry ~1 hour in the future
    expect(tokenExpiresAt).toBeInstanceOf(Date);
    expect(tokenExpiresAt.getTime()).toBeGreaterThan(Date.now() + 3500 * 1000);
  });

  it('preserves the existing refresh token when Microsoft does NOT return a new one', async () => {
    const mailbox = makeSampleMailbox();
    mockMailboxConnection.findUnique.mockResolvedValue(mailbox);
    fetchMock.mockResolvedValue({
      ok: true,
      json: async () => ({ access_token: 'new-access-token', expires_in: 3600 }),
    });
    mockMailboxConnection.update.mockResolvedValue({});

    await refreshAccessToken(mailbox.id);

    const { refreshToken } = mockMailboxConnection.update.mock.calls[0][0].data;
    // Original plaintext must be preserved (re-encrypted)
    expect(decrypt(refreshToken)).toBe('stored-refresh-token');
  });

  it('throws "MailboxConnection not found" when the mailbox does not exist', async () => {
    mockMailboxConnection.findUnique.mockResolvedValue(null);
    await expect(refreshAccessToken('no-such-id')).rejects.toThrow('MailboxConnection not found');
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it('throws on Microsoft token endpoint HTTP error and does NOT write to DB', async () => {
    const mailbox = makeSampleMailbox();
    mockMailboxConnection.findUnique.mockResolvedValue(mailbox);
    fetchMock.mockResolvedValue({ ok: false, status: 401, text: async () => 'invalid_grant' });

    await expect(refreshAccessToken(mailbox.id)).rejects.toThrow('Token refresh failed: 401');
    expect(mockMailboxConnection.update).not.toHaveBeenCalled();
  });

  it('throws on network error and does NOT write to DB', async () => {
    const mailbox = makeSampleMailbox();
    mockMailboxConnection.findUnique.mockResolvedValue(mailbox);
    fetchMock.mockRejectedValue(new Error('ECONNREFUSED'));

    await expect(refreshAccessToken(mailbox.id)).rejects.toThrow('ECONNREFUSED');
    expect(mockMailboxConnection.update).not.toHaveBeenCalled();
  });

  it('sends the decrypted refresh token to Microsoft — NOT the stored ciphertext', async () => {
    const mailbox = makeSampleMailbox();
    mockMailboxConnection.findUnique.mockResolvedValue(mailbox);
    fetchMock.mockResolvedValue({
      ok: true,
      json: async () => ({ access_token: 'a', expires_in: 3600 }),
    });
    mockMailboxConnection.update.mockResolvedValue({});

    await refreshAccessToken(mailbox.id);

    const fetchBody = fetchMock.mock.calls[0][1].body as string;
    const params = new URLSearchParams(fetchBody);
    expect(params.get('refresh_token')).toBe('stored-refresh-token');
    // The raw ciphertext must NOT appear in the HTTP request
    expect(fetchBody).not.toContain(mailbox.refreshToken);
  });
});
