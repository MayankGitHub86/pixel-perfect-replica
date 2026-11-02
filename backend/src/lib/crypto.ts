// backend/src/lib/crypto.ts
// AES-256-GCM encryption for sensitive tokens.
// Uses TOKEN_ENCRYPTION_KEY (base64 32-byte) from environment.

import crypto from "crypto";

const ALGO = "aes-256-gcm";
const KEY_ENV = "TOKEN_ENCRYPTION_KEY";
const IV_LENGTH = 12; // 96 bits
const TAG_LENGTH = 16; // 128 bits

function getKey(): Buffer {
  const b64 = process.env[KEY_ENV];
  if (!b64) {
    throw new Error(`Missing environment variable ${KEY_ENV}`);
  }
  const key = Buffer.from(b64, "base64");
  if (key.length !== 32) {
    throw new Error(`${KEY_ENV} must be a 32‑byte base64 string`);
  }
  return key;
}

/**
 * Encrypt a UTF‑8 string.
 * Returns `${iv}.${cipher}.${tag}` where each part is base64‑encoded.
 */
export function encrypt(plainText: string): string {
  const iv = crypto.randomBytes(IV_LENGTH);
  const key = getKey();
  const cipher = crypto.createCipheriv(ALGO, key, iv, { authTagLength: TAG_LENGTH });
  const encrypted = Buffer.concat([cipher.update(plainText, "utf8"), cipher.final()]);
  const tag = cipher.getAuthTag();
  return `${iv.toString("base64")}.${encrypted.toString("base64")}.${tag.toString("base64")}`;
}

/**
 * Decrypt a payload created by {@link encrypt}.
 * Throws on any failure (invalid format, auth tag mismatch, etc.).
 */
export function decrypt(payload: string): string {
  const [ivB64, encryptedB64, tagB64] = payload.split(".");
  if (!ivB64 || !encryptedB64 || !tagB64) {
    throw new Error("Invalid encrypted payload format");
  }
  const iv = Buffer.from(ivB64, "base64");
  const encrypted = Buffer.from(encryptedB64, "base64");
  const tag = Buffer.from(tagB64, "base64");
  const key = getKey();
  const decipher = crypto.createDecipheriv(ALGO, key, iv, { authTagLength: TAG_LENGTH });
  decipher.setAuthTag(tag);
  const decrypted = Buffer.concat([decipher.update(encrypted), decipher.final()]);
  return decrypted.toString("utf8");
}

