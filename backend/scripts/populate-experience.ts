#!/usr/bin/env node
// backend/scripts/populate-experience.ts
// Self-contained script: adds 'Experience' column to CandidateTable,
// reads each candidate's CV, populates experienceYears and Experience.
// No imports from shared src/ files.

import crypto from 'crypto';
import path from 'path';
import fs from 'fs';
import { PrismaClient } from '@prisma/client';
import dotenv from 'dotenv';

dotenv.config({ path: path.resolve(process.cwd(), '.env') });
const prisma = new PrismaClient();

// ─── Inline crypto ──────────────────────────────────────────────────────────
function decryptToken(payload: string): string {
  const [ivB64, encryptedB64, tagB64] = payload.split('.');
  const iv = Buffer.from(ivB64, 'base64');
  const encrypted = Buffer.from(encryptedB64, 'base64');
  const tag = Buffer.from(tagB64, 'base64');
  const key = Buffer.from(process.env.TOKEN_ENCRYPTION_KEY!, 'base64');
  const decipher = crypto.createDecipheriv('aes-256-gcm', key, iv, { authTagLength: 16 });
  decipher.setAuthTag(tag);
  return Buffer.concat([decipher.update(encrypted), decipher.final()]).toString('utf8');
}

// ─── Inline token ───────────────────────────────────────────────────────────
async function getValidAccessToken(): Promise<string> {
  const conn = await prisma.mailboxConnection.findFirst({ orderBy: { createdAt: 'desc' } });
  if (!conn) throw new Error('No mailbox connection found');
  const decryptedAccess = decryptToken(conn.accessToken);
  const [, payloadB64] = decryptedAccess.split('.');
  let payload: any = {};
  try { payload = JSON.parse(Buffer.from(payloadB64, 'base64url').toString('utf8')); } catch { payload = {}; }
  const expiresAt = payload.exp ? payload.exp * 1000 : (conn.tokenExpiresAt as any)?.getTime?.();
  const now = Date.now();
  if (!expiresAt || expiresAt - now < 5 * 60 * 1000) {
    const refreshToken = decryptToken(conn.refreshToken);
    const tokenRes = await fetch(
      `https://login.microsoftonline.com/${process.env.MICROSOFT_TENANT_ID}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id: process.env.MICROSOFT_CLIENT_ID!,
          client_secret: process.env.MICROSOFT_CLIENT_SECRET!,
          grant_type: 'refresh_token',
          refresh_token: refreshToken,
          redirect_uri: process.env.MICROSOFT_REDIRECT_URI || 'http://localhost:3000/api/integrations/microsoft/callback',
        }).toString(),
      }
    );
    if (!tokenRes.ok) throw new Error(`Token refresh failed: ${tokenRes.status}`);
    const tokenData: any = await tokenRes.json();
    return tokenData.access_token;
  }
  return decryptedAccess;
}

// ─── Graph helpers ───────────────────────────────────────────────────────────
async function graphGet(url: string, token: string) {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const text = await res.text();
  let data: any;
  try { data = JSON.parse(text); } catch { data = text; }
  return { status: res.status, data };
}

async function graphPost(url: string, token: string, body: any) {
  const res = await fetch(url, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  const text = await res.text();
  let data: any;
  try { data = JSON.parse(text); } catch { data = text; }
  return { status: res.status, data };
}

async function graphPatch(url: string, token: string, body: any) {
  const res = await fetch(url, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  const text = await res.text();
  let data: any;
  try { data = JSON.parse(text); } catch { data = text; }
  return { status: res.status, data };
}

// ─── CV text extraction ─────────────────────────────────────────────────────
async function extractTextFromBuffer(buffer: Buffer, filename: string): Promise<{ text: string; error?: string }> {
  const ext = path.extname(filename).toLowerCase();
  if (ext === '.pdf') {
    try {
      const pdfParse = (await import('pdf-parse')).default;
      const result = await pdfParse(buffer);
      return { text: result.text?.trim() || '' };
    } catch (err: any) {
      return { text: '', error: `PDF parse: ${err.message}` };
    }
  } else if (ext === '.docx' || ext === '.doc') {
    try {
      const mammoth = (await import('mammoth')).default;
      const result = await mammoth.extractRawText({ buffer });
      return { text: result.value?.trim() || '' };
    } catch (err: any) {
      return { text: '', error: `DOCX parse: ${err.message}` };
    }
  }
  return { text: '', error: `Unsupported extension: ${ext}` };
}

// ─── Experience extraction ──────────────────────────────────────────────────
function extractExperience(text: string): string {
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);

  // Try to find experience section boundaries
  const sectionHeaders = [
    /work\s*experience/i, /professional\s*experience/i, /employment\s*history/i,
    /experience/i, /work\s*history/i, /career\s*summary/i,
  ];
  const otherSections = [
    /^education$/i, /^skills$/i, /^certifications?$/i, /^projects?$/i,
    /^personal\s*details?$/i, /^objective$/i, /^summary$/i, /^hobbies$/i,
    /^achievements?$/i, /^languages?$/i, /^references?$/i, /^declaration$/i,
    /^academic/i, /^technical\s*skills/i,
  ];

  let startIdx = -1;
  for (let i = 0; i < lines.length; i++) {
    if (sectionHeaders.some(re => re.test(lines[i])) && lines[i].length < 60) {
      startIdx = i + 1;
      break;
    }
  }

  let endIdx = lines.length;
  if (startIdx !== -1) {
    for (let i = startIdx; i < lines.length; i++) {
      if (otherSections.some(re => re.test(lines[i]))) {
        endIdx = i;
        break;
      }
    }
  }

  if (startIdx !== -1 && startIdx < endIdx) {
    return lines.slice(startIdx, endIdx).join('\n');
  }

  // Fallback: collect lines with employment-related keywords
  const keywords = [
    /company/i, /employer/i, /organization/i, /organisation/i,
    /position/i, /role/i, /designation/i,
    /responsibilit/i, /duties/i,
    /\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\w*\s+\d{4}/i,
    /\d{4}\s*[-–—]\s*(\d{4}|present|current|till\s*date)/i,
  ];
  const blocks: string[] = [];
  let current: string[] = [];
  for (const line of lines) {
    if (keywords.some(re => re.test(line))) {
      current.push(line);
    } else if (current.length) {
      blocks.push(current.join('\n'));
      current = [];
    }
  }
  if (current.length) blocks.push(current.join('\n'));
  return blocks.join('\n\n');
}

function estimateYearsOfExperience(text: string): string {
  // Explicit statements
  const explicit = text.match(/(\d+(?:\.\d+)?)\+?\s*years?\s*(?:of\s*)?(?:experience|exp|professional)/i);
  if (explicit) return explicit[1];
  const total = text.match(/total\s*(?:work\s*)?experience\s*[:\-–]\s*(\d+(?:\.\d+)?)\s*years?/i);
  if (total) return total[1];
  // Sum date ranges
  const rangeRegex = /(\d{4})\s*[-–—]\s*(\d{4}|present|current|till\s*date|ongoing)/gi;
  let sum = 0;
  let m: RegExpExecArray | null;
  const currentYear = new Date().getFullYear();
  while ((m = rangeRegex.exec(text)) !== null) {
    const start = parseInt(m[1], 10);
    const endStr = m[2].toLowerCase();
    const end = /\d{4}/.test(m[2]) ? parseInt(m[2], 10) : currentYear;
    if (end >= start && end <= currentYear + 1) sum += end - start;
  }
  return sum > 0 ? sum.toString() : '';
}

// ─── Locate and read CV buffer for a candidate ─────────────────────────────
async function loadCvBuffer(candidateId: string, token: string): Promise<{ buffer: Buffer; filename: string } | { error: string }> {
  // 1. Check for local file via Attachment.storageKey
  const attachment = await prisma.attachment.findFirst({
    where: { candidateId },
    orderBy: { createdAt: 'desc' },
  });

  if (attachment?.storageKey) {
    // If storageKey is an absolute path to a local file, read it
    if (fs.existsSync(attachment.storageKey)) {
      const buffer = fs.readFileSync(attachment.storageKey);
      return { buffer, filename: attachment.filename };
    }
  }

  // 2. Check for OneDrive driveItem via Candidate.resumeRef
  const candidate = await prisma.candidate.findUnique({ where: { id: candidateId } });
  if (candidate?.resumeRef && candidate.resumeRef.includes('!')) {
    // It's a OneDrive driveItem ID like "5082520FEDE864C4!s9eafdf5f..."
    const driveItemId = candidate.resumeRef;
    const downloadRes = await graphGet(
      `https://graph.microsoft.com/v1.0/me/drive/items/${driveItemId}`,
      token,
    );
    if (downloadRes.status === 200 && downloadRes.data['@microsoft.graph.downloadUrl']) {
      const dlUrl = downloadRes.data['@microsoft.graph.downloadUrl'];
      const cvRes = await fetch(dlUrl);
      if (cvRes.ok) {
        const buffer = Buffer.from(await cvRes.arrayBuffer());
        return { buffer, filename: candidate.resumeFileName || downloadRes.data.name || 'resume.docx' };
      }
      return { error: `Download failed: ${cvRes.status}` };
    }
    // Try /content endpoint
    const contentRes = await fetch(
      `https://graph.microsoft.com/v1.0/me/drive/items/${driveItemId}/content`,
      { headers: { Authorization: `Bearer ${token}` }, redirect: 'follow' },
    );
    if (contentRes.ok) {
      const buffer = Buffer.from(await contentRes.arrayBuffer());
      return { buffer, filename: candidate.resumeFileName || 'resume.docx' };
    }
    return { error: `OneDrive item ${driveItemId}: metadata=${downloadRes.status}, content=${contentRes.status}` };
  }

  return { error: 'No CV storage reference found' };
}

// ─── Main ────────────────────────────────────────────────────────────────────
async function main() {
  console.log('=== populate-experience.ts ===\n');

  // Step 1: Get token and workbook info
  const token = await getValidAccessToken();
  console.log('[1/7] Token acquired');

  const wb = await prisma.excelWorkbook.findFirst();
  if (!wb) throw new Error('Excel workbook record not found');
  const base = `https://graph.microsoft.com/v1.0/me/drive/items/${wb.driveItemId}/workbook`;

  // Step 2: Get CandidateTable
  const tablesRes = await graphGet(`${base}/tables`, token);
  if (tablesRes.status !== 200) throw new Error(`Tables fetch failed: ${tablesRes.status}`);
  const candTable = (tablesRes.data.value || []).find((t: any) => t.name === 'CandidateTable');
  if (!candTable) throw new Error('CandidateTable not found');
  console.log('[2/7] CandidateTable found');

  // Step 3: Get current columns
  const colsRes = await graphGet(`${base}/tables/${candTable.id}/columns`, token);
  if (colsRes.status !== 200) throw new Error(`Columns fetch failed: ${colsRes.status}`);
  const columns: string[] = (colsRes.data.value || []).map((c: any) => c.name);
  console.log(`[3/7] Current columns (${columns.length}): ${JSON.stringify(columns)}`);

  const expYearsIdx = columns.indexOf('experienceYears');
  if (expYearsIdx === -1) throw new Error('experienceYears column not found');

  // Step 4: Add 'Experience' column if missing
  let experienceIdx = columns.indexOf('Experience');
  if (experienceIdx === -1) {
    console.log('[4/7] Adding "Experience" column...');
    const addColRes = await graphPost(`${base}/tables/${candTable.id}/columns/add`, token, {
      name: 'Experience',
      index: expYearsIdx + 1, // immediately after experienceYears
    });
    if (addColRes.status !== 200 && addColRes.status !== 201) {
      throw new Error(`Failed to add Experience column: ${addColRes.status} ${JSON.stringify(addColRes.data)}`);
    }
    console.log(`  Added "Experience" column (status ${addColRes.status})`);

    // Re-fetch columns to get updated indices
    const newColsRes = await graphGet(`${base}/tables/${candTable.id}/columns`, token);
    const newCols: string[] = (newColsRes.data.value || []).map((c: any) => c.name);
    experienceIdx = newCols.indexOf('Experience');
    columns.length = 0;
    columns.push(...newCols);
    console.log(`  Updated columns (${columns.length}): ${JSON.stringify(columns)}`);
  } else {
    console.log('[4/7] "Experience" column already exists');
  }

  // Step 5: Read current rows
  const rowsRes = await graphGet(`${base}/tables/${candTable.id}/rows`, token);
  if (rowsRes.status !== 200) throw new Error(`Rows fetch failed: ${rowsRes.status}`);
  const rows = rowsRes.data.value || [];
  console.log(`[5/7] Row count: ${rows.length}`);

  // Column indices we need
  const idIdx = columns.indexOf('id');
  const nameIdx = columns.indexOf('name');
  const expYearsNewIdx = columns.indexOf('experienceYears');
  const experienceNewIdx = columns.indexOf('Experience');

  // Step 6: Process each row
  console.log('[6/7] Processing candidates...\n');
  const results: { name: string; status: string; yearsSet: string; expLength: number }[] = [];
  const problematic: string[] = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const vals: any[] = row.values?.[0] || [];
    const candidateId = vals[idIdx];
    const candidateName = vals[nameIdx] || `Row ${i}`;
    const existingYears = vals[expYearsNewIdx];
    const existingExp = vals[experienceNewIdx];

    // Skip synthetic candidates (they are only in Prisma DB, not in Excel, but just in case)
    if (String(candidateName).includes('[DEV-SYNTHETIC]')) {
      console.log(`  SKIP ${candidateName} (synthetic)`);
      continue;
    }

    console.log(`  Processing: ${candidateName} (${candidateId})`);

    // Load the CV
    const cvResult = await loadCvBuffer(candidateId, token);
    if ('error' in cvResult) {
      problematic.push(`${candidateName}: ${cvResult.error}`);
      console.log(`    ⚠ ${cvResult.error}`);
      continue;
    }

    console.log(`    CV loaded: ${cvResult.filename} (${cvResult.buffer.length} bytes)`);

    // Parse text
    const parseResult = await extractTextFromBuffer(cvResult.buffer, cvResult.filename);
    if (parseResult.error) {
      problematic.push(`${candidateName}: ${parseResult.error}`);
      console.log(`    ⚠ ${parseResult.error}`);
      continue;
    }
    if (!parseResult.text) {
      problematic.push(`${candidateName}: empty text from CV (possibly scanned)`);
      console.log(`    ⚠ Empty text from CV`);
      continue;
    }

    console.log(`    Text extracted: ${parseResult.text.length} chars`);

    // Extract experience details
    const experienceText = extractExperience(parseResult.text);
    const yearsEstimate = estimateYearsOfExperience(parseResult.text);

    console.log(`    Experience text: ${experienceText.length} chars`);
    console.log(`    Years estimate: ${yearsEstimate || '(could not determine)'}`);

    // Build updated row values
    const newVals = [...vals];

    // Update experienceYears only if currently empty/blank and we have an estimate
    if (yearsEstimate && (!existingYears || String(existingYears).trim() === '')) {
      newVals[expYearsNewIdx] = yearsEstimate;
    }
    // Update Experience column
    newVals[experienceNewIdx] = experienceText || '';

    // PATCH the row
    const rowIndex = row.index;
    const patchRes = await graphPatch(
      `${base}/tables/${candTable.id}/rows/itemAt(index=${rowIndex})`,
      token,
      { values: [newVals] },
    );

    if (patchRes.status !== 200) {
      problematic.push(`${candidateName}: row update failed (${patchRes.status})`);
      console.log(`    ✗ Update failed: ${patchRes.status} ${JSON.stringify(patchRes.data).substring(0, 200)}`);
    } else {
      results.push({
        name: candidateName,
        status: 'updated',
        yearsSet: yearsEstimate || String(existingYears) || '',
        expLength: experienceText.length,
      });
      console.log(`    ✓ Updated`);
    }
  }

  // Step 7: Verification — read back
  console.log('\n[7/7] Verification read-back...\n');
  const verifyColsRes = await graphGet(`${base}/tables/${candTable.id}/columns`, token);
  const verifyCols: string[] = (verifyColsRes.data.value || []).map((c: any) => c.name);
  console.log(`Final columns (${verifyCols.length}): ${JSON.stringify(verifyCols)}`);

  // Show columns around experienceYears
  const eyIdx = verifyCols.indexOf('experienceYears');
  const exIdx = verifyCols.indexOf('Experience');
  console.log(`\nColumns around experienceYears:`);
  for (let i = Math.max(0, eyIdx - 1); i <= Math.min(verifyCols.length - 1, exIdx + 1); i++) {
    console.log(`  [${i}] ${verifyCols[i]}${i === eyIdx ? ' ← experienceYears' : i === exIdx ? ' ← Experience' : ''}`);
  }

  const verifyRowsRes = await graphGet(`${base}/tables/${candTable.id}/rows`, token);
  if (verifyRowsRes.status !== 200) {
    console.error('Failed to read back rows:', verifyRowsRes.status);
  } else {
    const verifyRows = verifyRowsRes.data.value || [];
    console.log(`\nFinal row count: ${verifyRows.length}`);

    for (const row of verifyRows) {
      const v = row.values?.[0] || [];
      const name = v[verifyCols.indexOf('name')];
      const uuid = v[verifyCols.indexOf('id')];
      const years = v[verifyCols.indexOf('experienceYears')];
      const exp = v[verifyCols.indexOf('Experience')];

      console.log(`\n  Candidate: ${name}`);
      console.log(`  UUID: ${uuid}`);
      console.log(`  experienceYears: ${years}`);
      console.log(`  Experience (first 500 chars): ${String(exp || '').substring(0, 500)}`);
    }
  }

  console.log('\n=== Summary ===');
  console.log(`Rows updated: ${results.length}`);
  console.log(`Problematic: ${problematic.length}`);
  if (problematic.length > 0) {
    console.log('Issues:');
    problematic.forEach(p => console.log(`  - ${p}`));
  }

  await prisma.$disconnect();
  console.log('\nDone.');
}

main().catch(err => {
  console.error('FATAL:', err);
  process.exit(1);
});
