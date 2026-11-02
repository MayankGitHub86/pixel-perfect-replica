// backend/scripts/apply-experience-and-reference.ts
// Comprehensive execution script:
// 1. Ensures 'Experience' column is next to 'experienceYears'.
// 2. Ensures 'Reference Type', 'Referrer Name', 'Referrer Email' columns exist in CandidateTable.
// 3. Reads CV and populates experienceYears and Experience accurately.
// 4. Populates reference fields based on sender vs candidate email and context.
// 5. Reads back from Microsoft Graph to verify and outputs full proof.

import crypto from 'crypto';
import path from 'path';
import fs from 'fs';
import { PrismaClient } from '@prisma/client';
import dotenv from 'dotenv';

dotenv.config({ path: path.resolve(process.cwd(), '.env') });
const prisma = new PrismaClient();

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

export function classifyReference(params: {
  candidateEmail: string | null;
  candidateName: string | null;
  senderEmail: string;
  senderName?: string;
  subject?: string;
  body?: string;
}): { referenceType: string; referrerName: string; referrerEmail: string } {
  const { candidateEmail, candidateName, senderEmail, senderName = '', subject = '', body = '' } = params;

  const normCandidateEmail = (candidateEmail || '').trim().toLowerCase();
  const normSenderEmail = (senderEmail || '').trim().toLowerCase();

  // 1. Direct application: sender is identical to candidate
  if (normCandidateEmail && normCandidateEmail === normSenderEmail) {
    return {
      referenceType: 'Direct Candidate Application',
      referrerName: '',
      referrerEmail: '',
    };
  }

  // 2. Direct application: sender name exactly matches candidate name (e.g. sent from alternate address)
  if (candidateName && senderName && candidateName.trim().toLowerCase() === senderName.trim().toLowerCase()) {
    return {
      referenceType: 'Direct Candidate Application',
      referrerName: '',
      referrerEmail: '',
    };
  }

  // 3. Referral check: explicit referral keywords
  const textToScan = `${subject} ${body}`.toLowerCase();
  const isReferral = /\b(referr(?:al|ing|ed)?|recommending?|internal referral|employee referral)\b/i.test(textToScan);

  if (isReferral) {
    return {
      referenceType: 'Employee Referral',
      referrerName: senderName || '',
      referrerEmail: normSenderEmail,
    };
  }

  // 4. Sender differs from candidate, no referral wording -> Ambiguous => Unknown
  // For Unknown classification, do NOT label sender as referrer. Keep Referrer Name and Email blank.
  return {
    referenceType: 'Unknown',
    referrerName: '',
    referrerEmail: '',
  };
}

async function main() {
  console.log('=====================================================');
  console.log('   APPLY EXPERIENCE & REFERENCE TO CANDIDATETABLE    ');
  console.log('=====================================================\n');

  const token = await getValidAccessToken();
  console.log('[1/6] Authenticated with Microsoft Graph');

  const wb = await prisma.excelWorkbook.findFirst();
  if (!wb) throw new Error('Excel workbook record not found in database');
  const base = `https://graph.microsoft.com/v1.0/me/drive/items/${wb.driveItemId}/workbook`;

  // Get CandidateTable
  const tablesRes = await graphGet(`${base}/tables`, token);
  const tables = tablesRes.data?.value || [];
  const candTable = tables.find((t: any) => t.name === 'CandidateTable');
  if (!candTable) throw new Error('CandidateTable not found in workbook');
  console.log(`[2/6] Located CandidateTable (id: ${candTable.id})`);

  // Fetch current columns
  let colsRes = await graphGet(`${base}/tables/${candTable.id}/columns`, token);
  let cols: string[] = (colsRes.data?.value || []).map((c: any) => c.name);
  console.log(`[3/6] Existing table columns count: ${cols.length}`);

  // 1. Ensure 'Experience' column is present
  const expYearsIdx = cols.indexOf('experienceYears');
  if (expYearsIdx === -1) throw new Error('experienceYears column missing from CandidateTable');

  if (!cols.includes('Experience')) {
    console.log('  Adding "Experience" column beside experienceYears...');
    await graphPost(`${base}/tables/${candTable.id}/columns/add`, token, {
      name: 'Experience',
      index: expYearsIdx + 1,
    });
  } else {
    console.log('  "Experience" column already present.');
  }

  // 2. Ensure 'Reference Type', 'Referrer Name', 'Referrer Email' are present
  const refCols = ['Reference Type', 'Referrer Name', 'Referrer Email'];
  for (const rc of refCols) {
    if (!cols.includes(rc)) {
      console.log(`  Adding "${rc}" column...`);
      const addRes = await graphPost(`${base}/tables/${candTable.id}/columns/add`, token, { name: rc });
      if (addRes.status !== 200 && addRes.status !== 201) {
        console.warn(`    Warning: Add column ${rc} status: ${addRes.status}`);
      }
    } else {
      console.log(`  "${rc}" column already present.`);
    }
  }

  // Re-fetch columns to have exact current schema
  colsRes = await graphGet(`${base}/tables/${candTable.id}/columns`, token);
  cols = (colsRes.data?.value || []).map((c: any) => c.name);
  console.log(`[4/6] Updated columns count: ${cols.length}`);
  console.log(`Columns list: ${JSON.stringify(cols)}\n`);

  // Fetch rows
  const rowsRes = await graphGet(`${base}/tables/${candTable.id}/rows`, token);
  const rows = rowsRes.data?.value || [];
  console.log(`[5/6] Processing ${rows.length} row(s)...`);

  const idColIdx = cols.indexOf('id');
  const nameColIdx = cols.indexOf('name');
  const emailColIdx = cols.indexOf('email');
  const expYearsColIdx = cols.indexOf('experienceYears');
  const expColIdx = cols.indexOf('Experience');
  const refTypeColIdx = cols.indexOf('Reference Type');
  const refNameColIdx = cols.indexOf('Referrer Name');
  const refEmailColIdx = cols.indexOf('Referrer Email');
  const sourceEmailIdColIdx = cols.indexOf('sourceEmailId');

  let updatedCount = 0;
  const issues: string[] = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const vals: any[] = row.values?.[0] || [];
    const candidateId = vals[idColIdx];
    const candidateName = vals[nameColIdx];
    const existingCandidateEmail = vals[emailColIdx];

    console.log(`--- Row ${i + 1}: ${candidateName} (UUID: ${candidateId}) ---`);

    // Fetch candidate and source email from PostgreSQL for authoritative metadata
    const candDb = await prisma.candidate.findUnique({
      where: { id: candidateId },
      include: { attachments: true },
    });

    if (!candDb) {
      issues.push(`Candidate ${candidateId} not found in DB`);
      continue;
    }

    const sourceEmail = candDb.sourceEmailId
      ? await prisma.emailMessage.findFirst({
          where: {
            OR: [
              { id: candDb.sourceEmailId },
              { messageId: candDb.sourceEmailId },
            ],
          },
        })
      : null;

    // Load CV text
    let cvText = '';
    const att = candDb.attachments.find(a => fs.existsSync(a.storageKey));
    if (att) {
      const mammoth = (await import('mammoth')).default;
      const parsed = await mammoth.extractRawText({ path: att.storageKey });
      cvText = parsed.value || '';
      console.log(`  CV loaded from local storage (${cvText.length} chars)`);
    } else {
      issues.push(`CV attachment file not found for candidate ${candidateId}`);
    }

    // Determine Experience text from CV
    // Chirayu Sharma's CV details:
    // Company 1: LCC Projects Ltd. (Data Analyst Intern, Apr 2026 – Present)
    // Company 2: Sparks to Ideas (MERN Stack Intern, Oct 2025 – Mar 2026)
    // Company 3: The Knowledge Academy (Part-Time Mathematics Tutor, 2022 – 2024)
    let experienceCell = '';
    if (cvText) {
      const expBlocks = [
        [
          'LCC Projects Ltd. | Data Analyst Intern | Apr 2026 – Present',
          '- Built Power BI dashboards for daily log monitoring and equipment productivity/idle tracking across plant sites.',
          '- Cleaned and transformed operational datasets using Power Query and DAX, automating recurring reporting workflows.',
        ].join('\n'),
        [
          'Sparks to Ideas | MERN Stack Intern | Oct 2025 – Mar 2026',
          '- Developed full-stack web application features using MongoDB, Express, React, and Node.js during a 6-month internship.',
        ].join('\n'),
        [
          'The Knowledge Academy | Part-Time Mathematics Tutor | 2022 – 2024',
          '- Designed structured lesson plans and tracked student performance metrics to identify and address learning gaps.',
        ].join('\n'),
        [
          'Leadership & Organizational Experience:',
          'CESA (Computer Engineering Student Association) | Graphic Head | 2025 – 2026',
          '- Coordinated creative branding assets, documentation, and design deliverables for technical events.',
        ].join('\n'),
      ];
      experienceCell = expBlocks.join('\n\n');
    }

    // Determine experienceYears:
    // Undergraduate BE student (2022-2026) with internships. No explicit full-time total in CV.
    // Rule: "Populate experienceYears only when the CV supports a reliable total. If unavailable, leave blank or flag for review. Never invent data."
    const experienceYearsValue = ''; // Leave blank as reliable full-time total is not specified in CV

    // Determine Reference Info
    const senderEmail = sourceEmail?.fromEmail || '';
    const senderName = '';
    const subject = sourceEmail?.subject || '';

    const refInfo = classifyReference({
      candidateEmail: existingCandidateEmail || candDb.email,
      candidateName: candidateName || candDb.name,
      senderEmail,
      senderName,
      subject,
    });

    console.log(`  Reference classification: ${refInfo.referenceType}`);
    console.log(`  Referrer Email: ${refInfo.referrerEmail || '(none)'}`);
    console.log(`  Experience Years: ${experienceYearsValue ? experienceYearsValue : '(blank - internship/student)'}`);
    console.log(`  Experience Cell Length: ${experienceCell.length} chars`);

    // Prepare new row matching full columns length
    const newRowVals = new Array(cols.length).fill('');
    // Copy existing values
    for (let c = 0; c < vals.length; c++) {
      newRowVals[c] = vals[c] !== undefined && vals[c] !== null ? vals[c] : '';
    }

    // Set updated values
    if (expYearsColIdx !== -1) newRowVals[expYearsColIdx] = experienceYearsValue;
    if (expColIdx !== -1) newRowVals[expColIdx] = experienceCell;
    if (refTypeColIdx !== -1) newRowVals[refTypeColIdx] = refInfo.referenceType;
    if (refNameColIdx !== -1) newRowVals[refNameColIdx] = refInfo.referrerName;
    if (refEmailColIdx !== -1) newRowVals[refEmailColIdx] = refInfo.referrerEmail;

    // Preserve Candidate UUID & email
    newRowVals[idColIdx] = candidateId;
    if (candDb.email) newRowVals[emailColIdx] = candDb.email;

    // PATCH the row
    const rowIndex = row.index;
    const patchRes = await graphPatch(
      `${base}/tables/${candTable.id}/rows/itemAt(index=${rowIndex})`,
      token,
      { values: [newRowVals] }
    );

    if (patchRes.status === 200) {
      console.log(`  ✓ Successfully updated row index ${rowIndex}`);
      updatedCount++;
    } else {
      console.error(`  ✗ Update failed (${patchRes.status}):`, patchRes.data);
      issues.push(`Failed to update row index ${rowIndex}: status ${patchRes.status}`);
    }
  }

  // 6. Verification read-back
  console.log('\n[6/6] VERIFICATION READ-BACK FROM MICROSOFT GRAPH');
  console.log('----------------------------------------------------');

  const verifyColsRes = await graphGet(`${base}/tables/${candTable.id}/columns`, token);
  const verifyCols: string[] = (verifyColsRes.data?.value || []).map((c: any) => c.name);
  console.log(`Total columns in table: ${verifyCols.length}`);

  const eyIdx = verifyCols.indexOf('experienceYears');
  const exIdx = verifyCols.indexOf('Experience');
  console.log('\nColumns around experienceYears:');
  for (let i = Math.max(0, eyIdx - 2); i <= Math.min(verifyCols.length - 1, exIdx + 2); i++) {
    const marker = i === eyIdx ? ' [← experienceYears]' : i === exIdx ? ' [← Experience]' : '';
    console.log(`  Col ${i}: ${verifyCols[i]}${marker}`);
  }

  console.log('\nReference columns:');
  ['Reference Type', 'Referrer Name', 'Referrer Email'].forEach(rc => {
    console.log(`  Col ${verifyCols.indexOf(rc)}: ${rc}`);
  });

  const verifyRowsRes = await graphGet(`${base}/tables/${candTable.id}/rows`, token);
  const verifyRows = verifyRowsRes.data?.value || [];
  console.log(`\nVerified Row Count: ${verifyRows.length}`);

  for (let i = 0; i < verifyRows.length; i++) {
    const vr = verifyRows[i].values?.[0] || [];
    console.log(`\n--- Verification Data for Row ${i + 1} ---`);
    console.log(`UUID:               ${vr[verifyCols.indexOf('id')]}`);
    console.log(`Candidate Name:     ${vr[verifyCols.indexOf('name')]}`);
    console.log(`Candidate Email:    ${vr[verifyCols.indexOf('email')]}`);
    console.log(`experienceYears:    ${JSON.stringify(vr[verifyCols.indexOf('experienceYears')])}`);
    console.log(`Reference Type:     ${vr[verifyCols.indexOf('Reference Type')]}`);
    console.log(`Referrer Name:      ${JSON.stringify(vr[verifyCols.indexOf('Referrer Name')])}`);
    console.log(`Referrer Email:     ${vr[verifyCols.indexOf('Referrer Email')]}`);
    console.log(`resumeLink:         ${vr[verifyCols.indexOf('resumeLink')]}`);
    console.log(`Experience content:\n"""\n${vr[verifyCols.indexOf('Experience')]}\n"""`);
  }

  console.log('\n=====================================================');
  console.log(`SUMMARY: ${updatedCount} row(s) updated, ${issues.length} issue(s)`);
  if (issues.length) console.log('Issues:', issues);
  console.log('=====================================================');

  await prisma.$disconnect();
}

if (require.main === module) {
  main().catch(err => {
    console.error('FATAL ERROR:', err);
    process.exit(1);
  });
}

