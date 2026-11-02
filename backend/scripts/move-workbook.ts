// move-workbook.ts
// Read‑only‑safe move of Candidate_Master_RealData.xlsx from OneDrive root to Documents folder.
// Uses the project's existing Graph auth helpers (same pattern as visibility‑diagnostics).

import { PrismaClient } from '@prisma/client';
import { decrypt } from '../src/lib/crypto.js';

const prisma = new PrismaClient({
  log: process.env.NODE_ENV === 'development' ? ['query', 'error', 'warn'] : ['error'],
});

async function getAccessToken() {
  const conn = await prisma.mailboxConnection.findFirst({ orderBy: { createdAt: 'desc' } });
  if (!conn) throw new Error('No mailbox connection found');
  return decrypt(conn.accessToken);
}

async function graphRequest(path: string, method = 'GET', token: string, body?: any) {
  const url = `https://graph.microsoft.com/v1.0${path}`;
  const opts: any = {
    method,
    headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
  };
  if (body) {
    opts.headers['Content-Type'] = 'application/json';
    opts.body = JSON.stringify(body);
  }
  const res = await fetch(url, opts);
  const result: any = { status: res.status };
  const txt = await res.text();
  if (!res.ok) {
    let code = null, message = null;
    try { const err = JSON.parse(txt); code = err.error?.code; message = err.error?.message; } catch (_) {}
    result.error = { code: code || 'UNKNOWN', message: message || txt.slice(0,200) };
    return result;
  }
  try { result.data = JSON.parse(txt); } catch (_) { result.data = txt; }
  return result;
}

async function main() {
  const token = await getAccessToken();
  const realId = '5082520FEDE864C4!s8dc280a78ce54f6b9ee4617268ede8bb';

  // -------- Pre‑move validation --------
  const curMeta = await graphRequest(`/me/drive/items/${realId}`, 'GET', token);
  if (curMeta.error) throw new Error(`Failed to fetch current metadata: ${curMeta.error.code} – ${curMeta.error.message}`);
  const { name, parentReference, webUrl } = curMeta.data;
  if (name !== 'Candidate_Master_RealData.xlsx' || curMeta.data.id !== realId) {
    throw new Error('Workbook name or ID does not match expected values');
  }
  if (!parentReference?.path?.startsWith('/drive/root:') || parentReference.path.includes('/Documents')) {
    throw new Error('Workbook is not currently located in the OneDrive root folder');
  }
  console.log('Pre‑move metadata verified.');

  // Verify Documents folder exists
  const docsFolder = await graphRequest('/me/drive/root:/Documents', 'GET', token);
  if (docsFolder.error) throw new Error(`Documents folder not accessible: ${docsFolder.error.code} – ${docsFolder.error.message}`);
  console.log('Documents folder exists.');

  // Verify DB pointer
  const dbRec = await prisma.excelWorkbook.findFirst();
  if (!dbRec) throw new Error('No ExcelWorkbook record found in DB');
  if (dbRec.driveItemId !== realId) throw new Error(`DB pointer mismatch (expected ${realId}, found ${dbRec.driveItemId})`);
  console.log('Database pointer verified.');

  // Capture worksheet/table counts before move
  const wsBefore = await graphRequest(`/me/drive/items/${realId}/workbook/worksheets`, 'GET', token);
  const tblBefore = await graphRequest(`/me/drive/items/${realId}/workbook/tables`, 'GET', token);
  const wsCountBefore = wsBefore.data?.value?.length ?? 0;
  const tblCountBefore = tblBefore.data?.value?.length ?? 0;

  // -------- Perform move --------
  const moveRes = await graphRequest(`/me/drive/items/${realId}`, 'PATCH', token, {
    parentReference: { path: '/drive/root:/Documents' },
  });
  if (moveRes.error) throw new Error(`Move operation failed: ${moveRes.error.code} – ${moveRes.error.message}`);
  console.log('Move operation succeeded.');

  // -------- Post‑move verification --------
  const postMeta = await graphRequest(`/me/drive/items/${realId}`, 'GET', token);
  if (postMeta.error) throw new Error(`Failed to fetch post‑move metadata: ${postMeta.error.code} – ${postMeta.error.message}`);
  const postParent = postMeta.data.parentReference?.path;
  if (!postParent?.includes('/Documents')) throw new Error('Post‑move parent path does not indicate Documents folder');
  console.log('Post‑move parent verified.');

  // Verify worksheet/table counts unchanged
  const wsAfter = await graphRequest(`/me/drive/items/${realId}/workbook/worksheets`, 'GET', token);
  const tblAfter = await graphRequest(`/me/drive/items/${realId}/workbook/tables`, 'GET', token);
  const wsCountAfter = wsAfter.data?.value?.length ?? 0;
  const tblCountAfter = tblAfter.data?.value?.length ?? 0;
  if (wsCountBefore !== wsCountAfter || tblCountBefore !== tblCountAfter) {
    throw new Error('Worksheet or table count changed after move');
  }
  console.log('Worksheet and table structure unchanged.');


  // Verify webUrl reachable (HEAD request)
  const webHead = await fetch(postMeta.data.webUrl, { method: 'HEAD' });
  if (!webHead.ok) throw new Error(`webUrl check failed: ${webHead.status}`);
  console.log('webUrl is reachable.');

  // Verify DB pointer still unchanged
  const dbRecAfter = await prisma.excelWorkbook.findFirst();
  if (!dbRecAfter) throw new Error('DB record missing after move');
  if (dbRecAfter.driveItemId !== realId) throw new Error('Database pointer changed after move');
  console.log('Database pointer unchanged.');

  // Final report
  console.log('\n=== MOVE VERIFICATION RESULT ===');
  console.log(JSON.stringify({
    id: postMeta.data.id,
    name: postMeta.data.name,
    parentPath: postParent,
    mimeType: postMeta.data.file?.mimeType,
    webUrl: postMeta.data.webUrl,
    worksheetCount: wsCountAfter,
    tableCount: tblCountAfter,
    dbDriveItemId: dbRecAfter.driveItemId,
  }, null, 2));
}

main().catch(err => {
  console.error('ERROR during move verification:', err.message);
  process.exit(1);
});

