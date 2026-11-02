// move-workbook.js
// Read‑only‑safe move of Candidate_Master_RealData.xlsx from OneDrive root to the Documents folder.
// No content changes, no DB updates, no permission changes.

import { prisma } from '../src/lib/db.js';
import { decrypt } from '../src/lib/crypto.js';

async function getAccessToken() {
  const conn = await prisma.mailboxConnection.findFirst({ orderBy: { createdAt: 'desc' } });
  if (!conn) throw new Error('No mailbox connection found');
  return decrypt(conn.accessToken);
}

async function graphRequest(path, method = 'GET', token, body = null) {
  const url = `https://graph.microsoft.com/v1.0${path}`;
  const opts = {
    method,
    headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
  };
  if (body) {
    opts.headers['Content-Type'] = 'application/json';
    opts.body = JSON.stringify(body);
  }
  const res = await fetch(url, opts);
  const result = { status: res.status };
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

  // 1️⃣ Verify current metadata
  const curMeta = await graphRequest(`/me/drive/items/${realId}`);
  if (curMeta.error) throw new Error(`Current metadata fetch failed: ${curMeta.error.code} – ${curMeta.error.message}`);
  const { name, parentReference, webUrl } = curMeta.data;
  if (name !== 'Candidate_Master_RealData.xlsx' || curMeta.data.id !== realId) {
    throw new Error('Current workbook name or ID does not match expected values');
  }
  if (!parentReference?.path?.startsWith('/drive/root:') || parentReference?.path?.includes('/Documents')) {
    throw new Error('Workbook is not currently in the OneDrive root folder as expected');
  }
  console.log('Pre‑move verification passed.');

  // 2️⃣ Verify Documents folder exists (GET its metadata)
  const docFolder = await graphRequest('/me/drive/root:/Documents');
  if (docFolder.error) throw new Error(`Documents folder not accessible: ${docFolder.error.code} – ${docFolder.error.message}`);
  console.log('Documents folder exists.');

  // 3️⃣ Capture worksheet/table counts before move for later comparison
  const wsBefore = await graphRequest(`/me/drive/items/${realId}/workbook/worksheets`);
  const tblBefore = await graphRequest(`/me/drive/items/${realId}/workbook/tables`);
  const wsCountBefore = wsBefore.data?.value?.length ?? 0;
  const tblCountBefore = tblBefore.data?.value?.length ?? 0;

  // 4️⃣ Perform the move (PATCH)
  const moveRes = await graphRequest(`/me/drive/items/${realId}`, 'PATCH', token, {
    parentReference: { path: '/drive/root:/Documents' }
  });
  if (moveRes.error) throw new Error(`Move operation failed: ${moveRes.error.code} – ${moveRes.error.message}`);
  console.log('Move operation succeeded.');

  // 5️⃣ Verify metadata after move
  const postMeta = await graphRequest(`/me/drive/items/${realId}`);
  if (postMeta.error) throw new Error(`Post‑move metadata fetch failed: ${postMeta.error.code} – ${postMeta.error.message}`);
  const postParent = postMeta.data.parentReference?.path;
  if (!postParent?.includes('/Documents')) {
    throw new Error('Post‑move parent path does not indicate Documents folder');
  }
  console.log('Post‑move metadata verification passed.');

  // 6️⃣ Verify worksheet/table counts unchanged
  const wsAfter = await graphRequest(`/me/drive/items/${realId}/workbook/worksheets`);
  const tblAfter = await graphRequest(`/me/drive/items/${realId}/workbook/tables`);
  const wsCountAfter = wsAfter.data?.value?.length ?? 0;
  const tblCountAfter = tblAfter.data?.value?.length ?? 0;
  if (wsCountBefore !== wsCountAfter || tblCountBefore !== tblCountAfter) {
    throw new Error('Worksheet or table count changed after move');
  }
  console.log('Worksheet/table structure unchanged.');

  // 7️⃣ Verify webUrl still works (simple GET – we do not follow redirects, just check HTTP 200)
  const webRes = await fetch(postMeta.data.webUrl, { method: 'HEAD' });
  if (!webRes.ok) {
    throw new Error(`webUrl check failed with status ${webRes.status}`);
  }
  console.log('webUrl is reachable.');

  // 8️⃣ Verify PostgreSQL pointer still references same DriveItem ID
  const dbRec = await prisma.excelWorkbook.findFirst();
  if (!dbRec) throw new Error('No ExcelWorkbook record found in DB');
  if (dbRec.driveItemId !== realId) {
    throw new Error(`DB pointer changed (expected ${realId}, got ${dbRec.driveItemId})`);
  }
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
    dbDriveItemId: dbRec.driveItemId
  }, null, 2));
}

main().catch(err => {
  console.error('ERROR during move verification:', err.message);
  process.exit(1);
});

