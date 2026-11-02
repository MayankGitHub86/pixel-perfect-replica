// visibility-diagnostics.js
// Read‑only diagnostic for Excel workbook visibility using the project's existing Graph auth helpers.
// No mutations are performed.

const { prisma } = require('../../src/lib/db');
const { decrypt } = require('../src/lib/crypto');

// Helper: obtain a valid Graph access token (assumes token is still valid; refresh logic omitted for safety)
async function getAccessToken() {
  const conn = await prisma.mailboxConnection.findFirst({ orderBy: { createdAt: 'desc' } });
  if (!conn) throw new Error('No mailbox connection found in the database');
  return decrypt(conn.accessToken);
}

async function graphGet(path, token) {
  const url = `https://graph.microsoft.com/v1.0${path}`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const result = { status: res.status };
  if (!res.ok) {
    const txt = await res.text();
    // Try to extract a Graph error code/message if JSON
    let code = null, message = null;
    try { const err = JSON.parse(txt); code = err.error?.code; message = err.error?.message; } catch (_) {}
    result.error = { code: code || 'UNKNOWN', message: message || txt.slice(0,200) };
    return result;
  }
  const data = await res.json();
  result.data = data;
  return result;
}

async function main() {
  const token = await getAccessToken();
  const realId = '5082520FEDE864C4!s8dc280a78ce54f6b9ee4617268ede8bb';

  // 1. Documents folder children
  const docsRes = await graphGet('/me/drive/root:/Documents/children', token);

  // 2. Permissions on real workbook
  const permsRes = await graphGet(`/me/drive/items/${realId}/permissions`, token);

  // 3. Metadata for real workbook
  const realMetaRes = await graphGet(`/me/drive/items/${realId}`, token);

  // 4. Find Recovery workbook ID and its metadata
  const rootChildren = await graphGet('/me/drive/root/children', token);
  let recoveryMetaRes = { status: null, data: null, error: null };
  if (rootChildren.data) {
    const rec = rootChildren.data.value?.find(i => i.name === 'Candidate_Master_Recovery_V2.xlsx');
    if (rec) {
      recoveryMetaRes = await graphGet(`/me/drive/items/${rec.id}`, token);
    } else {
      recoveryMetaRes.error = { code: 'NOT_FOUND', message: 'Recovery workbook not present in root' };
    }
  } else {
    recoveryMetaRes = { status: rootChildren.status, error: rootChildren.error };
  }

  function trimMeta(res) {
    if (!res.data) return { status: res.status, error: res.error };
    const d = res.data;
    return {
      status: res.status,
      name: d.name,
      id: d.id,
      parentPath: d.parentReference?.path,
      mimeType: d.file?.mimeType,
      hasDownloadUrl: !!d['@microsoft.graph.downloadUrl'],
      officeFacets: Object.keys(d).filter(k => k.startsWith('excel') || k.startsWith('package')).reduce((obj,k)=>{obj[k]=d[k];return obj;})
    };
  }

  const report = {
    documentsFolder: {
      status: docsRes.status,
      error: docsRes.error,
      containsRealWorkbook: docsRes.data?.value?.some(i => i.id === realId) || false,
    },
    realWorkbook: {
      metadata: trimMeta(realMetaRes),
      permissions: permsRes.data ? permsRes.data.value.map(p => ({
        role: p.roles?.[0] || null,
        grantedTo: p.grantedTo?.user?.displayName || null,
        isSelf: p.grantedTo?.user?.id === (realMetaRes.data?.createdBy?.user?.id || null)
      })) : { status: permsRes.status, error: permsRes.error }
    },
    recoveryWorkbook: { metadata: trimMeta(recoveryMetaRes) }
  };

  console.log(JSON.stringify(report, null, 2));
}

main().catch(err => {
  console.error('Unexpected error during diagnostics:', err.message);
  process.exit(1);
});
