// inspect-visibility.js
// Read‑only diagnostic script for Excel workbook visibility issues.
// Uses Microsoft Graph to fetch:
//   1. Children of the OneDrive "Documents" folder.
//   2. Permissions of the real workbook (Candidate_Master_RealData.xlsx).
//   3. Metadata of the real workbook and the older Recovery workbook.
// No mutation is performed.

const { prisma } = require('../src/lib/db');
const { decrypt } = require('../src/lib/crypto');

async function getAccessToken() {
  const conn = await prisma.mailboxConnection.findFirst({ orderBy: { createdAt: 'desc' } });
  if (!conn) throw new Error('No mailbox connection found');
  // Assume token is still valid; refresh logic omitted for safety.
  return decrypt(conn.accessToken);
}

async function graphRequest(path, token) {
  const url = `https://graph.microsoft.com/v1.0${path}`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`Graph request ${path} failed ${res.status}: ${txt}`);
  }
  return res.json();
}

async function main() {
  const token = await getAccessToken();

  // 1. List Documents folder children
  const documents = await graphRequest('/me/drive/root:/Documents/children', token);

  // 2. Permissions on the real workbook
  const realId = '5082520FEDE864C4!s8dc280a78ce54f6b9ee4617268ede8bb';
  const permissions = await graphRequest(`/me/drive/items/${realId}/permissions`, token);

  // 3. Fetch metadata for both workbooks (real and recovery)
  // First get root children to locate the recovery workbook ID
  const rootChildren = await graphRequest('/me/drive/root/children', token);
  const realMeta = await graphRequest(`/me/drive/items/${realId}`, token);
  const recovery = rootChildren.value?.find(i => i.name === 'Candidate_Master_Recovery_V2.xlsx');
  const recoveryMeta = recovery ? await graphRequest(`/me/drive/items/${recovery.id}`, token) : null;

  const output = {
    documentsFolderContainsReal: documents.value?.some(i => i.id === realId) || false,
    documentsFolderList: documents.value?.map(i => ({ id: i.id, name: i.name })) || [],
    realWorkbookPermissions: permissions.value?.map(p => ({ id: p.id, roles: p.roles, grantedTo: p.grantedTo?.user?.displayName })) || [],
    realWorkbookMeta: {
      id: realMeta.id,
      name: realMeta.name,
      parentReference: realMeta.parentReference,
      file: realMeta.file,
    },
    recoveryWorkbookMeta: recoveryMeta ? {
      id: recoveryMeta.id,
      name: recoveryMeta.name,
      parentReference: recoveryMeta.parentReference,
      file: recoveryMeta.file,
    } : null,
    // Redact download URL value, just indicate presence
    realHasDownloadUrl: !!realMeta['@microsoft.graph.downloadUrl'],
    recoveryHasDownloadUrl: recoveryMeta ? !!recoveryMeta['@microsoft.graph.downloadUrl'] : null,
  };

  console.log(JSON.stringify(output, null, 2));
}

main().catch(err => {
  console.error('Error during inspection:', err.message);
  process.exit(1);
});

