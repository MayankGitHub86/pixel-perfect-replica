// inspect-workbook.js
// Read‑only script: fetches metadata for a OneDrive DriveItem and prints selected fields.
// No mutation of files, DB, or tokens is performed. The access token is used only for the request
// and never printed.

const path = require('path');
// Using global fetch (Node 18+).

// Import helper to obtain a valid token (refreshes if needed)
const { getAccessToken } = require('../src/lib/excel'); // relative to repo root

async function main() {
  const driveItemId = process.argv[2];
  if (!driveItemId) {
    console.error('Usage: node inspect-workbook.js <driveItemId>');
    process.exit(1);
  }
  const token = await getAccessToken();
  const url = `https://graph.microsoft.com/v1.0/me/drive/items/${driveItemId}`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) {
    console.error('Graph request failed:', res.status, await res.text());
    process.exit(1);
  }
  const data = await res.json();
  // Build a trimmed output object with required fields, redacting the download URL value.
  const output = {
    id: data.id,
    name: data.name,
    webUrl: data.webUrl,
    parentReference: data.parentReference,
    file: data.file,
    hasDownloadUrl: !!data['@microsoft.graph.downloadUrl'],
    size: data.size,
    lastModifiedDateTime: data.lastModifiedDateTime,
    // Include drive and user info without exposing tokens
    driveId: data.parentReference?.driveId || null,
    owner: data.createdBy?.user?.displayName || null,
    ownerId: data.createdBy?.user?.id || null,
  };
  console.log(JSON.stringify(output, null, 2));
}

main().catch(err => {
  console.error('Unexpected error:', err);
  process.exit(1);
});
