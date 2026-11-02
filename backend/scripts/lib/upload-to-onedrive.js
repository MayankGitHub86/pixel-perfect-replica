// backend/scripts/lib/upload-to-onedrive.js
//
// Shared upload module: preflight existence check + conditional PUT with
// If-None-Match: * to prevent overwriting an existing OneDrive file.
//
// Used by create-real-workbook.js (production) and tested directly by
// src/tests/create-real-workbook-upload.test.ts (mocked unit tests).

'use strict';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

/**
 * Upload a file to OneDrive, refusing to overwrite any existing file.
 *
 * @param {object} opts
 * @param {string} opts.token      - Bearer access token (never logged).
 * @param {Buffer} opts.fileBuffer - Raw file bytes to upload.
 * @param {string} opts.fileName   - Target filename in the OneDrive root folder.
 * @returns {Promise<object>} Parsed Graph DriveItem metadata on success.
 * @throws {Error} If the file already exists, the preflight returns a non-404
 *   status, or the PUT fails for any reason.
 */
async function uploadToOneDrive({ token, fileBuffer, fileName }) {
  // --- Step 1: Preflight existence check ---
  const existenceUrl = `${GRAPH_BASE}/me/drive/root:/${fileName}`;
  const existenceRes = await fetch(existenceUrl, {
    method: 'GET',
    headers: { Authorization: `Bearer ${token}` },
  });

  if (existenceRes.ok) {
    throw new Error(
      `OneDrive file "${fileName}" already exists – abort to avoid overwrite`
    );
  }
  if (existenceRes.status !== 404) {
    const bodyTxt = await existenceRes.text();
    const snippet = bodyTxt.slice(0, 200).replace(/[\r\n]+/g, ' ');
    throw new Error(
      `Unexpected response checking OneDrive file existence: ${existenceRes.status} ${snippet}`
    );
  }

  // --- Step 2: Conditional PUT upload ---
  // If-None-Match: * tells Graph to reject the request with 412 if a resource
  // with this name already exists (race-condition protection).
  const uploadUrl = `${GRAPH_BASE}/me/drive/root:/${fileName}:/content`;
  const uploadRes = await fetch(uploadUrl, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type':
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Length': fileBuffer.length.toString(),
      'If-None-Match': '*',
    },
    body: fileBuffer,
  });

  if (uploadRes.status === 412) {
    const errBody = await uploadRes.text();
    const snippet = errBody.slice(0, 200).replace(/[\r\n]+/g, ' ');
    throw new Error(
      `Upload prevented – file already exists (412 Precondition Failed). ${snippet}`
    );
  }
  if (!uploadRes.ok) {
    const errBody = await uploadRes.text();
    const snippet = errBody.slice(0, 200).replace(/[\r\n]+/g, ' ');
    throw new Error(`Upload failed: ${uploadRes.status} ${snippet}`);
  }

  return uploadRes.json();
}

module.exports = { uploadToOneDrive };

