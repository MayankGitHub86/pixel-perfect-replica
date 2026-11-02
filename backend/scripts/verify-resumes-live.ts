import fs from 'fs';
import path from 'path';
import crypto from 'crypto';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

function sha256(buf: Buffer): string {
  return crypto.createHash('sha256').update(buf).digest('hex');
}

async function verifyResumeDownload() {
  console.log('=== RESUME DOWNLOAD VERIFICATION ===');

  const candidates = [
    {
      id: '106945ea-6ae3-4c64-b5bb-9c1148d72985',
      name: 'CHIRAYU SHARMA',
      expectedFilename: 'Chirayu_Sharma_Resume_TridhyaTech_BA.docx',
      storageDir: path.resolve(__dirname, '../storage/attachments/ca1341f7-7987-4cea-a3db-3ce2347425b8'),
    },
    {
      id: 'd5bd579a-309d-4697-aaa2-e27aadd2ce02',
      name: 'Mayank Pandey',
      expectedFilename: 'Mayank__CV.pdf',
      storageDir: path.resolve(__dirname, '../storage/attachments/4f8c8e9a-046f-482c-9d75-3bc301b72079'),
    },
  ];

  const results: any[] = [];

  for (const c of candidates) {
    const url = `http://localhost:3000/api/candidates/${c.id}/resume`;
    console.log(`\nTesting candidate: ${c.name} (${c.id})`);
    console.log(`URL: ${url}`);

    const res = await fetch(url);
    console.log(`HTTP status: ${res.status}`);
    console.log(`Content-Type: ${res.headers.get('content-type')}`);
    const contentDisp = res.headers.get('content-disposition');
    console.log(`Content-Disposition: ${contentDisp}`);

    if (res.status !== 200) {
      throw new Error(`Failed to download resume for ${c.name}: HTTP ${res.status}`);
    }

    const downloadedBuf = Buffer.from(await res.arrayBuffer());
    const downloadedHash = sha256(downloadedBuf);

    // Read stored file on disk
    const storedFile = fs.readdirSync(c.storageDir)[0];
    const storedBuf = fs.readFileSync(path.join(c.storageDir, storedFile));
    const storedHash = sha256(storedBuf);

    console.log(`Downloaded size: ${downloadedBuf.length} bytes`);
    console.log(`Stored file: ${storedFile} (${storedBuf.length} bytes)`);
    console.log(`Downloaded SHA-256: ${downloadedHash}`);
    console.log(`Stored SHA-256:     ${storedHash}`);
    const match = downloadedHash === storedHash;
    console.log(`SHA-256 Match: ${match ? 'PASS' : 'FAIL'}`);

    results.push({
      candidateId: c.id,
      name: c.name,
      status: res.status,
      filename: c.expectedFilename,
      downloadedHash,
      storedHash,
      match,
    });
  }

  // Cross-contamination check: Ensure Candidate A hash !== Candidate B hash
  const crossMatch = results[0].downloadedHash === results[1].downloadedHash;
  console.log(`\nCross-contamination check (Candidate A hash !== Candidate B hash): ${!crossMatch ? 'PASS (Distinct)' : 'FAIL'}`);

  // Test non-existent candidate ID
  const invalidRes = await fetch('http://localhost:3000/api/candidates/00000000-0000-0000-0000-000000000000/resume');
  console.log(`Non-existent candidate download returns HTTP ${invalidRes.status}: ${invalidRes.status === 404 ? 'PASS' : 'FAIL'}`);

  return { results, crossMatch: !crossMatch, nonExistent404: invalidRes.status === 404 };
}

verifyResumeDownload().catch(err => {
  console.error(err);
  process.exit(1);
});

