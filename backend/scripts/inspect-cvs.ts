import fs from 'fs';
import path from 'path';
import { parseResumeBuffer } from '../src/lib/resumeParser';
import { extractCandidateInfo } from '../src/lib/candidateExtractor';

async function main() {
  const dirMayank = path.resolve(__dirname, '../storage/attachments/4f8c8e9a-046f-482c-9d75-3bc301b72079');
  const fileMayank = fs.readdirSync(dirMayank)[0];
  const bufMayank = fs.readFileSync(path.join(dirMayank, fileMayank));
  const resMayank = await parseResumeBuffer(bufMayank, fileMayank);
  console.log('=== MAYANK CV TEXT ===');
  console.log(resMayank.text);
  console.log('=== EXTRACTED MAYANK ===');
  console.log(JSON.stringify(extractCandidateInfo(resMayank.text), null, 2));

  const dirChirayu = path.resolve(__dirname, '../storage/attachments/ca1341f7-7987-4cea-a3db-3ce2347425b8');
  const fileChirayu = fs.readdirSync(dirChirayu)[0];
  const bufChirayu = fs.readFileSync(path.join(dirChirayu, fileChirayu));
  const resChirayu = await parseResumeBuffer(bufChirayu, fileChirayu);
  console.log('=== CHIRAYU CV TEXT ===');
  console.log(resChirayu.text);
  console.log('=== EXTRACTED CHIRAYU ===');
  console.log(JSON.stringify(extractCandidateInfo(resChirayu.text), null, 2));
}

main().catch(console.error);
