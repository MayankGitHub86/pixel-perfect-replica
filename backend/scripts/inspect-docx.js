const fs = require('fs');
const mammoth = require('mammoth');

async function main() {
  const filePath = "C:\\Users\\ASUS\\Documents\\lcc outlook project\\pixel-perfect-replica\\backend\\storage\\attachments\\ca1341f7-7987-4cea-a3db-3ce2347425b8\\cv-att-1_Chirayu_Sharma_Resume_TridhyaTech_BA.docx";
  const buffer = fs.readFileSync(filePath);
  const result = await mammoth.extractRawText({ buffer });
  console.log('--- RESUME RAW TEXT ---');
  console.log(result.value);
  console.log('--- END RESUME RAW TEXT ---');
}

main().catch(console.error);

