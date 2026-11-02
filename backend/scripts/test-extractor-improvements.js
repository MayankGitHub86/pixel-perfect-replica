const fs = require('fs');
const mammoth = require('mammoth');

async function testExtraction() {
  const filePath = "C:\\Users\\ASUS\\Documents\\lcc outlook project\\pixel-perfect-replica\\backend\\storage\\attachments\\ca1341f7-7987-4cea-a3db-3ce2347425b8\\cv-att-1_Chirayu_Sharma_Resume_TridhyaTech_BA.docx";
  const buffer = fs.readFileSync(filePath);
  const { value: text } = await mammoth.extractRawText({ buffer });

  // Test phone regex
  const phoneRegex = /(?:\+91[\s.-]?)?[6-9]\d{4}[\s.-]?\d{5}\b|(?:\+?\d{1,3}[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}\b|\b(?:\+91[-\s]?)?[6-9]\d{9}\b/;
  const phoneMatch = text.match(phoneRegex);
  console.log('Extracted Phone:', phoneMatch ? phoneMatch[0] : null);

  // Test location extraction
  const locationRegex = /\b([A-Za-z0-9\s,.-]+,\s*(?:Ahmedabad|Mumbai|Pune|Bengaluru|Bangalore|Hyderabad|Delhi|Noida|Gurgaon|Gurugram|Chennai|Kolkata|Gujarat|Maharashtra|Karnataka|Delhi|Tamil Nadu|Telangana|Uttar Pradesh|Haryana|Punjab|India))\b/i;
  const locMatch = text.match(locationRegex);
  console.log('Extracted Location:', locMatch ? locMatch[0].trim() : null);

  // Test graduation year
  // In education section: 2022 – 2026
  const eduMatch = text.match(/Bachelor[^\n\r]*(\d{4})\s*[-–—]\s*(\d{4})/i) || text.match(/(\d{4})\s*[-–—]\s*(\d{4})/);
  console.log('Graduation Year:', eduMatch ? parseInt(eduMatch[2], 10) : null);

  // Test role
  const roleRegex = /\b(Business Analyst|Data Analyst|Software Engineer|Frontend Developer|Backend Developer|Full Stack Developer)\b/i;
  const roleMatch = text.match(roleRegex);
  console.log('Extracted Role:', roleMatch ? roleMatch[0] : null);
}

testExtraction();

