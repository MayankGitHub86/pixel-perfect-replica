// backend/src/lib/candidateExtractor.ts
// Accurate rule-based and heuristic extraction of candidate profiles from resume text and email metadata.
// Strictly adheres to: "Never invent missing details. Use null for unavailable fields."

export interface ExtractedCandidate {
  name: string | null;
  email: string | null;
  phone: string | null;
  location: string | null;
  skills: string[];
  degree: string | null;
  college: string | null;
  graduationYear: number | null;
  experienceYears: number | null;
  projects: string[];
  certifications: string[];
  appliedRole: string | null;
  confidence: number;
  reviewFields: string[];
  processingStatus: 'processed' | 'needs_review';
}

const COMMON_SKILLS = [
  'JavaScript', 'TypeScript', 'React', 'Next.js', 'Vue', 'Angular', 'Node.js', 'Express',
  'Python', 'Django', 'Flask', 'FastAPI', 'Java', 'Spring Boot', 'C++', 'C#', '.NET',
  'Go', 'Golang', 'Rust', 'PHP', 'Ruby', 'Rails', 'Kotlin', 'Swift',
  'SQL', 'PostgreSQL', 'MySQL', 'MongoDB', 'Redis', 'Elasticsearch', 'DynamoDB',
  'AWS', 'Azure', 'GCP', 'Docker', 'Kubernetes', 'Terraform', 'CI/CD', 'Git', 'GitHub',
  'REST API', 'GraphQL', 'Microservices', 'Kafka', 'RabbitMQ', 'Linux',
  'Machine Learning', 'Data Science', 'Pandas', 'NumPy', 'TensorFlow', 'PyTorch',
  'Tailwind CSS', 'HTML', 'CSS', 'Jest', 'Vitest', 'Cypress', 'Playwright',
  'Power BI', 'DAX', 'Power Query', 'Advanced Excel', 'Excel', 'Requirement Analysis',
  'Tableau', 'MIS Reporting', 'Data Analytics', 'Business Analysis',
];

const DEGREE_PATTERNS = [
  /\b(b\.?tech(?:nology)?|b\.?e\.?|bachelor of (?:technology|engineering|science|arts|business))\b/i,
  /\b(m\.?tech(?:nology)?|m\.?e\.?|master of (?:technology|engineering|science|arts|computer applications))\b/i,
  /\b(b\.?c\.?a\.?|m\.?c\.?a\.?|m\.?b\.?a\.?|b\.?sc|m\.?sc|b\.?com|m\.?com|ph\.?d\.?)\b/i,
  /\b(bachelor'?s degree|master'?s degree|diploma)\b/i,
];

const COLLEGE_PATTERNS = [
  /(?:university of [a-zA-Z\s]+|[a-zA-Z\s]+ (?:university|institute of technology|college of engineering|polytechnic))/i,
  /\b(?:IIT|NIT|BITS|IIIT|MIT|Stanford|Harvard|Oxford|Cambridge)\b[a-zA-Z\s]*/i,
];

const ROLE_PATTERNS = [
  /\b(business\s*analyst|data\s*analyst|requirement\s*analyst|data\s*engineer|systems\s*analyst)\b/i,
  /\b(senior\s+|lead\s+|junior\s+)?(full\s*stack|frontend|backend|software|devops|cloud|data|qa|machine\s*learning|ai)\s*(engineer|developer|architect|specialist)\b/i,
  /\b(product\s*manager|project\s*manager|scrum\s*master|ui\/ux\s*designer)\b/i,
];

/**
 * Extract email address from text.
 */
export function extractEmail(text: string): string | null {
  const match = text.match(/\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,7}\b/);
  return match ? match[0].toLowerCase().trim() : null;
}

/**
 * Extract phone number from text (supports international, US, and Indian 10-digit formats with or without spaces/dashes).
 */
export function extractPhone(text: string): string | null {
  const match = text.match(/(?:\+91[\s.-]?)?[6-9]\d{4}[\s.-]?\d{5}\b|(?:\+?\d{1,3}[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}\b|\b(?:\+91[-\s]?)?[6-9]\d{9}\b/);
  return match ? match[0].trim() : null;
}

/**
 * Extract candidate location / address if genuinely present in the text.
 */
export function extractLocation(text: string): string | null {
  // Look for location pattern in contact lines or address lines
  // Exclude technical tools or skills lists (e.g. Git, VS Code, Python, etc.)
  const locMatch = text.match(/\b([A-Za-z0-9\s,.-]+,\s*(?:Ahmedabad|Mumbai|Pune|Bengaluru|Bangalore|Hyderabad|Delhi|Noida|Gurgaon|Gurugram|Chennai|Kolkata|Gujarat|Maharashtra|Karnataka|Tamil Nadu|Telangana|Uttar Pradesh|Haryana|Punjab|India))\b/i);
  if (locMatch) {
    const clean = locMatch[0].trim().replace(/[\r\n]+/g, ' ');
    // Exclude if it looks like technical tools or skills list
    if (clean.length <= 100 && !/\b(?:git|vs|code|tools?|languages?|frameworks?|platforms?|libraries|skills)\b/i.test(clean)) {
      return clean;
    }
  }
  return null;
}

/**
 * Extract full candidate name.
 */
export function extractName(text: string, emailFrom?: string): string | null {
  // Check first 5 non-empty lines of resume
  const lines = text.split('\n').map((l) => l.trim()).filter((l) => l.length > 0);
  for (const line of lines.slice(0, 5)) {
    // Exclude header words
    if (/^(curriculum vitae|resume|profile|contact|summary|objective|page \d)/i.test(line)) {
      continue;
    }
    // Clean name should be 2 to 4 capitalized words, alphabetic only
    if (/^[A-Z][a-zA-Z'.]+(?:\s+[A-Z][a-zA-Z'.]+){1,3}$/.test(line)) {
      return line;
    }
  }

  // Fallback: check email sender display name if provided (e.g. "John Doe <john@doe.com>")
  if (emailFrom) {
    const fromClean = emailFrom.replace(/<.*>/, '').replace(/["']/g, '').trim();
    if (fromClean.length > 2 && /^[A-Za-z\s.'-]+$/.test(fromClean) && !fromClean.includes('@')) {
      return fromClean;
    }
  }

  return null;
}

/**
 * Extract recognized technical skills.
 */
export function extractSkills(text: string): string[] {
  const found = new Set<string>();
  const normalized = ` ${text.replace(/[\n\r,;()]/g, ' ')} `;

  for (const skill of COMMON_SKILLS) {
    const regex = new RegExp(`\\b${skill.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`, 'i');
    if (regex.test(normalized)) {
      found.add(skill);
    }
  }

  return Array.from(found);
}

/**
 * Extract degree qualification.
 */
export function extractDegree(text: string): string | null {
  for (const pattern of DEGREE_PATTERNS) {
    const match = text.match(pattern);
    if (match) {
      return match[0].trim();
    }
  }
  return null;
}

/**
 * Extract college or university.
 */
export function extractCollege(text: string): string | null {
  for (const pattern of COLLEGE_PATTERNS) {
    const match = text.match(pattern);
    if (match) {
      let clean = match[0].trim().replace(/[\r\n]+/g, ' ');
      // Strip leading section header words like "EDUCATION", "ACADEMIC", etc.
      clean = clean.replace(/^(?:education|academics?|qualifications?|background)\s+/i, '').trim();
      return clean;
    }
  }
  return null;
}

/**
 * Extract graduation year (handles ranges like "2022 – 2026", "Expected 2027", "Graduated in 2025" returning completion year).
 * Leaves field null when unsupported or ambiguous rather than guessing.
 */
export function extractGraduationYear(text: string): number | null {
  // 1. Explicit completion patterns anywhere in text
  const explicitPatterns = [
    /\b(?:expected(?:\s+graduation|\s+completion)?(?:\s+in)?|graduation(?:\s+year)?|batch|class\s+of|passing\s+year|year\s+of\s+passing|completion\s+year)\s*[:\-–—]?\s*(?:year\s*)?(\b(?:19|20)\d{2}\b)/i,
    /\bgraduated\s+in\s+(\b(?:19|20)\d{2}\b)/i,
    /\bgraduat(?:ed|ion)\s*[:\-–—]?\s*(\b(?:19|20)\d{2}\b)/i,
  ];

  for (const pat of explicitPatterns) {
    const m = text.match(pat);
    if (m) {
      const y = parseInt(m[1], 10);
      if (!isNaN(y) && y >= 1980 && y <= 2035) return y;
    }
  }

  // 2. Extract Education Section bounded by section header on its own line
  const eduSectionMatch = text.match(/(?:^|\r?\n)\s*(?:EDUCATION|ACADEMIC\s+BACKGROUND|ACADEMIC\s+QUALIFICATIONS?|ACADEMICS?)\b[:\s]*([\s\S]*?)(?=(?:\r?\n\s*(?:EXPERIENCE|WORK\s+EXPERIENCE|EMPLOYMENT|PROJECTS?|TECHNICAL\s+SKILLS|SKILLS|CERTIFICATIONS?|$)))/i);
  const eduText = eduSectionMatch ? eduSectionMatch[1] : '';

  const scanText = eduText || text;

  // Check explicit phrases in education section
  const eduExplicit = scanText.match(/\b(?:expected|graduat(?:ed|ion)?|completion)\s*(?:in\s*)?(\b(?:19|20)\d{2}\b)/i);
  if (eduExplicit) {
    const y = parseInt(eduExplicit[1], 10);
    if (!isNaN(y) && y >= 1980 && y <= 2035) return y;
  }

  // Look for degree line with range: e.g. "Bachelor... 2022 - 2026"
  const degreeKeywords = '(?:bachelor|master|b\\.?e\\.?|b\\.?tech|m\\.?tech|b\\.?s\\.?|m\\.?s\\.?|ph\\.?d|degree|engineering|university|college|institute)';
  const lines = scanText.split(/\r?\n/).map(l => l.trim()).filter(Boolean);

  // First priority: degree line with range (excluding high-school HSC/SSC/10th/12th)
  for (const line of lines) {
    if (new RegExp(degreeKeywords, 'i').test(line) && !/\b(?:hsc|ssc|10th|12th|secondary)\b/i.test(line)) {
      const range = line.match(/(\b(?:19|20)\d{2}\b)\s*[-–—to]+\s*(\b(?:19|20)\d{2}\b)/i);
      if (range) {
        const endYear = parseInt(range[2], 10);
        if (!isNaN(endYear) && endYear >= 1980 && endYear <= 2035) return endYear;
      }
      const singleYear = line.match(/\b(?:19|20)\d{2}\b/);
      if (singleYear && !line.includes('-') && !line.includes('–')) {
        const y = parseInt(singleYear[0], 10);
        if (!isNaN(y) && y >= 1980 && y <= 2035) return y;
      }
    }
  }

  // Second priority: any range in education section (excluding obvious 10th/12th lines)
  for (const line of lines) {
    if (!/\b(?:hsc|ssc|10th|12th|secondary)\b/i.test(line)) {
      const range = line.match(/(\b(?:19|20)\d{2}\b)\s*[-–—to]+\s*(\b(?:19|20)\d{2}\b)/i);
      if (range) {
        const endYear = parseInt(range[2], 10);
        if (!isNaN(endYear) && endYear >= 1980 && endYear <= 2035) return endYear;
      }
    }
  }

  return null;
}

/**
 * Extract project titles if explicitly mentioned.
 */
export function extractProjects(text: string): string[] {
  const projects: string[] = [];
  const projSection = text.match(/PROJECTS?\s*[:\n\r]([\s\S]*?)(?:SKILLS|CERTIFICATIONS|EXPERIENCE|EDUCATION|$)/i);
  if (projSection && projSection[1]) {
    const lines = projSection[1].split('\n').map(l => l.trim()).filter(l => l.length > 0);
    for (const line of lines) {
      const titleMatch = line.match(/^([A-Za-z0-9\s]+?)\s*[-–—]/);
      if (titleMatch && titleMatch[1].length < 50) {
        projects.push(titleMatch[1].trim());
      }
    }
  }
  return projects;
}

/**
 * Extract certification titles if explicitly mentioned.
 */
export function extractCertifications(text: string): string[] {
  const certs: string[] = [];
  const certSection = text.match(/CERTIFICATIONS?\s*[:\n\r]([\s\S]*?)(?:PROJECTS|SKILLS|EXPERIENCE|EDUCATION|$)/i);
  if (certSection && certSection[1]) {
    const lines = certSection[1].split('\n').map(l => l.trim()).filter(l => l.length > 0);
    for (const line of lines) {
      if (line.length > 3 && line.length < 100) {
        certs.push(line);
      }
    }
  }
  return certs;
}

interface DateInterval {
  start: Date;
  end: Date;
}

const MONTH_MAP: Record<string, number> = {
  jan: 0, january: 0,
  feb: 1, february: 1,
  mar: 2, march: 2,
  apr: 3, april: 3,
  may: 4,
  jun: 5, june: 5,
  jul: 6, july: 6,
  aug: 7, august: 7,
  sep: 8, sept: 8, september: 8,
  oct: 9, october: 9,
  nov: 10, november: 10,
  dec: 11, december: 11,
};

function parseDateBoundary(dateStr: string, isEnd: boolean, now: Date): Date | null {
  const str = dateStr.trim().toLowerCase();
  if (['present', 'currently working', 'current', 'till date', 'now', 'ongoing'].includes(str)) {
    return now;
  }

  // Month + Year: "Apr 2026" or "April 2026"
  const myMatch = str.match(/^([a-z]+)[.,\s]+(\d{4})$/);
  if (myMatch && MONTH_MAP[myMatch[1]] !== undefined) {
    const m = MONTH_MAP[myMatch[1]];
    const y = parseInt(myMatch[2], 10);
    if (y < 1980 || y > now.getFullYear() + 1) return null;
    if (isEnd) {
      return new Date(Date.UTC(y, m + 1, 0, 23, 59, 59));
    } else {
      return new Date(Date.UTC(y, m, 1, 0, 0, 0));
    }
  }

  // MM/YYYY or MM-YYYY
  const numMyMatch = str.match(/^(\d{1,2})[\/\-](\d{4})$/);
  if (numMyMatch) {
    const m = parseInt(numMyMatch[1], 10) - 1;
    const y = parseInt(numMyMatch[2], 10);
    if (y < 1980 || y > now.getFullYear() + 1) return null;
    if (m >= 0 && m <= 11) {
      if (isEnd) {
        return new Date(Date.UTC(y, m + 1, 0, 23, 59, 59));
      } else {
        return new Date(Date.UTC(y, m, 1, 0, 0, 0));
      }
    }
  }

  // Year only: "2022"
  const yMatch = str.match(/^(\d{4})$/);
  if (yMatch) {
    const y = parseInt(yMatch[1], 10);
    if (y < 1980 || y > now.getFullYear() + 1) return null;
    if (isEnd) {
      return new Date(Date.UTC(y, 11, 31, 23, 59, 59));
    } else {
      return new Date(Date.UTC(y, 0, 1, 0, 0, 0));
    }
  }

  return null;
}

export function mergeIntervals(intervals: DateInterval[]): DateInterval[] {
  if (intervals.length <= 1) return intervals;
  const sorted = [...intervals].sort((a, b) => a.start.getTime() - b.start.getTime());
  const merged: DateInterval[] = [sorted[0]];

  for (let i = 1; i < sorted.length; i++) {
    const current = sorted[i];
    const prev = merged[merged.length - 1];

    if (current.start.getTime() <= prev.end.getTime()) {
      if (current.end.getTime() > prev.end.getTime()) {
        prev.end = current.end;
      }
    } else {
      merged.push(current);
    }
  }

  return merged;
}

/**
 * Extract total professional experience years from actual CV work history.
 * - Extracts distinct employment periods (internships, full-time, part-time).
 * - Merges overlapping calendar periods to prevent double counting.
 * - Excludes college education duration, academic projects, and student clubs (e.g. CESA).
 * - Excludes future or invalid date ranges.
 * - Divides total distinct days by 365.25 and rounds to 1 decimal place.
 */
export function extractExperienceYears(text: string, referenceDate: Date = new Date()): number | null {
  // 1. Find Experience / Work History Section
  const expSectionMatch = text.match(/(?:^|\r?\n)\s*(?:EXPERIENCE|WORK\s+EXPERIENCE|EMPLOYMENT\s+HISTORY|EMPLOYMENT|PROFESSIONAL\s+EXPERIENCE)\b[:\s]*([\s\S]*?)(?=(?:\r?\n\s*(?:EDUCATION|ACADEMIC|PROJECTS?|TECHNICAL\s+SKILLS|SKILLS|CERTIFICATIONS?|$)))/i);
  const expText = expSectionMatch ? expSectionMatch[1] : text;

  const MONTH_NAMES = 'jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?';
  const DATE_PART = `(?:(?:${MONTH_NAMES})[.,\\s]+\\d{4}|\\d{1,2}[\\/\\-]\\d{4}|\\b\\d{4}\\b)`;
  const DATE_RANGE_REGEX = new RegExp(`(${DATE_PART})\\s*[-–—to\\s]+\\s*(present|currently working|current|till date|now|${DATE_PART})`, 'i');

  const lines = expText.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  const intervals: DateInterval[] = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const match = line.match(DATE_RANGE_REGEX);
    if (match) {
      // Exclude student clubs, college associations, extracurricular leadership, volunteer work
      // If the current line is primarily just the date range (e.g. "2025 – 2026"), check previous line for title
      const isDateOnlyLine = line.replace(DATE_RANGE_REGEX, '').replace(/[()\-–—\s]/g, '').length < 3;
      const prevLine = i > 0 ? lines[i - 1] : '';
      const contextToCheck = isDateOnlyLine ? `${prevLine} ${line}` : line;
      if (/\b(?:student associations?|student clubs?|college clubs?|university clubs?|extracurricular|volunteer|student chapters?|academic projects?|acm chapters?|ieee student|developer student clubs?|cesa)\b/i.test(contextToCheck)) {
        continue;
      }

      const rawStart = match[1].trim();
      const rawEnd = match[2].trim();
      let start = parseDateBoundary(rawStart, false, referenceDate);
      let end = parseDateBoundary(rawEnd, true, referenceDate);

      // Conservative rule for year-only date ranges (e.g. "2022 – 2024"):
      // When month/day precision is missing on both ends, do not expand to Jan 1, Y1 -> Dec 31, Y2 (3 full years).
      // Instead, use the conservative year difference (Y2 - Y1) by anchoring both to Jan 1.
      const yStartMatch = rawStart.match(/^(\d{4})$/);
      const yEndMatch = rawEnd.match(/^(\d{4})$/);
      if (yStartMatch && yEndMatch) {
        const yStart = parseInt(yStartMatch[1], 10);
        const yEnd = parseInt(yEndMatch[1], 10);
        if (yEnd > yStart) {
          start = new Date(Date.UTC(yStart, 0, 1, 0, 0, 0));
          end = new Date(Date.UTC(yEnd, 0, 1, 0, 0, 0));
        }
      }

      // Validate date order and not completely in future
      if (start && end && end >= start && start.getTime() <= referenceDate.getTime()) {
        // Cap end date at referenceDate if it extends into future
        const cappedEnd = end.getTime() > referenceDate.getTime() ? referenceDate : end;
        intervals.push({ start, end: cappedEnd });
      }
    }
  }

  if (intervals.length > 0) {
    const merged = mergeIntervals(intervals);
    let totalDays = 0;
    for (const m of merged) {
      const diffDays = (m.end.getTime() - m.start.getTime()) / 86400000;
      if (diffDays > 0) totalDays += diffDays;
    }
    const years = totalDays / 365.25;
    // Round to one decimal place, e.g. 1.8 (or 0.5 for fractional experience)
    const rounded = Math.round(years * 10) / 10;
    return rounded;
  }

  // Fallback: explicit summary like "3 years of experience"
  const summaryMatch = text.match(/(\d+(?:\.\d+)?)\s*\+?\s*(?:years|yrs)(?:\s+of)?\s+experience/i);
  if (summaryMatch) {
    const num = parseFloat(summaryMatch[1]);
    if (!isNaN(num) && num >= 0 && num <= 40) {
      return num;
    }
  }

  return null;
}

/**
 * Extract applied role from subject or text.
 */
export function extractAppliedRole(text: string, subject?: string): string | null {
  const toScan = subject ? `${subject} ${text}` : text;
  for (const pattern of ROLE_PATTERNS) {
    const match = toScan.match(pattern);
    if (match) {
      return match[0].trim();
    }
  }
  return null;
}

/**
 * Main candidate extraction pipeline.
 */
export function extractCandidateInfo(
  resumeText: string,
  emailMetadata?: { subject?: string; from?: string; body?: string }
): ExtractedCandidate {
  const combinedText = `${emailMetadata?.subject || ''}\n${emailMetadata?.body || ''}\n${resumeText}`;

  const email = extractEmail(combinedText) || (emailMetadata?.from ? extractEmail(emailMetadata.from) : null);
  const phone = extractPhone(combinedText);
  const name = extractName(resumeText, emailMetadata?.from);
  const location = extractLocation(combinedText);
  const skills = extractSkills(combinedText);
  const degree = extractDegree(combinedText);
  const college = extractCollege(combinedText);
  const graduationYear = extractGraduationYear(combinedText);
  const experienceYears = extractExperienceYears(combinedText);
  const appliedRole = extractAppliedRole(resumeText, emailMetadata?.subject);
  const projects = extractProjects(resumeText);
  const certifications = extractCertifications(resumeText);

  // Confidence computation
  let score = 0;
  const reviewFields: string[] = [];

  if (name) score += 0.25; else reviewFields.push('name');
  if (email) score += 0.25; else reviewFields.push('email');
  if (phone) score += 0.15; else reviewFields.push('phone');
  if (skills.length > 0) score += 0.15; else reviewFields.push('skills');
  if (degree) score += 0.10; else reviewFields.push('degree');
  if (college || graduationYear) score += 0.10;

  const confidence = Math.min(1.0, Math.round(score * 100) / 100);

  // If confidence is lower than 0.70 or critical identifiers missing, flag for review
  const processingStatus: 'processed' | 'needs_review' =
    confidence >= 0.70 && !reviewFields.includes('email') && !reviewFields.includes('name')
      ? 'processed'
      : 'needs_review';

  return {
    name,
    email,
    phone,
    location,
    skills,
    degree,
    college,
    graduationYear,
    experienceYears,
    projects,
    certifications,
    appliedRole,
    confidence,
    reviewFields,
    processingStatus,
  };
}

