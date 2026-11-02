// backend/src/__tests__/candidateExtractor.test.ts
import { describe, it, expect } from 'vitest';
import {
  extractEmail,
  extractPhone,
  extractName,
  extractSkills,
  extractDegree,
  extractCollege,
  extractGraduationYear,
  extractExperienceYears,
  extractCandidateInfo,
} from '../lib/candidateExtractor';

describe('Candidate Extractor', () => {
  const sampleResume = `
Johnathan Doe
Email: johnathan.doe@example.com
Phone: +1 (555) 234-5678
Location: San Francisco, CA

Professional Summary
Senior Full Stack Engineer with 6 years of experience building web applications.

Skills
JavaScript, TypeScript, React, Node.js, Express, PostgreSQL, Redis, Docker, AWS, Git

Education
Bachelor of Technology in Computer Science
Stanford University (Graduated: 2018)

Experience
Senior Software Developer at TechCorp (2020 - Present)
Software Engineer at StartupInc (2018 - 2020)
`;

  it('extracts candidate email', () => {
    const email = extractEmail(sampleResume);
    expect(email).toBe('johnathan.doe@example.com');
  });

  it('extracts candidate phone', () => {
    const phone = extractPhone(sampleResume);
    expect(phone).toBe('+1 (555) 234-5678');
  });

  it('extracts candidate name avoiding header words', () => {
    const name = extractName(sampleResume);
    expect(name).toBe('Johnathan Doe');
  });

  it('extracts recognized skills correctly', () => {
    const skills = extractSkills(sampleResume);
    expect(skills).toContain('JavaScript');
    expect(skills).toContain('TypeScript');
    expect(skills).toContain('React');
    expect(skills).toContain('Node.js');
    expect(skills).toContain('PostgreSQL');
    expect(skills).toContain('Docker');
  });

  it('extracts degree and college', () => {
    const degree = extractDegree(sampleResume);
    expect(degree).toMatch(/Bachelor of Technology/i);

    const college = extractCollege(sampleResume);
    expect(college).toMatch(/Stanford/i);
  });

  it('extracts graduation year and experience years', () => {
    const year = extractGraduationYear(sampleResume);
    expect(year).toBe(2018);

    const exp = extractExperienceYears(sampleResume);
    expect(exp).toBeGreaterThanOrEqual(6);
  });

  it('returns high confidence and status=processed for complete profile', () => {
    const result = extractCandidateInfo(sampleResume);
    expect(result.name).toBe('Johnathan Doe');
    expect(result.email).toBe('johnathan.doe@example.com');
    expect(result.phone).toBe('+1 (555) 234-5678');
    expect(result.confidence).toBeGreaterThanOrEqual(0.70);
    expect(result.processingStatus).toBe('processed');
    expect(result.reviewFields).not.toContain('email');
    expect(result.reviewFields).not.toContain('name');
  });

  it('does NOT invent missing fields (sets null)', () => {
    const incompleteResume = `
Jane Smith
jane.smith@example.com
Passionate coder interested in web technologies.
`;
    const result = extractCandidateInfo(incompleteResume);
    expect(result.name).toBe('Jane Smith');
    expect(result.email).toBe('jane.smith@example.com');
    expect(result.phone).toBeNull();
    expect(result.degree).toBeNull();
    expect(result.college).toBeNull();
    expect(result.graduationYear).toBeNull();
    expect(result.experienceYears).toBeNull();
    expect(result.reviewFields).toContain('phone');
    expect(result.reviewFields).toContain('degree');
    expect(result.processingStatus).toBe('needs_review');
  });

  describe('Strict Data Extraction & Integrity Regressions', () => {
    it('strips section headers like EDUCATION from college names', () => {
      const cvSnippet = `
EDUCATION
Lovely Professional University Phagwara, India
Bachelor of Technology
Aug’ 23 – Present
`;
      const college = extractCollege(cvSnippet);
      expect(college).toBe('Lovely Professional University');
      expect(college).not.toMatch(/^EDUCATION/i);
    });

    it('extracts degree completion year rather than 10th or 12th pass year', () => {
      const cvSnippet = `
EDUCATION
Bachelor of Engineering – Computer Engineering  —  Indus University 2022 – 2026
HSC (12th) – CBSE Board (75%)  —  St. Ann's School 2022
SSC (10th) – CBSE Board (85%)  —  Delhi Public School 2020
`;
      const gradYear = extractGraduationYear(cvSnippet);
      expect(gradYear).toBe(2026);
      expect(gradYear).not.toBe(2022);
      expect(gradYear).not.toBe(2020);
    });

    it('leaves graduation year null when only ongoing degree without completion date', () => {
      const cvSnippet = `
EDUCATION
Lovely Professional University
Bachelor of Technology
Aug’ 23 – Present
`;
      const gradYear = extractGraduationYear(cvSnippet);
      expect(gradYear).toBeNull();
    });

    it('extracts explicit expected graduation year', () => {
      const cvSnippet = `
EDUCATION
Bachelor of Technology in CS
MIT (Expected Graduation: 2027)
`;
      const gradYear = extractGraduationYear(cvSnippet);
      expect(gradYear).toBe(2027);
    });

    it('merges overlapping employment intervals to avoid double-counting', () => {
      const cvSnippet = `
EXPERIENCE
Software Intern at Company A
Jan 2023 - Jun 2023
Part-Time Developer at Company B
Apr 2023 - Dec 2023
`;
      // Jan 2023 to Dec 2023 = 1.0 year (not 0.5 + 0.7 = 1.2)
      const exp = extractExperienceYears(cvSnippet, new Date('2024-01-01'));
      expect(exp).toBe(1);
    });

    it('excludes student clubs, extracurriculars, and CESA from professional experience', () => {
      const cvSnippet = `
EXPERIENCE
Graphic Head – CESA (Computer Engineering Student Association)
2025 – 2026
Student Club Coordinator
2023 - 2024
`;
      const exp = extractExperienceYears(cvSnippet, new Date('2026-06-01'));
      expect(exp).toBeNull();
    });

    it('leaves experience null when no valid employment dates are present', () => {
      const cvSnippet = `
EXPERIENCE
Contributed to open source projects and participated in hackathons.
`;
      const exp = extractExperienceYears(cvSnippet);
      expect(exp).toBeNull();
    });
  });
});

