// SYNTHETIC SAMPLE DATA — every record below is fabricated for visual testing.
// No real people, CVs, or mailbox data. All emails use the reserved example.com domain.
import type {
  Application,
  ApplicationStatus,
  AuditLog,
  Candidate,
  EmailRecord,
  ExcelSyncStatus,
  InterviewStatus,
  MicrosoftStatus,
  ProcessingJob,
  ProcessingStatus,
  ReviewItem,
  JobStatus,
  JobType,
} from "./types";

let seed = 42;
const rand = () => {
  seed = (seed * 16807) % 2147483647;
  return (seed - 1) / 2147483646;
};
const pick = <T,>(arr: readonly T[]) => arr[Math.floor(rand() * arr.length)];

const FIRST = ["Asha", "Rohan", "Meera", "Kabir", "Ishita", "Dev", "Nina", "Arjun", "Tara", "Vikram", "Leela", "Sam", "Priya", "Neel", "Zoya", "Omar"];
const LAST = ["Sample", "Testwell", "Demo", "Placeholder", "Mockford", "Synthetic"];
const COLLEGES = ["Example Institute of Technology", "Sample State University", "Demo College of Engineering", "Placeholder University"];
const DEGREES = ["B.Tech CSE", "B.E. ECE", "MCA", "M.Tech AI", "B.Sc Statistics", "MBA"];
const ROLES = ["Frontend Engineer", "Backend Engineer", "Data Analyst", "ML Engineer", "Product Designer", "QA Engineer"];
const SKILLS = ["React", "TypeScript", "Node.js", "Python", "SQL", "PostgreSQL", "Figma", "Docker", "AWS", "Pandas", "Java", "Go"];
const RECRUITERS = ["R. Sharma (sample)", "J. Doe (sample)", null];

const DAY = 86_400_000;
const NOW = new Date("2026-10-29T09:00:00Z").getTime();
const iso = (ms: number) => new Date(ms).toISOString();

const procStatuses: ProcessingStatus[] = ["processed", "processed", "processed", "processed", "needs_review", "failed", "processing", "new"];
const interviewStatuses: InterviewStatus[] = ["not_scheduled", "not_scheduled", "scheduled", "completed", "no_show"];
const appStatuses: ApplicationStatus[] = ["new", "under_review", "shortlisted", "interview", "rejected", "hired"];

export function buildSampleData() {
  seed = 42;
  const candidates: Candidate[] = [];
  const applications: Application[] = [];
  const emails: EmailRecord[] = [];
  const jobs: ProcessingJob[] = [];
  const audit: AuditLog[] = [];

  for (let i = 0; i < 48; i++) {
    const first = pick(FIRST);
    const last = pick(LAST);
    const received = NOW - Math.floor(rand() * 28 * DAY) - Math.floor(rand() * DAY);
    const status = pick(procStatuses);
    const id = `CAND-${String(1001 + i)}`;
    const emailId = `MSG-${String(5001 + i)}`;
    const email = `${first}.${last}${i}`.toLowerCase() + "@example.com";
    const review = status === "needs_review" ? (rand() > 0.5 ? ["graduationYear", "phone"] : ["college"]) : [];
    const role = pick(ROLES);
    const skills = Array.from(new Set([pick(SKILLS), pick(SKILLS), pick(SKILLS), pick(SKILLS)]));
    const failed = status === "failed";
    candidates.push({
      id,
      name: failed ? null : `${first} ${last}`,
      email,
      phone: review.includes("phone") || failed ? null : `+1-555-01${String(10 + i).slice(-2)}`,
      college: review.includes("college") || failed ? null : pick(COLLEGES),
      degree: failed ? null : pick(DEGREES),
      graduationYear: review.includes("graduationYear") || failed ? null : 2018 + Math.floor(rand() * 8),
      skills: failed ? [] : skills,
      experienceYears: failed ? null : Math.floor(rand() * 7),
      appliedRole: role,
      receivedAt: iso(received),
      processingStatus: status,
      interviewStatus: pick(interviewStatuses),
      updatedAt: iso(received + Math.floor(rand() * 2 * DAY)),
      resumeRef: `res_${i.toString(36)}${(i * 7919).toString(36)}`,
      resumeFileName: `${first}_${last}_CV.${rand() > 0.3 ? "pdf" : "docx"}`,
      sourceEmailId: emailId,
      reviewFields: review,
      isSample: true,
    });
    applications.push({
      id: `APP-${2001 + i}`,
      candidateId: id,
      candidateName: failed ? null : `${first} ${last}`,
      role,
      sourceEmail: email,
      receivedAt: iso(received),
      status: failed ? "new" : pick(appStatuses),
      recruiter: pick(RECRUITERS),
      updatedAt: iso(received + DAY),
      isSample: true,
    });
    emails.push({
      id: emailId,
      from: email,
      subject: `Application for ${role}${rand() > 0.6 ? " – CV attached" : ""}`,
      receivedAt: iso(received),
      attachments: [
        { id: `ATT-${i}a`, fileName: `${first}_${last}_CV.pdf`, mimeType: "application/pdf", sizeBytes: 120_000 + Math.floor(rand() * 900_000) },
        ...(rand() > 0.75 ? [{ id: `ATT-${i}b`, fileName: "cover_letter.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", sizeBytes: 40_000 }] : []),
      ],
      status,
      candidateId: id,
      applicationId: `APP-${2001 + i}`,
      error: failed ? pick(["Attachment is password-protected", "OCR produced no readable text", "Unsupported file type: .pages"]) : null,
      history: [
        { at: iso(received + 20_000), step: "Email received", status: "succeeded" },
        { at: iso(received + 60_000), step: "Attachment parsed", status: failed ? "failed" : "succeeded" },
        ...(failed ? [] : [{ at: iso(received + 90_000), step: "Fields extracted", status: (status === "needs_review" ? "succeeded" : "succeeded") as JobStatus }]),
      ],
      isSample: true,
    });
  }

  const jobTypes: JobType[] = ["email_ingest", "attachment_parse", "ocr", "field_extraction", "excel_sync"];
  const jobStatuses: JobStatus[] = ["succeeded", "succeeded", "succeeded", "failed", "retrying", "queued", "processing", "cancelled"];
  for (let i = 0; i < 36; i++) {
    const st = pick(jobStatuses);
    const started = NOW - Math.floor(rand() * 10 * DAY);
    jobs.push({
      id: `JOB-${9001 + i}`,
      type: pick(jobTypes),
      related: `MSG-${5001 + Math.floor(rand() * 48)}`,
      status: st,
      attempts: st === "failed" ? 3 : st === "retrying" ? 2 : 1,
      startedAt: st === "queued" ? null : iso(started),
      completedAt: ["succeeded", "failed", "cancelled"].includes(st) ? iso(started + 45_000) : null,
      lastError: st === "failed" || st === "retrying" ? pick(["Timeout reading attachment", "Graph API 429: throttled", "Workbook locked by another session"]) : null,
      nextRetryAt: st === "retrying" ? iso(NOW + 10 * 60_000) : null,
      isSample: true,
    });
  }

  const actions = ["candidate.created", "candidate.updated", "email.ingested", "job.retried", "excel.sync", "application.status_changed"];
  for (let i = 0; i < 60; i++) {
    const a = pick(actions);
    audit.push({
      id: `LOG-${i}`,
      at: iso(NOW - Math.floor(rand() * 14 * DAY)),
      actor: a.startsWith("email") || a.startsWith("excel") ? "system" : "recruiter@example.com",
      action: a,
      related: a.startsWith("email") ? `MSG-${5001 + (i % 48)}` : a.startsWith("job") ? `JOB-${9001 + (i % 36)}` : `CAND-${1001 + (i % 48)}`,
      result: rand() > 0.85 ? "failure" : rand() > 0.8 ? "info" : "success",
      requestId: `req_${(i * 104729).toString(16)}`,
      isSample: true,
    });
  }
  audit.sort((a, b) => b.at.localeCompare(a.at));

  const review: ReviewItem[] = candidates
    .filter((c) => c.processingStatus === "needs_review")
    .map((c) => ({
      id: `REV-${c.id}`,
      candidateId: c.id,
      candidateName: c.name,
      attachment: c.resumeFileName ?? "unknown",
      isSample: true,
      fields: c.reviewFields.map((f) => ({
        field: f as ReviewItem["fields"][number]["field"],
        extracted: f === "graduationYear" ? "20l9" : null,
        confidence: f === "graduationYear" ? 0.41 : null,
        reason: f === "graduationYear" ? "OCR ambiguity between '1' and 'l'" : "Field not found in document",
        evidence: f === "graduationYear" ? "…Bachelor of Technology, class of 20l9, GPA…" : null,
      })),
    }));

  const excel: ExcelSyncStatus = {
    workbookName: "Candidate_Master.xlsx",
    state: "not_connected",
    storageProvider: null,
    lastSuccessfulSync: null,
    pendingRows: candidates.filter((c) => c.processingStatus === "processed").length,
    failedJobs: jobs.filter((j) => j.type === "excel_sync" && j.status === "failed").length,
    lastReconciliation: null,
    workbookUrl: null,
    history: [],
  };

  const microsoft: MicrosoftStatus = {
    state: "not_connected",
    accountEmail: null,
    mailboxPermission: "unknown",
    lastMailboxCheck: null,
    subscriptionStatus: "none",
    subscriptionExpiresAt: null,
    lastReconciliation: null,
  };

  return { candidates, applications, emails, jobs, audit, review, excel, microsoft };
}
