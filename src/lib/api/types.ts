// Shared request/response types for the proposed RecruitFlow backend contract.
// These describe the intended interface; the backend is developed separately.

export type ISODate = string;

export type ProcessingStatus = "new" | "processing" | "processed" | "needs_review" | "failed";
export type InterviewStatus = "not_scheduled" | "scheduled" | "completed" | "no_show";
export type ApplicationStatus =
  | "new"
  | "under_review"
  | "shortlisted"
  | "interview"
  | "rejected"
  | "hired";
export type JobStatus = "queued" | "processing" | "succeeded" | "retrying" | "failed" | "cancelled";
export type JobType = "email_ingest" | "attachment_parse" | "ocr" | "field_extraction" | "excel_sync";
export type ExcelConnectionState =
  | "not_connected"
  | "connecting"
  | "connected"
  | "syncing"
  | "attention_required"
  | "error";
export type MicrosoftConnectionState = "not_connected" | "connected" | "expired" | "error";

export interface Paginated<T> {
  items: T[];
  total: number;
  page: number;
  pageSize: number;
}

export interface ListQuery {
  page?: number;
  pageSize?: number;
  search?: string;
  sortBy?: string;
  sortDir?: "asc" | "desc";
  [filter: string]: string | number | undefined;
}

export interface Candidate {
  id: string;
  name: string | null;
  email: string | null;
  phone: string | null;
  college: string | null;
  degree: string | null;
  graduationYear: number | null;
  skills: string[];
  experienceYears: number | null;
  appliedRole: string | null;
  receivedAt: ISODate;
  processingStatus: ProcessingStatus;
  interviewStatus: InterviewStatus;
  updatedAt: ISODate;
  /** Opaque backend reference. Never a public storage URL. */
  resumeRef: string | null;
  resumeFileName: string | null;
  sourceEmailId: string | null;
  reviewFields: string[];
  isSample?: boolean;
}

export interface CandidateDetail extends Candidate {
  sourceEmail: { from: string; subject: string; receivedAt: ISODate; messageId: string } | null;
  applications: Application[];
  processingHistory: { at: ISODate; step: string; status: JobStatus; message?: string }[];
  audit: AuditLog[];
}

export type CandidatePatch = Partial<
  Pick<
    Candidate,
    "name" | "phone" | "college" | "degree" | "graduationYear" | "appliedRole" | "interviewStatus"
  >
>;

export interface Application {
  id: string;
  candidateId: string;
  candidateName: string | null;
  role: string;
  sourceEmail: string;
  receivedAt: ISODate;
  status: ApplicationStatus;
  recruiter: string | null;
  updatedAt: ISODate;
  isSample?: boolean;
}

export interface EmailAttachment {
  id: string;
  fileName: string;
  mimeType: string;
  sizeBytes: number;
}

export interface EmailRecord {
  id: string;
  from: string;
  subject: string;
  receivedAt: ISODate;
  attachments: EmailAttachment[];
  status: ProcessingStatus;
  candidateId: string | null;
  applicationId: string | null;
  error: string | null;
  history: { at: ISODate; step: string; status: JobStatus }[];
  isSample?: boolean;
}

export interface ProcessingJob {
  id: string;
  type: JobType;
  related: string;
  status: JobStatus;
  attempts: number;
  startedAt: ISODate | null;
  completedAt: ISODate | null;
  lastError: string | null;
  nextRetryAt: ISODate | null;
  isSample?: boolean;
}

export interface ExcelSyncRun {
  id: string;
  at: ISODate;
  rows: number;
  result: "success" | "partial" | "failed";
  message: string;
}

export interface ExcelSyncStatus {
  workbookName: string;
  state: ExcelConnectionState;
  storageProvider: string | null;
  lastSuccessfulSync: ISODate | null;
  pendingRows: number;
  failedJobs: number;
  lastReconciliation: string | null;
  workbookUrl: string | null;
  history: ExcelSyncRun[];
}

export interface MicrosoftStatus {
  state: MicrosoftConnectionState;
  accountEmail: string | null;
  mailboxPermission: "granted" | "missing" | "unknown";
  lastMailboxCheck: ISODate | null;
  subscriptionStatus: "active" | "expired" | "none";
  subscriptionExpiresAt: ISODate | null;
  lastReconciliation: ISODate | null;
}

export interface AuditLog {
  id: string;
  at: ISODate;
  actor: string;
  action: string;
  related: string;
  result: "success" | "failure" | "info";
  requestId: string;
  isSample?: boolean;
}

export interface ReviewItem {
  id: string;
  candidateId: string;
  candidateName: string | null;
  attachment: string;
  fields: {
    field: keyof CandidatePatch | "email" | "skills";
    extracted: string | null;
    confidence: number | null;
    reason: string | null;
    evidence: string | null;
  }[];
  isSample?: boolean;
}

export interface HealthStatus {
  status: "ok" | "degraded";
  version?: string;
}
