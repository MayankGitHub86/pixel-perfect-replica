// Centralized data-access layer. Pages call these functions only; switching
// from synthetic sample data to the real backend needs no page changes.
import { ApiError, apiRequest } from "./client";
import { buildSampleData } from "./sample-data";
import type {
  Application,
  ApplicationStatus,
  AuditLog,
  Candidate,
  CandidateDetail,
  CandidatePatch,
  EmailRecord,
  ExcelSyncStatus,
  HealthStatus,
  ListQuery,
  MicrosoftStatus,
  Paginated,
  ProcessingJob,
  ReviewItem,
} from "./types";

export type DataMode = "sample" | "live";

let store: ReturnType<typeof buildSampleData> | null = null;
const db = () => (store ??= buildSampleData());
const delay = (ms = 250) => new Promise((r) => setTimeout(r, ms));

function listLocal<T extends object>(rows: T[], q: ListQuery, searchKeys: (keyof T)[]): Paginated<T> {
  const { page = 1, pageSize = 10, search, sortBy, sortDir = "desc", ...filters } = q;
  let out = rows;
  if (search) {
    const s = search.toLowerCase();
    out = out.filter((r) =>
      searchKeys.some((k) => {
        const v = r[k];
        return (Array.isArray(v) ? v.join(" ") : String(v ?? "")).toLowerCase().includes(s);
      }),
    );
  }
  for (const [k, v] of Object.entries(filters)) {
    if (v === undefined || v === "" || v === "all") continue;
    if (k === "since") {
      const min = Date.now() - Number(v) * 86_400_000;
      out = out.filter((r) => new Date(String((r as Record<string, unknown>).receivedAt ?? (r as Record<string, unknown>).at)).getTime() >= min);
      continue;
    }
    out = out.filter((r) => String((r as Record<string, unknown>)[k]) === String(v));
  }
  if (sortBy) {
    out = [...out].sort((a, b) => {
      const av = (a as Record<string, unknown>)[sortBy];
      const bv = (b as Record<string, unknown>)[sortBy];
      const c = String(av ?? "").localeCompare(String(bv ?? ""), undefined, { numeric: true });
      return sortDir === "asc" ? c : -c;
    });
  }
  const total = out.length;
  return { items: out.slice((page - 1) * pageSize, page * pageSize), total, page, pageSize };
}

const sampleOnly = (action: string) =>
  new ApiError(501, `${action} requires the real backend. Sample mode never simulates this action.`);

export function createServices(mode: DataMode) {
  const live = mode === "live";
  return {
    mode,
    health: (): Promise<HealthStatus> => (live ? apiRequest("/api/health") : Promise.reject(sampleOnly("Health check"))),

    listCandidates: async (q: ListQuery): Promise<Paginated<Candidate>> =>
      live ? apiRequest("/api/candidates", { query: q }) : (await delay(), listLocal(db().candidates, q, ["id", "name", "email", "skills"])),

    getCandidate: async (id: string): Promise<CandidateDetail> => {
      if (live) return apiRequest(`/api/candidates/${encodeURIComponent(id)}`);
      await delay();
      const c = db().candidates.find((x) => x.id === id);
      if (!c) throw new ApiError(404, `Candidate ${id} was not found.`);
      const email = db().emails.find((e) => e.id === c.sourceEmailId);
      return {
        ...c,
        sourceEmail: email ? { from: email.from, subject: email.subject, receivedAt: email.receivedAt, messageId: email.id } : null,
        applications: db().applications.filter((a) => a.candidateId === id),
        processingHistory: (email?.history ?? []).map((h) => ({ ...h })),
        audit: db().audit.filter((a) => a.related === id),
      };
    },

    updateCandidate: async (id: string, patch: CandidatePatch): Promise<Candidate> => {
      if (live) return apiRequest(`/api/candidates/${encodeURIComponent(id)}`, { method: "PATCH", body: patch });
      await delay();
      const c = db().candidates.find((x) => x.id === id);
      if (!c) throw new ApiError(404, "Candidate not found");
      Object.assign(c, patch, { updatedAt: new Date().toISOString() });
      c.reviewFields = c.reviewFields.filter((f) => (patch as Record<string, unknown>)[f] == null);
      return { ...c };
    },

    listApplications: async (q: ListQuery): Promise<Paginated<Application>> =>
      live ? apiRequest("/api/applications", { query: q }) : (await delay(), listLocal(db().applications, q, ["id", "candidateName", "sourceEmail", "role"])),

    updateApplicationStatus: async (id: string, status: ApplicationStatus): Promise<Application> => {
      if (live) return apiRequest(`/api/applications/${encodeURIComponent(id)}/status`, { method: "PATCH", body: { status } });
      await delay();
      const a = db().applications.find((x) => x.id === id);
      if (!a) throw new ApiError(404, "Application not found");
      a.status = status;
      a.updatedAt = new Date().toISOString();
      return { ...a };
    },

    listEmails: async (q: ListQuery): Promise<Paginated<EmailRecord>> =>
      live ? apiRequest("/api/emails", { query: q }) : (await delay(), listLocal(db().emails, q, ["from", "subject", "id"])),

    listJobs: async (q: ListQuery): Promise<Paginated<ProcessingJob>> =>
      live ? apiRequest("/api/jobs", { query: q }) : (await delay(), listLocal(db().jobs, q, ["id", "related", "type"])),

    retryJob: (id: string): Promise<ProcessingJob> =>
      live ? apiRequest(`/api/jobs/${encodeURIComponent(id)}/retry`, { method: "POST" }) : Promise.reject(sampleOnly("Retrying a job")),

    excelStatus: async (): Promise<ExcelSyncStatus> => (live ? apiRequest("/api/excel-sync/status") : (await delay(), { ...db().excel })),
    excelTest: (): Promise<{ ok: boolean; message: string }> =>
      live ? apiRequest("/api/excel-sync/test", { method: "POST" }) : Promise.reject(sampleOnly("Testing the workbook connection")),
    excelRun: (): Promise<{ queued: number }> =>
      live ? apiRequest("/api/excel-sync/run", { method: "POST" }) : Promise.reject(sampleOnly("Syncing to Excel")),

    listAudit: async (q: ListQuery): Promise<Paginated<AuditLog>> =>
      live ? apiRequest("/api/audit-logs", { query: q }) : (await delay(), listLocal(db().audit, q, ["actor", "action", "related", "requestId"])),

    listReview: async (): Promise<ReviewItem[]> => {
      // Proposed: derived from candidates with processingStatus=needs_review.
      if (live) {
        const res = await apiRequest<Paginated<Candidate>>("/api/candidates", { query: { processingStatus: "needs_review", pageSize: 100 } });
        return res.items.map((c) => ({
          id: `REV-${c.id}`, candidateId: c.id, candidateName: c.name, attachment: c.resumeFileName ?? "unknown",
          fields: c.reviewFields.map((f) => ({ field: f as ReviewItem["fields"][number]["field"], extracted: null, confidence: null, reason: null, evidence: null })),
        }));
      }
      await delay();
      const open = new Set(db().candidates.filter((c) => c.reviewFields.length).map((c) => c.id));
      return db().review.filter((r) => open.has(r.candidateId));
    },

    microsoftStatus: async (): Promise<MicrosoftStatus> =>
      live ? apiRequest("/api/integrations/microsoft/status") : (await delay(), { ...db().microsoft }),
    /** Backend OAuth initiation URL — the browser navigates there; no credentials are handled here. */
    microsoftConnectPath: "/api/integrations/microsoft/connect",
  };
}

export type Services = ReturnType<typeof createServices>;
