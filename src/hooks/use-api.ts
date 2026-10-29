import { keepPreviousData, useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { toast } from "sonner";
import { describeApiError } from "@/lib/api/client";
import type { ApplicationStatus, CandidatePatch, ListQuery } from "@/lib/api/types";
import { useDataMode } from "@/lib/data-mode";

const opts = { placeholderData: keepPreviousData, retry: false } as const;

export function useCandidates(q: ListQuery) {
  const { api, mode } = useDataMode();
  return useQuery({ queryKey: [mode, "candidates", q], queryFn: () => api.listCandidates(q), ...opts });
}
export function useCandidate(id: string | null) {
  const { api, mode } = useDataMode();
  return useQuery({ queryKey: [mode, "candidate", id], queryFn: () => api.getCandidate(id!), enabled: !!id, retry: false });
}
export function useApplications(q: ListQuery) {
  const { api, mode } = useDataMode();
  return useQuery({ queryKey: [mode, "applications", q], queryFn: () => api.listApplications(q), ...opts });
}
export function useEmails(q: ListQuery) {
  const { api, mode } = useDataMode();
  return useQuery({ queryKey: [mode, "emails", q], queryFn: () => api.listEmails(q), ...opts });
}
export function useJobs(q: ListQuery) {
  const { api, mode } = useDataMode();
  return useQuery({ queryKey: [mode, "jobs", q], queryFn: () => api.listJobs(q), ...opts });
}
export function useAudit(q: ListQuery) {
  const { api, mode } = useDataMode();
  return useQuery({ queryKey: [mode, "audit", q], queryFn: () => api.listAudit(q), ...opts });
}
export function useReview() {
  const { api, mode } = useDataMode();
  return useQuery({ queryKey: [mode, "review"], queryFn: () => api.listReview(), retry: false });
}
export function useExcelStatus() {
  const { api, mode } = useDataMode();
  return useQuery({ queryKey: [mode, "excel"], queryFn: () => api.excelStatus(), retry: false });
}
export function useMicrosoftStatus() {
  const { api, mode } = useDataMode();
  return useQuery({ queryKey: [mode, "microsoft"], queryFn: () => api.microsoftStatus(), retry: false });
}
export function useHealth() {
  const { api, mode } = useDataMode();
  return useQuery({ queryKey: [mode, "health"], queryFn: () => api.health(), retry: false, enabled: mode === "live" });
}

const onError = (e: unknown) => {
  const { title, detail } = describeApiError(e);
  toast.error(title, { description: detail });
};

export function useUpdateCandidate() {
  const { api } = useDataMode();
  const qc = useQueryClient();
  return useMutation({
    mutationFn: ({ id, patch }: { id: string; patch: CandidatePatch }) => api.updateCandidate(id, patch),
    onSuccess: () => { qc.invalidateQueries(); toast.success("Candidate saved"); },
    onError,
  });
}
export function useUpdateApplicationStatus() {
  const { api } = useDataMode();
  const qc = useQueryClient();
  return useMutation({
    mutationFn: ({ id, status }: { id: string; status: ApplicationStatus }) => api.updateApplicationStatus(id, status),
    onSuccess: () => { qc.invalidateQueries(); toast.success("Status updated"); },
    onError,
  });
}
export function useRetryJob() {
  const { api } = useDataMode();
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (id: string) => api.retryJob(id),
    onSuccess: () => { qc.invalidateQueries(); toast.success("Retry accepted by backend"); },
    onError,
  });
}
export function useExcelAction(kind: "test" | "run") {
  const { api } = useDataMode();
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (): Promise<{ ok?: boolean; message?: string; queued?: number }> => (kind === "test" ? api.excelTest() : api.excelRun()),
    onSuccess: (r) => {
      qc.invalidateQueries();
      toast.success(kind === "test" ? (r.message ?? "Connection OK") : `Backend queued ${r.queued ?? 0} rows`);
    },
    onError,
  });
}
