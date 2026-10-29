import { createFileRoute } from "@tanstack/react-router";
import { Eye, RotateCw } from "lucide-react";
import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { EmptyState, ErrorState, Field, LoadingRows, PageHeader, Panel } from "@/components/app/states";
import { SampleTag, StatusBadge, STATUS_LABEL } from "@/components/app/status-badge";
import { FilterSelect, fmtDate, Pager, SearchInput } from "@/components/app/table-kit";
import { useJobs, useRetryJob } from "@/hooks/use-api";
import type { ProcessingJob } from "@/lib/api/types";

export const Route = createFileRoute("/jobs")({
  head: () => ({
    meta: [
      { title: "Processing Queue · RecruitFlow" },
      { name: "description", content: "Monitor background parsing, OCR, extraction, and sync jobs." },
      { property: "og:title", content: "Processing Queue · RecruitFlow" },
      { property: "og:description", content: "Monitor background parsing, OCR, extraction, and sync jobs." },
    ],
  }),
  component: JobsPage,
});

const typeLabel = (t: string) => t.replace(/_/g, " ").replace(/^\w/, (c) => c.toUpperCase());

function JobsPage() {
  const [status, setStatus] = useState("all");
  const [type, setType] = useState("all");
  const [search, setSearch] = useState("");
  const [page, setPage] = useState(1);
  const [detail, setDetail] = useState<ProcessingJob | null>(null);
  const { data, isLoading, error, refetch } = useJobs({ page, pageSize: 12, status, type, search, sortBy: "startedAt", sortDir: "desc" });
  const retry = useRetryJob();

  return (
    <>
      <PageHeader title="Processing Queue" description="Retries are executed by the backend. Without a backend response, no retry is reported as successful." />
      <Panel>
        <div className="flex flex-col gap-2 border-b p-3 sm:flex-row sm:flex-wrap">
          <SearchInput value={search} onChange={(v) => { setSearch(v); setPage(1); }} placeholder="Job ID or related email" />
          <FilterSelect label="Statuses" value={status} onChange={(v) => { setStatus(v); setPage(1); }} options={["queued", "processing", "succeeded", "retrying", "failed", "cancelled"].map((s) => ({ value: s, label: STATUS_LABEL(s) }))} />
          <FilterSelect label="Types" value={type} onChange={(v) => { setType(v); setPage(1); }} options={["email_ingest", "attachment_parse", "ocr", "field_extraction", "excel_sync"].map((s) => ({ value: s, label: typeLabel(s) }))} />
        </div>
        {error ? <ErrorState error={error} onRetry={refetch} /> : isLoading ? <LoadingRows /> : !data?.items.length ? <EmptyState title="Queue is empty" /> : (
          <>
            <Table>
              <TableHeader><TableRow><TableHead>Job</TableHead><TableHead>Type</TableHead><TableHead>Related</TableHead><TableHead>Status</TableHead><TableHead>Attempts</TableHead><TableHead>Started</TableHead><TableHead>Completed</TableHead><TableHead>Last error</TableHead><TableHead>Next retry</TableHead><TableHead className="text-right">Actions</TableHead></TableRow></TableHeader>
              <TableBody>
                {data.items.map((j) => (
                  <TableRow key={j.id}>
                    <TableCell className="font-mono text-xs"><div className="flex items-center gap-1.5">{j.id}{j.isSample && <SampleTag />}</div></TableCell>
                    <TableCell className="whitespace-nowrap text-xs">{typeLabel(j.type)}</TableCell>
                    <TableCell className="font-mono text-xs">{j.related}</TableCell>
                    <TableCell><StatusBadge status={j.status} /></TableCell>
                    <TableCell className="text-xs tabular-nums">{j.attempts}</TableCell>
                    <TableCell className="whitespace-nowrap text-xs">{fmtDate(j.startedAt, true)}</TableCell>
                    <TableCell className="whitespace-nowrap text-xs">{fmtDate(j.completedAt, true)}</TableCell>
                    <TableCell className="max-w-48 truncate text-xs text-destructive">{j.lastError ?? ""}</TableCell>
                    <TableCell className="whitespace-nowrap text-xs">{fmtDate(j.nextRetryAt, true)}</TableCell>
                    <TableCell className="text-right">
                      <div className="flex justify-end gap-1">
                        {(j.status === "failed" || j.status === "cancelled") && (
                          <Button size="sm" variant="outline" className="h-7" disabled={retry.isPending && retry.variables === j.id} onClick={() => retry.mutate(j.id)}><RotateCw /> Retry</Button>
                        )}
                        <Button size="icon" variant="ghost" className="size-7" onClick={() => setDetail(j)} aria-label={`View ${j.id}`}><Eye /></Button>
                      </div>
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
            <Pager page={page} pageSize={12} total={data.total} onPage={setPage} />
          </>
        )}
      </Panel>
      <Dialog open={!!detail} onOpenChange={(o) => !o && setDetail(null)}>
        <DialogContent>
          <DialogHeader><DialogTitle className="font-mono">{detail?.id}</DialogTitle></DialogHeader>
          {detail && (
            <dl className="divide-y">
              <Field label="Type">{typeLabel(detail.type)}</Field>
              <Field label="Related">{detail.related}</Field>
              <Field label="Status"><StatusBadge status={detail.status} /></Field>
              <Field label="Attempts">{detail.attempts}</Field>
              <Field label="Started">{fmtDate(detail.startedAt, true)}</Field>
              <Field label="Completed">{fmtDate(detail.completedAt, true)}</Field>
              <Field label="Next retry">{fmtDate(detail.nextRetryAt, true)}</Field>
              <Field label="Last error">{detail.lastError ?? "None"}</Field>
            </dl>
          )}
        </DialogContent>
      </Dialog>
    </>
  );
}
