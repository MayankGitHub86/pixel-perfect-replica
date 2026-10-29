import { createFileRoute } from "@tanstack/react-router";
import { ExternalLink, FileSpreadsheet, Link2, PlugZap, RefreshCw, RotateCw } from "lucide-react";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { EmptyState, ErrorState, Field, LoadingRows, PageHeader, Panel } from "@/components/app/states";
import { StatusBadge } from "@/components/app/status-badge";
import { fmtDate } from "@/components/app/table-kit";
import { useExcelAction, useExcelStatus } from "@/hooks/use-api";
import { useDataMode } from "@/lib/data-mode";

export const Route = createFileRoute("/excel-sync")({
  head: () => ({
    meta: [
      { title: "Excel Sync · RecruitFlow" },
      { name: "description", content: "Synchronization status for the Candidate_Master.xlsx workbook." },
      { property: "og:title", content: "Excel Sync · RecruitFlow" },
      { property: "og:description", content: "Synchronization status for the Candidate_Master.xlsx workbook." },
    ],
  }),
  component: ExcelSyncPage,
});

function ExcelSyncPage() {
  const { mode } = useDataMode();
  const { data: s, isLoading, error, refetch } = useExcelStatus();
  const test = useExcelAction("test");
  const run = useExcelAction("run");
  const connected = s?.state === "connected" || s?.state === "syncing" || s?.state === "attention_required";

  return (
    <>
      <PageHeader
        title="Excel Sync"
        description="The backend writes validated records to the workbook and decides which Microsoft Graph workbook APIs your account supports."
        actions={
          <>
            <Button size="sm" variant="outline" onClick={() => toast.info("Workbook connection is configured on the backend", { description: "It will expose a connect endpoint once workbook storage is chosen (OneDrive/SharePoint)." })}><Link2 /> Connect workbook</Button>
            <Button size="sm" variant="outline" disabled={test.isPending} onClick={() => test.mutate()}><PlugZap /> {test.isPending ? "Testing…" : "Test connection"}</Button>
            <Button size="sm" disabled={run.isPending} onClick={() => run.mutate()}><RefreshCw className={run.isPending ? "animate-spin" : ""} /> Sync pending</Button>
          </>
        }
      />
      {error ? <Panel><ErrorState error={error} onRetry={refetch} /></Panel> : isLoading || !s ? <Panel><LoadingRows /></Panel> : (
        <div className="grid gap-4 lg:grid-cols-3">
          <Panel className="lg:col-span-1">
            <div className="flex items-center gap-3 border-b p-4">
              <div className="flex size-10 items-center justify-center rounded-md bg-success/10 text-success"><FileSpreadsheet className="size-5" /></div>
              <div><p className="font-medium">{s.workbookName}</p><StatusBadge status={s.state} /></div>
            </div>
            <dl className="divide-y px-4">
              <Field label="Storage provider">{s.storageProvider}</Field>
              <Field label="Last successful sync">{s.lastSuccessfulSync ? fmtDate(s.lastSuccessfulSync, true) : "Never"}</Field>
              <Field label="Pending rows"><span className="tabular-nums">{s.pendingRows}</span></Field>
              <Field label="Failed sync jobs"><span className="tabular-nums">{s.failedJobs}</span></Field>
              <Field label="Last reconciliation">{s.lastReconciliation}</Field>
            </dl>
            <div className="flex flex-wrap gap-2 border-t p-4">
              <Button size="sm" variant="outline" disabled={!s.failedJobs || run.isPending} onClick={() => run.mutate()}><RotateCw /> Retry failed</Button>
              {s.workbookUrl ? (
                <Button size="sm" variant="outline" asChild><a href={s.workbookUrl} target="_blank" rel="noreferrer"><ExternalLink /> Open workbook</a></Button>
              ) : (
                <Button size="sm" variant="outline" disabled title="Available when the backend returns a workbook link"><ExternalLink /> Open workbook</Button>
              )}
            </div>
            {mode === "sample" && <p className="border-t px-4 py-3 text-xs text-muted-foreground">Sample mode: the workbook is shown as not connected and sync actions are refused rather than simulated.</p>}
          </Panel>
          <Panel title="Synchronization history" className="lg:col-span-2">
            {!s.history.length ? (
              <EmptyState icon={<FileSpreadsheet className="size-5" />} title={connected ? "No sync runs yet" : "Workbook not connected"} description="Sync runs reported by the backend will be listed here with row counts and results." />
            ) : (
              <Table>
                <TableHeader><TableRow><TableHead>Run</TableHead><TableHead>Time</TableHead><TableHead>Rows</TableHead><TableHead>Result</TableHead><TableHead>Details</TableHead></TableRow></TableHeader>
                <TableBody>{s.history.map((h) => <TableRow key={h.id}><TableCell className="font-mono text-xs">{h.id}</TableCell><TableCell className="text-xs">{fmtDate(h.at, true)}</TableCell><TableCell className="tabular-nums">{h.rows}</TableCell><TableCell><StatusBadge status={h.result} /></TableCell><TableCell className="text-xs text-muted-foreground">{h.message}</TableCell></TableRow>)}</TableBody>
              </Table>
            )}
          </Panel>
        </div>
      )}
    </>
  );
}
