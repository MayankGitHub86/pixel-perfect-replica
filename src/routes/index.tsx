import { createFileRoute, Link } from "@tanstack/react-router";
import { AlertTriangle, CheckCircle2, FileSpreadsheet, Inbox, Users, XCircle, type LucideIcon } from "lucide-react";
import { Area, AreaChart, Bar, BarChart, CartesianGrid, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";
import { useMemo } from "react";
import { Button } from "@/components/ui/button";
import { EmptyState, ErrorState, LoadingRows, PageHeader, Panel, Field } from "@/components/app/states";
import { StatusBadge } from "@/components/app/status-badge";
import { fmtDate } from "@/components/app/table-kit";
import { useAudit, useCandidates, useEmails, useExcelStatus, useJobs, useMicrosoftStatus } from "@/hooks/use-api";
import { useDataMode } from "@/lib/data-mode";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Overview · RecruitFlow" },
      { name: "description", content: "Pipeline health for CV ingestion, processing jobs, and Excel synchronization." },
      { property: "og:title", content: "Overview · RecruitFlow" },
      { property: "og:description", content: "Pipeline health for CV ingestion, processing jobs, and Excel synchronization." },
    ],
  }),
  component: Overview,
});

const ALL = { page: 1, pageSize: 1000 };

function Metric({ label, value, icon: Icon, to }: { label: string; value: number | string; icon: LucideIcon; to: string }) {
  return (
    <Link to={to} className="group rounded-lg border bg-card p-4 shadow-card transition-colors hover:border-primary/40 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring">
      <div className="flex items-center justify-between text-xs text-muted-foreground">
        {label}
        <Icon className="size-4 group-hover:text-primary" />
      </div>
      <p className="mt-2 text-2xl font-semibold tabular-nums tracking-tight">{value}</p>
    </Link>
  );
}

const tooltipStyle = { background: "var(--popover)", border: "1px solid var(--border)", borderRadius: 8, fontSize: 12 };

function Overview() {
  const { mode, setMode } = useDataMode();
  const cands = useCandidates(ALL);
  const emails = useEmails(ALL);
  const jobs = useJobs(ALL);
  const excel = useExcelStatus();
  const ms = useMicrosoftStatus();
  const audit = useAudit({ page: 1, pageSize: 6 });

  const m = useMemo(() => {
    const c = cands.data?.items ?? [];
    const e = emails.data?.items ?? [];
    const j = jobs.data?.items ?? [];
    const byDay = new Map<string, { day: string; emails: number }>();
    for (let i = 13; i >= 0; i--) {
      const d = new Date(Date.now() - i * 86_400_000).toISOString().slice(0, 10);
      byDay.set(d, { day: d.slice(5), emails: 0 });
    }
    e.forEach((x) => { const k = x.receivedAt.slice(0, 10); const b = byDay.get(k); if (b) b.emails++; });
    const types = ["email_ingest", "attachment_parse", "ocr", "field_extraction", "excel_sync"];
    return {
      total: cands.data?.total ?? 0,
      received: emails.data?.total ?? 0,
      processed: c.filter((x) => x.processingStatus === "processed").length,
      review: c.filter((x) => x.processingStatus === "needs_review").length,
      failedJobs: j.filter((x) => x.status === "failed").length,
      trend: [...byDay.values()],
      jobsByType: types.map((t) => ({
        type: t.replace("_", " "),
        succeeded: j.filter((x) => x.type === t && x.status === "succeeded").length,
        failed: j.filter((x) => x.type === t && x.status === "failed").length,
      })),
      jobStatus: ["queued", "processing", "retrying", "failed", "succeeded", "cancelled"].map((s) => [s, j.filter((x) => x.status === s).length] as const),
      recent: [...c].sort((a, b) => b.updatedAt.localeCompare(a.updatedAt)).slice(0, 6),
    };
  }, [cands.data, emails.data, jobs.data]);

  const loading = cands.isLoading || emails.isLoading || jobs.isLoading;
  const err = cands.error ?? emails.error ?? jobs.error;

  return (
    <>
      <PageHeader title="Overview" description="Pipeline health from Outlook ingestion through Candidate_Master.xlsx." />
      {err ? (
        <Panel><ErrorState error={err} onRetry={() => { cands.refetch(); emails.refetch(); jobs.refetch(); }} /></Panel>
      ) : loading ? (
        <Panel><LoadingRows rows={8} /></Panel>
      ) : m.total === 0 && m.received === 0 ? (
        <Panel>
          <EmptyState
            title="No applications yet"
            description="Connect Outlook in Settings. New CVs will appear here once the backend processes them."
            action={
              <div className="flex gap-2">
                <Button asChild size="sm"><Link to="/settings">Connect Outlook</Link></Button>
                {mode === "live" && <Button size="sm" variant="outline" onClick={() => setMode("sample")}>View sample data</Button>}
              </div>
            }
          />
        </Panel>
      ) : (
        <div className="space-y-4">
          <div className="grid grid-cols-2 gap-3 md:grid-cols-3 xl:grid-cols-6">
            <Metric label="Total Candidates" value={m.total} icon={Users} to="/candidates" />
            <Metric label="Applications Received" value={m.received} icon={Inbox} to="/emails" />
            <Metric label="Successfully Processed" value={m.processed} icon={CheckCircle2} to="/candidates" />
            <Metric label="Needs Review" value={m.review} icon={AlertTriangle} to="/review" />
            <Metric label="Failed Jobs" value={m.failedJobs} icon={XCircle} to="/jobs" />
            <Metric label="Excel Sync Pending" value={excel.data?.pendingRows ?? "—"} icon={FileSpreadsheet} to="/excel-sync" />
          </div>

          <div className="grid gap-4 lg:grid-cols-3">
            <Panel title="Email ingestion" description="Emails received, last 14 days" className="lg:col-span-2">
              <div className="h-56 p-3">
                <ResponsiveContainer>
                  <AreaChart data={m.trend}>
                    <defs>
                      <linearGradient id="g" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stopColor="var(--chart-1)" stopOpacity={0.25} />
                        <stop offset="100%" stopColor="var(--chart-1)" stopOpacity={0} />
                      </linearGradient>
                    </defs>
                    <CartesianGrid stroke="var(--border)" vertical={false} />
                    <XAxis dataKey="day" tick={{ fontSize: 11, fill: "var(--muted-foreground)" }} tickLine={false} axisLine={false} />
                    <YAxis allowDecimals={false} width={24} tick={{ fontSize: 11, fill: "var(--muted-foreground)" }} tickLine={false} axisLine={false} />
                    <Tooltip contentStyle={tooltipStyle} />
                    <Area dataKey="emails" stroke="var(--chart-1)" strokeWidth={2} fill="url(#g)" />
                  </AreaChart>
                </ResponsiveContainer>
              </div>
            </Panel>
            <Panel title="Connections">
              <dl className="divide-y px-4">
                <Field label="Outlook">{ms.data ? <StatusBadge status={ms.data.state} /> : ms.isError ? <StatusBadge status="error" /> : "…"}</Field>
                <Field label="Account">{ms.data?.accountEmail}</Field>
                <Field label="Excel workbook">{excel.data ? <StatusBadge status={excel.data.state} /> : excel.isError ? <StatusBadge status="error" /> : "…"}</Field>
                <Field label="Last sync">{excel.data?.lastSuccessfulSync ? fmtDate(excel.data.lastSuccessfulSync, true) : null}</Field>
                <Field label="Failed sync jobs">{excel.data?.failedJobs ?? null}</Field>
              </dl>
            </Panel>
          </div>

          <div className="grid gap-4 lg:grid-cols-3">
            <Panel title="CV processing outcomes" description="Succeeded vs failed by job type" className="lg:col-span-2">
              <div className="h-56 p-3">
                <ResponsiveContainer>
                  <BarChart data={m.jobsByType}>
                    <CartesianGrid stroke="var(--border)" vertical={false} />
                    <XAxis dataKey="type" tick={{ fontSize: 11, fill: "var(--muted-foreground)" }} tickLine={false} axisLine={false} />
                    <YAxis allowDecimals={false} width={24} tick={{ fontSize: 11, fill: "var(--muted-foreground)" }} tickLine={false} axisLine={false} />
                    <Tooltip contentStyle={tooltipStyle} cursor={{ fill: "var(--muted)" }} />
                    <Bar dataKey="succeeded" fill="var(--chart-2)" radius={[3, 3, 0, 0]} />
                    <Bar dataKey="failed" fill="var(--chart-3)" radius={[3, 3, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </Panel>
            <Panel title="Processing jobs" actions={<Link to="/jobs" className="text-xs text-primary">View queue</Link>}>
              <ul className="divide-y px-4">
                {m.jobStatus.map(([s, n]) => (
                  <li key={s} className="flex items-center justify-between py-2 text-sm">
                    <StatusBadge status={s} />
                    <span className="tabular-nums">{n}</span>
                  </li>
                ))}
              </ul>
            </Panel>
          </div>

          <div className="grid gap-4 lg:grid-cols-2">
            <Panel title="Recent candidate activity" actions={<Link to="/candidates" className="text-xs text-primary">All candidates</Link>}>
              <ul className="divide-y">
                {m.recent.map((c) => (
                  <li key={c.id}>
                    <Link to="/candidates/$id" params={{ id: c.id }} className="flex items-center justify-between gap-3 px-4 py-2.5 text-sm hover:bg-muted/60">
                      <div className="min-w-0">
                        <p className="truncate font-medium">{c.name ?? <span className="italic text-muted-foreground">Name unknown</span>}</p>
                        <p className="truncate text-xs text-muted-foreground">{c.appliedRole} · {fmtDate(c.updatedAt, true)}</p>
                      </div>
                      <StatusBadge status={c.processingStatus} />
                    </Link>
                  </li>
                ))}
              </ul>
            </Panel>
            <Panel title="Recent audit events" actions={<Link to="/activity" className="text-xs text-primary">Activity logs</Link>}>
              {audit.isError ? <ErrorState error={audit.error} /> : (
                <ul className="divide-y">
                  {(audit.data?.items ?? []).map((a) => (
                    <li key={a.id} className="flex items-center justify-between gap-3 px-4 py-2.5 text-sm">
                      <div className="min-w-0">
                        <p className="truncate font-mono text-xs">{a.action}</p>
                        <p className="truncate text-xs text-muted-foreground">{a.actor} · {a.related} · {fmtDate(a.at, true)}</p>
                      </div>
                      <StatusBadge status={a.result} />
                    </li>
                  ))}
                </ul>
              )}
            </Panel>
          </div>
        </div>
      )}
    </>
  );
}
