import { createFileRoute, Link } from "@tanstack/react-router";
import { Paperclip } from "lucide-react";
import { useState } from "react";
import { Sheet, SheetContent, SheetDescription, SheetHeader, SheetTitle } from "@/components/ui/sheet";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Tabs, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { EmptyState, ErrorState, Field, LoadingRows, PageHeader, Panel } from "@/components/app/states";
import { SampleTag, StatusBadge } from "@/components/app/status-badge";
import { fmtDate, Pager, SearchInput } from "@/components/app/table-kit";
import { useEmails } from "@/hooks/use-api";
import type { EmailRecord } from "@/lib/api/types";

export const Route = createFileRoute("/emails")({
  head: () => ({
    meta: [
      { title: "Email Inbox · RecruitFlow" },
      { name: "description", content: "Application emails ingested from Outlook via Microsoft Graph." },
      { property: "og:title", content: "Email Inbox · RecruitFlow" },
      { property: "og:description", content: "Application emails ingested from Outlook via Microsoft Graph." },
    ],
  }),
  component: EmailsPage,
});

const kb = (b: number) => (b > 1e6 ? `${(b / 1e6).toFixed(1)} MB` : `${Math.round(b / 1e3)} KB`);

function EmailsPage() {
  const [status, setStatus] = useState("all");
  const [search, setSearch] = useState("");
  const [page, setPage] = useState(1);
  const [open, setOpen] = useState<EmailRecord | null>(null);
  const { data, isLoading, error, refetch } = useEmails({ page, pageSize: 12, status, search, sortBy: "receivedAt", sortDir: "desc" });

  return (
    <>
      <PageHeader title="Email Inbox" description="Ingestion records created by the backend from Microsoft Graph. This page never reads your mailbox directly." />
      <Panel>
        <div className="flex flex-col gap-2 border-b p-3 md:flex-row md:items-center md:justify-between">
          <Tabs value={status} onValueChange={(v) => { setStatus(v); setPage(1); }}>
            <TabsList className="flex-wrap">
              {[["all", "All"], ["new", "New"], ["processing", "Processing"], ["processed", "Processed"], ["needs_review", "Needs Review"], ["failed", "Failed"]].map(([v, l]) => <TabsTrigger key={v} value={v!}>{l}</TabsTrigger>)}
            </TabsList>
          </Tabs>
          <SearchInput value={search} onChange={(v) => { setSearch(v); setPage(1); }} placeholder="Sender or subject" />
        </div>
        {error ? <ErrorState error={error} onRetry={refetch} /> : isLoading ? <LoadingRows /> : !data?.items.length ? <EmptyState title="No emails" description="Emails appear once the backend's Graph subscription delivers them." /> : (
          <>
            <Table>
              <TableHeader><TableRow><TableHead>Sender</TableHead><TableHead>Subject</TableHead><TableHead>Received</TableHead><TableHead>Files</TableHead><TableHead>Status</TableHead><TableHead>Linked</TableHead><TableHead>Error</TableHead></TableRow></TableHeader>
              <TableBody>
                {data.items.map((e) => (
                  <TableRow key={e.id} className="cursor-pointer" onClick={() => setOpen(e)} tabIndex={0} onKeyDown={(k) => k.key === "Enter" && setOpen(e)}>
                    <TableCell className="text-xs"><div className="flex items-center gap-1.5">{e.from}{e.isSample && <SampleTag />}</div></TableCell>
                    <TableCell className="max-w-64 truncate text-sm font-medium">{e.subject}</TableCell>
                    <TableCell className="whitespace-nowrap text-xs">{fmtDate(e.receivedAt, true)}</TableCell>
                    <TableCell className="text-xs"><span className="inline-flex items-center gap-1"><Paperclip className="size-3" />{e.attachments.length}</span></TableCell>
                    <TableCell><StatusBadge status={e.status} /></TableCell>
                    <TableCell className="font-mono text-[11px]">{e.candidateId ?? "—"}</TableCell>
                    <TableCell className="max-w-48 truncate text-xs text-destructive">{e.error ?? ""}</TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
            <Pager page={page} pageSize={12} total={data.total} onPage={setPage} />
          </>
        )}
      </Panel>
      <Sheet open={!!open} onOpenChange={(o) => !o && setOpen(null)}>
        <SheetContent className="w-full overflow-y-auto sm:max-w-md">
          {open && (
            <>
              <SheetHeader><SheetTitle>{open.subject}</SheetTitle><SheetDescription>{open.from}</SheetDescription></SheetHeader>
              <div className="space-y-5 px-4 pb-6">
                <dl className="divide-y">
                  <Field label="Message ID"><span className="font-mono text-xs">{open.id}</span></Field>
                  <Field label="Received">{fmtDate(open.receivedAt, true)}</Field>
                  <Field label="Status"><StatusBadge status={open.status} /></Field>
                  <Field label="Candidate">{open.candidateId ? <Link to="/candidates/$id" params={{ id: open.candidateId }} className="text-primary hover:underline">{open.candidateId}</Link> : null}</Field>
                  <Field label="Application">{open.applicationId}</Field>
                </dl>
                {open.error && <p className="rounded-md border border-destructive/30 bg-destructive/10 p-3 text-sm text-destructive">{open.error}</p>}
                <div>
                  <p className="mb-2 text-xs font-medium text-muted-foreground">Attachments</p>
                  <ul className="space-y-2">{open.attachments.map((a) => <li key={a.id} className="rounded-md border p-2.5 text-sm"><p className="font-medium">{a.fileName}</p><p className="text-xs text-muted-foreground">{a.mimeType} · {kb(a.sizeBytes)}</p></li>)}</ul>
                </div>
                <div>
                  <p className="mb-2 text-xs font-medium text-muted-foreground">Processing history</p>
                  <ol className="space-y-2">{open.history.map((h, i) => <li key={i} className="flex items-center justify-between text-sm"><span>{h.step}<span className="ml-2 text-xs text-muted-foreground">{fmtDate(h.at, true)}</span></span><StatusBadge status={h.status} /></li>)}</ol>
                </div>
              </div>
            </>
          )}
        </SheetContent>
      </Sheet>
    </>
  );
}
