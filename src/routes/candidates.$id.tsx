import { createFileRoute, Link } from "@tanstack/react-router";
import { ArrowLeft, Download, FileText, Pencil } from "lucide-react";
import { useState } from "react";
import { Button } from "@/components/ui/button";
import { ErrorState, Field, LoadingRows, Panel, Unknown } from "@/components/app/states";
import { SampleTag, StatusBadge } from "@/components/app/status-badge";
import { fmtDate } from "@/components/app/table-kit";
import { CandidateEditDialog } from "@/components/app/candidate-edit-dialog";
import { useCandidate } from "@/hooks/use-api";
import { API_BASE_URL } from "@/lib/api/client";
import type { Candidate } from "@/lib/api/types";

export const Route = createFileRoute("/candidates/$id")({
  head: ({ params }) => ({
    meta: [
      { title: `${params.id} · Candidates · RecruitFlow` },
      { name: "description", content: "Candidate profile, extracted CV data, and processing history." },
      { property: "og:title", content: `${params.id} · RecruitFlow` },
      { property: "og:description", content: "Candidate profile, extracted CV data, and processing history." },
    ],
  }),
  component: CandidateDetailPage,
});

function Flag({ c, k, children }: { c: Candidate; k: string; children: React.ReactNode }) {
  return c.reviewFields.includes(k) ? (
    <span className="inline-flex items-center gap-1.5">{children ?? <Unknown />}<StatusBadge status="needs_review" /></span>
  ) : (children ?? <Unknown />);
}

function CandidateDetailPage() {
  const { id } = Route.useParams();
  const { data: c, isLoading, error, refetch } = useCandidate(id);
  const [edit, setEdit] = useState(false);

  return (
    <>
      <Link to="/candidates" className="mb-4 inline-flex items-center gap-1 text-sm text-muted-foreground hover:text-foreground"><ArrowLeft className="size-4" /> Candidates</Link>
      {error ? <Panel><ErrorState error={error} onRetry={refetch} /></Panel> : isLoading || !c ? <Panel><LoadingRows /></Panel> : (
        <>
          <div className="mb-6 flex flex-col justify-between gap-3 sm:flex-row sm:items-center">
            <div>
              <div className="flex items-center gap-2"><h1 className="text-xl font-semibold tracking-tight">{c.name ?? "Name unknown"}</h1>{c.isSample && <SampleTag />}</div>
              <p className="mt-1 font-mono text-xs text-muted-foreground">{c.id} · {c.email ?? "no email"}</p>
            </div>
            <div className="flex items-center gap-2">
              <StatusBadge status={c.processingStatus} />
              <Button size="sm" variant="outline" onClick={() => setEdit(true)}><Pencil /> Edit</Button>
            </div>
          </div>
          <div className="grid gap-4 lg:grid-cols-3">
            <Panel title="Extracted CV information" className="lg:col-span-2">
              <dl className="grid divide-y px-4 sm:grid-cols-2 sm:gap-x-8 sm:divide-y-0">
                <Field label="Phone"><Flag c={c} k="phone">{c.phone}</Flag></Field>
                <Field label="College"><Flag c={c} k="college">{c.college}</Flag></Field>
                <Field label="Degree"><Flag c={c} k="degree">{c.degree}</Flag></Field>
                <Field label="Graduation year"><Flag c={c} k="graduationYear">{c.graduationYear}</Flag></Field>
                <Field label="Experience">{c.experienceYears != null ? `${c.experienceYears} years` : null}</Field>
                <Field label="Applied role">{c.appliedRole}</Field>
                <Field label="Interview"><StatusBadge status={c.interviewStatus} /></Field>
                <Field label="Received">{fmtDate(c.receivedAt, true)}</Field>
              </dl>
              <div className="border-t px-4 py-3">
                <p className="mb-2 text-xs text-muted-foreground">Skills</p>
                {c.skills.length ? <div className="flex flex-wrap gap-1.5">{c.skills.map((s) => <span key={s} className="rounded-md border bg-muted px-2 py-0.5 text-xs">{s}</span>)}</div> : <Unknown />}
              </div>
            </Panel>
            <div className="space-y-4">
              <Panel title="Resume">
                <div className="flex items-center justify-between gap-3 p-4 text-sm">
                  <div className="flex items-center gap-3 min-w-0">
                    <FileText className="size-8 text-muted-foreground shrink-0" />
                    <div className="min-w-0">
                      <p className="truncate font-medium">{c.resumeFileName ?? "No file"}</p>
                      <p className="text-xs text-muted-foreground">{c.resumeRef ? "Verified document" : "Not provided"}</p>
                    </div>
                  </div>
                  {c.resumeRef && (
                    <Button size="sm" variant="outline" asChild>
                      <a
                        href={`${API_BASE_URL}/api/candidates/${encodeURIComponent(c.id)}/resume`}
                        target="_blank"
                        rel="noopener noreferrer"
                        download={c.resumeFileName ?? "resume.docx"}
                      >
                        <Download className="mr-1.5 size-4" /> Download CV
                      </a>
                    </Button>
                  )}
                </div>
                <p className="border-t px-4 py-2 text-xs text-muted-foreground">Files are served only through authorized, short-lived backend downloads — never public URLs.</p>
              </Panel>
              <Panel title="Source email">
                <dl className="divide-y px-4">
                  <Field label="From">{c.sourceEmail?.from}</Field>
                  <Field label="Subject">{c.sourceEmail?.subject}</Field>
                  <Field label="Received">{c.sourceEmail ? fmtDate(c.sourceEmail.receivedAt, true) : null}</Field>
                  <Field label="Status">{c.sourceEmail ? <StatusBadge status="processed" /> : null}</Field>
                </dl>
                {c.sourceEmail ? (
                  <div className="border-t p-4">
                    <p className="text-xs font-medium text-muted-foreground mb-1.5">Message Content</p>
                    {c.sourceEmail.bodyText ? (
                      <div className="rounded-md border bg-muted/30 p-3 text-xs leading-relaxed whitespace-pre-wrap break-words max-h-48 overflow-y-auto">
                        {c.sourceEmail.bodyText}
                      </div>
                    ) : (
                      <p className="text-xs text-muted-foreground italic">No text content in source email.</p>
                    )}
                  </div>
                ) : (
                  <p className="border-t p-4 text-xs text-muted-foreground italic">No source email recorded.</p>
                )}
              </Panel>
            </div>
            <Panel title="Applications">
              <ul className="divide-y">{c.applications.map((a) => <li key={a.id} className="flex items-center justify-between px-4 py-2.5 text-sm"><span>{a.role}<span className="ml-2 font-mono text-xs text-muted-foreground">{a.id}</span></span><StatusBadge status={a.status} /></li>)}</ul>
              {!c.applications.length && <p className="p-4 text-sm text-muted-foreground">No applications.</p>}
            </Panel>
            <Panel title="Processing history">
              <ol className="space-y-3 p-4">{c.processingHistory.map((h, i) => <li key={i} className="flex items-start justify-between gap-2 text-sm"><div><p>{h.step}</p><p className="text-xs text-muted-foreground">{fmtDate(h.at, true)}</p></div><StatusBadge status={h.status} /></li>)}</ol>
              {!c.processingHistory.length && <p className="p-4 text-sm text-muted-foreground">No history.</p>}
            </Panel>
            <Panel title="Audit history">
              <ul className="divide-y">{c.audit.map((a) => <li key={a.id} className="px-4 py-2.5 text-sm"><p className="font-mono text-xs">{a.action}</p><p className="text-xs text-muted-foreground">{a.actor} · {fmtDate(a.at, true)}</p></li>)}</ul>
              {!c.audit.length && <p className="p-4 text-sm text-muted-foreground">No audit events.</p>}
            </Panel>
          </div>
          <CandidateEditDialog candidate={edit ? c : null} onClose={() => setEdit(false)} />
        </>
      )}
    </>
  );
}
