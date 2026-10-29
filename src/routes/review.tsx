import { createFileRoute, Link } from "@tanstack/react-router";
import { CheckCircle2 } from "lucide-react";
import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { EmptyState, ErrorState, LoadingRows, PageHeader, Panel } from "@/components/app/states";
import { SampleTag } from "@/components/app/status-badge";
import { useReview, useUpdateCandidate } from "@/hooks/use-api";
import type { CandidatePatch, ReviewItem } from "@/lib/api/types";

export const Route = createFileRoute("/review")({
  head: () => ({
    meta: [
      { title: "Review Required · RecruitFlow" },
      { name: "description", content: "Confirm or correct uncertain CV extraction results." },
      { property: "og:title", content: "Review Required · RecruitFlow" },
      { property: "og:description", content: "Confirm or correct uncertain CV extraction results." },
    ],
  }),
  component: ReviewPage,
});

const LABEL: Record<string, string> = { graduationYear: "Graduation year", phone: "Phone", college: "College", degree: "Degree", name: "Name", appliedRole: "Applied role", email: "Email", skills: "Skills" };

function ReviewCard({ item }: { item: ReviewItem }) {
  const mut = useUpdateCandidate();
  const [vals, setVals] = useState<Record<string, string>>({});
  const [err, setErr] = useState<string | null>(null);
  const editable = item.fields.filter((f) => f.field !== "email" && f.field !== "skills");

  const save = () => {
    const patch: Record<string, unknown> = {};
    for (const f of editable) {
      const v = (vals[f.field] ?? "").trim();
      if (!v) continue;
      if (f.field === "graduationYear") {
        if (!/^\d{4}$/.test(v)) return setErr("Graduation year must be a 4-digit year.");
        patch[f.field] = +v;
      } else patch[f.field] = v;
    }
    if (!Object.keys(patch).length) return setErr("Enter at least one verified value. Leave unknown fields blank.");
    setErr(null);
    mut.mutate({ id: item.candidateId, patch: patch as CandidatePatch });
  };

  return (
    <Panel
      title={item.candidateName ?? "Name unknown"}
      description={item.attachment}
      actions={<div className="flex items-center gap-2">{item.isSample && <SampleTag />}<Link to="/candidates/$id" params={{ id: item.candidateId }} className="font-mono text-xs text-primary">{item.candidateId}</Link></div>}
    >
      <div className="divide-y">
        {item.fields.map((f) => (
          <div key={f.field} className="grid gap-3 p-4 md:grid-cols-[160px_1fr_1fr]">
            <div>
              <p className="text-sm font-medium">{LABEL[f.field] ?? f.field}</p>
              {f.confidence != null && <p className="text-xs text-muted-foreground">Confidence {(f.confidence * 100).toFixed(0)}%</p>}
            </div>
            <div className="space-y-1 text-sm">
              <p><span className="text-muted-foreground">Extracted: </span>{f.extracted ? <code className="font-mono">{f.extracted}</code> : <span className="italic text-muted-foreground">nothing found</span>}</p>
              {f.reason && <p className="text-xs text-muted-foreground">{f.reason}</p>}
              {f.evidence && <blockquote className="border-l-2 border-warning pl-2 font-mono text-xs text-muted-foreground">{f.evidence}</blockquote>}
            </div>
            {f.field === "email" || f.field === "skills" ? (
              <p className="text-xs text-muted-foreground">Not editable here.</p>
            ) : (
              <Input placeholder="Verified value (leave blank if unknown)" value={vals[f.field] ?? ""} onChange={(e) => setVals((s) => ({ ...s, [f.field]: e.target.value }))} aria-label={`Corrected ${LABEL[f.field]}`} />
            )}
          </div>
        ))}
      </div>
      <div className="flex items-center justify-end gap-3 border-t px-4 py-3">
        {err && <p className="mr-auto text-xs text-destructive">{err}</p>}
        <Button size="sm" disabled={mut.isPending || !editable.length} onClick={save}><CheckCircle2 /> {mut.isPending ? "Saving…" : "Approve corrections"}</Button>
      </div>
    </Panel>
  );
}

function ReviewPage() {
  const { data, isLoading, error, refetch } = useReview();
  return (
    <>
      <PageHeader title="Review Required" description="Only enter values you can verify from the CV. Missing information stays unknown — nothing is guessed." />
      {error ? <Panel><ErrorState error={error} onRetry={refetch} /></Panel> : isLoading ? <Panel><LoadingRows /></Panel> : !data?.length ? (
        <Panel><EmptyState icon={<CheckCircle2 className="size-5" />} title="Nothing to review" description="All extracted fields have sufficient evidence." /></Panel>
      ) : <div className="space-y-4">{data.map((r) => <ReviewCard key={r.id} item={r} />)}</div>}
    </>
  );
}
