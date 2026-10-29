import { cva, type VariantProps } from "class-variance-authority";
import { cn } from "@/lib/utils";

const tone = cva(
  "inline-flex items-center gap-1.5 whitespace-nowrap rounded-full border px-2 py-0.5 text-xs font-medium",
  {
    variants: {
      tone: {
        neutral: "border-border bg-muted text-muted-foreground",
        success: "border-success/25 bg-success/10 text-success",
        warning: "border-warning/30 bg-warning/15 text-foreground",
        danger: "border-destructive/25 bg-destructive/10 text-destructive",
        info: "border-info/25 bg-info/10 text-info",
        primary: "border-primary/25 bg-primary/10 text-primary",
      },
    },
    defaultVariants: { tone: "neutral" },
  },
);
const dot = { neutral: "bg-muted-foreground", success: "bg-success", warning: "bg-warning", danger: "bg-destructive", info: "bg-info", primary: "bg-primary" };

type Tone = NonNullable<VariantProps<typeof tone>["tone"]>;

const MAP: Record<string, [string, Tone]> = {
  new: ["New", "info"],
  processing: ["Processing", "primary"],
  processed: ["Processed", "success"],
  needs_review: ["Needs review", "warning"],
  failed: ["Failed", "danger"],
  queued: ["Queued", "neutral"],
  succeeded: ["Succeeded", "success"],
  retrying: ["Retrying", "warning"],
  cancelled: ["Cancelled", "neutral"],
  under_review: ["Under review", "primary"],
  shortlisted: ["Shortlisted", "info"],
  interview: ["Interview", "primary"],
  rejected: ["Rejected", "neutral"],
  hired: ["Hired", "success"],
  not_scheduled: ["Not scheduled", "neutral"],
  scheduled: ["Scheduled", "info"],
  completed: ["Completed", "success"],
  no_show: ["No show", "danger"],
  not_connected: ["Not connected", "neutral"],
  connecting: ["Connecting", "primary"],
  connected: ["Connected", "success"],
  syncing: ["Syncing", "primary"],
  attention_required: ["Attention required", "warning"],
  error: ["Error", "danger"],
  expired: ["Expired", "warning"],
  success: ["Success", "success"],
  failure: ["Failure", "danger"],
  info: ["Info", "info"],
  partial: ["Partial", "warning"],
};

export const STATUS_LABEL = (s: string) => MAP[s]?.[0] ?? s;

export function StatusBadge({ status, className }: { status: string; className?: string }) {
  const [label, t] = MAP[status] ?? [status, "neutral" as Tone];
  return (
    <span className={cn(tone({ tone: t }), className)}>
      <span className={cn("size-1.5 rounded-full", dot[t])} />
      {label}
    </span>
  );
}

export function SampleTag() {
  return (
    <span className="rounded border border-warning/40 bg-warning/15 px-1 py-px font-mono text-[10px] font-semibold uppercase tracking-wide text-foreground">
      Sample
    </span>
  );
}
