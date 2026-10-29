import { createFileRoute } from "@tanstack/react-router";
import { LogOut, Mail, RefreshCw, ShieldCheck } from "lucide-react";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Switch } from "@/components/ui/switch";
import { ErrorState, Field, LoadingRows, PageHeader, Panel } from "@/components/app/states";
import { StatusBadge } from "@/components/app/status-badge";
import { fmtDate } from "@/components/app/table-kit";
import { useExcelStatus, useHealth, useMicrosoftStatus } from "@/hooks/use-api";
import { API_BASE_URL, isApiConfigured } from "@/lib/api/client";
import { useDataMode } from "@/lib/data-mode";

export const Route = createFileRoute("/settings")({
  head: () => ({
    meta: [
      { title: "Settings · RecruitFlow" },
      { name: "description", content: "Outlook, Excel workbook, processing, and security settings." },
      { property: "og:title", content: "Settings · RecruitFlow" },
      { property: "og:description", content: "Outlook, Excel workbook, processing, and security settings." },
    ],
  }),
  component: SettingsPage,
});

const ReadOnly = () => <span className="ml-2 rounded border px-1.5 py-px text-[10px] font-normal uppercase tracking-wide text-muted-foreground">Read-only</span>;

function SettingsPage() {
  const { mode, setMode } = useDataMode();
  const ms = useMicrosoftStatus();
  const excel = useExcelStatus();
  const health = useHealth();

  const connectOutlook = () => {
    if (mode === "sample" || !isApiConfigured()) {
      toast.error("Backend required", { description: "Outlook sign-in runs through the backend's Microsoft OAuth flow. Configure VITE_API_BASE_URL and switch to live mode." });
      return;
    }
    // Full-page navigation: Microsoft handles sign-in; this app never sees the password or OTP.
    window.location.href = `${API_BASE_URL}/api/integrations/microsoft/connect`;
  };
  const unavailable = (what: string) => toast.info(`${what} is not yet available`, { description: "The backend has not published this endpoint." });

  return (
    <>
      <PageHeader title="Settings" description="Controls appear only where the backend supports them; everything else is marked read-only." />
      <div className="grid gap-4 lg:grid-cols-2">
        <Panel title="Outlook integration" description="Intended mailbox: mayankpandey5515@outlook.com" className="lg:row-span-2">
          {ms.isError ? <ErrorState error={ms.error} onRetry={ms.refetch} /> : ms.isLoading || !ms.data ? <LoadingRows rows={4} /> : (
            <>
              <dl className="divide-y px-4">
                <Field label="Connection"><StatusBadge status={ms.data.state} /></Field>
                <Field label="Connected account">{ms.data.accountEmail}</Field>
                <Field label="Mailbox permission">{ms.data.mailboxPermission === "unknown" ? null : <StatusBadge status={ms.data.mailboxPermission === "granted" ? "success" : "attention_required"} />}</Field>
                <Field label="Last mailbox check">{ms.data.lastMailboxCheck ? fmtDate(ms.data.lastMailboxCheck, true) : null}</Field>
                <Field label="Subscription">{ms.data.subscriptionStatus === "none" ? "None" : <StatusBadge status={ms.data.subscriptionStatus === "active" ? "connected" : "expired"} />}</Field>
                <Field label="Subscription expiry">{ms.data.subscriptionExpiresAt ? fmtDate(ms.data.subscriptionExpiresAt, true) : null}</Field>
                <Field label="Last reconciliation">{ms.data.lastReconciliation ? fmtDate(ms.data.lastReconciliation, true) : null}</Field>
              </dl>
              <div className="flex flex-wrap gap-2 border-t p-4">
                {ms.data.state === "connected" ? (
                  <>
                    <Button size="sm" variant="outline" onClick={connectOutlook}><RefreshCw /> Reconnect</Button>
                    <Button size="sm" variant="outline" onClick={() => unavailable("Disconnect")}><LogOut /> Disconnect</Button>
                  </>
                ) : (
                  <Button size="sm" onClick={connectOutlook}><Mail /> Connect Outlook</Button>
                )}
              </div>
              <p className="border-t px-4 py-3 text-xs text-muted-foreground">You'll sign in on Microsoft's own page. RecruitFlow never asks for your Microsoft password or one-time code, and tokens stay on the backend.</p>
            </>
          )}
        </Panel>

        <Panel title="Backend connection">
          <dl className="divide-y px-4">
            <Field label="API base URL">{isApiConfigured() ? <span className="font-mono text-xs">{API_BASE_URL}</span> : "Not configured"}</Field>
            <Field label="Health">{mode === "sample" ? "Not checked in sample mode" : health.isSuccess ? <StatusBadge status="connected" /> : health.isLoading ? "Checking…" : <StatusBadge status="error" />}</Field>
            <Field label="Synthetic sample data"><Switch checked={mode === "sample"} onCheckedChange={(v) => setMode(v ? "sample" : "live")} aria-label="Use sample data" /></Field>
          </dl>
        </Panel>

        <Panel title="Excel workbook">
          <dl className="divide-y px-4">
            <Field label="Workbook">{excel.data?.workbookName ?? "Candidate_Master.xlsx"}</Field>
            <Field label="State">{excel.data ? <StatusBadge status={excel.data.state} /> : excel.isError ? <StatusBadge status="error" /> : "…"}</Field>
            <Field label="Storage provider">{excel.data?.storageProvider}</Field>
          </dl>
        </Panel>

        <Panel title="Processing preferences">
          <dl className="divide-y px-4">
            <Field label={"Supported attachments"}><span>PDF, DOCX<ReadOnly /></span></Field>
            <Field label="OCR for scanned PDFs"><span>Enabled<ReadOnly /></span></Field>
            <Field label="Maximum file size"><span>Set by backend<ReadOnly /></span></Field>
            <Field label="Automatic hiring decisions"><span>Never<ReadOnly /></span></Field>
          </dl>
        </Panel>

        <Panel title="Data retention & security">
          <dl className="divide-y px-4">
            <Field label="Resume retention"><span>Set by backend<ReadOnly /></span></Field>
            <Field label="Resume access">Authorized backend downloads only</Field>
            <Field label="Tokens in browser">None stored</Field>
          </dl>
          <p className="flex items-start gap-2 border-t px-4 py-3 text-xs text-muted-foreground"><ShieldCheck className="mt-0.5 size-3.5 shrink-0" /> Screens in this app are a convenience. Every privileged action is authorized again by the backend.</p>
        </Panel>

        <Panel title="Account & profile">
          <dl className="divide-y px-4">
            <Field label="Profile">Provided by backend session</Field>
            <Field label="Role">{null}</Field>
          </dl>
        </Panel>
      </div>
    </>
  );
}
