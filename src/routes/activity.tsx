import { createFileRoute } from "@tanstack/react-router";
import { useState } from "react";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { EmptyState, ErrorState, LoadingRows, PageHeader, Panel } from "@/components/app/states";
import { SampleTag, StatusBadge } from "@/components/app/status-badge";
import { FilterSelect, fmtDate, Pager, SearchInput } from "@/components/app/table-kit";
import { useAudit } from "@/hooks/use-api";

export const Route = createFileRoute("/activity")({
  head: () => ({
    meta: [
      { title: "Activity Logs · RecruitFlow" },
      { name: "description", content: "Searchable audit trail of user and system actions." },
      { property: "og:title", content: "Activity Logs · RecruitFlow" },
      { property: "og:description", content: "Searchable audit trail of user and system actions." },
    ],
  }),
  component: ActivityPage,
});

function ActivityPage() {
  const [search, setSearch] = useState("");
  const [result, setResult] = useState("all");
  const [page, setPage] = useState(1);
  const { data, isLoading, error, refetch } = useAudit({ page, pageSize: 15, search, result, sortBy: "at", sortDir: "desc" });
  return (
    <>
      <PageHeader title="Activity Logs" description="Audit events recorded by the backend." />
      <Panel>
        <div className="flex flex-col gap-2 border-b p-3 sm:flex-row">
          <SearchInput value={search} onChange={(v) => { setSearch(v); setPage(1); }} placeholder="Actor, action, related ID, request ID" />
          <FilterSelect label="Results" value={result} onChange={(v) => { setResult(v); setPage(1); }} options={[{ value: "success", label: "Success" }, { value: "failure", label: "Failure" }, { value: "info", label: "Info" }]} />
        </div>
        {error ? <ErrorState error={error} onRetry={refetch} /> : isLoading ? <LoadingRows /> : !data?.items.length ? <EmptyState title="No audit events" /> : (
          <>
            <Table>
              <TableHeader><TableRow><TableHead>Timestamp</TableHead><TableHead>Actor</TableHead><TableHead>Action</TableHead><TableHead>Related</TableHead><TableHead>Result</TableHead><TableHead>Request ID</TableHead></TableRow></TableHeader>
              <TableBody>
                {data.items.map((a) => (
                  <TableRow key={a.id}>
                    <TableCell className="whitespace-nowrap text-xs">{fmtDate(a.at, true)}</TableCell>
                    <TableCell className="text-xs"><div className="flex items-center gap-1.5">{a.actor}{a.isSample && <SampleTag />}</div></TableCell>
                    <TableCell className="font-mono text-xs">{a.action}</TableCell>
                    <TableCell className="font-mono text-xs">{a.related}</TableCell>
                    <TableCell><StatusBadge status={a.result} /></TableCell>
                    <TableCell className="font-mono text-xs text-muted-foreground">{a.requestId}</TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
            <Pager page={page} pageSize={15} total={data.total} onPage={setPage} />
          </>
        )}
      </Panel>
    </>
  );
}
