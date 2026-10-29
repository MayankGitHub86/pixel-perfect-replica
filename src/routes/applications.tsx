import { createFileRoute, Link } from "@tanstack/react-router";
import { useState } from "react";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { EmptyState, ErrorState, LoadingRows, PageHeader, Panel, Unknown } from "@/components/app/states";
import { SampleTag, STATUS_LABEL } from "@/components/app/status-badge";
import { FilterSelect, fmtDate, Pager, SearchInput, SortHead, type SortState } from "@/components/app/table-kit";
import { useApplications, useUpdateApplicationStatus } from "@/hooks/use-api";
import type { ApplicationStatus } from "@/lib/api/types";

export const Route = createFileRoute("/applications")({
  head: () => ({
    meta: [
      { title: "Applications · RecruitFlow" },
      { name: "description", content: "Track each application's status from New to Hired." },
      { property: "og:title", content: "Applications · RecruitFlow" },
      { property: "og:description", content: "Track each application's status from New to Hired." },
    ],
  }),
  component: ApplicationsPage,
});

const STATUSES: ApplicationStatus[] = ["new", "under_review", "shortlisted", "interview", "rejected", "hired"];

function ApplicationsPage() {
  const [search, setSearch] = useState("");
  const [status, setStatus] = useState("all");
  const [page, setPage] = useState(1);
  const [sort, setSort] = useState<SortState>({ sortBy: "receivedAt", sortDir: "desc" });
  const { data, isLoading, error, refetch } = useApplications({ page, pageSize: 12, search, status, ...sort });
  const mut = useUpdateApplicationStatus();

  return (
    <>
      <PageHeader title="Applications" description="Status changes are made by recruiters. RecruitFlow never rejects or hires automatically." />
      <Panel>
        <div className="flex flex-col gap-2 border-b p-3 sm:flex-row">
          <SearchInput value={search} onChange={(v) => { setSearch(v); setPage(1); }} placeholder="Candidate, email, or role" />
          <FilterSelect label="Statuses" value={status} onChange={(v) => { setStatus(v); setPage(1); }} options={STATUSES.map((s) => ({ value: s, label: STATUS_LABEL(s) }))} />
        </div>
        {error ? <ErrorState error={error} onRetry={refetch} /> : isLoading ? <LoadingRows /> : !data?.items.length ? <EmptyState title="No applications" /> : (
          <>
            <Table>
              <TableHeader><TableRow>
                <SortHead col="candidateName" label="Candidate" sort={sort} onSort={setSort} />
                <SortHead col="role" label="Applied role" sort={sort} onSort={setSort} />
                <TableHead>Source email</TableHead>
                <SortHead col="receivedAt" label="Received" sort={sort} onSort={setSort} />
                <TableHead>Status</TableHead>
                <TableHead>Recruiter</TableHead>
                <SortHead col="updatedAt" label="Updated" sort={sort} onSort={setSort} />
              </TableRow></TableHeader>
              <TableBody>
                {data.items.map((a) => (
                  <TableRow key={a.id}>
                    <TableCell>
                      <Link to="/candidates/$id" params={{ id: a.candidateId }} className="font-medium hover:underline">{a.candidateName ?? <Unknown />}</Link>
                      <div className="flex items-center gap-1.5 font-mono text-[11px] text-muted-foreground">{a.id}{a.isSample && <SampleTag />}</div>
                    </TableCell>
                    <TableCell className="text-sm">{a.role}</TableCell>
                    <TableCell className="text-xs text-muted-foreground">{a.sourceEmail}</TableCell>
                    <TableCell className="whitespace-nowrap text-xs">{fmtDate(a.receivedAt)}</TableCell>
                    <TableCell>
                      <Select value={a.status} onValueChange={(v) => mut.mutate({ id: a.id, status: v as ApplicationStatus })}>
                        <SelectTrigger className="h-8 w-36 text-xs" aria-label={`Status for ${a.id}`}><SelectValue /></SelectTrigger>
                        <SelectContent>{STATUSES.map((s) => <SelectItem key={s} value={s}>{STATUS_LABEL(s)}</SelectItem>)}</SelectContent>
                      </Select>
                    </TableCell>
                    <TableCell className="text-xs">{a.recruiter ?? <span className="text-muted-foreground">Unassigned</span>}</TableCell>
                    <TableCell className="whitespace-nowrap text-xs">{fmtDate(a.updatedAt)}</TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
            <Pager page={page} pageSize={12} total={data.total} onPage={setPage} />
          </>
        )}
      </Panel>
    </>
  );
}
