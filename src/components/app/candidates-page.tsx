import { Link } from "@tanstack/react-router";
import { Download, Eye, Pencil } from "lucide-react";
import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { EmptyState, ErrorState, LoadingRows, PageHeader, Panel, Unknown } from "./states";
import { SampleTag, StatusBadge } from "./status-badge";
import { downloadText, FilterSelect, fmtDate, Pager, SearchInput, SortHead, toCsv, type SortState } from "./table-kit";
import { CandidateEditDialog } from "./candidate-edit-dialog";
import { useCandidates } from "@/hooks/use-api";
import type { Candidate } from "@/lib/api/types";

const ROLES = ["Frontend Engineer", "Backend Engineer", "Data Analyst", "ML Engineer", "Product Designer", "QA Engineer"];
const DEGREES = ["B.Tech CSE", "B.E. ECE", "MCA", "M.Tech AI", "B.Sc Statistics", "MBA"];
const opt = (xs: string[]) => xs.map((x) => ({ value: x, label: x }));

export function CandidatesPage({ initialSearch }: { initialSearch: string }) {
  const [search, setSearch] = useState(initialSearch);
  const [page, setPage] = useState(1);
  const [sort, setSort] = useState<SortState>({ sortBy: "receivedAt", sortDir: "desc" });
  const [f, setF] = useState({ since: "all", appliedRole: "all", degree: "all", processingStatus: "all", interviewStatus: "all" });
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [editing, setEditing] = useState<Candidate | null>(null);
  const q = { page, pageSize: 10, search, ...sort, ...f };
  const { data, isLoading, error, refetch, isFetching } = useCandidates(q);
  const setFilter = (k: keyof typeof f) => (v: string) => { setF((s) => ({ ...s, [k]: v })); setPage(1); };
  const items = data?.items ?? [];
  const allSel = items.length > 0 && items.every((c) => selected.has(c.id));

  const exportRows = () => {
    const rows = (selected.size ? items.filter((c) => selected.has(c.id)) : items).map(({ reviewFields: _r, resumeRef: _ref, ...c }) => c);
    downloadText("candidates.csv", toCsv(rows as unknown as Record<string, unknown>[]));
  };

  return (
    <>
      <PageHeader
        title="Candidates"
        description="Records extracted from CV attachments. Missing fields stay blank until evidence exists."
        actions={<Button variant="outline" size="sm" onClick={exportRows} disabled={!items.length}><Download /> Export {selected.size ? `${selected.size} selected` : "page"} (CSV)</Button>}
      />
      <Panel>
        <div className="flex flex-col gap-2 border-b p-3 lg:flex-row lg:flex-wrap lg:items-center">
          <SearchInput value={search} onChange={(v) => { setSearch(v); setPage(1); }} placeholder="Name, email, skill, or ID" />
          <FilterSelect label="Dates" value={f.since} onChange={setFilter("since")} options={[{ value: "7", label: "Last 7 days" }, { value: "14", label: "Last 14 days" }, { value: "30", label: "Last 30 days" }]} />
          <FilterSelect label="Roles" value={f.appliedRole} onChange={setFilter("appliedRole")} options={opt(ROLES)} />
          <FilterSelect label="Degrees" value={f.degree} onChange={setFilter("degree")} options={opt(DEGREES)} />
          <FilterSelect label="Processing" value={f.processingStatus} onChange={setFilter("processingStatus")} options={[["new","New"],["processing","Processing"],["processed","Processed"],["needs_review","Needs review"],["failed","Failed"]].map(([value,label])=>({value: value!, label: label!}))} />
          <FilterSelect label="Interview" value={f.interviewStatus} onChange={setFilter("interviewStatus")} options={[["not_scheduled","Not scheduled"],["scheduled","Scheduled"],["completed","Completed"],["no_show","No show"]].map(([value,label])=>({value: value!, label: label!}))} />
        </div>
        {error ? <ErrorState error={error} onRetry={refetch} /> : isLoading ? <LoadingRows /> : !items.length ? (
          <EmptyState title="No candidates match" description="Adjust search or filters. New CVs appear after the backend processes incoming email." />
        ) : (
          <div className={isFetching ? "opacity-70 transition-opacity" : ""}>
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead className="w-8"><Checkbox checked={allSel} onCheckedChange={(v) => setSelected(v ? new Set(items.map((c) => c.id)) : new Set())} aria-label="Select all rows" /></TableHead>
                  <SortHead col="id" label="ID" sort={sort} onSort={setSort} />
                  <SortHead col="name" label="Candidate" sort={sort} onSort={setSort} />
                  <TableHead>Phone</TableHead>
                  <SortHead col="college" label="College" sort={sort} onSort={setSort} />
                  <SortHead col="degree" label="Degree" sort={sort} onSort={setSort} />
                  <SortHead col="graduationYear" label="Grad" sort={sort} onSort={setSort} />
                  <TableHead>Skills</TableHead>
                  <SortHead col="experienceYears" label="Exp" sort={sort} onSort={setSort} />
                  <SortHead col="appliedRole" label="Role" sort={sort} onSort={setSort} />
                  <SortHead col="receivedAt" label="Received" sort={sort} onSort={setSort} />
                  <TableHead>Processing</TableHead>
                  <TableHead>Interview</TableHead>
                  <SortHead col="updatedAt" label="Updated" sort={sort} onSort={setSort} />
                  <TableHead className="text-right">Actions</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {items.map((c) => (
                  <TableRow key={c.id} data-state={selected.has(c.id) ? "selected" : undefined}>
                    <TableCell><Checkbox checked={selected.has(c.id)} onCheckedChange={(v) => setSelected((s) => { const n = new Set(s); v ? n.add(c.id) : n.delete(c.id); return n; })} aria-label={`Select ${c.id}`} /></TableCell>
                    <TableCell className="whitespace-nowrap font-mono text-xs"><div className="flex items-center gap-1.5">{c.id}{c.isSample && <SampleTag />}</div></TableCell>
                    <TableCell>
                      <p className="font-medium">{c.name ?? <Unknown />}</p>
                      <p className="text-xs text-muted-foreground">{c.email ?? "—"}</p>
                    </TableCell>
                    <TableCell className="whitespace-nowrap text-xs">{c.phone ?? <Unknown />}</TableCell>
                    <TableCell className="max-w-40 truncate text-xs">{c.college ?? <Unknown />}</TableCell>
                    <TableCell className="whitespace-nowrap text-xs">{c.degree ?? <Unknown />}</TableCell>
                    <TableCell className="text-xs tabular-nums">{c.graduationYear ?? <Unknown />}</TableCell>
                    <TableCell><div className="flex items-center gap-1 whitespace-nowrap">{c.skills.slice(0, 3).map((s) => <span key={s} className="rounded bg-muted px-1.5 py-0.5 text-[11px]">{s}</span>)}{c.skills.length > 3 && <span className="text-[11px] text-muted-foreground">+{c.skills.length - 3}</span>}</div></TableCell>
                    <TableCell className="text-xs tabular-nums">{c.experienceYears != null ? `${c.experienceYears}y` : <Unknown />}</TableCell>
                    <TableCell className="whitespace-nowrap text-xs">{c.appliedRole ?? <Unknown />}</TableCell>
                    <TableCell className="whitespace-nowrap text-xs">{fmtDate(c.receivedAt)}</TableCell>
                    <TableCell><StatusBadge status={c.processingStatus} /></TableCell>
                    <TableCell><StatusBadge status={c.interviewStatus} /></TableCell>
                    <TableCell className="whitespace-nowrap text-xs">{fmtDate(c.updatedAt)}</TableCell>
                    <TableCell className="text-right">
                      <div className="flex justify-end gap-1">
                        <Button variant="ghost" size="icon" className="size-7" asChild aria-label={`View ${c.id}`}>
                          <Link to="/candidates/$id" params={{ id: c.id }}><Eye /></Link>
                        </Button>
                        <Button variant="ghost" size="icon" className="size-7" onClick={() => setEditing(c)} aria-label={`Edit ${c.id}`}><Pencil /></Button>
                      </div>
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
            <Pager page={page} pageSize={10} total={data?.total ?? 0} onPage={setPage} />
          </div>
        )}
      </Panel>
      <CandidateEditDialog candidate={editing} onClose={() => setEditing(null)} />
    </>
  );
}
