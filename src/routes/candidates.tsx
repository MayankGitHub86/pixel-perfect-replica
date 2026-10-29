import { createFileRoute, Outlet, useMatch } from "@tanstack/react-router";
import { CandidatesPage } from "@/components/app/candidates-page";

type Search = { q?: string };

export const Route = createFileRoute("/candidates")({
  validateSearch: (s: Record<string, unknown>): Search => (typeof s.q === "string" && s.q ? { q: s.q } : {}),
  head: () => ({
    meta: [
      { title: "Candidates · RecruitFlow" },
      { name: "description", content: "Search, filter, and review candidate records extracted from CVs." },
      { property: "og:title", content: "Candidates · RecruitFlow" },
      { property: "og:description", content: "Search, filter, and review candidate records extracted from CVs." },
    ],
  }),
  component: CandidatesLayout,
});

function CandidatesLayout() {
  const child = useMatch({ from: "/candidates/$id", shouldThrow: false });
  const { q } = Route.useSearch();
  return child ? <Outlet /> : <CandidatesPage initialSearch={q ?? ""} />;
}
