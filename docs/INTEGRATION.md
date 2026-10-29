# RecruitFlow — Backend integration (PROPOSED contract)

These endpoints are **proposed**. The frontend is ready for them; the Node.js/Express backend must implement them.

Configure: `VITE_API_BASE_URL=https://api.example.com` (public URL only — never secrets). Requests use `credentials: "include"` (session cookie issued by the backend). The backend should return `x-request-id` headers and JSON errors `{ "message": "..." }`.

| Method | Path | Used by | Notes |
|---|---|---|---|
| GET | /api/health | top bar, Settings | `{ status: "ok" \| "degraded" }` |
| GET | /api/candidates | Candidates, Overview, Review | query: page, pageSize, search, sortBy, sortDir, since (days), appliedRole, degree, processingStatus, interviewStatus → `Paginated<Candidate>` |
| GET | /api/candidates/:id | Candidate detail | `CandidateDetail` |
| PATCH | /api/candidates/:id | Edit dialog, Review | body `CandidatePatch` |
| GET | /api/applications | Applications | query: page, pageSize, search, status, sortBy, sortDir |
| PATCH | /api/applications/:id/status | Applications | body `{ status }` |
| GET | /api/emails | Email Inbox, Overview | query: page, pageSize, status, search |
| GET | /api/jobs | Processing Queue, Overview | query: page, pageSize, status, type, search |
| POST | /api/jobs/:id/retry | Processing Queue | returns updated job |
| GET | /api/excel-sync/status | Excel Sync, Overview | `ExcelSyncStatus` |
| POST | /api/excel-sync/test | Excel Sync | `{ ok, message }` |
| POST | /api/excel-sync/run | Excel Sync | `{ queued }` |
| GET | /api/audit-logs | Activity Logs | query: page, pageSize, search, result |
| GET | /api/integrations/microsoft/status | Settings, top bar | `MicrosoftStatus` |
| GET | /api/integrations/microsoft/connect | Settings | browser navigates here; backend redirects to Microsoft OAuth |

Types live in `src/lib/api/types.ts`. Error codes 401/403/404/429/500 are shown as distinct error states.

## Steps to connect
1. Implement the endpoints above with CORS allowing the app origin and credentials.
2. Set `VITE_API_BASE_URL`.
3. Turn off "Sample data" in the sidebar.

## Known limitations
- Overview metrics are computed from list endpoints (pageSize=1000); a `/api/stats` endpoint would scale better.
- Review queue in live mode derives from `processingStatus=needs_review`; per-field evidence needs a dedicated endpoint.
- "Connect workbook" and "Disconnect Outlook" have no proposed endpoint yet and are marked unavailable.
