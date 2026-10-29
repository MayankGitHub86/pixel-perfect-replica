<!-- LOVABLE:BEGIN -->
> [!IMPORTANT]
> This project is connected to [Lovable](https://lovable.dev). Avoid rewriting
> published git history — force pushing, or rebasing/amending/squashing commits
> that are already pushed — as it rewrites history on Lovable's side and the
> user will likely lose their project history.
>
> Commits you push to the connected branch sync back to Lovable and show up in
> the editor, so keep the branch in a working state.
<!-- LOVABLE:END -->

## Architecture rules
- Pages read/write data only through `src/hooks/use-api.ts` → `src/lib/api/services.ts`; why: sample vs live backend swaps without page changes.
- Sample mode must never simulate backend side effects (retry, Excel sync, OAuth); why: the product must not report fake success.
- Backend URL comes only from `VITE_API_BASE_URL`; no secrets or tokens in the browser; why: auth lives in the separate Express backend.
- Proposed backend contract is documented in `docs/INTEGRATION.md`; keep it in sync with `src/lib/api/types.ts`.
