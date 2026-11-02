// Typed HTTP client for the RecruitFlow backend. The only configuration is the
// public VITE_API_BASE_URL; no secrets ever live in the browser.

export const API_BASE_URL: string = ((import.meta.env["VITE_API_BASE_URL"] as string | undefined) ?? "").replace(/\/$/, "");

export const isApiConfigured = () => API_BASE_URL.length > 0;

export class ApiError extends Error {
  constructor(
    public status: number,
    message: string,
    public requestId?: string,
  ) {
    super(message);
    this.name = "ApiError";
  }
}

export function describeApiError(err: unknown): { title: string; detail: string } {
  if (err instanceof ApiError) {
    const map: Record<number, string> = {
      0: "Backend unreachable",
      401: "You are signed out",
      403: "You don't have permission",
      404: "Not found",
      429: "Too many requests",
      500: "Server error",
      501: "Backend not configured",
    };
    return {
      title: map[err.status] ?? `Request failed (${err.status})`,
      detail: err.message + (err.requestId ? ` · Request ${err.requestId}` : ""),
    };
  }
  return { title: "Unexpected error", detail: err instanceof Error ? err.message : String(err) };
}

type Query = Record<string, string | number | undefined>;

export async function apiRequest<T>(
  path: string,
  opts: { method?: string; query?: Query; body?: unknown } = {},
): Promise<T> {
  if (!isApiConfigured()) {
    throw new ApiError(501, "VITE_API_BASE_URL is not set. Configure the backend URL to use live data.");
  }
  const url = new URL(API_BASE_URL + path);
  for (const [k, v] of Object.entries(opts.query ?? {})) {
    if (v !== undefined && v !== "" && v !== "all") url.searchParams.set(k, String(v));
  }
  let res: Response;
  try {
    res = await fetch(url, {
      method: opts.method ?? "GET",
      credentials: "include", // session cookie issued by the backend
      ...(opts.body !== undefined ? { headers: { "Content-Type": "application/json" }, body: JSON.stringify(opts.body) } : {}),
    });
  } catch {
    throw new ApiError(0, `Could not reach ${API_BASE_URL}.`);
  }
  const requestId = res.headers.get("x-request-id") ?? undefined;
  if (!res.ok) {
    let message = res.statusText;
    try {
      const data = await res.json();
      message = data.message ?? (typeof data.error === "object" ? data.error?.message : data.error) ?? message;
    } catch {
      /* non-JSON error body */
    }
    throw new ApiError(res.status, message || "Request failed", requestId);
  }
  if (res.status === 204) return undefined as T;
  return (await res.json()) as T;
}
