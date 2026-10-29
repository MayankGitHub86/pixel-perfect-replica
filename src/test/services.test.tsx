import { describe, expect, it } from "vitest";
import { createServices } from "@/lib/api/services";

describe("data layer safety rules", () => {
  const sample = createServices("sample");

  it("sample mode refuses job retries instead of faking success", async () => {
    await expect(sample.retryJob("JOB-9001")).rejects.toMatchObject({ status: 501 });
  });

  it("sample mode refuses Excel sync instead of faking success", async () => {
    await expect(sample.excelRun()).rejects.toMatchObject({ status: 501 });
  });

  it("sample workbook is never reported as connected and has no URL", async () => {
    const s = await sample.excelStatus();
    expect(s.state).toBe("not_connected");
    expect(s.workbookUrl).toBeNull();
  });

  it("every sample candidate is labelled synthetic and uses example.com", async () => {
    const res = await sample.listCandidates({ pageSize: 1000 });
    expect(res.items.every((c) => c.isSample && (c.email ?? "").endsWith("@example.com"))).toBe(true);
  });

  it("live mode without VITE_API_BASE_URL errors rather than returning data", async () => {
    await expect(createServices("live").listCandidates({})).rejects.toMatchObject({ status: 501 });
  });
});
