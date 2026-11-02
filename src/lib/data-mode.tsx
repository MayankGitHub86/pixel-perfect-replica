import { createContext, useContext, useEffect, useMemo, useState, type ReactNode } from "react";
import { useQueryClient } from "@tanstack/react-query";
import { createServices, type DataMode, type Services } from "./api/services";
import { isApiConfigured } from "./api/client";

const Ctx = createContext<{ mode: DataMode; setMode: (m: DataMode) => void; api: Services } | null>(null);
const KEY = "recruitflow:data-mode"; // UI preference only — never tokens or secrets

export function DataModeProvider({ children }: { children: ReactNode }) {
  const [mode, setModeState] = useState<DataMode>(() => {
    if (typeof window !== "undefined") {
      const saved = localStorage.getItem(KEY);
      if (saved === "live" || saved === "sample") return saved;
    }
    return isApiConfigured() ? "live" : "sample";
  });
  const qc = useQueryClient();
  useEffect(() => {
    const saved = localStorage.getItem(KEY);
    if (saved === "live" || saved === "sample") {
      setModeState(saved);
    } else if (isApiConfigured()) {
      setModeState("live");
    }
  }, []);
  const setMode = (m: DataMode) => {
    localStorage.setItem(KEY, m);
    setModeState(m);
    qc.clear();
  };
  const api = useMemo(() => createServices(mode), [mode]);
  return <Ctx.Provider value={{ mode, setMode, api }}>{children}</Ctx.Provider>;
}

export function useDataMode() {
  const v = useContext(Ctx);
  if (!v) throw new Error("useDataMode must be used inside DataModeProvider");
  return v;
}
