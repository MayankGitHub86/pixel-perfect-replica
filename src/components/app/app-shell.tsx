import { Link, useNavigate, useRouterState } from "@tanstack/react-router";
import {
  Activity, Bell, ClipboardCheck, FileSpreadsheet, FlaskConical, Inbox, LayoutDashboard, ListChecks,
  Moon, Search, Settings, Sun, Users, Workflow, Briefcase,
} from "lucide-react";
import { useEffect, useState, type ReactNode } from "react";
import {
  Sidebar, SidebarContent, SidebarFooter, SidebarGroup, SidebarGroupContent, SidebarGroupLabel, SidebarHeader,
  SidebarInset, SidebarMenu, SidebarMenuButton, SidebarMenuItem, SidebarProvider, SidebarTrigger,
} from "@/components/ui/sidebar";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Switch } from "@/components/ui/switch";
import { Avatar, AvatarFallback } from "@/components/ui/avatar";
import {
  DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuLabel, DropdownMenuSeparator, DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { useDataMode } from "@/lib/data-mode";
import { isApiConfigured, API_BASE_URL } from "@/lib/api/client";
import { useHealth, useMicrosoftStatus } from "@/hooks/use-api";
import { StatusBadge } from "./status-badge";

const NAV = [
  { title: "Overview", url: "/", icon: LayoutDashboard },
  { title: "Candidates", url: "/candidates", icon: Users },
  { title: "Applications", url: "/applications", icon: Briefcase },
  { title: "Email Inbox", url: "/emails", icon: Inbox },
  { title: "Processing Queue", url: "/jobs", icon: Workflow },
  { title: "Excel Sync", url: "/excel-sync", icon: FileSpreadsheet },
  { title: "Review Required", url: "/review", icon: ClipboardCheck },
  { title: "Activity Logs", url: "/activity", icon: Activity },
  { title: "Settings", url: "/settings", icon: Settings },
] as const;

function AppSidebar() {
  const path = useRouterState({ select: (r) => r.location.pathname });
  const isActive = (u: string) => (u === "/" ? path === "/" : path.startsWith(u));
  const { mode, setMode } = useDataMode();
  return (
    <Sidebar collapsible="icon">
      <SidebarHeader className="px-3 py-3">
        <div className="flex items-center gap-2.5">
          <div className="flex size-8 shrink-0 items-center justify-center rounded-md bg-primary text-primary-foreground">
            <ListChecks className="size-4" />
          </div>
          <div className="leading-tight group-data-[collapsible=icon]:hidden">
            <p className="text-sm font-semibold tracking-tight">RecruitFlow</p>
            <p className="text-[11px] text-muted-foreground">Recruitment Automation</p>
          </div>
        </div>
      </SidebarHeader>
      <SidebarContent>
        <SidebarGroup>
          <SidebarGroupLabel>Workspace</SidebarGroupLabel>
          <SidebarGroupContent>
            <SidebarMenu>
              {NAV.map((n) => (
                <SidebarMenuItem key={n.url}>
                  <SidebarMenuButton asChild isActive={isActive(n.url)} tooltip={n.title}>
                    <Link to={n.url}>
                      <n.icon />
                      <span>{n.title}</span>
                    </Link>
                  </SidebarMenuButton>
                </SidebarMenuItem>
              ))}
            </SidebarMenu>
          </SidebarGroupContent>
        </SidebarGroup>
      </SidebarContent>
      <SidebarFooter className="group-data-[collapsible=icon]:hidden">
        <label className="flex items-center justify-between gap-2 rounded-md border bg-card p-2.5 text-xs">
          <span className="flex items-center gap-1.5">
            <FlaskConical className="size-3.5 text-warning" /> Sample data
          </span>
          <Switch checked={mode === "sample"} onCheckedChange={(v) => setMode(v ? "sample" : "live")} aria-label="Toggle synthetic sample data" />
        </label>
      </SidebarFooter>
    </Sidebar>
  );
}

function ThemeToggle() {
  const [dark, setDark] = useState(false);
  useEffect(() => setDark(localStorage.getItem("recruitflow:theme") === "dark"), []);
  useEffect(() => {
    document.documentElement.classList.toggle("dark", dark);
    localStorage.setItem("recruitflow:theme", dark ? "dark" : "light");
  }, [dark]);
  return (
    <Button variant="ghost" size="icon" onClick={() => setDark((d) => !d)} aria-label="Toggle theme">
      {dark ? <Sun /> : <Moon />}
    </Button>
  );
}

function ConnectionPill() {
  const { mode } = useDataMode();
  const health = useHealth();
  const ms = useMicrosoftStatus();
  if (mode === "sample") return <StatusBadge status="attention_required" className="hidden md:inline-flex" />;
  if (!isApiConfigured()) return <span className="hidden text-xs text-muted-foreground md:inline">Backend not configured</span>;
  return (
    <div className="hidden items-center gap-2 md:flex">
      <span className="text-xs text-muted-foreground">API</span>
      <StatusBadge status={health.isSuccess ? "connected" : health.isLoading ? "connecting" : "error"} />
      <span className="text-xs text-muted-foreground">Outlook</span>
      <StatusBadge status={ms.data?.state ?? (ms.isError ? "error" : "connecting")} />
    </div>
  );
}

function TopBar() {
  const navigate = useNavigate();
  const [q, setQ] = useState("");
  return (
    <header className="sticky top-0 z-20 flex h-14 items-center gap-2 border-b bg-background/85 px-3 backdrop-blur md:px-5">
      <SidebarTrigger />
      <form
        className="relative ml-1 max-w-md flex-1"
        onSubmit={(e) => {
          e.preventDefault();
          navigate({ to: "/candidates", search: { q } });
        }}
      >
        <Search className="absolute left-2.5 top-1/2 size-4 -translate-y-1/2 text-muted-foreground" />
        <Input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search candidates by name, email, skill, ID…" className="h-9 bg-card pl-8" aria-label="Global candidate search" />
      </form>
      <div className="ml-auto flex items-center gap-1">
        <ConnectionPill />
        <Popover>
          <PopoverTrigger asChild>
            <Button variant="ghost" size="icon" aria-label="Notifications">
              <Bell />
            </Button>
          </PopoverTrigger>
          <PopoverContent align="end" className="w-72 text-sm">
            <p className="font-medium">Notifications</p>
            <p className="mt-1 text-muted-foreground">Real-time notifications will arrive from the backend once it is connected.</p>
          </PopoverContent>
        </Popover>
        <ThemeToggle />
        <DropdownMenu>
          <DropdownMenuTrigger asChild>
            <Button variant="ghost" size="icon" className="rounded-full" aria-label="Profile menu">
              <Avatar className="size-7">
                <AvatarFallback className="bg-accent text-[11px] text-accent-foreground">MP</AvatarFallback>
              </Avatar>
            </Button>
          </DropdownMenuTrigger>
          <DropdownMenuContent align="end" className="w-56">
            <DropdownMenuLabel className="font-normal">
              <p className="text-sm font-medium">Recruiter</p>
              <p className="text-xs text-muted-foreground">Signed-in profile comes from the backend</p>
            </DropdownMenuLabel>
            <DropdownMenuSeparator />
            <DropdownMenuItem asChild>
              <Link to="/settings">Settings</Link>
            </DropdownMenuItem>
            <DropdownMenuItem disabled>Sign out (backend not connected)</DropdownMenuItem>
          </DropdownMenuContent>
        </DropdownMenu>
      </div>
    </header>
  );
}

function ModeBanner() {
  const { mode, setMode } = useDataMode();
  if (mode === "sample")
    return (
      <div className="flex flex-wrap items-center gap-2 border-b border-warning/30 bg-warning/10 px-5 py-2 text-xs">
        <FlaskConical className="size-3.5" />
        <span><strong>Synthetic sample data.</strong> No real candidates, mailbox, or workbook. Backend actions are not simulated.</span>
        <button className="ml-auto underline underline-offset-2" onClick={() => setMode("live")}>Switch to live backend</button>
      </div>
    );
  if (!isApiConfigured())
    return (
      <div className="border-b bg-muted px-5 py-2 text-xs text-muted-foreground">
        Live mode · Set <code className="font-mono">VITE_API_BASE_URL</code> to connect the RecruitFlow backend.
      </div>
    );
  return <div className="border-b bg-muted px-5 py-1.5 font-mono text-[11px] text-muted-foreground">Live · {API_BASE_URL}</div>;
}

export function AppShell({ children }: { children: ReactNode }) {
  return (
    <SidebarProvider>
      <div className="flex min-h-screen w-full">
        <AppSidebar />
        <SidebarInset className="min-w-0">
          <TopBar />
          <ModeBanner />
          <main className="mx-auto w-full max-w-[1400px] flex-1 px-4 py-6 md:px-6">{children}</main>
        </SidebarInset>
      </div>
    </SidebarProvider>
  );
}
