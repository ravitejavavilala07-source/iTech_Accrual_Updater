import { useEffect, useRef, useState } from "react";
import {
  FolderOpen, FileSpreadsheet, Play, StopCircle,
  CheckCircle2, AlertCircle, Settings2,
} from "lucide-react";
import { PayDateEditor, PayDate } from "./PayDateEditor";
import { pushToast } from "./Toast";

const MONTHS = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];

interface Config {
  master_path: string;
  paysheets_folder: string;
  month: string;
  year: number;
  dry_run: boolean;
  enable_accrual: boolean;
  enable_admin_fee: boolean;
  enable_carryforward: boolean;
  pay_dates: PayDate[];
}

const STORAGE_KEY = "itech-accrual-config";

function loadConfig(): Config {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) return { ...defaultConfig(), ...JSON.parse(raw) };
  } catch {
    /* ignore */
  }
  return defaultConfig();
}

function defaultConfig(): Config {
  return {
    master_path: "",
    paysheets_folder: "",
    month: "January",
    year: new Date().getFullYear(),
    dry_run: true,
    enable_accrual: true,
    enable_admin_fee: true,
    enable_carryforward: false,
    pay_dates: [],
  };
}

interface RunStats {
  updated: number;
  skipped: number;
  errors: number;
  cellsWritten: number;
  elapsedSec: number;
}

type RunPhase = "idle" | "validating" | "processing" | "writing" | "done" | "error";

export function AccrualForm() {
  const [cfg, setCfg] = useState<Config>(defaultConfig);
  const [logs, setLogs] = useState<string[]>([]);
  const [running, setRunning] = useState(false);
  const [progress, setProgress] = useState(0); // 0-100
  const [progressLabel, setProgressLabel] = useState("");
  const [phase, setPhase] = useState<RunPhase>("idle");
  const [stats, setStats] = useState<RunStats | null>(null);
  const [pickerBusy, setPickerBusy] = useState<null | "file" | "folder">(null);
  const logRef = useRef<HTMLDivElement>(null);
  const esRef = useRef<EventSource | null>(null);
  const runStartRef = useRef<number>(0);

  useEffect(() => {
    setCfg(loadConfig());
  }, []);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(cfg));
  }, [cfg]);

  useEffect(() => {
    if (logRef.current) logRef.current.scrollTop = logRef.current.scrollHeight;
  }, [logs]);

  const update = <K extends keyof Config>(k: K, v: Config[K]) =>
    setCfg((c) => ({ ...c, [k]: v }));

  const openNativePicker = async (mode: "file" | "folder") => {
    setPickerBusy(mode);
    try {
      const title =
        mode === "file" ? "Select Master Excel File" : "Select Paysheets Folder";
      const r = await fetch(
        `/api/native-picker?mode=${mode}&title=${encodeURIComponent(title)}`
      );
      if (!r.ok) {
        const err = await r.text();
        pushToast(`Picker error: ${err}`, "error");
        return;
      }
      const data = await r.json();
      if (data.cancelled || !data.path) return;
      if (mode === "file") {
        update("master_path", data.path);
        pushToast("Master file selected", "success");
      } else {
        update("paysheets_folder", data.path);
        pushToast("Paysheets folder selected", "success");
      }
    } catch (e: any) {
      pushToast(`Picker failed: ${String(e.message || e)}`, "error");
    } finally {
      setPickerBusy(null);
    }
  };

  // Parse [N/total] from log to drive progress bar
  const parseProgress = (line: string): { done: number; total: number } | null => {
    const m = line.match(/\[\s*(\d+)\s*\/\s*(\d+)\s*\]/);
    if (!m) return null;
    return { done: parseInt(m[1], 10), total: parseInt(m[2], 10) };
  };

  const run = async () => {
    if (!cfg.master_path) {
      pushToast("Select a master file first", "error");
      return;
    }
    if (!cfg.paysheets_folder) {
      pushToast("Select a paysheets folder first", "error");
      return;
    }
    if (!cfg.dry_run) {
      if (!confirm("Modify the master file (not a dry run)?")) return;
    }

    setLogs([]);
    setProgress(0);
    setProgressLabel("Starting…");
    setPhase("validating");
    setStats(null);
    setRunning(true);
    runStartRef.current = Date.now();

    try {
      const r = await fetch("/api/run", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(cfg),
      });
      if (!r.ok) {
        const err = await r.text();
        setLogs((l) => [...l, `✗ ${err}`]);
        setRunning(false);
        setPhase("error");
        if (r.status === 409) {
          pushToast("Another run is in progress. Wait for it to finish.", "error");
        } else {
          pushToast(`Run failed: ${err.slice(0, 100)}`, "error");
        }
        return;
      }
      const { job_id } = await r.json();
      streamLogs(job_id);
    } catch (e: any) {
      const msg = String(e.message || e);
      setLogs((l) => [...l, `✗ ${msg}`]);
      setRunning(false);
      setPhase("error");
      pushToast(msg, "error");
    }
  };

  // Parse summary stats from log block at end of run
  const extractStats = (allLogs: string[]): RunStats | null => {
    const updatedM = allLogs.find((l) => /✓\s*Updated:\s*(\d+)/i.test(l))?.match(/(\d+)/);
    const skippedM = allLogs.find((l) => /No\s*Match|skipped/i.test(l))?.match(/(\d+)/);
    const errorsM = allLogs.find((l) => /Failed:\s*(\d+)/i.test(l))?.match(/(\d+)/);
    const cellsM = allLogs.find((l) => /(\d+)\s*cells?\s*written/i.test(l))?.match(/(\d+)/);
    if (!updatedM && !cellsM) return null;
    return {
      updated: updatedM ? parseInt(updatedM[1], 10) : 0,
      skipped: skippedM ? parseInt(skippedM[1], 10) : 0,
      errors: errorsM ? parseInt(errorsM[1], 10) : 0,
      cellsWritten: cellsM ? parseInt(cellsM[1], 10) : 0,
      elapsedSec: Math.round((Date.now() - runStartRef.current) / 100) / 10,
    };
  };

  const streamLogs = (id: string) => {
    const es = new EventSource(`/api/logs/${id}`);
    esRef.current = es;
    const collected: string[] = [];
    es.onmessage = (ev) => {
      try {
        const data = JSON.parse(ev.data);
        if (typeof data.line === "string") {
          collected.push(data.line);
          setLogs((l) => [...l, data.line]);
          const prog = parseProgress(data.line);
          if (prog && prog.total > 0) {
            const pct = Math.min(100, Math.round((prog.done / prog.total) * 100));
            setProgress(pct);
            setProgressLabel(`${prog.done}/${prog.total} paysheets`);
            setPhase("processing");
          } else if (data.line.includes("Pre-validation")) {
            setProgressLabel("Validating prior months");
            setPhase("validating");
          } else if (data.line.includes("Opening Excel") || data.line.includes("Writing")) {
            setProgressLabel("Writing to Excel");
            setProgress((p) => Math.max(p, 92));
            setPhase("writing");
          } else if (data.line.includes("Excel saved") || data.line.includes("Saved")) {
            setProgress(100);
            setProgressLabel("Saved");
          }
        }
      } catch {
        /* ignore */
      }
    };
    es.addEventListener("done", () => {
      es.close();
      esRef.current = null;
      setProgress(100);
      setProgressLabel("Complete");
      setPhase("done");
      setStats(extractStats(collected));
      setRunning(false);
      pushToast("Run complete", "success");
    });
    es.onerror = () => {
      es.close();
      esRef.current = null;
      setRunning(false);
      if (phase !== "done") {
        setPhase("error");
        pushToast("Connection lost", "error");
      }
    };
  };

  const stop = () => {
    if (esRef.current) {
      esRef.current.close();
      esRef.current = null;
    }
    setRunning(false);
  };

  return (
    <div className="relative z-10 max-w-5xl mx-auto px-6 py-12 space-y-6">
      {/* Header */}
      <div className="mb-8 anim-in" style={{ animationDelay: "0ms" }}>
        <h1
          className="font-display text-5xl font-normal leading-tight"
          style={{ letterSpacing: "-0.024em" }}
        >
          <span className="text-foreground">Accrual </span>
          <span className="brand-gradient-text">Updater</span>
        </h1>
        <p className="text-foreground/60 text-base mt-2">
          powered by 7t.ai for Smartworks and Itech
        </p>
      </div>

      {/* Wrong-month banner */}
      <WrongMonthBanner cfgMonth={cfg.month} cfgYear={cfg.year} />

      {/* Master file */}
      <div className="anim-in" style={{ animationDelay: "60ms" }}>
        <Section title="Master File">
          <PathInput
            icon={<FileSpreadsheet className="w-4 h-4" />}
            value={cfg.master_path}
            placeholder="Select Excel master file (.xlsx)"
            onBrowse={() => openNativePicker("file")}
            busy={pickerBusy === "file"}
          />
        </Section>
      </div>

      {/* Paysheets */}
      <div className="anim-in" style={{ animationDelay: "120ms" }}>
        <Section title="Paysheets Folder">
          <PathInput
            icon={<FolderOpen className="w-4 h-4" />}
            value={cfg.paysheets_folder}
            placeholder="Select folder containing paysheets (searches subfolders)"
            onBrowse={() => openNativePicker("folder")}
            busy={pickerBusy === "folder"}
          />
        </Section>
      </div>

      {/* Period */}
      <div className="anim-in" style={{ animationDelay: "180ms" }}>
        <Section title="Period">
        <div className="flex items-center gap-4 flex-wrap">
          <label className="text-sm text-foreground/70 flex items-center gap-2">
            Month:
            <select
              value={cfg.month}
              onChange={(e) => update("month", e.target.value)}
              className="liquid-glass rounded-lg px-3 py-2 text-foreground text-sm bg-black/30 outline-none"
            >
              {MONTHS.map((m) => (
                <option key={m} value={m} className="bg-background">
                  {m}
                </option>
              ))}
            </select>
          </label>
          <label className="text-sm text-foreground/70 flex items-center gap-2">
            Year:
            <input
              type="number"
              value={cfg.year}
              min={2020}
              max={2100}
              onChange={(e) => update("year", parseInt(e.target.value) || 2026)}
              className="liquid-glass rounded-lg px-3 py-2 text-foreground text-sm bg-black/30 outline-none w-28"
            />
          </label>
          <label className="text-sm text-foreground/70 flex items-center gap-2">
            Pick from calendar:
            <input
              type="month"
              value={`${cfg.year}-${String(MONTHS.indexOf(cfg.month) + 1).padStart(2, "0")}`}
              onChange={(e) => {
                const [y, m] = e.target.value.split("-").map((v) => parseInt(v, 10));
                if (!isNaN(y)) update("year", y);
                if (!isNaN(m)) update("month", MONTHS[m - 1]);
              }}
              className="liquid-glass rounded-lg px-3 py-2 text-foreground text-sm bg-black/30 outline-none"
            />
          </label>
        </div>
      </Section>
      </div>

      {/* Pay dates */}
      <div className="anim-in" style={{ animationDelay: "240ms" }}>
        <Section title="Pay Date Multipliers">
        <PayDateEditor
          rows={cfg.pay_dates}
          onChange={(rows) => update("pay_dates", rows)}
        />
      </Section>
      </div>

      {/* Options */}
      <div className="anim-in" style={{ animationDelay: "300ms" }}>
        <Section title="Options">
        <div className="flex flex-wrap gap-6">
          <Checkbox
            label="Dry Run"
            checked={cfg.dry_run}
            onChange={(v) => update("dry_run", v)}
          />
          <Checkbox
            label="Accrual (Hours / Billed / Salary Paid)"
            checked={cfg.enable_accrual}
            onChange={(v) => update("enable_accrual", v)}
          />
          <Checkbox
            label="Admin Fee"
            checked={cfg.enable_admin_fee}
            onChange={(v) => update("enable_admin_fee", v)}
          />
          <Checkbox
            label="Carryforward (force on)"
            checked={cfg.enable_carryforward}
            onChange={(v) => update("enable_carryforward", v)}
            title="Carryforward runs automatically in January. Tick this to force it for any other month."
          />
        </div>
      </Section>
      </div>

      {/* Run button + progress */}
      <div className="pt-2 space-y-3 anim-in" style={{ animationDelay: "360ms" }}>
        {!running ? (
          <button
            onClick={run}
            className="liquid-glass rounded-full px-8 py-4 text-foreground font-medium flex items-center gap-2 hover:text-white transition press hover-lift focus-ring group"
          >
            <Play className="w-4 h-4 transition-transform group-hover:translate-x-0.5" />
            Run Update
            {cfg.dry_run && (
              <span className="ml-1 text-[10px] uppercase tracking-wider text-amber-300/80 border border-amber-300/30 rounded-full px-2 py-0.5">
                dry
              </span>
            )}
          </button>
        ) : (
          <button
            onClick={stop}
            className="liquid-glass rounded-full px-8 py-4 text-foreground font-medium flex items-center gap-2 hover:text-red-300 transition press focus-ring"
          >
            <StopCircle className="w-4 h-4" /> Stop
          </button>
        )}

        {(running || progress > 0) && (
          <div className="space-y-3">
            {/* Step pills */}
            <div className="flex items-center gap-2 text-[11px] uppercase tracking-wider">
              <StepPill label="Validate" active={phase === "validating"} done={["processing","writing","done"].includes(phase)} />
              <StepArrow />
              <StepPill label="Process" active={phase === "processing"} done={["writing","done"].includes(phase)} />
              <StepArrow />
              <StepPill label="Write" active={phase === "writing"} done={phase === "done"} />
              <StepArrow />
              <StepPill label="Done" active={false} done={phase === "done"} />
            </div>

            <div className="flex justify-between text-xs text-foreground/70 font-mono">
              <span className="flex items-center gap-1.5">
                {running && <span className="w-1.5 h-1.5 rounded-full bg-indigo-400 pulse-soft" />}
                {progressLabel}
              </span>
              <span className="tabular-nums">{progress}%</span>
            </div>
            <div className="relative h-2 rounded-full bg-black/40 border border-foreground/10 overflow-hidden">
              <div
                className="h-full brand-gradient transition-all duration-500 ease-out"
                style={{ width: `${progress}%` }}
              />
              {running && progress > 0 && progress < 100 && (
                <div
                  className="absolute inset-y-0 left-0 shimmer"
                  style={{ width: `${progress}%` }}
                />
              )}
            </div>
          </div>
        )}

        {/* Summary card on completion */}
        {phase === "done" && stats && (
          <SummaryCard
            stats={stats}
            dryRun={cfg.dry_run}
            onCommit={() => {
              update("dry_run", false);
              // Re-run after toggle settles
              setTimeout(run, 50);
            }}
          />
        )}
        {phase === "error" && !running && (
          <div className="liquid-glass rounded-xl px-4 py-3 flex items-start gap-3 toast-in"
               style={{ background: "rgba(60, 10, 10, 0.5)" }}>
            <AlertCircle className="w-4 h-4 text-red-400 mt-0.5 shrink-0" />
            <div className="text-sm text-foreground/85">
              Run failed. Check log below.
            </div>
          </div>
        )}
      </div>

      {/* Log panel */}
      <div className="anim-in" style={{ animationDelay: "420ms" }}>
        <Section title="Log">
        <div
          ref={logRef}
          className="rounded-xl bg-black/55 border border-foreground/10 p-4 h-80 overflow-y-auto pretty-scroll font-mono text-xs text-foreground/90 whitespace-pre-wrap"
        >
          {logs.length === 0 && (
            <div className="text-foreground/30">
              {running ? "Starting…" : "No output yet. Click Run Update."}
            </div>
          )}
          {logs.map((line, i) => (
            <div key={i} className={logLineClass(line)}>{line || " "}</div>
          ))}
        </div>
      </Section>
      </div>
    </div>
  );
}

function logLineClass(line: string): string {
  if (/^\s*✓|✅|Saved|complete/i.test(line)) return "text-emerald-300/90";
  if (/^\s*⚠|WARN|warning/i.test(line)) return "text-amber-300/90";
  if (/^\s*✗|❌|ERROR|failed/i.test(line)) return "text-red-300/90";
  if (/^={5,}|^STEP\s/i.test(line)) return "text-foreground/70 font-semibold";
  return "text-foreground/85";
}

function StepPill({ label, active, done }: { label: string; active: boolean; done: boolean }) {
  return (
    <span
      className={`px-2.5 py-1 rounded-full border text-[10px] flex items-center gap-1 transition-all duration-300 ${
        done
          ? "bg-emerald-500/10 border-emerald-400/40 text-emerald-300"
          : active
          ? "bg-indigo-500/15 border-indigo-400/50 text-indigo-200"
          : "bg-black/30 border-foreground/10 text-foreground/40"
      }`}
    >
      {done && <CheckCircle2 className="w-3 h-3" />}
      {active && !done && <span className="w-1.5 h-1.5 rounded-full bg-indigo-400 pulse-soft" />}
      {label}
    </span>
  );
}

function StepArrow() {
  return <span className="text-foreground/20 text-xs">→</span>;
}

function SummaryCard({ stats, dryRun, onCommit }: { stats: RunStats; dryRun: boolean; onCommit: () => void }) {
  return (
    <div
      className="liquid-glass rounded-2xl p-5 toast-in space-y-3"
      style={{ background: "rgba(20, 30, 50, 0.5)" }}
    >
      <div className="flex items-center gap-2">
        <CheckCircle2 className="w-5 h-5 text-emerald-400" />
        <h3 className="text-foreground font-semibold">
          {dryRun ? "Dry run complete" : "Run complete"}
        </h3>
        <span className="ml-auto text-xs text-foreground/50 tabular-nums">
          {stats.elapsedSec}s
        </span>
      </div>
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        <Stat label="Updated" value={stats.updated} accent="emerald" />
        <Stat label="Cells written" value={stats.cellsWritten} accent="indigo" />
        <Stat label="Skipped" value={stats.skipped} accent="amber" />
        <Stat label="Errors" value={stats.errors} accent={stats.errors > 0 ? "red" : "neutral"} />
      </div>
      {dryRun && (
        <div className="flex items-center justify-between gap-3 pt-1">
          <div className="text-xs text-amber-300/80 flex items-center gap-1.5">
            <Settings2 className="w-3 h-3" />
            Dry run — no cells written.
          </div>
          <button
            onClick={() => {
              if (confirm("Commit changes to master file? This writes for real.")) onCommit();
            }}
            className="liquid-glass rounded-full px-4 py-2 text-xs font-medium text-foreground hover:text-white transition press hover-lift focus-ring brand-gradient-text"
          >
            Commit changes →
          </button>
        </div>
      )}
    </div>
  );
}

function WrongMonthBanner({ cfgMonth, cfgYear }: { cfgMonth: string; cfgYear: number }) {
  const months = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December",
  ];
  const now = new Date();
  const currentMonth = months[now.getMonth()];
  const currentYear = now.getFullYear();
  const prevMonth = months[(now.getMonth() + 11) % 12];
  const prevYear = now.getMonth() === 0 ? currentYear - 1 : currentYear;
  // OK: matches current month or immediately-prior month (typical accrual cadence)
  const ok =
    (cfgMonth === currentMonth && cfgYear === currentYear) ||
    (cfgMonth === prevMonth && cfgYear === prevYear);
  if (ok) return null;
  return (
    <div className="liquid-glass rounded-xl px-4 py-3 flex items-start gap-3 anim-in"
         style={{ background: "rgba(60, 45, 10, 0.4)" }}>
      <AlertCircle className="w-4 h-4 text-amber-400 mt-0.5 shrink-0" />
      <div className="text-sm text-amber-100/90">
        <strong>Heads up:</strong> You're set to <strong>{cfgMonth} {cfgYear}</strong>, but today is{" "}
        <strong>{currentMonth} {currentYear}</strong>. Confirm this is the right period before running.
      </div>
    </div>
  );
}

function Stat({ label, value, accent }: { label: string; value: number; accent: "emerald"|"indigo"|"amber"|"red"|"neutral" }) {
  const color = {
    emerald: "text-emerald-300",
    indigo: "text-indigo-300",
    amber: "text-amber-300",
    red: "text-red-300",
    neutral: "text-foreground/70",
  }[accent];
  return (
    <div className="liquid-glass rounded-lg px-3 py-2.5">
      <div className="text-[10px] uppercase tracking-wider text-foreground/50">{label}</div>
      <div className={`text-2xl font-semibold tabular-nums ${color}`}>{value}</div>
    </div>
  );
}

function Section({
  title,
  children,
}: {
  title: string;
  children: React.ReactNode;
}) {
  return (
    <div>
      <h3 className="text-foreground/60 text-xs uppercase tracking-wider mb-2 font-medium">
        {title}
      </h3>
      {children}
    </div>
  );
}

function PathInput({
  icon,
  value,
  placeholder,
  onBrowse,
  busy,
}: {
  icon: React.ReactNode;
  value: string;
  placeholder: string;
  onBrowse: () => void;
  busy: boolean;
}) {
  return (
    <div className="liquid-glass rounded-xl flex items-center gap-3 px-3 py-3 hover-lift">
      <div className={`shrink-0 transition-colors ${value ? "text-emerald-300/80" : "text-foreground/40"}`}>{icon}</div>
      <div className="flex-1 text-sm text-foreground/90 truncate">
        {value || <span className="text-foreground/30">{placeholder}</span>}
      </div>
      <button
        onClick={onBrowse}
        disabled={busy}
        className="liquid-glass rounded-lg px-3 py-1.5 text-xs text-foreground hover:text-white transition press disabled:opacity-50 focus-ring"
      >
        {busy ? <span className="inline-flex items-center gap-1.5"><span className="w-1.5 h-1.5 rounded-full bg-indigo-300 pulse-soft" />Opening…</span> : "Browse"}
      </button>
    </div>
  );
}

function Checkbox({
  label,
  checked,
  onChange,
  title,
}: {
  label: string;
  checked: boolean;
  onChange: (v: boolean) => void;
  title?: string;
}) {
  return (
    <label
      className="flex items-center gap-2 text-sm text-foreground/80 cursor-pointer select-none"
      title={title}
    >
      <span
        className={`w-4 h-4 rounded border flex items-center justify-center transition ${
          checked
            ? "bg-indigo-500 border-indigo-400"
            : "bg-black/30 border-foreground/30"
        }`}
      >
        {checked && (
          <svg
            viewBox="0 0 16 16"
            className="w-3 h-3 text-white"
            fill="currentColor"
          >
            <path d="M13.854 3.854a.5.5 0 0 0-.708-.708L6 10.293 2.854 7.146a.5.5 0 1 0-.708.708l3.5 3.5a.5.5 0 0 0 .708 0l7.5-7.5Z" />
          </svg>
        )}
      </span>
      <input
        type="checkbox"
        checked={checked}
        onChange={(e) => onChange(e.target.checked)}
        className="hidden"
      />
      {label}
    </label>
  );
}
