import { X, BookOpen, CheckCircle2, AlertTriangle } from "lucide-react";

interface Props {
  open: boolean;
  onClose: () => void;
}

export function FormatRulesModal({ open, onClose }: Props) {
  if (!open) return null;

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 backdrop-blur-sm p-6"
      onClick={onClose}
    >
      <div
        className="liquid-glass rounded-2xl w-[860px] max-w-[95vw] max-h-[90vh] flex flex-col overflow-hidden"
        style={{ background: "rgba(20, 10, 40, 0.95)" }}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between px-5 py-4 border-b border-foreground/10">
          <div className="flex items-center gap-2">
            <BookOpen className="w-4 h-4 text-indigo-400" />
            <h2 className="text-foreground font-semibold text-base">
              Accrual Master File Format Rules
            </h2>
          </div>
          <button
            onClick={onClose}
            className="text-foreground/60 hover:text-foreground transition"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        <div className="flex-1 overflow-y-auto p-6 space-y-6 text-sm text-foreground/85">
          {/* Quick rules */}
          <Section title="Golden Rules" icon={<CheckCircle2 className="w-4 h-4 text-emerald-400" />}>
            <ul className="list-disc pl-5 space-y-1.5 text-foreground/80">
              <li>Sheet name must be <Code>Profit Sharing</Code></li>
              <li>Header row must be <Code>row 3</Code></li>
              <li>Data starts at <Code>row 4</Code></li>
              <li>File ID column holds 5- or 6-digit numbers</li>
              <li>Column position can change anytime — code finds them by name</li>
              <li>Each year's paysheet tab must be named with the year (e.g. <Code>2026</Code>, <Code>2026-Closed</Code>, <Code>FY 2026</Code>)</li>
            </ul>
          </Section>

          {/* Required columns */}
          <Section title="Required Column Headers" icon={<BookOpen className="w-4 h-4 text-indigo-400" />}>
            <p className="text-xs text-foreground/60 mb-2">
              Code searches row 3 for any of these phrases (case-insensitive, whitespace flexible):
            </p>
            <div className="rounded-lg overflow-hidden border border-foreground/10">
              <table className="w-full text-xs">
                <thead className="bg-white/[0.03] text-foreground/60">
                  <tr>
                    <th className="text-left px-3 py-2 font-medium">Master Column</th>
                    <th className="text-left px-3 py-2 font-medium">Keywords (any one works)</th>
                    <th className="text-left px-3 py-2 font-medium">Required?</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-foreground/5">
                  <Row col="File ID" kws='"applicant number", "app id", "appid", "file #", "file number", "app no"' req="Yes" />
                  <Row col="Employee Name" kws='"employee name", "payroll name", "employee", "payroll"' req="Yes" />
                  <Row col="Hours" kws='"{Month} hours", "{Mon} hours", "{Month} hrs", "{Mon} hrs"' req="Yes" />
                  <Row col="Billed" kws='"{Month} billed", "{Mon} billed", "{Month} billing", "{Mon} billing"' req="Yes" />
                  <Row col="Admin Fee" kws='"admin fee", "adminfee"' req="Optional" />
                  <Row col="Salary Paid" kws='"salary paid"' req="Optional" />
                  <Row col="Wages Earned" kws='"wages earned"' req="Hourly only" />
                  <Row col="Carryforward" kws='"carryforward", "balance forward", "accrued payroll per audit", "starting pi balance"' req="Jan only" />
                  <Row col="Gross Salary" kws='"total gross salary", "gross salary", "gross pay"' req="Optional" />
                </tbody>
              </table>
            </div>
            <p className="text-xs text-foreground/50 mt-2">
              <strong>Month placeholder:</strong> code substitutes the run month and its 3-letter abbreviation.
              For <Code>February</Code>, it tries: <Code>"february hours"</Code>, <Code>"feb hours"</Code>, <Code>"february hrs"</Code>, <Code>"feb hrs"</Code>.
            </p>
          </Section>

          {/* Examples */}
          <Section title="Header Examples (all valid)" icon={<CheckCircle2 className="w-4 h-4 text-emerald-400" />}>
            <ul className="list-disc pl-5 space-y-1 text-foreground/80">
              <li><Code>Feb Hours</Code> ✅ &nbsp;<Code>February Hours</Code> ✅ &nbsp;<Code>FEB HRS</Code> ✅</li>
              <li><Code>February Billed to the Client on PS</Code> ✅</li>
              <li><Code>12/31/2025 Accrued Payroll per Audit</Code> ✅ (carryforward)</li>
              <li><Code>February Starting PI Balance (January Owe)</Code> ✅</li>
              <li><Code>Ultra-Staff Applicant     Number</Code> ✅ (multi-space OK)</li>
            </ul>
          </Section>

          {/* When it breaks */}
          <Section title="When it Breaks" icon={<AlertTriangle className="w-4 h-4 text-amber-400" />}>
            <ul className="list-disc pl-5 space-y-1.5 text-foreground/80">
              <li>Header renamed to phrase with NO recognized keyword.<br/>
                <span className="text-xs text-foreground/55">Example: <Code>"Feb Time"</Code> instead of <Code>"Feb Hours"</Code> — no "hours"/"hrs" word</span>
              </li>
              <li>Header row not in row 3 (currently hardcoded)</li>
              <li>Sheet name not <Code>Profit Sharing</Code></li>
              <li>Two columns both match same keyword — picks first hit</li>
              <li>File ID column has values that aren't 5- or 6-digit numbers</li>
            </ul>
          </Section>

          {/* Cells we never touch */}
          <Section title="Always Protected (code never overwrites)" icon={<CheckCircle2 className="w-4 h-4 text-emerald-400" />}>
            <ul className="list-disc pl-5 space-y-1 text-foreground/80">
              <li>Any cell starting with <Code>=</Code> (formula)</li>
              <li>Pivot tables, queries, external data ranges, named ranges</li>
              <li>All columns OUTSIDE the 5–6 we write to</li>
              <li>All other sheets in the workbook</li>
            </ul>
          </Section>

          {/* Recommendation */}
          <Section title="Recommended Practice" icon={<BookOpen className="w-4 h-4 text-indigo-400" />}>
            <ul className="list-disc pl-5 space-y-1 text-foreground/80">
              <li>Keep month + "Hours" or "Hrs" word in hours columns: <Code>"Feb Hours"</Code>, <Code>"Mar Hrs"</Code></li>
              <li>Keep month + "Billed" or "Billing": <Code>"Feb Billed"</Code></li>
              <li>Keep <Code>"Admin Fee"</Code> exactly</li>
              <li>Keep <Code>"Salary Paid"</Code> exactly</li>
              <li>Keep <Code>"Wages Earned"</Code> exactly (hourly section)</li>
              <li>Year tabs in paysheets named with year: <Code>2026</Code>, <Code>2027</Code> (suffixes OK)</li>
            </ul>
          </Section>
        </div>

        <div className="px-5 py-3 border-t border-foreground/10 flex justify-end">
          <button
            onClick={onClose}
            className="liquid-glass rounded-lg px-4 py-2 text-sm text-foreground hover:text-white transition"
          >
            Close
          </button>
        </div>
      </div>
    </div>
  );
}

function Section({
  title,
  icon,
  children,
}: {
  title: string;
  icon?: React.ReactNode;
  children: React.ReactNode;
}) {
  return (
    <div>
      <h3 className="flex items-center gap-2 text-foreground font-semibold text-sm mb-2">
        {icon}
        {title}
      </h3>
      {children}
    </div>
  );
}

function Code({ children }: { children: React.ReactNode }) {
  return (
    <code className="bg-black/40 border border-foreground/10 rounded px-1.5 py-0.5 text-xs font-mono text-amber-200">
      {children}
    </code>
  );
}

function Row({ col, kws, req }: { col: string; kws: string; req: string }) {
  const reqColor =
    req === "Yes" ? "text-emerald-400" :
    req === "Hourly only" || req === "Jan only" ? "text-amber-300" :
    "text-foreground/50";
  return (
    <tr>
      <td className="px-3 py-2 text-foreground font-medium whitespace-nowrap">{col}</td>
      <td className="px-3 py-2 text-foreground/70 font-mono">{kws}</td>
      <td className={`px-3 py-2 ${reqColor} whitespace-nowrap`}>{req}</td>
    </tr>
  );
}
