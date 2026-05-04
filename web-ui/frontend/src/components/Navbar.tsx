import { useEffect, useRef, useState } from "react";
import { ChevronDown, Bug, HelpCircle, BookOpen } from "lucide-react";
import { Button } from "./Button";
import { BugReportModal } from "./BugReportModal";
import { FormatRulesModal } from "./FormatRulesModal";

export function Navbar() {
  const [helpOpen, setHelpOpen] = useState(false);
  const [bugOpen, setBugOpen] = useState(false);
  const [rulesOpen, setRulesOpen] = useState(false);
  const helpRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const close = (e: MouseEvent) => {
      if (helpRef.current && !helpRef.current.contains(e.target as Node)) {
        setHelpOpen(false);
      }
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, []);

  return (
    <div className="relative">
      <nav className="w-full py-5 px-8 flex flex-row justify-between items-center">
        <div className="flex items-center gap-2">
          <div className="liquid-glass w-8 h-8 rounded-lg flex items-center justify-center">
            <span className="text-foreground text-sm font-bold">7</span>
          </div>
          <span
            className="font-display text-xl font-semibold text-foreground"
            style={{ letterSpacing: "-0.02em" }}
          >
            7t.ai
          </span>
        </div>

        <div className="flex items-center gap-8">
          <div ref={helpRef} className="relative">
            <button
              onClick={() => setHelpOpen((v) => !v)}
              className="flex items-center gap-1.5 text-foreground/90 hover:text-foreground transition-colors text-sm"
            >
              <HelpCircle className="w-4 h-4" />
              Help
              <ChevronDown
                className={`w-4 h-4 transition-transform ${helpOpen ? "rotate-180" : ""}`}
              />
            </button>

            {helpOpen && (
              <div
                className="absolute top-full right-0 mt-2 w-56 liquid-glass rounded-xl overflow-hidden z-30"
                style={{ background: "rgba(20, 10, 40, 0.95)" }}
              >
                <button
                  onClick={() => {
                    setHelpOpen(false);
                    setRulesOpen(true);
                  }}
                  className="w-full flex items-center gap-2 px-4 py-3 text-sm text-foreground/90 hover:text-foreground hover:bg-foreground/5 transition"
                >
                  <BookOpen className="w-4 h-4 text-indigo-400" />
                  Format Rules
                </button>
                <button
                  onClick={() => {
                    setHelpOpen(false);
                    setBugOpen(true);
                  }}
                  className="w-full flex items-center gap-2 px-4 py-3 text-sm text-foreground/90 hover:text-foreground hover:bg-foreground/5 transition border-t border-foreground/5"
                >
                  <Bug className="w-4 h-4 text-amber-400" />
                  Bug Report
                </button>
              </div>
            )}
          </div>
        </div>

        <div>
          <Button variant="heroSecondary" className="rounded-full px-4 py-2">
            Kristina Woodstock
          </Button>
        </div>
      </nav>
      <div className="mt-[3px] h-px bg-gradient-to-r from-transparent via-foreground/20 to-transparent" />

      <BugReportModal open={bugOpen} onClose={() => setBugOpen(false)} />
      <FormatRulesModal open={rulesOpen} onClose={() => setRulesOpen(false)} />
    </div>
  );
}
