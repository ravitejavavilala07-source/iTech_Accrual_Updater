import { useState } from "react";
import { X, Bug, CheckCircle2, AlertCircle } from "lucide-react";

interface Props {
  open: boolean;
  onClose: () => void;
}

type Status = "idle" | "sending" | "success" | "error";

export function BugReportModal({ open, onClose }: Props) {
  const [description, setDescription] = useState("");
  const [senderEmail, setSenderEmail] = useState("");
  const [status, setStatus] = useState<Status>("idle");
  const [error, setError] = useState<string | null>(null);

  if (!open) return null;

  const openMailtoFallback = () => {
    const subject = encodeURIComponent("7t.ai Accrual Updater — Bug Report");
    const body = encodeURIComponent(
      `From: ${senderEmail || "anonymous"}\n` +
        `URL: ${window.location.href}\n` +
        `\n` +
        `Description:\n${description}\n`
    );
    window.location.href = `mailto:ravi.vavilala@riseits.com?subject=${subject}&body=${body}`;
  };

  const submit = async () => {
    if (!description.trim()) {
      setError("Please describe the bug.");
      return;
    }
    setStatus("sending");
    setError(null);
    try {
      const r = await fetch("/api/bug-report", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          description,
          sender_email: senderEmail || null,
          user_agent: navigator.userAgent,
          url: window.location.href,
        }),
      });
      if (!r.ok) throw new Error(await r.text());
      const data = await r.json();

      if (data.emailed_to) {
        setStatus("success");
        setTimeout(() => {
          setDescription("");
          setSenderEmail("");
          setStatus("idle");
          onClose();
        }, 1500);
      } else {
        // Backend logged but couldn't send. Show the actual error then mailto fallback.
        const reason = data.error || "Outlook unavailable";
        setError(`Email send failed (${String(reason).slice(0, 120)}). Opening your mail app as fallback…`);
        setTimeout(() => {
          openMailtoFallback();
          setStatus("success");
          setTimeout(() => {
            setDescription("");
            setSenderEmail("");
            setStatus("idle");
            setError(null);
            onClose();
          }, 1500);
        }, 1200);
      }
    } catch (e: any) {
      // Backend unreachable — fallback straight to mailto
      openMailtoFallback();
      setStatus("error");
      setError(
        "Backend unreachable — opened your mail app instead. Click Send in your email client."
      );
    }
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 backdrop-blur-sm"
      onClick={onClose}
    >
      <div
        className="liquid-glass rounded-2xl w-[520px] max-w-[90vw] flex flex-col overflow-hidden"
        style={{ background: "rgba(20, 10, 40, 0.92)" }}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between px-5 py-4 border-b border-foreground/10">
          <div className="flex items-center gap-2">
            <Bug className="w-4 h-4 text-indigo-400" />
            <h2 className="text-foreground font-semibold text-base">
              Report a Bug
            </h2>
          </div>
          <button
            onClick={onClose}
            className="text-foreground/60 hover:text-foreground transition"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        <div className="p-5 space-y-4">
          <div>
            <label className="block text-xs text-foreground/50 uppercase tracking-wider mb-1.5">
              Your email (optional)
            </label>
            <input
              type="email"
              value={senderEmail}
              onChange={(e) => setSenderEmail(e.target.value)}
              placeholder="you@example.com"
              className="w-full bg-black/40 border border-foreground/15 rounded-lg px-3 py-2 text-foreground text-sm outline-none focus:border-indigo-400 transition"
              disabled={status === "sending"}
            />
          </div>

          <div>
            <label className="block text-xs text-foreground/50 uppercase tracking-wider mb-1.5">
              Describe the bug
            </label>
            <textarea
              value={description}
              onChange={(e) => setDescription(e.target.value)}
              placeholder="What did you expect? What happened? How can we reproduce it?"
              rows={6}
              className="w-full bg-black/40 border border-foreground/15 rounded-lg px-3 py-2 text-foreground text-sm outline-none focus:border-indigo-400 transition resize-none"
              disabled={status === "sending"}
            />
          </div>

          {error && (
            <div className="flex items-start gap-2 text-red-400 text-xs">
              <AlertCircle className="w-4 h-4 shrink-0 mt-0.5" />
              <span>{error}</span>
            </div>
          )}

          {status === "success" && (
            <div className="flex items-center gap-2 text-emerald-400 text-sm">
              <CheckCircle2 className="w-4 h-4" />
              <span>Report sent. Thank you!</span>
            </div>
          )}
        </div>

        <div className="px-5 py-4 border-t border-foreground/10 flex items-center justify-end gap-2">
          <button
            onClick={onClose}
            disabled={status === "sending"}
            className="px-4 py-2 text-sm text-foreground/70 hover:text-foreground transition disabled:opacity-50"
          >
            Cancel
          </button>
          <button
            onClick={submit}
            disabled={status === "sending" || status === "success"}
            className="liquid-glass rounded-lg px-5 py-2 text-sm text-foreground font-medium hover:text-white transition disabled:opacity-50"
          >
            {status === "sending" ? "Sending…" : "Send Report"}
          </button>
        </div>
      </div>
    </div>
  );
}
