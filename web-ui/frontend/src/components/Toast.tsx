import { useEffect, useState } from "react";
import { CheckCircle2, AlertCircle, Info, X } from "lucide-react";

type Variant = "success" | "error" | "info";

interface ToastItem {
  id: number;
  message: string;
  variant: Variant;
}

let _push: ((msg: string, variant?: Variant) => void) | null = null;

export function pushToast(msg: string, variant: Variant = "info") {
  if (_push) _push(msg, variant);
  else console.log(`[toast/${variant}]`, msg);
}

const ICONS: Record<Variant, React.ReactNode> = {
  success: <CheckCircle2 className="w-4 h-4 text-emerald-400" />,
  error: <AlertCircle className="w-4 h-4 text-red-400" />,
  info: <Info className="w-4 h-4 text-indigo-400" />,
};

export function ToastHost() {
  const [items, setItems] = useState<ToastItem[]>([]);

  useEffect(() => {
    _push = (message, variant = "info") => {
      const id = Date.now() + Math.random();
      setItems((prev) => [...prev, { id, message, variant }]);
      setTimeout(() => {
        setItems((prev) => prev.filter((t) => t.id !== id));
      }, 4000);
    };
    return () => {
      _push = null;
    };
  }, []);

  return (
    <div className="fixed bottom-6 right-6 z-[60] flex flex-col gap-2 pointer-events-none">
      {items.map((t) => (
        <div
          key={t.id}
          className="liquid-glass rounded-xl px-4 py-3 flex items-start gap-3 min-w-[280px] max-w-[420px] toast-in pointer-events-auto"
          style={{ background: "rgba(20, 10, 40, 0.95)" }}
        >
          {ICONS[t.variant]}
          <div className="flex-1 text-sm text-foreground/90">{t.message}</div>
          <button
            onClick={() =>
              setItems((prev) => prev.filter((x) => x.id !== t.id))
            }
            className="text-foreground/40 hover:text-foreground transition"
          >
            <X className="w-4 h-4" />
          </button>
        </div>
      ))}
    </div>
  );
}
