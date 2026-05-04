import { useEffect, useState } from "react";
import { ChevronRight, Folder, FileSpreadsheet, ArrowUp, X } from "lucide-react";

interface Entry {
  name: string;
  path: string;
  is_dir: boolean;
}

interface BrowseResp {
  path: string;
  parent: string | null;
  entries: Entry[];
}

interface Props {
  open: boolean;
  mode: "file" | "folder";
  initialPath?: string;
  onClose: () => void;
  onSelect: (path: string) => void;
}

const API = "/api";

export function FolderPicker({ open, mode, initialPath, onClose, onSelect }: Props) {
  const [data, setData] = useState<BrowseResp | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);

  const load = async (path?: string) => {
    setLoading(true);
    setError(null);
    try {
      const qs = new URLSearchParams();
      if (path) qs.set("path", path);
      if (mode === "file") qs.set("files", "true");
      const r = await fetch(`${API}/browse?${qs}`);
      if (!r.ok) throw new Error(await r.text());
      setData(await r.json());
    } catch (e: any) {
      setError(String(e.message || e));
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    if (open) load(initialPath);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  if (!open) return null;

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 backdrop-blur-sm"
      onClick={onClose}
    >
      <div
        className="liquid-glass rounded-2xl w-[720px] max-w-[90vw] h-[560px] max-h-[85vh] flex flex-col overflow-hidden"
        style={{ background: "rgba(20, 10, 40, 0.85)" }}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between px-5 py-4 border-b border-foreground/10">
          <div>
            <h2 className="text-foreground font-semibold text-base">
              {mode === "file" ? "Select Master File" : "Select Paysheets Folder"}
            </h2>
            <p className="text-foreground/50 text-xs mt-0.5 truncate max-w-[600px]">
              {data?.path || "Loading..."}
            </p>
          </div>
          <button
            onClick={onClose}
            className="text-foreground/60 hover:text-foreground transition"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        <div className="flex-1 overflow-y-auto">
          {loading && (
            <div className="p-5 text-foreground/60 text-sm">Loading…</div>
          )}
          {error && (
            <div className="p-5 text-red-400 text-sm">Error: {error}</div>
          )}
          {data && !loading && (
            <div className="p-2">
              {data.parent && (
                <button
                  onClick={() => load(data.parent!)}
                  className="w-full flex items-center gap-3 px-3 py-2 rounded-lg hover:bg-foreground/5 transition text-foreground/80 text-sm"
                >
                  <ArrowUp className="w-4 h-4" />
                  <span>..</span>
                </button>
              )}
              {data.entries.length === 0 && (
                <div className="p-4 text-foreground/40 text-sm">Empty</div>
              )}
              {data.entries.map((e) => (
                <button
                  key={e.path}
                  onDoubleClick={() => {
                    if (e.is_dir) load(e.path);
                    else onSelect(e.path);
                  }}
                  onClick={() => {
                    if (e.is_dir) load(e.path);
                  }}
                  className="w-full flex items-center gap-3 px-3 py-2 rounded-lg hover:bg-foreground/5 transition text-foreground/90 text-sm group"
                >
                  {e.is_dir ? (
                    <Folder className="w-4 h-4 text-indigo-400" />
                  ) : (
                    <FileSpreadsheet className="w-4 h-4 text-emerald-400" />
                  )}
                  <span className="flex-1 text-left truncate">{e.name}</span>
                  {e.is_dir && (
                    <ChevronRight className="w-4 h-4 text-foreground/30 opacity-0 group-hover:opacity-100" />
                  )}
                </button>
              ))}
            </div>
          )}
        </div>

        <div className="px-5 py-4 border-t border-foreground/10 flex items-center justify-between gap-3">
          <span className="text-foreground/40 text-xs">
            {mode === "folder"
              ? "Navigate into a folder and click Select"
              : "Double-click a file to select it"}
          </span>
          <div className="flex gap-2">
            <button
              onClick={onClose}
              className="px-4 py-2 text-sm text-foreground/70 hover:text-foreground transition"
            >
              Cancel
            </button>
            {mode === "folder" && data && (
              <button
                onClick={() => onSelect(data.path)}
                className="liquid-glass px-4 py-2 rounded-lg text-sm text-foreground font-medium hover:text-white transition"
              >
                Select This Folder
              </button>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
