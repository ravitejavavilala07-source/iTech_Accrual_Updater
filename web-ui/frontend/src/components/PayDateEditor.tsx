import { useState } from "react";
import { Plus, Trash2, Pencil } from "lucide-react";
import { pushToast } from "./Toast";

export interface PayDate {
  date: string; // MM/DD/YYYY
  multiplier: number;
}

interface Props {
  rows: PayDate[];
  onChange: (rows: PayDate[]) => void;
}

function todayStr(): string {
  const d = new Date();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${mm}/${dd}/${d.getFullYear()}`;
}

// MM/DD/YYYY -> YYYY-MM-DD (HTML5 date input format)
function toIsoDate(mdy: string): string {
  const m = mdy.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) return "";
  return `${m[3]}-${m[1]}-${m[2]}`;
}

// YYYY-MM-DD -> MM/DD/YYYY
function fromIsoDate(iso: string): string {
  const m = iso.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return "";
  return `${m[2]}/${m[3]}/${m[1]}`;
}

export function PayDateEditor({ rows, onChange }: Props) {
  const [editIdx, setEditIdx] = useState<number | null>(null);
  const [formDate, setFormDate] = useState<string>(todayStr());
  const [formMul, setFormMul] = useState<string>("1.00");
  const [showForm, setShowForm] = useState(false);

  const openAdd = () => {
    setEditIdx(null);
    setFormDate(todayStr());
    setFormMul("1.00");
    setShowForm(true);
  };

  const openEdit = (i: number) => {
    setEditIdx(i);
    setFormDate(rows[i].date);
    setFormMul(String(rows[i].multiplier));
    setShowForm(true);
  };

  const save = () => {
    const m = parseFloat(formMul);
    if (!/^\d{2}\/\d{2}\/\d{4}$/.test(formDate) || isNaN(m)) {
      pushToast("Enter date as MM/DD/YYYY and a numeric multiplier.", "error");
      return;
    }
    const newRows = [...rows];
    const row = { date: formDate, multiplier: m };
    if (editIdx === null) newRows.push(row);
    else newRows[editIdx] = row;
    onChange(newRows);
    setShowForm(false);
  };

  const remove = (i: number) => {
    const newRows = rows.filter((_, idx) => idx !== i);
    onChange(newRows);
  };

  return (
    <div className="space-y-3">
      <div className="flex items-center gap-2">
        <button
          onClick={openAdd}
          className="liquid-glass rounded-lg px-3 py-1.5 text-sm text-foreground flex items-center gap-1.5 hover:text-white transition"
        >
          <Plus className="w-3.5 h-3.5" /> Add Date
        </button>
      </div>

      {showForm && (
        <div className="liquid-glass rounded-lg p-3 flex items-center gap-3 text-sm flex-wrap">
          <input
            type="date"
            value={toIsoDate(formDate)}
            onChange={(e) => setFormDate(fromIsoDate(e.target.value) || formDate)}
            className="bg-black/30 border border-foreground/10 rounded px-2 py-1.5 text-foreground w-44 outline-none focus:border-foreground/40"
          />
          <input
            type="number"
            step="0.01"
            min="0"
            max="10"
            value={formMul}
            onChange={(e) => setFormMul(e.target.value)}
            placeholder="1.00"
            className="bg-black/30 border border-foreground/10 rounded px-2 py-1.5 text-foreground w-24 outline-none focus:border-foreground/40"
          />
          <button
            onClick={save}
            className="liquid-glass rounded px-3 py-1.5 text-xs text-foreground hover:text-white transition"
          >
            Save
          </button>
          <button
            onClick={() => setShowForm(false)}
            className="text-foreground/60 hover:text-foreground text-xs"
          >
            Cancel
          </button>
        </div>
      )}

      <div className="rounded-lg overflow-hidden border border-foreground/10">
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-white/[0.02] text-foreground/60 text-xs">
              <th className="text-left px-3 py-2 font-medium">Date</th>
              <th className="text-left px-3 py-2 font-medium">Multiplier</th>
              <th className="w-20" />
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 && (
              <tr>
                <td colSpan={3} className="px-3 py-4 text-foreground/40 text-center text-xs">
                  No pay dates added.
                </td>
              </tr>
            )}
            {rows.map((r, i) => (
              <tr key={i} className="border-t border-foreground/5">
                <td className="px-3 py-2 text-foreground/90">{r.date}</td>
                <td className="px-3 py-2 text-foreground/90">{r.multiplier.toFixed(2)}</td>
                <td className="px-3 py-2">
                  <div className="flex gap-2 justify-end">
                    <button
                      onClick={() => openEdit(i)}
                      className="text-foreground/50 hover:text-foreground transition"
                    >
                      <Pencil className="w-3.5 h-3.5" />
                    </button>
                    <button
                      onClick={() => remove(i)}
                      className="text-foreground/50 hover:text-red-400 transition"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
