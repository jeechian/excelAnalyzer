"use client";
import { useState, useRef, useEffect } from "react";
import { ChevronDown, X, Check } from "lucide-react";

interface ValueDropdownProps {
  label: string;
  values: string[];
  selected: string[];
  onChange: (vals: string[]) => void;
}

export default function ValueDropdown({ label, values, selected, onChange }: ValueDropdownProps) {
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState("");
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handler = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, []);

  const filtered = values.filter((v) =>
    v.toLowerCase().includes(search.toLowerCase())
  );

  const toggle = (v: string) => {
    if (selected.includes(v)) onChange(selected.filter((s) => s !== v));
    else onChange([...selected, v]);
  };

  const toggleAll = () => {
    if (selected.length === values.length) onChange([]);
    else onChange([...values]);
  };

  return (
    <div className="relative" ref={ref}>
      <button
        onClick={() => setOpen((o) => !o)}
        className="flex items-center gap-2 px-3 py-2 rounded-lg bg-slate-800 border border-slate-700 hover:border-violet-500 transition-colors text-sm text-slate-300 min-w-[180px] justify-between"
      >
        <span className="truncate max-w-[140px]">
          {selected.length === 0
            ? `Select ${label}…`
            : selected.length === values.length
            ? `All ${label}`
            : `${selected.length} selected`}
        </span>
        <div className="flex items-center gap-1">
          {selected.length > 0 && (
            <span
              className="text-xs bg-violet-600 text-white rounded-full px-1.5 py-0.5 font-mono"
              onClick={(e) => { e.stopPropagation(); onChange([]); }}
            >
              ×
            </span>
          )}
          <ChevronDown className={`w-3.5 h-3.5 text-slate-500 transition-transform ${open ? "rotate-180" : ""}`} />
        </div>
      </button>

      {open && (
        <div className="absolute z-50 mt-1 w-64 rounded-lg border border-slate-700 bg-slate-900 shadow-xl overflow-hidden">
          <div className="p-2 border-b border-slate-800">
            <input
              autoFocus
              type="text"
              placeholder="Search…"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              className="w-full bg-slate-800 border border-slate-700 rounded px-2.5 py-1.5 text-sm text-slate-200 placeholder-slate-500 focus:outline-none focus:border-violet-500"
            />
          </div>
          <div className="p-1.5 border-b border-slate-800">
            <button
              onClick={toggleAll}
              className="w-full text-left text-xs text-slate-400 hover:text-violet-400 px-2 py-1 rounded hover:bg-slate-800 transition-colors"
            >
              {selected.length === values.length ? "Deselect all" : "Select all"}
            </button>
          </div>
          <div className="max-h-52 overflow-y-auto">
            {filtered.length === 0 && (
              <p className="text-center text-slate-600 text-xs py-4">No matches</p>
            )}
            {filtered.map((v) => (
              <button
                key={v}
                onClick={() => toggle(v)}
                className="w-full flex items-center gap-2.5 px-3 py-2 text-sm text-slate-300 hover:bg-slate-800 transition-colors text-left"
              >
                <span className={`w-4 h-4 rounded border flex items-center justify-center flex-shrink-0 transition-colors ${
                  selected.includes(v) ? "bg-violet-600 border-violet-600" : "border-slate-600"
                }`}>
                  {selected.includes(v) && <Check className="w-2.5 h-2.5 text-white" />}
                </span>
                <span className="truncate">{v}</span>
              </button>
            ))}
          </div>
          <div className="p-2 border-t border-slate-800 text-right">
            <button
              onClick={() => setOpen(false)}
              className="text-xs text-violet-400 hover:text-violet-300 px-2 py-1"
            >
              Done
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
