"use client";
import { ArrowUpDown, ArrowUp, ArrowDown, X } from "lucide-react";
import { useState } from "react";
import type { SortEntry } from "@/lib/excel";

interface SortControlsProps {
  sortableColumns: string[];
  sortOrder: SortEntry[];
  onSortOrder: (order: SortEntry[]) => void;
}

export default function SortControls({ sortableColumns, sortOrder, onSortOrder }: SortControlsProps) {
  const [search, setSearch] = useState("");

  const filtered = sortableColumns.filter((c) =>
    c.toLowerCase().includes(search.toLowerCase())
  );

  const handleColClick = (col: string) => {
    const idx = sortOrder.findIndex((e) => e.col === col);
    if (idx === -1) {
      onSortOrder([...sortOrder, { col, dir: "asc" }]);
    } else {
      onSortOrder(sortOrder.filter((e) => e.col !== col));
    }
  };

  const toggleDir = (col: string) => {
    onSortOrder(
      sortOrder.map((e) =>
        e.col === col ? { ...e, dir: e.dir === "asc" ? "desc" : "asc" } : e
      )
    );
  };

  if (sortableColumns.length === 0) return null;

  return (
    <div className="rounded-xl border border-slate-800 overflow-hidden">
      <div className="px-5 py-4 bg-slate-900 border-b border-slate-800 flex items-center gap-3">
        <ArrowUpDown className="w-4 h-4 text-violet-400" />
        <span className="font-semibold text-slate-200 text-sm">Sort Results</span>
        <span className="text-xs text-slate-500">Click to add · click arrow to toggle direction · click again to remove</span>
        {sortOrder.length > 0 && (
          <button
            onClick={() => onSortOrder([])}
            className="ml-auto flex items-center gap-1 text-xs text-slate-500 hover:text-slate-300 transition-colors"
          >
            <X className="w-3 h-3" /> Clear
          </button>
        )}
      </div>

      <div className="p-4 bg-slate-950">
        <input
          type="text"
          placeholder="Search column…"
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          className="mb-3 bg-slate-800 border border-slate-700 rounded-lg px-3 py-1.5 text-sm text-slate-300 placeholder-slate-600 focus:outline-none focus:border-violet-500 w-52"
        />
        <div className="flex flex-wrap gap-1.5">
          {filtered.map((col) => {
            const idx = sortOrder.findIndex((e) => e.col === col);
            const entry = idx !== -1 ? sortOrder[idx] : null;
            return (
              <button
                key={col}
                onClick={() => handleColClick(col)}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg border text-xs transition-all ${
                  entry
                    ? "bg-violet-900/60 border-violet-700 text-violet-200"
                    : "bg-transparent border-slate-800 text-slate-500 hover:border-slate-700 hover:text-slate-400"
                }`}
              >
                {entry && sortOrder.length > 1 && (
                  <span className="w-4 h-4 rounded-full bg-violet-600 text-white text-[10px] font-bold flex items-center justify-center flex-shrink-0">
                    {idx + 1}
                  </span>
                )}
                <span>{col}</span>
                {entry && (
                  <span
                    onClick={(e) => { e.stopPropagation(); toggleDir(col); }}
                    className="text-violet-300 hover:text-white transition-colors"
                  >
                    {entry.dir === "asc"
                      ? <ArrowUp className="w-3 h-3" />
                      : <ArrowDown className="w-3 h-3" />}
                  </span>
                )}
              </button>
            );
          })}
        </div>
      </div>
    </div>
  );
}
