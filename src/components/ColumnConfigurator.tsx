"use client";
import { getColumnStats, isNumeric } from "@/lib/excel";
import type { ColumnConfig, ColumnMode } from "@/lib/excel";
import type { ExcelData } from "@/lib/excel";
import ValueDropdown from "./ValueDropdown";
import { Hash, Type, ToggleLeft } from "lucide-react";

interface ColumnConfiguratorProps {
  data: ExcelData;
  configs: ColumnConfig[];
  onChange: (configs: ColumnConfig[]) => void;
}

const MODE_OPTIONS: { value: ColumnMode; label: string; color: string }[] = [
  { value: "ignore", label: "Ignore", color: "bg-slate-800 text-slate-400 border-slate-700" },
  { value: "select", label: "Filter", color: "bg-blue-950 text-blue-300 border-blue-800" },
  { value: "separate", label: "Group by", color: "bg-violet-950 text-violet-300 border-violet-800" },
];

export default function ColumnConfigurator({ data, configs, onChange }: ColumnConfiguratorProps) {
  const setMode = (col: string, mode: ColumnMode) => {
    onChange(
      configs.map((c) =>
        c.col === col ? { ...c, mode, selectedValues: mode === "ignore" ? [] : c.selectedValues } : c
      )
    );
  };

  const setSelected = (col: string, vals: string[]) => {
    onChange(configs.map((c) => (c.col === col ? { ...c, selectedValues: vals } : c)));
  };

  return (
    <div className="rounded-xl border border-slate-800 overflow-hidden">
      <div className="px-5 py-4 bg-slate-900 border-b border-slate-800">
        <div className="flex items-center gap-3">
          <ToggleLeft className="w-4 h-4 text-violet-400" />
          <span className="font-semibold text-slate-200 text-sm">Column Configuration</span>
          <span className="text-xs text-slate-500">Choose how each column is used in the summary</span>
        </div>
      </div>

      <div className="divide-y divide-slate-800/60">
        {configs.map((cfg) => {
          const stats = getColumnStats(data.rows, cfg.col);
          return (
            <div key={cfg.col} className="flex flex-wrap items-center gap-3 px-5 py-3.5 bg-slate-950 hover:bg-slate-900/50 transition-colors">
              {/* Column name + type badge */}
              <div className="flex items-center gap-2 w-52 flex-shrink-0">
                {stats.isNumeric
                  ? <Hash className="w-3.5 h-3.5 text-amber-400 flex-shrink-0" />
                  : <Type className="w-3.5 h-3.5 text-sky-400 flex-shrink-0" />
                }
                <span className="text-sm font-medium text-slate-200 truncate">{cfg.col}</span>
                <span className="text-xs text-slate-600 font-mono whitespace-nowrap">
                  {stats.uniqueCount} unique
                </span>
              </div>

              {/* Mode toggle */}
              <div className="flex items-center gap-1">
                {MODE_OPTIONS.map((opt) => (
                  <button
                    key={opt.value}
                    onClick={() => setMode(cfg.col, opt.value)}
                    className={`px-3 py-1.5 rounded-md border text-xs font-medium transition-all ${
                      cfg.mode === opt.value
                        ? opt.color + " shadow-sm"
                        : "bg-transparent text-slate-600 border-slate-800 hover:border-slate-700 hover:text-slate-400"
                    }`}
                  >
                    {opt.label}
                  </button>
                ))}
              </div>

              {/* Value selector for "Filter" mode */}
              {cfg.mode === "select" && !stats.isNumeric && (
                <ValueDropdown
                  label={cfg.col}
                  values={stats.uniqueValues}
                  selected={cfg.selectedValues}
                  onChange={(vals) => setSelected(cfg.col, vals)}
                />
              )}

              {cfg.mode === "select" && stats.isNumeric && (
                <span className="text-xs text-amber-400/60 italic">
                  Numeric column — filtering not available, use Group by or Ignore
                </span>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}
