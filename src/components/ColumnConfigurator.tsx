"use client";
import { getColumnStats } from "@/lib/excel";
import type { ColumnConfig, ColumnMode } from "@/lib/excel";
import type { ExcelData } from "@/lib/excel";
import ValueDropdown from "./ValueDropdown";
import { Hash, Type, ToggleLeft, Calendar, RotateCcw } from "lucide-react";

interface ColumnConfiguratorProps {
  data: ExcelData;
  configs: ColumnConfig[];
  onChange: (configs: ColumnConfig[]) => void;
}

function colNameHasStartOrEnd(col: string): boolean {
  const n = col.toLowerCase().replace(/[\s_-]/g, "");
  return n.includes("startdate") || n.includes("enddate");
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
        c.col === col
          ? {
              ...c,
              mode,
              selectedValues: mode === "ignore" ? [] : c.selectedValues,
              dateRange: mode === "select" ? c.dateRange : undefined,
              numericRange: mode === "select" ? c.numericRange : undefined,
            }
          : c
      )
    );
  };

  const setSelected = (col: string, vals: string[]) => {
    onChange(configs.map((c) => (c.col === col ? { ...c, selectedValues: vals } : c)));
  };

  const setDateRange = (col: string, range: { start?: string; end?: string }) => {
    onChange(configs.map((c) => (c.col === col ? { ...c, dateRange: range } : c)));
  };

  const setDateGroupMode = (col: string, mode: "split" | "combine") => {
    onChange(configs.map((c) => (c.col === col ? { ...c, dateGroupMode: mode } : c)));
  };

  const setNumericRange = (col: string, range: { min?: string; max?: string }) => {
    onChange(configs.map((c) => (c.col === col ? { ...c, numericRange: range } : c)));
  };

  const setNumericGroupMode = (col: string, mode: "split" | "combine") => {
    onChange(configs.map((c) => (c.col === col ? { ...c, numericGroupMode: mode } : c)));
  };

  return (
    <div className="rounded-xl border border-slate-800 overflow-hidden">
      <div className="px-5 py-4 bg-slate-900 border-b border-slate-800 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <ToggleLeft className="w-4 h-4 text-violet-400" />
          <span className="font-semibold text-slate-200 text-sm">Column Configuration</span>
          <span className="text-xs text-slate-500">Choose how each column is used in the summary</span>
        </div>
        <button
          onClick={() =>
            onChange(
              configs.map((c) => ({
                ...c,
                mode: "ignore",
                selectedValues: [],
                dateRange: undefined,
                numericRange: undefined,
              }))
            )
          }
          className="flex items-center gap-1.5 text-xs text-slate-500 hover:text-slate-300 px-2.5 py-1.5 rounded-lg border border-slate-800 hover:border-slate-700 transition-all"
        >
          <RotateCcw className="w-3 h-3" />
          Reset
        </button>
      </div>

      <div className="divide-y divide-slate-800/60">
        {configs.map((cfg) => {
          const stats = getColumnStats(data.rows, cfg.col);
          const isDateCol = stats.isDate || (data.dateColumns ?? []).includes(cfg.col);
          const skipDateRange = colNameHasStartOrEnd(cfg.col);
          const showDateRange = isDateCol && !skipDateRange;
          return (
            <div key={cfg.col} className="flex flex-wrap items-center gap-3 px-5 py-3.5 bg-slate-950 hover:bg-slate-900/50 transition-colors">
              {/* Column name + type badge */}
              <div className="flex items-center gap-2 w-52 flex-shrink-0">
                {isDateCol
                  ? <Calendar className="w-3.5 h-3.5 text-emerald-400 flex-shrink-0" />
                  : stats.isNumeric
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

              {/* Date range inputs for detected date columns (not already named start/end date) */}
              {cfg.mode === "select" && showDateRange && (
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="text-xs text-slate-500">From</span>
                  <input
                    type="date"
                    value={cfg.dateRange?.start ?? ""}
                    onChange={(e) =>
                      setDateRange(cfg.col, { ...cfg.dateRange, start: e.target.value || undefined })
                    }
                    className="bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs text-slate-200 focus:outline-none focus:border-emerald-600 [color-scheme:dark]"
                  />
                  <span className="text-xs text-slate-500">To</span>
                  <input
                    type="date"
                    value={cfg.dateRange?.end ?? ""}
                    onChange={(e) =>
                      setDateRange(cfg.col, { ...cfg.dateRange, end: e.target.value || undefined })
                    }
                    className="bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs text-slate-200 focus:outline-none focus:border-emerald-600 [color-scheme:dark]"
                  />
                  {(cfg.dateRange?.start || cfg.dateRange?.end) && (
                    <div className="flex items-center gap-1 ml-1">
                      {(["split", "combine"] as const).map((m) => {
                        const active = (cfg.dateGroupMode ?? "split") === m;
                        return (
                          <button
                            key={m}
                            onClick={() => setDateGroupMode(cfg.col, m)}
                            className={`px-2.5 py-1 rounded-md border text-xs font-medium transition-all capitalize ${
                              active
                                ? "bg-emerald-950 text-emerald-300 border-emerald-800 shadow-sm"
                                : "bg-transparent text-slate-600 border-slate-800 hover:border-slate-700 hover:text-slate-400"
                            }`}
                          >
                            {m}
                          </button>
                        );
                      })}
                    </div>
                  )}
                </div>
              )}

              {/* Value selector for "Filter" mode on non-date text columns */}
              {cfg.mode === "select" && !showDateRange && !stats.isNumeric && (
                <ValueDropdown
                  label={cfg.col}
                  values={stats.uniqueValues}
                  selected={cfg.selectedValues}
                  onChange={(vals) => setSelected(cfg.col, vals)}
                />
              )}

              {/* Numeric range inputs — mirrors date range UI */}
              {cfg.mode === "select" && stats.isNumeric && (
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="text-xs text-slate-500">From</span>
                  <input
                    type="number"
                    value={cfg.numericRange?.min ?? ""}
                    onChange={(e) =>
                      setNumericRange(cfg.col, { ...cfg.numericRange, min: e.target.value || undefined })
                    }
                    placeholder="min"
                    className="w-24 bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs text-slate-200 focus:outline-none focus:border-amber-600 [appearance:textfield] [&::-webkit-outer-spin-button]:appearance-none [&::-webkit-inner-spin-button]:appearance-none"
                  />
                  <span className="text-xs text-slate-500">To</span>
                  <input
                    type="number"
                    value={cfg.numericRange?.max ?? ""}
                    onChange={(e) =>
                      setNumericRange(cfg.col, { ...cfg.numericRange, max: e.target.value || undefined })
                    }
                    placeholder="max"
                    className="w-24 bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs text-slate-200 focus:outline-none focus:border-amber-600 [appearance:textfield] [&::-webkit-outer-spin-button]:appearance-none [&::-webkit-inner-spin-button]:appearance-none"
                  />
                  {(cfg.numericRange?.min || cfg.numericRange?.max) && (
                    <div className="flex items-center gap-1 ml-1">
                      {(["split", "combine"] as const).map((m) => {
                        const active = (cfg.numericGroupMode ?? "split") === m;
                        return (
                          <button
                            key={m}
                            onClick={() => setNumericGroupMode(cfg.col, m)}
                            className={`px-2.5 py-1 rounded-md border text-xs font-medium transition-all capitalize ${
                              active
                                ? "bg-amber-950 text-amber-300 border-amber-800 shadow-sm"
                                : "bg-transparent text-slate-600 border-slate-800 hover:border-slate-700 hover:text-slate-400"
                            }`}
                          >
                            {m}
                          </button>
                        );
                      })}
                    </div>
                  )}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}
