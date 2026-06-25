"use client";
import React, { useState, useMemo, useCallback } from "react";
import { Network, Play, Plus, X, Download, ChevronDown, RotateCcw, HelpCircle } from "lucide-react";
import * as XLSX from "xlsx";
import { getColumnStats, type ExcelData } from "@/lib/excel";
import {
  computeContribution,
  type DimensionConfig,
  type MetricConfig,
  type ContributionResult,
} from "@/lib/association";

interface Props {
  data: ExcelData;
}

const COL_HEADER_COLORS = ["text-violet-400", "text-sky-400", "text-emerald-400"];
const COL_CELL_COLORS = ["text-violet-200", "text-sky-200", "text-emerald-200"];
const COL_BADGE_COLORS = [
  "bg-violet-950 text-violet-300 border-violet-800",
  "bg-sky-950 text-sky-300 border-sky-800",
  "bg-emerald-950 text-emerald-300 border-emerald-800",
];
const COL_FILTER_ACTIVE = [
  "bg-violet-950 text-violet-300 border-violet-700",
  "bg-sky-950 text-sky-300 border-sky-700",
  "bg-emerald-950 text-emerald-300 border-emerald-700",
];

const DATE_INPUT_CLS =
  "bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs text-slate-200 focus:outline-none focus:border-emerald-600 [color-scheme:dark]";
const NUM_INPUT_CLS =
  "w-20 bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs text-slate-200 focus:outline-none focus:border-violet-600 [appearance:textfield] [&::-webkit-outer-spin-button]:appearance-none [&::-webkit-inner-spin-button]:appearance-none";
const NUM_INPUT_AMBER =
  "w-20 bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs text-slate-200 focus:outline-none focus:border-amber-600 [appearance:textfield] [&::-webkit-outer-spin-button]:appearance-none [&::-webkit-inner-spin-button]:appearance-none";

const PAGE_SIZE = 20;
let _id = 0;
const uid = () => String(++_id);

interface DimState {
  id: string;
  col: string;
  mode: "groupby" | "filter";
  filterValues: string[];
  numericRange?: { min?: number; max?: number };
  numericGroupMode?: "split" | "combine";
  dateRange?: { start?: string; end?: string };
  dateGroupMode?: "split" | "combine";
}

interface MetricState {
  id: string;
  col: string; // "__count__" or column name
  filterValues: string[];
  numericRange?: { min?: number; max?: number }; // for numeric metric: filter groups by sum range
  dateFilterCol?: string;
  dateRange?: { start?: string; end?: string };   // for date metric: defines date window
}

function fmtNum(n: number, isCount: boolean): string {
  if (isCount || Number.isInteger(n)) return n.toLocaleString();
  return n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

export default function AssociationMiner({ data }: Props) {
  const colStats = useMemo(
    () => Object.fromEntries(data.headers.map((h) => [h, getColumnStats(data.rows, h)])),
    [data]
  );

  const numericColsSet = useMemo(
    () => new Set(data.headers.filter((h) => colStats[h].isNumeric)),
    [data, colStats]
  );

  const dateCols = useMemo(
    () => data.headers.filter((h) => colStats[h].isDate || (data.dateColumns ?? []).includes(h)),
    [data, colStats]
  );

  const allCols = useMemo(() => data.headers, [data]);

  const categoricalCols = useMemo(
    () =>
      data.headers.filter(
        (h) => !colStats[h].isNumeric && !colStats[h].isDate && !(data.dateColumns ?? []).includes(h)
      ),
    [data, colStats]
  );

  const isDateCol = useCallback(
    (col: string) => colStats[col]?.isDate || (data.dateColumns ?? []).includes(col),
    [colStats, data.dateColumns]
  );

  // ── State ──
  const [dims, setDims] = useState<DimState[]>(() => {
    const first = categoricalCols[0];
    return first ? [{ id: uid(), col: first, mode: "groupby", filterValues: [] }] : [];
  });

  const [metrics, setMetrics] = useState<MetricState[]>([
    { id: uid(), col: "__count__", filterValues: [] },
  ]);

  const [result, setResult] = useState<ContributionResult | null>(null);
  const [running, setRunning] = useState(false);
  const [visible, setVisible] = useState(PAGE_SIZE);
  const [showHelp, setShowHelp] = useState(false);

  // ── Dim helpers ──
  const addDim = () => {
    const used = new Set(dims.map((d) => d.col));
    const next = allCols.find((c) => !used.has(c));
    if (!next) return;
    setDims((prev) => [...prev, { id: uid(), col: next, mode: "groupby", filterValues: [] }]);
  };
  const removeDim = (id: string) => setDims((prev) => prev.filter((d) => d.id !== id));
  const updateDim = (id: string, patch: Partial<DimState>) =>
    setDims((prev) => prev.map((d) => (d.id === id ? { ...d, ...patch } : d)));

  // ── Metric helpers ──
  const addMetric = () =>
    setMetrics((prev) => [...prev, { id: uid(), col: "__count__", filterValues: [] }]);
  const removeMetric = (id: string) => setMetrics((prev) => prev.filter((m) => m.id !== id));
  const updateMetric = (id: string, patch: Partial<MetricState>) =>
    setMetrics((prev) => prev.map((m) => (m.id === id ? { ...m, ...patch } : m)));

  const resetParams = () => {
    const first = categoricalCols[0];
    setDims(first ? [{ id: uid(), col: first, mode: "groupby", filterValues: [] }] : []);
    setMetrics([{ id: uid(), col: "__count__", filterValues: [] }]);
    setResult(null);
  };

  const run = useCallback(() => {
    setRunning(true);
    setResult(null);
    setVisible(PAGE_SIZE);
    setTimeout(() => {
      const dimConfigs: DimensionConfig[] = dims.map((d) => ({
        col: d.col,
        mode: d.mode,
        filterValues: d.filterValues,
        numericRange: d.numericRange,
        numericGroupMode: d.numericGroupMode,
        dateRange: d.dateRange,
        dateGroupMode: d.dateGroupMode,
      }));
      const metricConfigs: MetricConfig[] = metrics.map((m) => {
        const isDateMetricCol = m.col !== "__count__" && isDateCol(m.col);
        return {
          id: m.id,
          col: m.col,
          filterValues: m.filterValues,
          numericRange: m.numericRange,
          dateFilterCol: isDateMetricCol ? m.col : m.dateFilterCol,
          dateRange: m.dateRange,
        };
      });
      const res = computeContribution(data.rows, dimConfigs, metricConfigs, numericColsSet);
      setResult(res);
      setRunning(false);
    }, 10);
  }, [data.rows, dims, metrics, numericColsSet]);

  // ── Derived display values ──
  const effectiveGroupbyCols = dims
    .filter((d) => {
      if (d.mode === "groupby") return true;
      // categorical expand: filter with ≥2 values, no range constraints
      if (d.mode === "filter" && !d.numericRange && !d.dateRange && d.filterValues.length >= 2) return true;
      return false;
    })
    .map((d) => d.col);

  const shownFilterDims = dims.filter((d) => {
    if (d.mode !== "filter") return false;
    if (!d.numericRange && !d.dateRange && d.filterValues.length >= 2) return false;
    return (
      d.filterValues.length > 0 ||
      (d.numericRange && (d.numericRange.min !== undefined || d.numericRange.max !== undefined)) ||
      (d.dateRange && (d.dateRange.start || d.dateRange.end))
    );
  });

  const totalLabel =
    result == null
      ? ""
      : result.metricMetas.map((m) => `${m.label}: ${fmtNum(m.total, !m.isNumeric)}`).join(" · ");

  // ── Export ──
  const handleExport = () => {
    if (!result) return;
    const exportRows = result.rows.map((r) => {
      const obj: Record<string, unknown> = {};
      for (const c of effectiveGroupbyCols) obj[c] = r.labels[c] ?? "";
      if (result.pivotValues) {
        for (const v of result.pivotValues) obj[v] = r.breakdown?.[v] ?? 0;
        obj["% of Total"] = parseFloat(
          (r.percentages[result.metricMetas[0].id] ?? 0).toFixed(2)
        );
      } else {
        for (const meta of result.metricMetas) {
          obj[meta.label] = r.values[meta.id] ?? 0;
          obj[`${meta.label} %`] = parseFloat((r.percentages[meta.id] ?? 0).toFixed(2));
        }
      }
      return obj;
    });
    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Contribution");
    XLSX.writeFile(wb, "contribution.xlsx");
  };

  // ─── Render ─────────────────────────────────────────────────────────────────

  return (
    <div className="rounded-xl border border-slate-800 overflow-hidden">
      {/* Header */}
      <div className="px-5 py-4 bg-slate-900 border-b border-slate-800">
        <div className="flex items-center gap-3">
          <Network className="w-4 h-4 text-violet-400" />
          <span className="font-semibold text-slate-200 text-sm">Contribution Analysis</span>
          <span className="text-xs text-slate-500">
            See how much each group contributes to a metric
          </span>
        </div>
      </div>

      <div className="p-5 space-y-5 bg-slate-950">
        {/* ── Dimensions ── */}
        <div>
          <p className="text-xs font-medium text-slate-400 uppercase tracking-wider mb-3">
            Dimensions
          </p>

          <div className="space-y-3">
            {dims.map((dim, i) => {
              const stats = colStats[dim.col];
              const isDate = isDateCol(dim.col);
              const isNumeric = numericColsSet.has(dim.col);
              const badgeColor = COL_BADGE_COLORS[i % COL_BADGE_COLORS.length];
              const filterActive = COL_FILTER_ACTIVE[i % COL_FILTER_ACTIVE.length];
              const hasDateRange = dim.dateRange?.start || dim.dateRange?.end;
              const hasNumericRange =
                dim.numericRange?.min !== undefined || dim.numericRange?.max !== undefined;

              return (
                <div key={dim.id} className="flex items-start gap-2 flex-wrap">
                  {/* Index badge */}
                  <span
                    className={`mt-1.5 w-5 h-5 rounded text-[10px] font-bold flex items-center justify-center flex-shrink-0 border ${badgeColor}`}
                  >
                    {i + 1}
                  </span>

                  {/* Column picker */}
                  <div className="relative">
                    <select
                      value={dim.col}
                      onChange={(e) =>
                        updateDim(dim.id, {
                          col: e.target.value,
                          filterValues: [],
                          numericRange: undefined,
                          dateRange: undefined,
                        })
                      }
                      className="appearance-none bg-slate-900 border border-slate-700 rounded-lg pl-3 pr-7 py-1.5 text-xs text-slate-200 focus:outline-none focus:border-violet-600 cursor-pointer"
                    >
                      {allCols.map((c) => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))}
                    </select>
                    <ChevronDown className="absolute right-2 top-1/2 -translate-y-1/2 w-3 h-3 text-slate-500 pointer-events-none" />
                  </div>

                  {/* Mode toggle */}
                  <div className="flex items-center gap-1 mt-0.5">
                    {(["groupby", "filter"] as const).map((m) => (
                      <button
                        key={m}
                        onClick={() => updateDim(dim.id, { mode: m, filterValues: [] })}
                        className={`px-2.5 py-1 rounded-md border text-xs font-medium transition-all ${
                          dim.mode === m
                            ? m === "groupby"
                              ? "bg-violet-950 text-violet-300 border-violet-700"
                              : "bg-blue-950 text-blue-300 border-blue-700"
                            : "bg-transparent text-slate-500 border-slate-700 hover:text-slate-400 hover:border-slate-600"
                        }`}
                      >
                        {m === "groupby" ? "Group by" : "Filter"}
                      </button>
                    ))}
                  </div>

                  {/* Date column controls */}
                  {isDate && (
                    <div className="flex items-center gap-2 flex-wrap mt-0.5">
                      <span className="text-xs text-slate-500">From</span>
                      <input
                        type="date"
                        value={dim.dateRange?.start ?? ""}
                        onChange={(e) =>
                          updateDim(dim.id, {
                            dateRange: { ...dim.dateRange, start: e.target.value || undefined },
                          })
                        }
                        className={DATE_INPUT_CLS}
                      />
                      <span className="text-xs text-slate-500">To</span>
                      <input
                        type="date"
                        value={dim.dateRange?.end ?? ""}
                        onChange={(e) =>
                          updateDim(dim.id, {
                            dateRange: { ...dim.dateRange, end: e.target.value || undefined },
                          })
                        }
                        className={DATE_INPUT_CLS}
                      />
                      {/* combine/split toggle for groupby date (not for filter mode) */}
                      {dim.mode === "groupby" && hasDateRange && (
                        <div className="flex items-center gap-1 ml-1">
                          {(["split", "combine"] as const).map((m) => {
                            const active = (dim.dateGroupMode ?? "split") === m;
                            return (
                              <button
                                key={m}
                                onClick={() => updateDim(dim.id, { dateGroupMode: m })}
                                className={`px-2.5 py-1 rounded-md border text-[11px] font-medium transition-all capitalize ${
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

                  {/* Numeric column controls */}
                  {isNumeric && (
                    <div className="flex items-center gap-2 flex-wrap mt-0.5">
                      {dim.mode === "groupby" && (
                        <span className="text-[11px] text-slate-600">Range (optional):</span>
                      )}
                      <input
                        type="number"
                        value={dim.numericRange?.min ?? ""}
                        onChange={(e) =>
                          updateDim(dim.id, {
                            numericRange: {
                              ...dim.numericRange,
                              min: e.target.value !== "" ? Number(e.target.value) : undefined,
                            },
                          })
                        }
                        placeholder="min"
                        className={NUM_INPUT_CLS}
                      />
                      <span className="text-xs text-slate-500">–</span>
                      <input
                        type="number"
                        value={dim.numericRange?.max ?? ""}
                        onChange={(e) =>
                          updateDim(dim.id, {
                            numericRange: {
                              ...dim.numericRange,
                              max: e.target.value !== "" ? Number(e.target.value) : undefined,
                            },
                          })
                        }
                        placeholder="max"
                        className={NUM_INPUT_CLS}
                      />
                      {dim.mode === "groupby" && hasNumericRange && (
                        <div className="flex items-center gap-1">
                          {(["split", "combine"] as const).map((m) => {
                            const active = (dim.numericGroupMode ?? "split") === m;
                            return (
                              <button
                                key={m}
                                onClick={() => updateDim(dim.id, { numericGroupMode: m })}
                                className={`px-2.5 py-1 rounded-md border text-[11px] font-medium transition-all capitalize ${
                                  active
                                    ? "bg-violet-950 text-violet-300 border-violet-700"
                                    : "bg-transparent text-slate-500 border-slate-700 hover:border-slate-600 hover:text-slate-400"
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

                  {/* Categorical filter: value chips */}
                  {!isDate && !isNumeric && dim.mode === "filter" && (
                    <div className="flex flex-wrap gap-1 mt-0.5 max-w-lg">
                      {(stats?.uniqueValues ?? []).slice(0, 40).map((v) => {
                        const active = dim.filterValues.includes(v);
                        return (
                          <button
                            key={v}
                            onClick={() =>
                              updateDim(dim.id, {
                                filterValues: active
                                  ? dim.filterValues.filter((x) => x !== v)
                                  : [...dim.filterValues, v],
                              })
                            }
                            className={`px-2 py-0.5 rounded border text-[11px] font-medium transition-all ${
                              active
                                ? filterActive
                                : "bg-transparent text-slate-500 border-slate-700 hover:border-slate-600 hover:text-slate-400"
                            }`}
                          >
                            {v}
                          </button>
                        );
                      })}
                    </div>
                  )}

                  {/* Remove */}
                  <button
                    onClick={() => removeDim(dim.id)}
                    className="mt-1.5 p-1 rounded text-slate-600 hover:text-slate-300 transition-colors flex-shrink-0"
                  >
                    <X className="w-3.5 h-3.5" />
                  </button>
                </div>
              );
            })}
          </div>

          <button
            onClick={addDim}
            disabled={dims.length >= allCols.length}
            className="mt-3 flex items-center gap-1.5 text-xs text-slate-500 hover:text-slate-300 border border-dashed border-slate-700 hover:border-slate-600 rounded-lg px-3 py-1.5 transition-all disabled:opacity-40 disabled:cursor-not-allowed"
          >
            <Plus className="w-3 h-3" />
            Add dimension
          </button>
        </div>

        {/* ── Metrics ── */}
        <div className="pt-4 border-t border-slate-800/60">
          <p className="text-xs font-medium text-slate-400 uppercase tracking-wider mb-3">
            Metrics — what to measure (defines 100%)
          </p>

          <div className="space-y-5">
            {metrics.map((m, mi) => {
              const isCount = m.col === "__count__";
              const isDateMetric = !isCount && isDateCol(m.col);
              const isNumericM = !isCount && !isDateMetric && numericColsSet.has(m.col);
              const isCategorical = !isCount && !isNumericM && !isDateMetric;
              const stats = colStats[m.col];

              return (
                <div key={m.id} className="space-y-2">
                  <div className="flex items-start gap-2 flex-wrap">
                    {/* Index badge */}
                    <span className="mt-1.5 w-5 h-5 rounded text-[10px] font-bold flex items-center justify-center flex-shrink-0 border bg-amber-950 text-amber-300 border-amber-800">
                      {mi + 1}
                    </span>

                    {/* Column picker */}
                    <div className="relative">
                      <select
                        value={m.col}
                        onChange={(e) =>
                          updateMetric(m.id, {
                            col: e.target.value,
                            filterValues: [],
                            numericRange: undefined,
                            dateRange: undefined,
                          })
                        }
                        className="appearance-none bg-slate-900 border border-slate-700 rounded-lg pl-3 pr-7 py-1.5 text-xs text-slate-200 focus:outline-none focus:border-amber-600 cursor-pointer"
                      >
                        <option value="__count__">Row count</option>
                        {allCols.map((h) => (
                          <option key={h} value={h}>
                            {h}
                            {colStats[h].isNumeric
                              ? " (numeric)"
                              : ` — ${colStats[h].uniqueCount} values`}
                          </option>
                        ))}
                      </select>
                      <ChevronDown className="absolute right-2 top-1/2 -translate-y-1/2 w-3 h-3 text-slate-500 pointer-events-none" />
                    </div>

                    {isNumericM && (
                      <span className="text-xs text-amber-300/80 mt-1.5">
                        → Sum of {m.col} per group
                      </span>
                    )}
                    {isCount && (
                      <span className="text-xs text-slate-500 mt-1.5">
                        → Count of rows per group
                      </span>
                    )}
                    {isDateMetric && (
                      <span className="text-xs text-emerald-400/80 mt-1.5">
                        → Count of rows per group (filtered by date range)
                      </span>
                    )}

                    {/* Remove */}
                    {metrics.length > 1 && (
                      <button
                        onClick={() => removeMetric(m.id)}
                        className="mt-1.5 p-1 rounded text-slate-600 hover:text-slate-300 transition-colors flex-shrink-0"
                      >
                        <X className="w-3.5 h-3.5" />
                      </button>
                    )}
                  </div>

                  {/* Categorical: value filter chips */}
                  {isCategorical && (
                    <div className="ml-7 space-y-1.5">
                      <p className="text-[11px] text-slate-500">
                        Select values to count — or leave empty to count all rows:
                      </p>
                      <div className="flex flex-wrap gap-1 max-w-lg">
                        {(stats?.uniqueValues ?? []).slice(0, 40).map((v) => {
                          const active = m.filterValues.includes(v);
                          return (
                            <button
                              key={v}
                              onClick={() =>
                                updateMetric(m.id, {
                                  filterValues: active
                                    ? m.filterValues.filter((x) => x !== v)
                                    : [...m.filterValues, v],
                                })
                              }
                              className={`px-2 py-0.5 rounded border text-[11px] font-medium transition-all ${
                                active
                                  ? "bg-amber-950 text-amber-300 border-amber-700"
                                  : "bg-transparent text-slate-500 border-slate-700 hover:border-slate-600 hover:text-slate-400"
                              }`}
                            >
                              {v}
                            </button>
                          );
                        })}
                      </div>
                      {m.filterValues.length > 0 ? (
                        <p className="text-[11px] text-amber-300/70">
                          → Count rows where {m.col} = {m.filterValues.join(" or ")}
                        </p>
                      ) : (
                        <p className="text-[11px] text-slate-600">→ Count all rows</p>
                      )}
                    </div>
                  )}

                  {/* Numeric metric: value range filter */}
                  {isNumericM && (
                    <div className="ml-7 flex items-center gap-2 flex-wrap">
                      <span className="text-[11px] text-slate-500">Value range (optional):</span>
                      <input
                        type="number"
                        value={m.numericRange?.min ?? ""}
                        onChange={(e) =>
                          updateMetric(m.id, {
                            numericRange: {
                              ...m.numericRange,
                              min: e.target.value !== "" ? Number(e.target.value) : undefined,
                            },
                          })
                        }
                        placeholder="min"
                        className={NUM_INPUT_AMBER}
                      />
                      <span className="text-xs text-slate-500">–</span>
                      <input
                        type="number"
                        value={m.numericRange?.max ?? ""}
                        onChange={(e) =>
                          updateMetric(m.id, {
                            numericRange: {
                              ...m.numericRange,
                              max: e.target.value !== "" ? Number(e.target.value) : undefined,
                            },
                          })
                        }
                        placeholder="max"
                        className={NUM_INPUT_AMBER}
                      />
                      {(m.numericRange?.min !== undefined || m.numericRange?.max !== undefined) && (
                        <span className="text-[11px] text-amber-300/70">
                          → only show groups where sum of {m.col} is in this range
                        </span>
                      )}
                    </div>
                  )}

                  {/* Date metric: From/To pickers define the 100% base */}
                  {isDateMetric && (
                    <div className="ml-7 flex items-center gap-2 flex-wrap">
                      <span className="text-[11px] text-slate-500">From</span>
                      <input
                        type="date"
                        value={m.dateRange?.start ?? ""}
                        onChange={(e) =>
                          updateMetric(m.id, {
                            dateRange: { ...m.dateRange, start: e.target.value || undefined },
                          })
                        }
                        className={DATE_INPUT_CLS}
                      />
                      <span className="text-xs text-slate-500">To</span>
                      <input
                        type="date"
                        value={m.dateRange?.end ?? ""}
                        onChange={(e) =>
                          updateMetric(m.id, {
                            dateRange: { ...m.dateRange, end: e.target.value || undefined },
                          })
                        }
                        className={DATE_INPUT_CLS}
                      />
                      {(m.dateRange?.start || m.dateRange?.end) && (
                        <span className="text-[11px] text-emerald-400/60">
                          → defines 100%
                        </span>
                      )}
                    </div>
                  )}

                </div>
              );
            })}
          </div>

          <button
            onClick={addMetric}
            className="mt-4 flex items-center gap-1.5 text-xs text-slate-500 hover:text-slate-300 border border-dashed border-slate-700 hover:border-slate-600 rounded-lg px-3 py-1.5 transition-all"
          >
            <Plus className="w-3 h-3" />
            Add metric
          </button>
        </div>

        {/* ── Run ── */}
        <div className="flex items-center gap-3 pt-3 border-t border-slate-800/60">
          <button
            onClick={run}
            disabled={running || dims.length === 0}
            className="flex items-center gap-2 px-5 py-2 rounded-lg bg-violet-600 hover:bg-violet-500 disabled:bg-slate-800 disabled:text-slate-600 text-white text-xs font-semibold transition-all"
          >
            {running ? (
              <div className="w-3.5 h-3.5 border-2 border-white/50 border-t-white rounded-full animate-spin" />
            ) : (
              <Play className="w-3.5 h-3.5 fill-white" />
            )}
            {running ? "Running…" : "Run Analysis"}
          </button>
          <button
            onClick={resetParams}
            className="flex items-center gap-1.5 text-xs text-slate-500 hover:text-slate-300 px-3 py-2 rounded-lg border border-slate-800 hover:border-slate-700 transition-all"
          >
            <RotateCcw className="w-3 h-3" />
            Reset
          </button>
          <button
            onClick={() => setShowHelp(true)}
            className="flex items-center gap-1.5 text-xs text-slate-500 hover:text-slate-300 px-3 py-2 rounded-lg border border-slate-800 hover:border-slate-700 transition-all"
          >
            <HelpCircle className="w-3 h-3" />
            How to use
          </button>
          {dims.length === 0 && (
            <span className="text-xs text-slate-500">Add at least one dimension first.</span>
          )}
        </div>

        {/* ── Results ── */}
        {result !== null && (
          <div className="pt-4 border-t border-slate-800/60 space-y-3">
            <div className="flex items-center justify-between flex-wrap gap-2">
              <div className="space-y-1">
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="text-xs font-medium text-slate-400 uppercase tracking-wider">
                    Results
                  </span>
                  <span className="text-xs text-slate-500 font-mono">
                    {result.rows.length} group{result.rows.length !== 1 ? "s" : ""}
                    {" · "}
                    {totalLabel}
                  </span>
                </div>
                {shownFilterDims.length > 0 && (
                  <div className="flex items-center gap-1.5 flex-wrap">
                    <span className="text-[11px] text-slate-500">Filtered to:</span>
                    {shownFilterDims.map((d) => {
                      const lbl = d.numericRange
                        ? [
                            d.numericRange.min !== undefined ? `≥ ${d.numericRange.min}` : "",
                            d.numericRange.max !== undefined ? `≤ ${d.numericRange.max}` : "",
                          ]
                            .filter(Boolean)
                            .join(" & ")
                        : d.dateRange
                        ? [d.dateRange.start ?? "", d.dateRange.end ?? ""]
                            .filter(Boolean)
                            .join(" – ")
                        : d.filterValues.join(", ");
                      return (
                        <span
                          key={d.col}
                          className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md border text-[11px] bg-blue-950 text-blue-300 border-blue-800"
                        >
                          <span className="opacity-60">{d.col}</span>
                          {lbl}
                        </span>
                      );
                    })}
                  </div>
                )}
              </div>

              {result.rows.length > 0 && (
                <button
                  onClick={handleExport}
                  className="flex items-center gap-1.5 text-xs text-slate-400 hover:text-slate-200 px-3 py-1.5 rounded-lg border border-slate-700 hover:border-slate-600 transition-all"
                >
                  <Download className="w-3.5 h-3.5" />
                  Export
                </button>
              )}
            </div>

            {result.rows.length === 0 ? (
              <div className="rounded-lg border border-dashed border-slate-800 p-8 text-center">
                <p className="text-sm text-slate-500">No results. Check your filters.</p>
              </div>
            ) : (
              <div className="rounded-lg border border-slate-800 overflow-hidden overflow-x-auto">
                <table className="w-full">
                  <thead>
                    <tr className="bg-slate-900 border-b border-slate-800">
                      {effectiveGroupbyCols.map((c, i) => (
                        <th
                          key={c}
                          className={`px-4 py-2.5 text-left text-[11px] font-medium uppercase tracking-wider whitespace-nowrap ${
                            COL_HEADER_COLORS[i % COL_HEADER_COLORS.length]
                          }`}
                        >
                          {c}
                        </th>
                      ))}

                      {result.pivotValues ? (
                        <>
                          {result.pivotValues.map((v) => (
                            <th
                              key={v}
                              className="px-4 py-2.5 text-right text-[11px] font-medium text-amber-400 uppercase tracking-wider whitespace-nowrap"
                            >
                              {v}
                            </th>
                          ))}
                          <th className="px-4 py-2.5 text-right text-[11px] font-medium text-slate-400 uppercase tracking-wider w-44 whitespace-nowrap">
                            % of Total
                          </th>
                        </>
                      ) : (
                        result.metricMetas.map((meta) => (
                          <React.Fragment key={meta.id}>
                            <th className="px-4 py-2.5 text-right text-[11px] font-medium text-amber-400 uppercase tracking-wider whitespace-nowrap">
                              {meta.label}
                            </th>
                            <th className="px-4 py-2.5 text-right text-[11px] font-medium text-slate-400 uppercase tracking-wider whitespace-nowrap">
                              {result.metricMetas.length > 1 ? `${meta.label} %` : "% of Total"}
                            </th>
                          </React.Fragment>
                        ))
                      )}
                    </tr>
                  </thead>

                  <tbody className="divide-y divide-slate-800/40">
                    {result.rows.slice(0, visible).map((row, i) => (
                      <tr key={i} className="hover:bg-slate-900/50 transition-colors">
                        {effectiveGroupbyCols.map((c, ci) => (
                          <td
                            key={c}
                            className={`px-4 py-2.5 text-sm font-medium whitespace-nowrap ${
                              COL_CELL_COLORS[ci % COL_CELL_COLORS.length]
                            }`}
                          >
                            {row.labels[c] ?? "—"}
                          </td>
                        ))}

                        {result.pivotValues ? (
                          <>
                            {result.pivotValues.map((v) => (
                              <td
                                key={v}
                                className="px-4 py-2.5 text-right font-mono text-xs text-amber-200"
                              >
                                {fmtNum(row.breakdown?.[v] ?? 0, true)}
                              </td>
                            ))}
                            <td className="px-4 py-2.5">
                              <div className="flex items-center justify-end gap-2">
                                <div className="w-20 bg-slate-800 rounded-full h-1.5 flex-shrink-0 overflow-hidden">
                                  <div
                                    className="h-full bg-violet-500 rounded-full"
                                    style={{
                                      width: `${Math.min(100, row.percentages[result.metricMetas[0].id] ?? 0)}%`,
                                    }}
                                  />
                                </div>
                                <span className="font-mono text-xs text-slate-200 w-14 text-right tabular-nums">
                                  {(row.percentages[result.metricMetas[0].id] ?? 0).toFixed(1)}%
                                </span>
                              </div>
                            </td>
                          </>
                        ) : (
                          result.metricMetas.map((meta) => (
                            <React.Fragment key={meta.id}>
                              <td className="px-4 py-2.5 text-right font-mono text-xs text-amber-200">
                                {fmtNum(row.values[meta.id] ?? 0, !meta.isNumeric)}
                              </td>
                              <td className="px-4 py-2.5">
                                {result.metricMetas.length === 1 ? (
                                  <div className="flex items-center justify-end gap-2">
                                    <div className="w-20 bg-slate-800 rounded-full h-1.5 flex-shrink-0 overflow-hidden">
                                      <div
                                        className="h-full bg-violet-500 rounded-full"
                                        style={{
                                          width: `${Math.min(100, row.percentages[meta.id] ?? 0)}%`,
                                        }}
                                      />
                                    </div>
                                    <span className="font-mono text-xs text-slate-200 w-14 text-right tabular-nums">
                                      {(row.percentages[meta.id] ?? 0).toFixed(1)}%
                                    </span>
                                  </div>
                                ) : (
                                  <span className="font-mono text-xs text-slate-200 block text-right tabular-nums">
                                    {(row.percentages[meta.id] ?? 0).toFixed(1)}%
                                  </span>
                                )}
                              </td>
                            </React.Fragment>
                          ))
                        )}
                      </tr>
                    ))}
                  </tbody>

                  <tfoot>
                    <tr className="bg-slate-900/60 border-t border-slate-700">
                      {effectiveGroupbyCols.length > 0 && (
                        <td
                          colSpan={effectiveGroupbyCols.length}
                          className="px-4 py-2 text-right text-[11px] text-slate-500 font-medium uppercase tracking-wider"
                        >
                          Total
                        </td>
                      )}
                      {result.pivotValues ? (
                        <>
                          {result.pivotValues.map((v) => (
                            <td
                              key={v}
                              className="px-4 py-2 text-right font-mono text-xs text-amber-300 font-semibold"
                            >
                              {fmtNum(
                                result.rows.reduce((a, r) => a + (r.breakdown?.[v] ?? 0), 0),
                                true
                              )}
                            </td>
                          ))}
                          <td className="px-4 py-2 text-right font-mono text-xs text-slate-500">
                            100%
                          </td>
                        </>
                      ) : (
                        result.metricMetas.map((meta) => (
                          <React.Fragment key={meta.id}>
                            <td className="px-4 py-2 text-right font-mono text-xs text-amber-300 font-semibold">
                              {fmtNum(meta.total, !meta.isNumeric)}
                            </td>
                            <td className="px-4 py-2 text-right font-mono text-xs text-slate-500">
                              100%
                            </td>
                          </React.Fragment>
                        ))
                      )}
                    </tr>
                  </tfoot>
                </table>

                {result.rows.length > visible && (
                  <div className="px-4 py-3 bg-slate-900/50 border-t border-slate-800 flex items-center gap-2">
                    <button
                      onClick={() => setVisible((v) => v + PAGE_SIZE)}
                      className="px-3 py-1.5 rounded-lg border border-slate-700 text-xs text-slate-400 hover:border-slate-600 hover:text-slate-200 transition-all"
                    >
                      Show {Math.min(result.rows.length - visible, PAGE_SIZE)} more
                    </button>
                    <button
                      onClick={() => setVisible(result.rows.length)}
                      className="px-3 py-1.5 rounded-lg border border-slate-700 text-xs text-slate-400 hover:border-slate-600 hover:text-slate-200 transition-all"
                    >
                      Show all {result.rows.length}
                    </button>
                    <span className="text-xs text-slate-600">
                      {result.rows.length - visible} hidden
                    </span>
                  </div>
                )}
              </div>
            )}
          </div>
        )}
      </div>

      {/* How to use modal */}
      {showHelp && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4"
          onClick={() => setShowHelp(false)}
        >
          <div
            className="bg-slate-900 border border-slate-700 rounded-2xl max-w-lg w-full p-6 space-y-5 shadow-2xl"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between">
              <h2 className="text-sm font-semibold text-slate-100">How to use Contribution Analysis</h2>
              <button onClick={() => setShowHelp(false)} className="text-slate-500 hover:text-slate-300">
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="space-y-4 text-xs text-slate-400 leading-relaxed">
              <div>
                <p className="text-slate-200 font-medium mb-1">1. Dimensions — what column to analyze</p>
                <p><span className="text-violet-400 font-medium">Group by</span> — split results into one row per unique value (e.g. one row per Region).</p>
                <p className="mt-1"><span className="text-blue-400 font-medium">Filter</span> — restrict which rows are shown, but the 100% total stays the same. Use this to zoom into a subset without changing the denominator.</p>
              </div>

              <div>
                <p className="text-slate-200 font-medium mb-1">2. Metrics — define the 100% population</p>
                <p>Each metric adds a condition. All conditions are combined (AND) to define which rows count as 100%.</p>
                <ul className="mt-1 space-y-1 list-disc list-inside">
                  <li><span className="text-emerald-400">Date column</span> — set a From / To date range.</li>
                  <li><span className="text-amber-400">Numeric column</span> — set a min / max value range.</li>
                  <li><span className="text-amber-400">Categorical column</span> — pick specific values to include.</li>
                </ul>
                <p className="mt-1 text-slate-500">Example: Date Jan–Dec AND ShippingCost ≥ 1000 → only rows matching both conditions form the 100%.</p>
              </div>

              <div>
                <p className="text-slate-200 font-medium mb-1">3. Run Analysis</p>
                <p>Each group shows its row count and what percentage it contributes to the total population defined by your metrics.</p>
              </div>

              <div className="pt-2 border-t border-slate-800 text-slate-500">
                Click anywhere outside this box to close.
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
