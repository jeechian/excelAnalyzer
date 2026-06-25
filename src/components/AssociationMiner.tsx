"use client";
import React, { useState, useMemo, useCallback } from "react";
import { Network, Play, Plus, X, Download, ChevronDown, RotateCcw } from "lucide-react";
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

  const selectableCols = useMemo(
    () =>
      data.headers.filter(
        (h) => !colStats[h].isDate && !(data.dateColumns ?? []).includes(h)
      ),
    [data, colStats]
  );

  const categoricalCols = useMemo(
    () => selectableCols.filter((h) => !colStats[h].isNumeric),
    [selectableCols, colStats]
  );

  // Dimensions
  const [dims, setDims] = useState<DimState[]>(() => {
    const first = categoricalCols[0];
    return first ? [{ id: uid(), col: first, mode: "groupby", filterValues: [] }] : [];
  });

  // Metric
  const firstNumeric = selectableCols.find((h) => colStats[h].isNumeric);
  const [metricCol, setMetricCol] = useState<string>("__count__");
  const [metricFilter, setMetricFilter] = useState<string[]>([]);

  const metricIsNumeric = metricCol !== "__count__" && numericColsSet.has(metricCol);
  const metricIsCategorical = metricCol !== "__count__" && !metricIsNumeric;

  // Results
  const [result, setResult] = useState<ContributionResult | null>(null);
  const [running, setRunning] = useState(false);
  const [visible, setVisible] = useState(PAGE_SIZE);

  const addDim = () => {
    const used = new Set(dims.map((d) => d.col));
    const next = selectableCols.find((c) => !used.has(c));
    if (!next) return;
    setDims((prev) => [...prev, { id: uid(), col: next, mode: "groupby", filterValues: [] }]);
  };

  const removeDim = (id: string) => setDims((prev) => prev.filter((d) => d.id !== id));

  const updateDim = (id: string, patch: Partial<DimState>) =>
    setDims((prev) => prev.map((d) => (d.id === id ? { ...d, ...patch } : d)));

  const resetParams = () => {
    const first = categoricalCols[0];
    setDims(first ? [{ id: uid(), col: first, mode: "groupby", filterValues: [] }] : []);
    setMetricCol("__count__");
    setMetricFilter([]);
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
      }));
      const metricConfig: MetricConfig = {
        col: metricCol,
        filterValues: metricIsCategorical ? metricFilter : [],
      };
      const res = computeContribution(data.rows, dimConfigs, metricConfig, numericColsSet);
      setResult(res);
      setRunning(false);
    }, 10);
  }, [data.rows, dims, metricCol, metricFilter, metricIsCategorical, numericColsSet]);

  const handleExport = () => {
    if (!result) return;
    const metricLabel =
      metricCol === "__count__"
        ? "Count"
        : metricIsNumeric
        ? `${metricCol} Sum`
        : `${metricCol} Count`;
    const exportRows = result.pivotValues
      ? result.rows.map((r) => ({
          ...Object.fromEntries(effectiveGroupbyCols.map((c) => [c, r.labels[c] ?? ""])),
          ...Object.fromEntries(result.pivotValues!.map((v) => [v, r.breakdown?.[v] ?? 0])),
          "% of Total": parseFloat(r.percentage.toFixed(2)),
        }))
      : result.rows.map((r) => ({
          ...Object.fromEntries(effectiveGroupbyCols.map((c) => [c, r.labels[c] ?? ""])),
          [metricLabel]: r.value,
          "% of Total": parseFloat(r.percentage.toFixed(2)),
        }));
    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Contribution");
    XLSX.writeFile(wb, "contribution.xlsx");
  };

  const isExpandDim = (d: DimState) =>
    d.mode === "filter" && !d.numericRange && d.filterValues.length >= 2;
  const effectiveGroupbyCols = dims
    .filter((d) => d.mode === "groupby" || isExpandDim(d))
    .map((d) => d.col);
  const shownFilterDims = dims.filter(
    (d) =>
      d.mode === "filter" &&
      !isExpandDim(d) &&
      (d.filterValues.length > 0 ||
        (d.numericRange &&
          (d.numericRange.min !== undefined || d.numericRange.max !== undefined)))
  );

  const metricHeaderLabel =
    metricCol === "__count__"
      ? "Count"
      : metricIsNumeric
      ? `${metricCol} (sum)`
      : metricFilter.length > 0
      ? `${metricCol} = ${metricFilter.join(" / ")} (count)`
      : `${metricCol} (count)`;

  const totalLabel =
    result == null
      ? ""
      : `Total: ${fmtNum(result.total, !result.isNumericMetric)}`;

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
              const badgeColor = COL_BADGE_COLORS[i % COL_BADGE_COLORS.length];
              const filterActive = COL_FILTER_ACTIVE[i % COL_FILTER_ACTIVE.length];

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
                        updateDim(dim.id, { col: e.target.value, filterValues: [] })
                      }
                      className="appearance-none bg-slate-900 border border-slate-700 rounded-lg pl-3 pr-7 py-1.5 text-xs text-slate-200 focus:outline-none focus:border-violet-600 cursor-pointer"
                    >
                      {selectableCols.map((c) => (
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

                  {/* Filter: numeric range inputs */}
                  {dim.mode === "filter" && numericColsSet.has(dim.col) && (
                    <div className="flex items-center gap-2 flex-wrap mt-0.5">
                      <span className="text-xs text-slate-500">From</span>
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
                        className="w-20 bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs text-slate-200 focus:outline-none focus:border-violet-600 [appearance:textfield] [&::-webkit-outer-spin-button]:appearance-none [&::-webkit-inner-spin-button]:appearance-none"
                      />
                      <span className="text-xs text-slate-500">To</span>
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
                        className="w-20 bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs text-slate-200 focus:outline-none focus:border-violet-600 [appearance:textfield] [&::-webkit-outer-spin-button]:appearance-none [&::-webkit-inner-spin-button]:appearance-none"
                      />
                    </div>
                  )}

                  {/* Filter: categorical value picker */}
                  {dim.mode === "filter" && !numericColsSet.has(dim.col) && (
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

                  {/* Group by numeric: optional range + split/combine */}
                  {dim.mode === "groupby" && numericColsSet.has(dim.col) && (
                    <div className="flex items-center gap-2 flex-wrap mt-0.5">
                      <span className="text-[11px] text-slate-600">Range (optional):</span>
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
                        className="w-20 bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs text-slate-200 focus:outline-none focus:border-violet-600 [appearance:textfield] [&::-webkit-outer-spin-button]:appearance-none [&::-webkit-inner-spin-button]:appearance-none"
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
                        className="w-20 bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs text-slate-200 focus:outline-none focus:border-violet-600 [appearance:textfield] [&::-webkit-outer-spin-button]:appearance-none [&::-webkit-inner-spin-button]:appearance-none"
                      />
                      {(dim.numericRange?.min !== undefined || dim.numericRange?.max !== undefined) && (
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
            disabled={dims.length >= selectableCols.length}
            className="mt-3 flex items-center gap-1.5 text-xs text-slate-500 hover:text-slate-300 border border-dashed border-slate-700 hover:border-slate-600 rounded-lg px-3 py-1.5 transition-all disabled:opacity-40 disabled:cursor-not-allowed"
          >
            <Plus className="w-3 h-3" />
            Add dimension
          </button>
        </div>

        {/* ── Metric ── */}
        <div className="pt-4 border-t border-slate-800/60">
          <p className="text-xs font-medium text-slate-400 uppercase tracking-wider mb-3">
            Metric — what to measure
          </p>

          <div className="flex flex-wrap items-start gap-3">
            {/* Column picker */}
            <div className="relative">
              <select
                value={metricCol}
                onChange={(e) => {
                  setMetricCol(e.target.value);
                  setMetricFilter([]);
                }}
                className="appearance-none bg-slate-900 border border-slate-700 rounded-lg pl-3 pr-7 py-1.5 text-xs text-slate-200 focus:outline-none focus:border-amber-600 cursor-pointer"
              >
                <option value="__count__">Row count</option>
                {selectableCols.map((h) => (
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

            {/* Mode hint */}
            {metricIsNumeric && (
              <span className="text-xs text-amber-300/80 mt-1.5">
                → Sum of {metricCol} per group
              </span>
            )}
            {metricCol === "__count__" && (
              <span className="text-xs text-slate-500 mt-1.5">→ Count of rows per group</span>
            )}

            {/* Categorical metric: optional value filter */}
            {metricIsCategorical && (
              <div className="space-y-2">
                <p className="text-[11px] text-slate-500">
                  Select specific values to count — or leave empty to count all rows:
                </p>
                <div className="flex flex-wrap gap-1 max-w-lg">
                  {(colStats[metricCol]?.uniqueValues ?? []).slice(0, 40).map((v) => {
                    const active = metricFilter.includes(v);
                    return (
                      <button
                        key={v}
                        onClick={() =>
                          setMetricFilter((prev) =>
                            active ? prev.filter((x) => x !== v) : [...prev, v]
                          )
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
                {metricFilter.length > 0 ? (
                  <p className="text-[11px] text-amber-300/70">
                    → Count rows where {metricCol} = {metricFilter.join(" or ")}
                    {" · "}% relative to all such rows in dataset
                  </p>
                ) : (
                  <p className="text-[11px] text-slate-600">
                    → Count all rows · % relative to total row count
                  </p>
                )}
              </div>
            )}
          </div>
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
          {dims.length === 0 && (
            <span className="text-xs text-slate-500">Add at least one dimension first.</span>
          )}
        </div>

        {/* ── Results ── */}
        {result !== null && (
          <div className="pt-4 border-t border-slate-800/60 space-y-3">
            {/* Summary line */}
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
                {/* Applied filters */}
                {shownFilterDims.length > 0 && (
                  <div className="flex items-center gap-1.5 flex-wrap">
                    <span className="text-[11px] text-slate-500">Filtered to:</span>
                    {shownFilterDims.map((d) => {
                      const lbl = d.numericRange
                        ? [
                            d.numericRange.min !== undefined ? `≥ ${d.numericRange.min}` : "",
                            d.numericRange.max !== undefined ? `≤ ${d.numericRange.max}` : "",
                          ].filter(Boolean).join(" & ")
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
              <div className="rounded-lg border border-slate-800 overflow-hidden">
                <table className="w-full">
                  <thead>
                    <tr className="bg-slate-900 border-b border-slate-800">
                      {effectiveGroupbyCols.map((c, i) => (
                        <th
                          key={c}
                          className={`px-4 py-2.5 text-left text-[11px] font-medium uppercase tracking-wider ${
                            COL_HEADER_COLORS[i % COL_HEADER_COLORS.length]
                          }`}
                        >
                          {c}
                        </th>
                      ))}
                      {result.pivotValues ? (
                        result.pivotValues.map((v) => (
                          <th key={v} className="px-4 py-2.5 text-right text-[11px] font-medium text-amber-400 uppercase tracking-wider">
                            {v}
                          </th>
                        ))
                      ) : (
                        <th className="px-4 py-2.5 text-right text-[11px] font-medium text-amber-400 uppercase tracking-wider">
                          {metricHeaderLabel}
                        </th>
                      )}
                      <th className="px-4 py-2.5 text-right text-[11px] font-medium text-slate-400 uppercase tracking-wider w-44">
                        % of Total
                      </th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/40">
                    {result.rows.slice(0, visible).map((row, i) => (
                      <tr key={i} className="hover:bg-slate-900/50 transition-colors">
                        {effectiveGroupbyCols.map((c, ci) => (
                          <td
                            key={c}
                            className={`px-4 py-2.5 text-sm font-medium ${
                              COL_CELL_COLORS[ci % COL_CELL_COLORS.length]
                            }`}
                          >
                            {row.labels[c] ?? "—"}
                          </td>
                        ))}
                        {result.pivotValues ? (
                          result.pivotValues.map((v) => (
                            <td key={v} className="px-4 py-2.5 text-right font-mono text-xs text-amber-200">
                              {fmtNum(row.breakdown?.[v] ?? 0, true)}
                            </td>
                          ))
                        ) : (
                          <td className="px-4 py-2.5 text-right font-mono text-xs text-amber-200">
                            {fmtNum(row.value, !result.isNumericMetric)}
                          </td>
                        )}
                        <td className="px-4 py-2.5">
                          <div className="flex items-center justify-end gap-2">
                            <div className="w-20 bg-slate-800 rounded-full h-1.5 flex-shrink-0 overflow-hidden">
                              <div
                                className="h-full bg-violet-500 rounded-full"
                                style={{ width: `${Math.min(100, row.percentage)}%` }}
                              />
                            </div>
                            <span className="font-mono text-xs text-slate-200 w-14 text-right tabular-nums">
                              {row.percentage.toFixed(1)}%
                            </span>
                          </div>
                        </td>
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
                        result.pivotValues.map((v) => (
                          <td key={v} className="px-4 py-2 text-right font-mono text-xs text-amber-300 font-semibold">
                            {fmtNum(result.rows.reduce((a, r) => a + (r.breakdown?.[v] ?? 0), 0), true)}
                          </td>
                        ))
                      ) : (
                        <td className="px-4 py-2 text-right font-mono text-xs text-amber-300 font-semibold">
                          {fmtNum(result.total, !result.isNumericMetric)}
                        </td>
                      )}
                      <td className="px-4 py-2 text-right font-mono text-xs text-slate-500">
                        100%
                      </td>
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
    </div>
  );
}
