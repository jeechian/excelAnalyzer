"use client";
import { BarChart3, Download } from "lucide-react";
import { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import type { SummaryRow, ColumnConfig } from "@/lib/excel";

interface SummaryTableProps {
  results: SummaryRow[];
  configs: ColumnConfig[];
  numericCols: string[];
}

function fmt(n: number): string {
  if (Number.isInteger(n)) return n.toLocaleString();
  return n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatDateRangeLabel(dateRange: { start?: string; end?: string }): string {
  const fmtDate = (s: string) => {
    const [y, m, d] = s.split("-").map(Number);
    return new Date(y, m - 1, d).toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" });
  };
  const { start, end } = dateRange;
  if (start && end) return `${fmtDate(start)} — ${fmtDate(end)}`;
  if (start) return `From ${fmtDate(start)}`;
  if (end) return `Until ${fmtDate(end)}`;
  return "";
}

function formatNumericRangeLabel(range: { min?: string; max?: string }): string {
  const { min, max } = range;
  if (min && max) return `${min} – ${max}`;
  if (min) return `≥ ${min}`;
  if (max) return `≤ ${max}`;
  return "";
}

type ColKind = "single" | "multi" | "daterange" | "numericrange" | "separate";
interface DisplayCol { col: string; kind: ColKind; label?: string }

const PAGE_SIZE = 10;

export default function SummaryTable({ results, configs, numericCols }: SummaryTableProps) {
  const [visible, setVisible] = useState(PAGE_SIZE);

  useEffect(() => { setVisible(PAGE_SIZE); }, [results]);

  // Build display columns in config order — skipping ignored and numeric columns
  const displayCols: DisplayCol[] = configs.flatMap((c): DisplayCol[] => {
    if (c.mode === "ignore" || numericCols.includes(c.col)) return [];
    if (c.mode === "select") {
      if (c.selectedValues.length === 1) return [{ col: c.col, kind: "single", label: c.selectedValues[0] }];
      if (c.selectedValues.length > 1)   return [{ col: c.col, kind: "multi" }];
      if (c.dateRange && (c.dateRange.start || c.dateRange.end)) {
        return [{ col: c.col, kind: "daterange", label: formatDateRangeLabel(c.dateRange) }];
      }
      if (c.numericRange && (c.numericRange.min || c.numericRange.max)) {
        return [{ col: c.col, kind: "numericrange", label: formatNumericRangeLabel(c.numericRange) }];
      }
      return [];
    }
    if (c.mode === "separate") return [{ col: c.col, kind: "separate" }];
    return [];
  });

  const handleExport = () => {
    const header = [
      ...displayCols.map((dc) => dc.col),
      "Count",
      ...numericCols.map((c) => `${c} (sum)`),
    ];
    const dataRows = results.map((row) => [
      ...displayCols.map((dc) =>
        dc.kind === "single" ? (dc.label ?? "") : (row.labels[dc.col] ?? "")
      ),
      row.count,
      ...numericCols.map((c) => row.metrics[c] ?? 0),
    ]);
    const ws = XLSX.utils.aoa_to_sheet([header, ...dataRows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Summary");
    XLSX.writeFile(wb, "summary.xlsx");
  };

  const visibleRows = results.slice(0, visible);
  const remaining = results.length - visible;
  const colSpan = displayCols.length + 1 + numericCols.length;

  return (
    <div className="rounded-xl border border-slate-800 overflow-hidden">
      <div className="px-5 py-4 bg-slate-900 border-b border-slate-800 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <BarChart3 className="w-4 h-4 text-violet-400" />
          <span className="font-semibold text-slate-200 text-sm">Summary Results</span>
          <span className="text-xs text-slate-500 font-mono">
            {results.length} group{results.length !== 1 ? "s" : ""}
          </span>
        </div>
        <button
          onClick={handleExport}
          className="flex items-center gap-1.5 text-xs text-slate-400 hover:text-slate-200 px-3 py-1.5 rounded-lg border border-slate-700 hover:border-slate-600 transition-all"
        >
          <Download className="w-3.5 h-3.5" />
          Export Excel
        </button>
      </div>

      {results.length === 0 ? (
        <div className="p-12 text-center text-slate-500">
          <p className="text-sm">No data matches your current filters.</p>
          <p className="text-xs mt-1 text-slate-600">Try adjusting your Filter selections.</p>
        </div>
      ) : (
        <div className="overflow-x-auto">
          <table className="excel-table w-full">
            <thead>
              <tr>
                {displayCols.map((dc) => (
                  <th key={dc.col} className={dc.kind === "separate" ? "text-violet-300" : "text-blue-300"}>
                    {dc.col}
                    {(dc.kind === "daterange" || dc.kind === "numericrange") && dc.label && (
                      <span className="block text-[10px] font-normal text-slate-500 mt-0.5">{dc.label}</span>
                    )}
                  </th>
                ))}
                <th>Count</th>
                {numericCols.map((c) => (
                  <th key={c} className="text-amber-300">{c} (sum)</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {visibleRows.map((row, i) => (
                <tr key={i}>
                  {displayCols.map((dc) => {
                    const value =
                      dc.kind === "single"
                        ? (dc.label ?? "—")
                        : (row.labels[dc.col] ?? "—");
                    return (
                      <td key={dc.col} className={dc.kind === "separate" ? "text-slate-200 font-medium" : "text-blue-200 font-medium"}>
                        {value}
                      </td>
                    );
                  })}
                  <td className="text-slate-400 font-mono text-right">{row.count}</td>
                  {numericCols.map((c) => (
                    <td key={c} className="text-amber-200 font-mono text-right">{fmt(row.metrics[c] ?? 0)}</td>
                  ))}
                </tr>
              ))}
            </tbody>
            {results.length > 1 && numericCols.length > 0 && (
              <tfoot>
                <tr>
                  {displayCols.length > 0 && (
                    <td colSpan={displayCols.length} className="text-right text-slate-500 text-xs uppercase tracking-wider pr-3">
                      Total
                    </td>
                  )}
                  <td className="text-slate-300 font-mono text-right font-semibold">
                    {results.reduce((a, b) => a + b.count, 0)}
                  </td>
                  {numericCols.map((c) => (
                    <td key={c} className="text-amber-300 font-mono text-right font-semibold border-t border-slate-700">
                      {fmt(results.reduce((a, b) => a + (b.metrics[c] ?? 0), 0))}
                    </td>
                  ))}
                </tr>
              </tfoot>
            )}
            {remaining > 0 && (
              <tfoot>
                <tr>
                  <td colSpan={colSpan} className="px-4 py-3 bg-slate-900/50">
                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => setVisible((v) => v + PAGE_SIZE)}
                        className="px-3 py-1.5 rounded-lg border border-slate-700 text-xs text-slate-400 hover:border-slate-600 hover:text-slate-200 transition-all"
                      >
                        Show {Math.min(remaining, PAGE_SIZE)} more
                      </button>
                      <button
                        onClick={() => setVisible(results.length)}
                        className="px-3 py-1.5 rounded-lg border border-slate-700 text-xs text-slate-400 hover:border-slate-600 hover:text-slate-200 transition-all"
                      >
                        Show all {results.length}
                      </button>
                      <span className="text-xs text-slate-600">{remaining} row{remaining !== 1 ? "s" : ""} hidden</span>
                    </div>
                  </td>
                </tr>
              </tfoot>
            )}
          </table>
        </div>
      )}
    </div>
  );
}
