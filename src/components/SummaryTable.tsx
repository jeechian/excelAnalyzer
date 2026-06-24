"use client";
import { BarChart3 } from "lucide-react";
import { useState, useEffect } from "react";
import type { SummaryRow, ColumnConfig } from "@/lib/excel";

interface SummaryTableProps {
  results: SummaryRow[];
  separateCols: string[];
  multiSelectCols: string[];
  numericCols: string[];
  configs: ColumnConfig[];
}

function fmt(n: number): string {
  if (Number.isInteger(n)) return n.toLocaleString();
  return n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

const PAGE_SIZE = 10;

export default function SummaryTable({ results, separateCols, multiSelectCols, numericCols, configs }: SummaryTableProps) {
  const selectCols = configs.filter((c) => c.mode === "select" && c.selectedValues.length > 0);
  const [visible, setVisible] = useState(PAGE_SIZE);

  // Reset to first page whenever results change
  useEffect(() => { setVisible(PAGE_SIZE); }, [results]);

  const visibleRows = results.slice(0, visible);
  const remaining = results.length - visible;
  const colSpan = separateCols.length + multiSelectCols.length + 1 + numericCols.length;

  return (
    <div className="rounded-xl border border-slate-800 overflow-hidden">
      <div className="px-5 py-4 bg-slate-900 border-b border-slate-800 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <BarChart3 className="w-4 h-4 text-violet-400" />
          <span className="font-semibold text-slate-200 text-sm">Summary Results</span>
        </div>
        <div className="flex items-center gap-3">
          {selectCols.length > 0 && (
            <div className="flex flex-wrap gap-2">
              {selectCols.map((c) => (
                <span key={c.col} className="text-xs bg-blue-950 text-blue-300 border border-blue-800 rounded-full px-2.5 py-0.5">
                  {c.col}: {c.selectedValues.join(", ")}
                </span>
              ))}
            </div>
          )}
          <span className="text-xs text-slate-500 font-mono">
            {visible < results.length ? `${visible} / ` : ""}{results.length} group{results.length !== 1 ? "s" : ""}
          </span>
        </div>
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
                {separateCols.map((c) => (
                  <th key={c} className="text-violet-300">{c}</th>
                ))}
                {multiSelectCols.map((c) => (
                  <th key={c} className="text-blue-300">{c}</th>
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
                  {separateCols.map((c) => (
                    <td key={c} className="text-slate-200 font-medium">{row.labels[c] ?? "—"}</td>
                  ))}
                  {multiSelectCols.map((c) => (
                    <td key={c} className="text-blue-200 font-medium">{row.labels[c] ?? "—"}</td>
                  ))}
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
                  {(separateCols.length + multiSelectCols.length) > 0 && (
                    <td colSpan={separateCols.length + multiSelectCols.length} className="text-right text-slate-500 text-xs uppercase tracking-wider pr-3">
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
