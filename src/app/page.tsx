"use client";
import { useState, useMemo } from "react";
import { FileSpreadsheet, RefreshCw, HelpCircle, X } from "lucide-react";
import FileUpload from "@/components/FileUpload";
import DataPreview from "@/components/DataPreview";
import ColumnConfigurator from "@/components/ColumnConfigurator";
import SortControls from "@/components/SortControls";
import SummaryTable from "@/components/SummaryTable";
import SummaryChart from "@/components/SummaryChart";
import AssociationMiner from "@/components/AssociationMiner";
import {
  parseExcelFile,
  getColumnStats,
  computeSummary,
  type ExcelData,
  type ColumnConfig,
  type SortEntry,
} from "@/lib/excel";

export default function Home() {
  const [data, setData] = useState<ExcelData | null>(null);
  const [configs, setConfigs] = useState<ColumnConfig[]>([]);
  const [sortOrder, setSortOrder] = useState<SortEntry[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [mode, setMode] = useState<"summary" | "association">("summary");
  const [showSummaryHelp, setShowSummaryHelp] = useState(false);

  const handleFile = async (file: File) => {
    setLoading(true);
    setError(null);
    try {
      const parsed = await parseExcelFile(file);
      setData(parsed);
      setConfigs(
        parsed.headers.map((col) => ({
          col,
          mode: "ignore",
          selectedValues: [],
        }))
      );
      setSortOrder([]);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setLoading(false);
    }
  };

  const reset = () => {
    setData(null);
    setConfigs([]);
    setSortOrder([]);
    setError(null);
    setMode("summary");
  };

  // Numeric columns for summing — exclude cols used as range filters (they become group dimensions)
  const numericCols = useMemo(() => {
    if (!data) return [];
    return data.headers.filter((h) => {
      const stats = getColumnStats(data.rows, h);
      if (!stats.isNumeric) return false;
      const cfg = configs.find((c) => c.col === h);
      return !(cfg?.mode === "select" && (cfg.numericRange?.min || cfg.numericRange?.max));
    });
  }, [data, configs]);

  // Sortable columns: follow config order, then Count, then numeric sums
  const sortableColumns = useMemo(() => {
    const activeCols = configs
      .filter((c) => c.mode !== "ignore" && !numericCols.includes(c.col))
      .filter((c) =>
        c.mode === "separate" ||
        (c.mode === "select" && (
          c.selectedValues.length > 0 ||
          (c.dateRange && (c.dateRange.start || c.dateRange.end)) ||
          (c.numericRange && (c.numericRange.min || c.numericRange.max))
        ))
      )
      .map((c) => c.col);
    return [...activeCols, "Count", ...numericCols];
  }, [configs, numericCols]);

  // Summary
  const summaryResults = useMemo(() => {
    if (!data) return [];
    const activeNumericCols = numericCols;
    return computeSummary(data.rows, configs, activeNumericCols, sortOrder);
  }, [data, configs, numericCols, sortOrder]);

  const showSummary = configs.some((c) => c.mode !== "ignore");

  return (
    <div className="min-h-screen bg-slate-950">
      {/* Header */}
      <header className="border-b border-slate-800/60 bg-slate-950/80 backdrop-blur sticky top-0 z-40">
        <div className="max-w-6xl mx-auto px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="p-1.5 rounded-lg bg-violet-600/20">
              <FileSpreadsheet className="w-5 h-5 text-violet-400" />
            </div>
            <div>
              <h1 className="font-bold text-slate-100 text-base leading-none">Excel Analyzer</h1>
              <p className="text-xs text-slate-500 mt-0.5">Upload · Configure · Summarize</p>
            </div>
          </div>
          {data && (
            <button
              onClick={reset}
              className="flex items-center gap-2 text-sm text-slate-400 hover:text-slate-200 px-3 py-1.5 rounded-lg border border-slate-800 hover:border-slate-700 transition-all"
            >
              <RefreshCw className="w-3.5 h-3.5" /> New file
            </button>
          )}
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-6 py-8 space-y-6">
        {/* Upload */}
        {!data && (
          <div className="max-w-xl mx-auto mt-12">
            <h2 className="text-2xl font-bold text-slate-100 text-center mb-2">
              Analyze your spreadsheet
            </h2>
            <p className="text-slate-500 text-center text-sm mb-8">
              Upload an Excel file to filter, group, and summarize your data instantly.
            </p>
            {error && (
              <div className="mb-4 p-3 rounded-lg bg-red-950 border border-red-800 text-red-300 text-sm">
                {error}
              </div>
            )}
            {loading ? (
              <div className="flex items-center justify-center h-52 rounded-xl border-2 border-dashed border-slate-700">
                <div className="flex items-center gap-3 text-slate-400">
                  <div className="w-5 h-5 border-2 border-violet-400 border-t-transparent rounded-full animate-spin" />
                  Parsing file…
                </div>
              </div>
            ) : (
              <FileUpload onFile={handleFile} />
            )}
          </div>
        )}

        {/* Data loaded */}
        {data && (
          <>
            {/* Preview */}
            <DataPreview data={data} />

            {/* Mode tabs */}
            <div className="flex items-center gap-3">
              <div className="flex gap-1 p-1 rounded-xl border border-slate-800 bg-slate-900/50 w-fit">
                {(["summary", "association"] as const).map((m) => (
                  <button
                    key={m}
                    onClick={() => setMode(m)}
                    className={`px-4 py-2 rounded-lg text-xs font-semibold transition-all ${
                      mode === m
                        ? "bg-slate-700 text-slate-100 shadow-sm"
                        : "text-slate-500 hover:text-slate-300"
                    }`}
                  >
                    {m === "summary" ? "Summary" : "Association Analysis"}
                  </button>
                ))}
              </div>
              {mode === "summary" && (
                <button
                  onClick={() => setShowSummaryHelp(true)}
                  className="flex items-center gap-1.5 text-xs text-slate-500 hover:text-slate-300 px-3 py-2 rounded-lg border border-slate-800 hover:border-slate-700 transition-all"
                >
                  <HelpCircle className="w-3 h-3" />
                  How to use
                </button>
              )}
            </div>

            {/* Summary tab */}
            {mode === "summary" && (
              <>
                {/* Column config */}
                <ColumnConfigurator
                  data={data}
                  configs={configs}
                  onChange={setConfigs}
                />

                {/* Sort controls - only when something is active */}
                {showSummary && (
                  <SortControls
                    sortableColumns={sortableColumns}
                    sortOrder={sortOrder}
                    onSortOrder={setSortOrder}
                  />
                )}

                {/* Summary */}
                {showSummary && (
                  <SummaryTable
                    results={summaryResults}
                    configs={configs}
                    numericCols={numericCols}
                  />
                )}

                {/* Chart */}
                {showSummary && summaryResults.length > 0 && (
                  <SummaryChart
                    results={summaryResults}
                    configs={configs}
                    numericCols={numericCols}
                  />
                )}

                {/* Hint when nothing configured */}
                {!showSummary && (
                  <div className="rounded-xl border border-dashed border-slate-800 p-10 text-center">
                    <p className="text-slate-500 text-sm">
                      Set at least one column to <span className="text-blue-400">Filter</span> or <span className="text-violet-400">Group by</span> to see a summary.
                    </p>
                  </div>
                )}
              </>
            )}

            {/* Association tab */}
            {mode === "association" && <AssociationMiner data={data} />}
          </>
        )}
      </main>

      {/* Summary how to use modal */}
      {showSummaryHelp && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4"
          onClick={() => setShowSummaryHelp(false)}
        >
          <div
            className="bg-slate-900 border border-slate-700 rounded-2xl max-w-lg w-full p-6 space-y-5 shadow-2xl"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between">
              <h2 className="text-sm font-semibold text-slate-100">How to use Summary</h2>
              <button onClick={() => setShowSummaryHelp(false)} className="text-slate-500 hover:text-slate-300">
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="space-y-4 text-xs text-slate-400 leading-relaxed">
              <div>
                <p className="text-slate-200 font-medium mb-1">1. Upload your Excel file</p>
                <p>Click the upload area or drag and drop an Excel (.xlsx, .xls) or CSV file. Your data loads instantly — nothing is sent to a server.</p>
              </div>

              <div>
                <p className="text-slate-200 font-medium mb-1">2. Configure columns</p>
                <ul className="space-y-1 list-disc list-inside">
                  <li><span className="text-violet-400 font-medium">Group by</span> — split the data into groups by this column's values (e.g. one row per Region).</li>
                  <li><span className="text-blue-400 font-medium">Filter</span> — narrow the data to specific values, a date range, or a numeric range.</li>
                  <li><span className="text-slate-500 font-medium">Ignore</span> — exclude the column from analysis.</li>
                </ul>
              </div>

              <div>
                <p className="text-slate-200 font-medium mb-1">3. Read the results</p>
                <p>Each row in the summary table is one combination of your grouped columns. You can see the row count and sums of any numeric columns. The chart visualizes the top groups.</p>
              </div>

              <div>
                <p className="text-slate-200 font-medium mb-1">4. Sort</p>
                <p>Use the sort controls to order results by any column — ascending or descending.</p>
              </div>

              <div className="pt-2 border-t border-slate-800 text-slate-500">
                Switch to <span className="text-slate-300">Association Analysis</span> to see how much each group contributes to a metric.
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
