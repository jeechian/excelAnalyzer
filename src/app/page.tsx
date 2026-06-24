"use client";
import { useState, useMemo } from "react";
import { FileSpreadsheet, RefreshCw } from "lucide-react";
import FileUpload from "@/components/FileUpload";
import DataPreview from "@/components/DataPreview";
import ColumnConfigurator from "@/components/ColumnConfigurator";
import SortControls from "@/components/SortControls";
import SummaryTable from "@/components/SummaryTable";
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
  };

  // Numeric columns (for summing)
  const numericCols = useMemo(() => {
    if (!data) return [];
    return data.headers.filter((h) => {
      const stats = getColumnStats(data.rows, h);
      return stats.isNumeric;
    });
  }, [data]);

  // Columns in "separate" (group-by) mode
  const separateCols = configs.filter((c) => c.mode === "separate").map((c) => c.col);

  // Filter columns with 2+ values selected — shown as table columns too
  const multiSelectCols = configs
    .filter((c) => c.mode === "select" && c.selectedValues.length > 1)
    .map((c) => c.col);

  // Sortable: separate cols + multi-select cols + Count + numeric cols
  const sortableColumns = useMemo(() => {
    return [...separateCols, ...multiSelectCols, "Count", ...numericCols];
  }, [separateCols, multiSelectCols, numericCols]);

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
                separateCols={separateCols}
                multiSelectCols={multiSelectCols}
                numericCols={numericCols}
                configs={configs}
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
      </main>
    </div>
  );
}
