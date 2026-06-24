"use client";
import { useState } from "react";
import { ChevronDown, ChevronUp, TableIcon } from "lucide-react";
import type { ExcelData } from "@/lib/excel";

interface DataPreviewProps {
  data: ExcelData;
}

export default function DataPreview({ data }: DataPreviewProps) {
  const [open, setOpen] = useState(true);
  const previewRows = data.rows.slice(0, 10);

  return (
    <div className="rounded-xl border border-slate-800 overflow-hidden">
      <button
        className="w-full flex items-center justify-between px-5 py-4 bg-slate-900 hover:bg-slate-800 transition-colors"
        onClick={() => setOpen((v) => !v)}
      >
        <div className="flex items-center gap-3">
          <TableIcon className="w-4 h-4 text-violet-400" />
          <span className="font-semibold text-slate-200 text-sm">
            Data Preview
          </span>
          <span className="text-xs text-slate-500 font-mono">
            {data.rows.length} rows · {data.headers.length} columns · sheet: {data.sheetName}
          </span>
        </div>
        {open
          ? <ChevronUp className="w-4 h-4 text-slate-500" />
          : <ChevronDown className="w-4 h-4 text-slate-500" />}
      </button>

      {open && (
        <div className="overflow-x-auto">
          <table className="excel-table w-full">
            <thead>
              <tr>
                <th className="text-slate-600 w-10">#</th>
                {data.headers.map((h) => (
                  <th key={h}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {previewRows.map((row, i) => (
                <tr key={i}>
                  <td className="text-slate-600 text-right">{i + 1}</td>
                  {data.headers.map((h) => (
                    <td key={h}>{String(row[h] ?? "")}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
          {data.rows.length > 10 && (
            <p className="text-center text-xs text-slate-600 py-2 font-mono">
              Showing 10 of {data.rows.length} rows
            </p>
          )}
        </div>
      )}
    </div>
  );
}
