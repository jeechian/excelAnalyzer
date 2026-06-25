"use client";
import { useState } from "react";
import { BarChart2 } from "lucide-react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip,
  ResponsiveContainer, Cell, LabelList,
} from "recharts";
import type { SummaryRow, ColumnConfig } from "@/lib/excel";

interface SummaryChartProps {
  results: SummaryRow[];
  configs: ColumnConfig[];
  numericCols: string[];
}

const MAX_BARS = 50;

function buildRowLabel(row: SummaryRow, configs: ColumnConfig[], numericCols: string[]): string {
  const parts: string[] = [];
  for (const c of configs) {
    if (c.mode === "ignore" || numericCols.includes(c.col)) continue;
    if (c.mode === "select") {
      if (c.selectedValues.length === 1) continue; // constant — not useful in label
      if (c.selectedValues.length > 1 || (c.dateRange && (c.dateRange.start || c.dateRange.end))) {
        parts.push(row.labels[c.col] ?? "—");
      }
    } else if (c.mode === "separate") {
      parts.push(row.labels[c.col] ?? "—");
    }
  }
  return parts.join(" · ") || "Total";
}

interface TooltipProps {
  active?: boolean;
  payload?: Array<{ value: number }>;
  label?: string;
}

function CustomTooltip({ active, payload, label }: TooltipProps) {
  if (!active || !payload?.length) return null;
  return (
    <div className="bg-slate-900 border border-slate-700 rounded-lg px-3 py-2.5 shadow-xl pointer-events-none">
      <p className="text-slate-400 text-xs mb-1.5 max-w-72 break-words">{label}</p>
      <p className="text-violet-300 text-lg font-bold leading-none">{payload[0].value.toLocaleString()}</p>
      <p className="text-slate-500 text-xs mt-0.5">count</p>
    </div>
  );
}

export default function SummaryChart({ results, configs, numericCols }: SummaryChartProps) {
  const [activeIndex, setActiveIndex] = useState<number | null>(null);

  const chartData = results.slice(0, MAX_BARS).map((row) => ({
    name: buildRowLabel(row, configs, numericCols),
    count: row.count,
  }));

  if (chartData.length === 0) return null;

  const handleClick = (_: unknown, index: number) => {
    setActiveIndex((prev) => (prev === index ? null : index));
  };

  return (
    <div className="rounded-xl border border-slate-800 overflow-hidden">
      <div className="px-5 py-4 bg-slate-900 border-b border-slate-800 flex items-center gap-3">
        <BarChart2 className="w-4 h-4 text-violet-400" />
        <span className="font-semibold text-slate-200 text-sm">Count Distribution</span>
        {results.length > MAX_BARS && (
          <span className="text-xs text-slate-500">
            showing first {MAX_BARS} of {results.length} groups
          </span>
        )}
      </div>

      <div className="bg-slate-950 px-4 pt-4 pb-2">
        <ResponsiveContainer width="100%" height={300}>
          <BarChart
            data={chartData}
            margin={{ top: 24, right: 16, left: 0, bottom: chartData.length > 8 ? 64 : 24 }}
          >
            <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" vertical={false} />
            <XAxis
              dataKey="name"
              tick={{ fill: "#64748b", fontSize: 11 }}
              axisLine={{ stroke: "#1e293b" }}
              tickLine={false}
              angle={chartData.length > 8 ? -35 : 0}
              textAnchor={chartData.length > 8 ? "end" : "middle"}
              interval={0}
            />
            <YAxis
              tick={{ fill: "#64748b", fontSize: 11 }}
              axisLine={false}
              tickLine={false}
              width={40}
              allowDecimals={false}
            />
            <Tooltip
              content={<CustomTooltip />}
              cursor={{ fill: "#1e293b" }}
            />
            <Bar
              dataKey="count"
              radius={[4, 4, 0, 0]}
              onClick={handleClick}
              style={{ cursor: "pointer" }}
            >
              <LabelList
                dataKey="count"
                position="top"
                style={{ fill: "#475569", fontSize: 10 }}
                formatter={(v) => Number(v).toLocaleString()}
              />
              {chartData.map((_, i) => (
                <Cell
                  key={i}
                  fill={i === activeIndex ? "#a78bfa" : "#7c3aed"}
                />
              ))}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}
