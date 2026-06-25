export interface DimensionConfig {
  col: string;
  mode: "groupby" | "filter";
  filterValues: string[];
  numericRange?: { min?: number; max?: number };
  numericGroupMode?: "split" | "combine";
}

export interface MetricConfig {
  col: string; // "__count__" = pure row count
  filterValues: string[]; // for categorical: filter to these values; empty = count all rows
}

export interface ContributionRow {
  labels: Record<string, string>;
  value: number;
  percentage: number;
}

export interface ContributionResult {
  rows: ContributionRow[];
  total: number; // global denominator
  isNumericMetric: boolean;
}

function sumRows(rows: Record<string, unknown>[], col: string): number {
  return rows.reduce((acc, r) => {
    const v = Number(r[col]);
    return acc + (isNaN(v) ? 0 : v);
  }, 0);
}

export function computeContribution(
  allRows: Record<string, unknown>[],
  dimensions: DimensionConfig[],
  metric: MetricConfig,
  numericCols: Set<string>
): ContributionResult {
  const isNumeric = metric.col !== "__count__" && numericCols.has(metric.col);

  // Global denominator: total metric across ALL rows (with label filter, ignoring dim filters)
  let globalRows = allRows;
  if (!isNumeric && metric.col !== "__count__" && metric.filterValues.length > 0) {
    globalRows = allRows.filter((r) =>
      metric.filterValues.includes(String(r[metric.col] ?? ""))
    );
  }
  const total = isNumeric ? sumRows(globalRows, metric.col) : globalRows.length;

  // Filter dims: categorical value filter OR numeric range filter
  const filterDims = dimensions.filter(
    (d) =>
      d.mode === "filter" &&
      (d.filterValues.length > 0 ||
        (d.numericRange &&
          (d.numericRange.min !== undefined || d.numericRange.max !== undefined)))
  );
  const groupbyDims = dimensions.filter((d) => d.mode === "groupby");

  function applyNumericRange(
    rows: Record<string, unknown>[],
    col: string,
    range: { min?: number; max?: number }
  ) {
    return rows.filter((r) => {
      const raw = r[col];
      if (raw === "" || raw === null || raw === undefined) return true;
      const v = Number(raw);
      if (isNaN(v)) return true;
      if (range.min !== undefined && v < range.min) return false;
      if (range.max !== undefined && v > range.max) return false;
      return true;
    });
  }

  function dimLabel(dim: DimensionConfig): string {
    if (dim.numericRange) {
      const { min, max } = dim.numericRange;
      const parts: string[] = [];
      if (min !== undefined) parts.push(`≥ ${min}`);
      if (max !== undefined) parts.push(`≤ ${max}`);
      return parts.join(" & ");
    }
    return dim.filterValues.join(", ");
  }

  // Apply dimension filters, then label filter
  let working = allRows;
  for (const dim of filterDims) {
    working = dim.numericRange
      ? applyNumericRange(working, dim.col, dim.numericRange)
      : working.filter((r) => dim.filterValues.includes(String(r[dim.col] ?? "")));
  }
  // Apply numeric range filter on groupby dims (to limit scope before grouping)
  for (const dim of groupbyDims) {
    if (dim.numericRange) {
      working = applyNumericRange(working, dim.col, dim.numericRange);
    }
  }
  if (!isNumeric && metric.col !== "__count__" && metric.filterValues.length > 0) {
    working = working.filter((r) => metric.filterValues.includes(String(r[metric.col] ?? "")));
  }

  function metricValue(rows: Record<string, unknown>[]): number {
    return isNumeric ? sumRows(rows, metric.col) : rows.length;
  }

  // Pre-compute combine labels for groupby numeric dims
  const numericGroupLabels: Record<string, string> = {};
  for (const dim of groupbyDims) {
    if (dim.numericRange && (dim.numericGroupMode ?? "split") === "combine") {
      const { min, max } = dim.numericRange;
      const parts: string[] = [];
      if (min !== undefined) parts.push(`≥ ${min}`);
      if (max !== undefined) parts.push(`≤ ${max}`);
      numericGroupLabels[dim.col] = parts.join(" & ");
    }
  }

  // No group-by: single row showing the filtered combo
  if (groupbyDims.length === 0) {
    const value = metricValue(working);
    const labels: Record<string, string> = {};
    for (const dim of filterDims) labels[dim.col] = dimLabel(dim);
    return {
      rows: [{ labels, value, percentage: total > 0 ? (value / total) * 100 : 0 }],
      total,
      isNumericMetric: isNumeric,
    };
  }

  // Group by
  const groups = new Map<string, Record<string, unknown>[]>();
  for (const row of working) {
    const key = groupbyDims
      .map((d) => {
        const label = numericGroupLabels[d.col];
        return label !== undefined ? `${d.col}|||${label}` : `${d.col}|||${String(row[d.col] ?? "")}`;
      })
      .join("^^^");
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key)!.push(row);
  }

  const rows: ContributionRow[] = [];
  for (const [key, groupRows] of groups) {
    const labels: Record<string, string> = {};
    for (const dim of filterDims) labels[dim.col] = dimLabel(dim);
    for (const part of key.split("^^^")) {
      const sep = part.indexOf("|||");
      if (sep !== -1) labels[part.slice(0, sep)] = part.slice(sep + 3);
    }
    const value = metricValue(groupRows);
    rows.push({ labels, value, percentage: total > 0 ? (value / total) * 100 : 0 });
  }

  rows.sort((a, b) => b.percentage - a.percentage);
  return { rows, total, isNumericMetric: isNumeric };
}
