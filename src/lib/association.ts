export interface DimensionConfig {
  col: string;
  mode: "groupby" | "filter";
  filterValues: string[];
  numericRange?: { min?: number; max?: number };
  numericGroupMode?: "split" | "combine";
}

export interface MetricConfig {
  col: string; // "__count__" = pure row count
  filterValues: string[]; // categorical: filter to these values; empty = count all rows
}

export interface ContributionRow {
  labels: Record<string, string>;
  value: number;
  percentage: number;
  breakdown?: Record<string, number>; // per pivot value; set when ContributionResult.pivotValues is set
}

export interface ContributionResult {
  rows: ContributionRow[];
  total: number;
  isNumericMetric: boolean;
  pivotValues?: string[]; // categorical metric with ≥2 selected values → one column per value
}

function sumRows(rows: Record<string, unknown>[], col: string): number {
  return rows.reduce((acc, r) => {
    const v = Number(r[col]);
    return acc + (isNaN(v) ? 0 : v);
  }, 0);
}

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

export function computeContribution(
  allRows: Record<string, unknown>[],
  dimensions: DimensionConfig[],
  metric: MetricConfig,
  numericCols: Set<string>
): ContributionResult {
  const isNumeric = metric.col !== "__count__" && numericCols.has(metric.col);
  const isCategorical = metric.col !== "__count__" && !isNumeric;

  // Pivot: categorical metric with ≥2 selected values → one column per value in output
  const pivotValues =
    isCategorical && metric.filterValues.length >= 2 ? metric.filterValues : undefined;

  const allFilterDims = dimensions.filter(
    (d) =>
      d.mode === "filter" &&
      (d.filterValues.length > 0 ||
        (d.numericRange &&
          (d.numericRange.min !== undefined || d.numericRange.max !== undefined)))
  );
  const groupbyDims = dimensions.filter((d) => d.mode === "groupby");

  // Expand dims: categorical filter with ≥2 values → restrict AND create one row per value
  const expandDims = allFilterDims.filter((d) => !d.numericRange && d.filterValues.length >= 2);
  // Restrict dims: numeric range or ≤1 categorical value → just filter data
  const restrictDims = allFilterDims.filter((d) => d.numericRange || d.filterValues.length <= 1);

  // Global denominator (before dimension filters, consistent with original behavior)
  let globalRows = allRows;
  if (pivotValues) {
    globalRows = allRows.filter((r) => pivotValues.includes(String(r[metric.col] ?? "")));
  } else if (isCategorical && metric.filterValues.length === 1) {
    globalRows = allRows.filter((r) =>
      metric.filterValues.includes(String(r[metric.col] ?? ""))
    );
  }
  const total = isNumeric ? sumRows(globalRows, metric.col) : globalRows.length;

  // Build working set
  let working = allRows;
  for (const dim of restrictDims) {
    working = dim.numericRange
      ? applyNumericRange(working, dim.col, dim.numericRange)
      : working.filter((r) => dim.filterValues.includes(String(r[dim.col] ?? "")));
  }
  for (const dim of expandDims) {
    working = working.filter((r) => dim.filterValues.includes(String(r[dim.col] ?? "")));
  }
  for (const dim of groupbyDims) {
    if (dim.numericRange) working = applyNumericRange(working, dim.col, dim.numericRange);
  }
  if (pivotValues) {
    working = working.filter((r) => pivotValues.includes(String(r[metric.col] ?? "")));
  } else if (isCategorical && metric.filterValues.length === 1) {
    working = working.filter((r) =>
      metric.filterValues.includes(String(r[metric.col] ?? ""))
    );
  }

  // Effective groupby: original groupby dims + expand dims (preserving dimension order)
  const effectiveGroupDims = dimensions.filter(
    (d) =>
      d.mode === "groupby" ||
      (d.mode === "filter" && !d.numericRange && d.filterValues.length >= 2)
  );

  // Pre-compute combine labels for numeric groupby dims
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

  function metricValue(rows: Record<string, unknown>[]): number {
    return isNumeric ? sumRows(rows, metric.col) : rows.length;
  }

  function computeBreakdown(rows: Record<string, unknown>[]): Record<string, number> {
    const bd: Record<string, number> = {};
    for (const v of pivotValues!) {
      bd[v] = rows.filter((r) => String(r[metric.col] ?? "") === v).length;
    }
    return bd;
  }

  // No effective groupby: single row
  if (effectiveGroupDims.length === 0) {
    const value = metricValue(working);
    const labels: Record<string, string> = {};
    for (const dim of restrictDims) labels[dim.col] = dimLabel(dim);
    const row: ContributionRow = {
      labels,
      value,
      percentage: total > 0 ? (value / total) * 100 : 0,
    };
    if (pivotValues) row.breakdown = computeBreakdown(working);
    return { rows: [row], total, isNumericMetric: isNumeric, pivotValues };
  }

  // Group by effective group dims
  const groups = new Map<string, Record<string, unknown>[]>();
  for (const row of working) {
    const key = effectiveGroupDims
      .map((d) => {
        const label = numericGroupLabels[d.col];
        return label !== undefined
          ? `${d.col}|||${label}`
          : `${d.col}|||${String(row[d.col] ?? "")}`;
      })
      .join("^^^");
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key)!.push(row);
  }

  const rows: ContributionRow[] = [];
  for (const [key, groupRows] of groups) {
    const labels: Record<string, string> = {};
    for (const dim of restrictDims) labels[dim.col] = dimLabel(dim);
    for (const part of key.split("^^^")) {
      const sep = part.indexOf("|||");
      if (sep !== -1) labels[part.slice(0, sep)] = part.slice(sep + 3);
    }
    const value = pivotValues ? groupRows.length : metricValue(groupRows);
    const row: ContributionRow = {
      labels,
      value,
      percentage: total > 0 ? (value / total) * 100 : 0,
    };
    if (pivotValues) row.breakdown = computeBreakdown(groupRows);
    rows.push(row);
  }

  rows.sort((a, b) => b.value - a.value);
  return { rows, total, isNumericMetric: isNumeric, pivotValues };
}
