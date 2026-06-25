export interface DimensionConfig {
  col: string;
  mode: "groupby" | "filter";
  filterValues: string[];
  numericRange?: { min?: number; max?: number };
  numericGroupMode?: "split" | "combine";
  dateRange?: { start?: string; end?: string };
  dateGroupMode?: "split" | "combine";
}

export interface MetricConfig {
  id: string;
  col: string; // "__count__" = pure row count
  filterValues: string[]; // categorical: filter to these values
  numericRange?: { min?: number; max?: number }; // filter rows by numeric metric value
  dateFilterCol?: string; // date column to apply date filter on
  dateRange?: { start?: string; end?: string }; // date range filter (always combine)
}

export interface MetricMeta {
  id: string;
  label: string;
  total: number;
  isNumeric: boolean;
}

export interface ContributionRow {
  labels: Record<string, string>;
  values: Record<string, number>;      // metric.id → aggregate value
  percentages: Record<string, number>; // metric.id → %
  breakdown?: Record<string, number>;  // pivot: filterValue → count
}

export interface ContributionResult {
  rows: ContributionRow[];
  metricMetas: MetricMeta[];
  pivotValues?: string[]; // only when 1 categorical metric with ≥2 filterValues
}

// ─── Helpers ────────────────────────────────────────────────────────────────

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
): Record<string, unknown>[] {
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

function parseStoredDate(v: string): Date | null {
  // "01 Jan 2026" or "01 Jan 2026 12:19"
  const m1 = v.match(/^(\d{2}) ([A-Z][a-z]{2}) (\d{4})/);
  if (m1) {
    const months: Record<string, number> = {
      Jan: 0, Feb: 1, Mar: 2, Apr: 3, May: 4, Jun: 5,
      Jul: 6, Aug: 7, Sep: 8, Oct: 9, Nov: 10, Dec: 11,
    };
    return new Date(Number(m1[3]), months[m1[2]] ?? 0, Number(m1[1]));
  }
  // "YYYY-MM-DD" or "YYYY-MM-DD HH:MM"
  const m2 = v.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m2) return new Date(Number(m2[1]), Number(m2[2]) - 1, Number(m2[3]));
  return null;
}

function parseInputDate(s: string): Date | null {
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function isInDateRange(
  rowVal: string,
  range: { start?: string; end?: string }
): boolean {
  const d = parseStoredDate(rowVal);
  if (!d) return true;
  if (range.start) {
    const s = parseInputDate(range.start);
    if (s && d < s) return false;
  }
  if (range.end) {
    const e = parseInputDate(range.end);
    if (e) {
      e.setHours(23, 59, 59, 999);
      if (d > e) return false;
    }
  }
  return true;
}

function formatDateRangeLabel(range: { start?: string; end?: string }): string {
  const fmt = (s: string): string => {
    const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return s;
    return `${m[3]} ${months[Number(m[2]) - 1]} ${m[1]}`;
  };
  const parts: string[] = [];
  if (range.start) parts.push(fmt(range.start));
  if (range.end) parts.push(fmt(range.end));
  return parts.join(" – ");
}

function numericRangeLabel(range: { min?: number; max?: number }): string {
  const parts: string[] = [];
  if (range.min !== undefined) parts.push(`≥ ${range.min}`);
  if (range.max !== undefined) parts.push(`≤ ${range.max}`);
  return parts.join(" & ");
}

// ─── Metric helpers ──────────────────────────────────────────────────────────

function applyMetricFilters(
  rows: Record<string, unknown>[],
  m: MetricConfig
): Record<string, unknown>[] {
  let filtered = rows;
  if (m.dateFilterCol && m.dateRange && (m.dateRange.start || m.dateRange.end)) {
    filtered = filtered.filter((r) =>
      isInDateRange(String(r[m.dateFilterCol!] ?? ""), m.dateRange!)
    );
  }
  return filtered;
}

function applyPostAggregationFilter(
  rows: ContributionRow[],
  metrics: MetricConfig[],
  metricMetas: MetricMeta[]
): ContributionRow[] {
  let filtered = rows;
  for (let i = 0; i < metrics.length; i++) {
    const m = metrics[i];
    const meta = metricMetas[i];
    if (!meta.isNumeric || !m.numericRange) continue;
    const { min, max } = m.numericRange;
    if (min === undefined && max === undefined) continue;
    filtered = filtered.filter((row) => {
      const v = row.values[m.id] ?? 0;
      if (min !== undefined && v < min) return false;
      if (max !== undefined && v > max) return false;
      return true;
    });
    meta.total = filtered.reduce((acc, row) => acc + (row.values[m.id] ?? 0), 0);
    for (const row of filtered) {
      row.percentages[m.id] = meta.total > 0 ? ((row.values[m.id] ?? 0) / meta.total) * 100 : 0;
    }
  }
  return filtered;
}

function computeMetricVal(
  m: MetricConfig,
  isNumericM: boolean,
  rows: Record<string, unknown>[]
): number {
  const filtered = applyMetricFilters(rows, m);
  if (isNumericM) return sumRows(filtered, m.col);
  if (m.col === "__count__") return filtered.length;
  if (m.filterValues.length === 0) return filtered.length;
  return filtered.filter((r) => m.filterValues.includes(String(r[m.col] ?? ""))).length;
}

function computeBreakdown(
  rows: Record<string, unknown>[],
  metricCol: string,
  pivotValues: string[]
): Record<string, number> {
  const bd: Record<string, number> = {};
  for (const v of pivotValues) {
    bd[v] = rows.filter((r) => String(r[metricCol] ?? "") === v).length;
  }
  return bd;
}

// ─── Main function ───────────────────────────────────────────────────────────

// Returns true if dim contributes a column to the results table
function isDimGroupKey(dim: DimensionConfig): boolean {
  if (dim.mode === "groupby") return true;
  // categorical expand: filter with ≥2 values and no range constraints
  if (dim.mode === "filter" && !dim.numericRange && !dim.dateRange && dim.filterValues.length >= 2) return true;
  return false;
}

function getGroupKey(row: Record<string, unknown>, dim: DimensionConfig): string {
  // Numeric combine
  if (dim.numericRange && (dim.numericRange.min !== undefined || dim.numericRange.max !== undefined)) {
    if ((dim.numericGroupMode ?? "split") === "combine") {
      return numericRangeLabel(dim.numericRange);
    }
    return String(row[dim.col] ?? "");
  }
  // Date combine
  if (dim.dateRange && (dim.dateRange.start || dim.dateRange.end)) {
    if ((dim.dateGroupMode ?? "split") === "combine") {
      return formatDateRangeLabel(dim.dateRange);
    }
    return String(row[dim.col] ?? "");
  }
  return String(row[dim.col] ?? "");
}

function getDimFilterLabel(dim: DimensionConfig): string {
  if (dim.numericRange) return numericRangeLabel(dim.numericRange);
  if (dim.dateRange) return formatDateRangeLabel(dim.dateRange);
  return dim.filterValues.join(", ");
}

export function computeContribution(
  allRows: Record<string, unknown>[],
  dimensions: DimensionConfig[],
  metrics: MetricConfig[],
  numericCols: Set<string>
): ContributionResult {
  // Pivot: only when exactly 1 metric, categorical, ≥2 filterValues
  const onlyMetric = metrics.length === 1 ? metrics[0] : null;
  const onlyIsNum = onlyMetric
    ? onlyMetric.col !== "__count__" && numericCols.has(onlyMetric.col)
    : false;
  const onlyIsCat = onlyMetric ? onlyMetric.col !== "__count__" && !onlyIsNum : false;
  const pivotValues =
    onlyIsCat && onlyMetric!.filterValues.length >= 2 ? onlyMetric!.filterValues : undefined;

  // Build working set — dimension filters narrow what gets grouped/displayed.
  // The denominator (100%) is always derived from allRows + metric conditions only,
  // so dimension filters only affect the numerator, never the percentage base.
  let working = allRows;
  for (const dim of dimensions) {
    if (dim.numericRange && (dim.numericRange.min !== undefined || dim.numericRange.max !== undefined)) {
      working = applyNumericRange(working, dim.col, dim.numericRange);
    }
    if (dim.dateRange && (dim.dateRange.start || dim.dateRange.end)) {
      working = working.filter((r) =>
        isInDateRange(String(r[dim.col] ?? ""), dim.dateRange!)
      );
    }
    if (!dim.numericRange && !dim.dateRange && dim.filterValues.length > 0) {
      working = working.filter((r) => dim.filterValues.includes(String(r[dim.col] ?? "")));
    }
  }

  // For pivot: filter working to pivot values
  if (pivotValues) {
    working = working.filter((r) =>
      pivotValues.includes(String(r[onlyMetric!.col] ?? ""))
    );
  }

  // Multi-metric mode: AND all metric conditions → population (100% base).
  // Dimension group-by splits the population. Dimension filter restricts the numerator only — NOT the denominator.
  if (metrics.length > 1) {
    // 1. Build population from allRows + all metric conditions (dimension filters do NOT affect 100%)
    let population = allRows;
    const condParts: string[] = [];

    for (const m of metrics) {
      if (m.col === "__count__") continue;
      if (m.dateFilterCol && m.dateRange && (m.dateRange.start || m.dateRange.end)) {
        population = population.filter((r) =>
          isInDateRange(String(r[m.dateFilterCol!] ?? ""), m.dateRange!)
        );
        condParts.push(formatDateRangeLabel(m.dateRange));
      }
      if (numericCols.has(m.col) && m.numericRange &&
          (m.numericRange.min !== undefined || m.numericRange.max !== undefined)) {
        population = applyNumericRange(population, m.col, m.numericRange);
        condParts.push(`${m.col} ${numericRangeLabel(m.numericRange)}`);
      }
      if (!numericCols.has(m.col) && !m.dateFilterCol && m.filterValues.length > 0) {
        population = population.filter((r) =>
          m.filterValues.includes(String(r[m.col] ?? ""))
        );
        condParts.push(`${m.col} ∈ {${m.filterValues.join(", ")}}`);
      }
    }

    const synthId = metrics.map((m) => m.id).join("+");
    const total = population.length; // denominator = full metric population

    const groupKeyDimsM = dimensions.filter(isDimGroupKey);
    const filterOnlyDimsM = dimensions.filter((d) => !isDimGroupKey(d));

    // 2. Apply dimension filters to population — restricts the numerator only
    let visible = population;
    for (const dim of filterOnlyDimsM) {
      if (dim.numericRange && (dim.numericRange.min !== undefined || dim.numericRange.max !== undefined)) {
        visible = applyNumericRange(visible, dim.col, dim.numericRange);
      } else if (dim.dateRange && (dim.dateRange.start || dim.dateRange.end)) {
        visible = visible.filter((r) => isInDateRange(String(r[dim.col] ?? ""), dim.dateRange!));
      } else if (dim.filterValues.length > 0) {
        visible = visible.filter((r) => dim.filterValues.includes(String(r[dim.col] ?? "")));
      }
    }

    const synthMeta: MetricMeta = {
      id: synthId,
      label: `Count${condParts.length ? ` · ${condParts.join(" & ")}` : ""}`,
      total,
      isNumeric: false,
    };

    const filterLabelsM: Record<string, string> = {};
    for (const dim of filterOnlyDimsM) filterLabelsM[dim.col] = getDimFilterLabel(dim);

    function buildMultiRow(
      groupRows: Record<string, unknown>[],
      extraLabels: Record<string, string>
    ): ContributionRow {
      const count = groupRows.length;
      return {
        labels: { ...extraLabels },
        values: { [synthId]: count },
        percentages: { [synthId]: total > 0 ? (count / total) * 100 : 0 },
      };
    }

    // 3. Group visible rows by group-by dimensions
    if (groupKeyDimsM.length === 0) {
      return { rows: [buildMultiRow(visible, filterLabelsM)], metricMetas: [synthMeta], pivotValues: undefined };
    }

    const groupsM = new Map<string, Record<string, unknown>[]>();
    for (const row of visible) {
      const key = groupKeyDimsM.map((d) => `${d.col}|||${getGroupKey(row, d)}`).join("^^^");
      if (!groupsM.has(key)) groupsM.set(key, []);
      groupsM.get(key)!.push(row);
    }

    const multiRows: ContributionRow[] = [];
    for (const [key, groupRows] of groupsM) {
      const labels = { ...filterLabelsM };
      for (const part of key.split("^^^")) {
        const sep = part.indexOf("|||");
        if (sep !== -1) labels[part.slice(0, sep)] = part.slice(sep + 3);
      }
      multiRows.push(buildMultiRow(groupRows, labels));
    }

    multiRows.sort((a, b) => (b.values[synthId] ?? 0) - (a.values[synthId] ?? 0));
    return { rows: multiRows, metricMetas: [synthMeta], pivotValues: undefined };
  }

  // Metric metadata and global totals — always computed from allRows (metric conditions only).
  // Dimension filters must NOT change the denominator; they only affect grouping.
  const metricMetas: MetricMeta[] = metrics.map((m) => {
    const isNumericM = m.col !== "__count__" && numericCols.has(m.col);
    const isCategorical = m.col !== "__count__" && !isNumericM;

    let globalRows: Record<string, unknown>[];
    if (pivotValues && m === onlyMetric) {
      globalRows = allRows.filter((r) =>
        pivotValues.includes(String(r[m.col] ?? ""))
      );
    } else {
      globalRows = applyMetricFilters(allRows, m);
      if (isCategorical && m.filterValues.length > 0) {
        globalRows = globalRows.filter((r) =>
          m.filterValues.includes(String(r[m.col] ?? ""))
        );
      }
    }
    const total = isNumericM ? sumRows(globalRows, m.col) : globalRows.length;

    let label: string;
    if (m.col === "__count__") {
      label = "Count";
    } else if (isNumericM) {
      const rangePart =
        m.numericRange && (m.numericRange.min !== undefined || m.numericRange.max !== undefined)
          ? ` [${numericRangeLabel(m.numericRange)}]`
          : "";
      label = `${m.col}${rangePart} (sum)`;
    } else if (m.filterValues.length > 0) {
      label = `${m.col} = ${m.filterValues.join("/")} (count)`;
    } else {
      label = `${m.col} (count)`;
    }
    if (m.dateFilterCol && m.dateRange && (m.dateRange.start || m.dateRange.end)) {
      label += ` · ${formatDateRangeLabel(m.dateRange)}`;
    }

    return { id: m.id, label, total, isNumeric: isNumericM };
  });

  const groupKeyDims = dimensions.filter(isDimGroupKey);
  const filterOnlyDims = dimensions.filter((d) => !isDimGroupKey(d));

  function buildRow(
    groupRows: Record<string, unknown>[],
    extraLabels: Record<string, string>
  ): ContributionRow {
    const labels = { ...extraLabels };
    const values: Record<string, number> = {};
    const percentages: Record<string, number> = {};
    for (let i = 0; i < metrics.length; i++) {
      const m = metrics[i];
      const meta = metricMetas[i];
      const v = pivotValues
        ? groupRows.length
        : computeMetricVal(m, meta.isNumeric, groupRows);
      values[m.id] = v;
      percentages[m.id] = meta.total > 0 ? (v / meta.total) * 100 : 0;
    }
    const row: ContributionRow = { labels, values, percentages };
    if (pivotValues) row.breakdown = computeBreakdown(groupRows, onlyMetric!.col, pivotValues);
    return row;
  }

  const filterLabels: Record<string, string> = {};
  for (const dim of filterOnlyDims) filterLabels[dim.col] = getDimFilterLabel(dim);

  // No group key dims → single aggregate row
  if (groupKeyDims.length === 0) {
    return {
      rows: applyPostAggregationFilter([buildRow(working, filterLabels)], metrics, metricMetas),
      metricMetas,
      pivotValues,
    };
  }

  // Group by group key dims
  const groups = new Map<string, Record<string, unknown>[]>();
  for (const row of working) {
    const key = groupKeyDims.map((d) => `${d.col}|||${getGroupKey(row, d)}`).join("^^^");
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key)!.push(row);
  }

  const rows: ContributionRow[] = [];
  for (const [key, groupRows] of groups) {
    const labels = { ...filterLabels };
    for (const part of key.split("^^^")) {
      const sep = part.indexOf("|||");
      if (sep !== -1) labels[part.slice(0, sep)] = part.slice(sep + 3);
    }
    rows.push(buildRow(groupRows, labels));
  }

  rows.sort((a, b) => (b.values[metrics[0]?.id] ?? 0) - (a.values[metrics[0]?.id] ?? 0));
  return { rows: applyPostAggregationFilter(rows, metrics, metricMetas), metricMetas, pivotValues };
}
