import * as XLSX from "xlsx";

export interface ExcelData {
  headers: string[];
  rows: Record<string, unknown>[];
  sheetName: string;
  dateColumns: string[];
}

export function parseExcelFile(file: File): Promise<ExcelData> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: "array", cellDates: true });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Get all rows as arrays (including blanks) to find the real header row
        const rawRows: unknown[][] = XLSX.utils.sheet_to_json(worksheet, {
          header: 1,
          defval: null,
          blankrows: true,
        });

        if (rawRows.length === 0) {
          reject(new Error("The Excel file appears to be empty."));
          return;
        }

        // Find the header row: pick the row (within first 30) with the MOST non-empty string cells.
        // This handles files with metadata rows that may also have 1-2 string cells.
        const scanLimit = Math.min(rawRows.length, 30);
        let bestScore = -1;
        let headerRowIndex = -1;
        for (let i = 0; i < scanLimit; i++) {
          const row = rawRows[i];
          const stringCells = row.filter(
            (v) => v !== null && v !== undefined && typeof v === "string" && String(v).trim() !== ""
          );
          if (stringCells.length >= 3 && stringCells.length > bestScore) {
            bestScore = stringCells.length;
            headerRowIndex = i;
          }
        }

        if (headerRowIndex === -1) {
          reject(new Error("Could not find a header row. Ensure the column header row has at least 3 text columns."));
          return;
        }

        const headerRow = rawRows[headerRowIndex] as unknown[];

        // Build header list, skipping empty columns
        const headers: string[] = [];
        const colIndices: number[] = [];
        for (let c = 0; c < headerRow.length; c++) {
          const h = headerRow[c];
          if (h !== null && h !== undefined && String(h).trim() !== "") {
            headers.push(String(h).trim());
            colIndices.push(c);
          }
        }

        // Build data rows from rows after the header
        const rows: Record<string, unknown>[] = [];
        const dateColCount: Record<number, number> = {};
        for (let r = headerRowIndex + 1; r < rawRows.length; r++) {
          const rawRow = rawRows[r];

          // Stop at footer rows: 2+ cells starting with "Total", "Grand Total", "Subtotal", etc.
          const footerCells = rawRow.filter(
            (v) => typeof v === "string" && /^(total|grand\s*total|sub\s*total|subtotal)\b/i.test(String(v).trim())
          );
          if (footerCells.length >= 2) break;

          // Skip completely empty rows
          const hasData = rawRow.some(
            (v) => v !== null && v !== undefined && String(v).trim() !== ""
          );
          if (!hasData) continue;

          const obj: Record<string, unknown> = {};
          for (let ci = 0; ci < colIndices.length; ci++) {
            const colIdx = colIndices[ci];
            let val = rawRow[colIdx] ?? "";
            // Format dates/datetimes as readable strings, track which header indices are dates
            if (val instanceof Date) {
              dateColCount[ci] = (dateColCount[ci] ?? 0) + 1;
              const hasTime = val.getHours() !== 0 || val.getMinutes() !== 0 || val.getSeconds() !== 0;
              const datePart = val.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
              val = hasTime
                ? datePart + " " + val.toLocaleTimeString("en-GB", { hour: "2-digit", minute: "2-digit", hour12: false })
                : datePart;
            }
            obj[headers[ci]] = val;
          }
          rows.push(obj);
        }

        if (rows.length === 0) {
          reject(new Error("No data rows found after the header."));
          return;
        }

        const dateColumns = headers.filter(
          (_, i) => (dateColCount[i] ?? 0) > rows.length * 0.5
        );

        resolve({ headers, rows, sheetName, dateColumns });
      } catch (err) {
        reject(new Error("Failed to parse Excel file: " + (err as Error).message));
      }
    };
    reader.onerror = () => reject(new Error("Failed to read file."));
    reader.readAsArrayBuffer(file);
  });
}

export function isNumeric(value: unknown): boolean {
  if (value === null || value === undefined || value === "") return false;
  return !isNaN(Number(value));
}

function looksLikeDate(v: unknown): boolean {
  if (!v || typeof v !== "string") return false;
  // "01 Jan 2026" or "01 Jan 2026 12:19" (our formatted output)
  if (/^\d{2} [A-Z][a-z]{2} \d{4}/.test(v)) return true;
  // "2026-05-01" or "2026-05-01 12:19" or "2026-05-01T12:19"
  if (/^\d{4}-\d{2}-\d{2}([ T]\d{2}:\d{2})?$/.test(v)) return true;
  return false;
}

export function getColumnStats(rows: Record<string, unknown>[], col: string) {
  const values = rows.map((r) => r[col]);
  const numericValues = values.filter(isNumeric).map(Number);
  const isNumericCol = numericValues.length > rows.length * 0.5;
  const uniqueValues = [...new Set(values.map((v) => String(v ?? "")))];
  const dateLikeCount = values.filter(looksLikeDate).length;
  const isDateCol = !isNumericCol && dateLikeCount > rows.length * 0.5;

  return {
    isNumeric: isNumericCol,
    isDate: isDateCol,
    uniqueCount: uniqueValues.length,
    uniqueValues: uniqueValues.slice(0, 50),
    numericValues,
  };
}

export type ColumnMode = "ignore" | "select" | "separate";

export interface ColumnConfig {
  col: string;
  mode: ColumnMode;
  selectedValues: string[];
  dateRange?: { start?: string; end?: string };
  dateGroupMode?: "split" | "combine";
  numericRange?: { min?: string; max?: string };
  numericGroupMode?: "split" | "combine";
}

function parseLocalDate(dateStr: string): Date | null {
  if (!dateStr) return null;
  // "YYYY-MM-DD" — avoid UTC interpretation
  const m1 = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m1) return new Date(Number(m1[1]), Number(m1[2]) - 1, Number(m1[3]));
  // "YYYY-MM-DD HH:MM" or "YYYY-MM-DDTHH:MM" — local datetime
  const m2 = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})/);
  if (m2) return new Date(Number(m2[1]), Number(m2[2]) - 1, Number(m2[3]), Number(m2[4]), Number(m2[5]));
  const d = new Date(dateStr);
  return isNaN(d.getTime()) ? null : d;
}

export interface SummaryRow {
  labels: Record<string, string>;
  metrics: Record<string, number>;
  count: number;
}

export interface SortEntry {
  col: string;
  dir: "asc" | "desc";
}

export function computeSummary(
  rows: Record<string, unknown>[],
  columnConfigs: ColumnConfig[],
  numericCols: string[],
  sortOrder: SortEntry[],
): SummaryRow[] {
  const separateCols = columnConfigs.filter((c) => c.mode === "separate");
  const selectCols = columnConfigs.filter(
    (c) => c.mode === "select" && c.selectedValues.length > 0
  );
  // Filter cols with 2+ values are also grouped so each value gets its own row
  const multiSelectCols = selectCols.filter((c) => c.selectedValues.length > 1);

  // Filter rows by selected values
  let filtered = rows;
  for (const sc of selectCols) {
    filtered = filtered.filter((row) =>
      sc.selectedValues.includes(String(row[sc.col] ?? ""))
    );
  }

  // Filter rows by date ranges
  for (const cfg of columnConfigs) {
    if (cfg.mode !== "select" || !cfg.dateRange) continue;
    const { start, end } = cfg.dateRange;
    if (!start && !end) continue;
    filtered = filtered.filter((row) => {
      const rowDate = parseLocalDate(String(row[cfg.col] ?? ""));
      if (!rowDate) return true;
      if (start) {
        const s = parseLocalDate(start);
        if (s && rowDate < s) return false;
      }
      if (end) {
        const e = parseLocalDate(end);
        if (e) {
          e.setHours(23, 59, 59, 999);
          if (rowDate > e) return false;
        }
      }
      return true;
    });
  }

  // Filter rows by numeric ranges
  for (const cfg of columnConfigs) {
    if (cfg.mode !== "select" || !cfg.numericRange) continue;
    const { min, max } = cfg.numericRange;
    if (!min && !max) continue;
    filtered = filtered.filter((row) => {
      const raw = row[cfg.col];
      if (raw === "" || raw === null || raw === undefined) return true;
      const v = Number(raw);
      if (isNaN(v)) return true;
      if (min && v < Number(min)) return false;
      if (max && v > Number(max)) return false;
      return true;
    });
  }

  // Date range filter cols grouped into summary rows
  const dateRangeCols = columnConfigs.filter(
    (c) => c.mode === "select" && !!c.dateRange && (!!c.dateRange.start || !!c.dateRange.end)
  );

  // Numeric range filter cols grouped into summary rows
  const numericRangeCols = columnConfigs.filter(
    (c) => c.mode === "select" && !!c.numericRange && (!!c.numericRange.min || !!c.numericRange.max)
  );

  // For combine-mode date cols, pre-compute the "date1 + date2" label from actual filtered data
  const combinedDateLabels: Record<string, string> = {};
  for (const c of dateRangeCols) {
    if (c.dateGroupMode === "combine") {
      const unique = [...new Set(filtered.map((r) => String(r[c.col] ?? "")).filter(Boolean))].sort();
      combinedDateLabels[c.col] = unique.length > 1
        ? `${unique[0]} - ${unique[unique.length - 1]}`
        : (unique[0] ?? "");
    }
  }

  // For combine-mode numeric cols, use the range bounds as the label
  const combinedNumericLabels: Record<string, string> = {};
  for (const c of numericRangeCols) {
    if ((c.numericGroupMode ?? "split") === "combine") {
      const { min, max } = c.numericRange!;
      const parts: string[] = [];
      if (min) parts.push(`≥ ${min}`);
      if (max) parts.push(`≤ ${max}`);
      combinedNumericLabels[c.col] = parts.join(" & ");
    }
  }

  // Group by separate + multi-select + date range + numeric range columns
  const groupCols = [...separateCols, ...multiSelectCols, ...dateRangeCols, ...numericRangeCols];
  const groups: Map<string, Record<string, unknown>[]> = new Map();
  for (const row of filtered) {
    const key = groupCols
      .map((c) => {
        const combined = combinedDateLabels[c.col] ?? combinedNumericLabels[c.col];
        return combined !== undefined
          ? `${c.col}:::${combined}`
          : `${c.col}:::${String(row[c.col] ?? "")}`;
      })
      .join("||");
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key)!.push(row);
  }

  const results: SummaryRow[] = [];
  for (const [key, groupRows] of groups) {
    const labels: Record<string, string> = {};
    if (key) {
      for (const part of key.split("||")) {
        const sepIdx = part.indexOf(":::");
        if (sepIdx !== -1) labels[part.slice(0, sepIdx)] = part.slice(sepIdx + 3);
      }
    }

    const metrics: Record<string, number> = {};
    for (const nc of numericCols) {
      const vals = groupRows
        .map((r) => r[nc])
        .filter(isNumeric)
        .map(Number);
      metrics[nc] = vals.reduce((a, b) => a + b, 0);
    }

    results.push({ labels, metrics, count: groupRows.length });
  }

  // Multi-column sort: iterate sort entries in priority order
  if (sortOrder.length > 0) {
    results.sort((a, b) => {
      for (const { col, dir } of sortOrder) {
        let cmp = 0;
        if (col === "Count") {
          cmp = a.count - b.count;
        } else if (groupCols.some((c) => c.col === col)) {
          const av = a.labels[col] ?? "";
          const bv = b.labels[col] ?? "";
          cmp = isNumeric(av) && isNumeric(bv)
            ? Number(av) - Number(bv)
            : av.localeCompare(bv);
        } else if (numericCols.includes(col)) {
          cmp = (a.metrics[col] ?? 0) - (b.metrics[col] ?? 0);
        }
        if (cmp !== 0) return dir === "asc" ? cmp : -cmp;
      }
      return 0;
    });
  }

  return results;
}
