import * as XLSX from "xlsx";

export interface ExcelData {
  headers: string[];
  rows: Record<string, unknown>[];
  sheetName: string;
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
        for (let r = headerRowIndex + 1; r < rawRows.length; r++) {
          const rawRow = rawRows[r];
          // Skip completely empty rows
          const hasData = rawRow.some(
            (v) => v !== null && v !== undefined && String(v).trim() !== ""
          );
          if (!hasData) continue;

          const obj: Record<string, unknown> = {};
          for (let ci = 0; ci < colIndices.length; ci++) {
            const colIdx = colIndices[ci];
            let val = rawRow[colIdx] ?? "";
            // Format dates as readable strings
            if (val instanceof Date) {
              val = val.toLocaleDateString("en-GB", {
                day: "2-digit", month: "short", year: "numeric",
              });
            }
            obj[headers[ci]] = val;
          }
          rows.push(obj);
        }

        if (rows.length === 0) {
          reject(new Error("No data rows found after the header."));
          return;
        }

        resolve({ headers, rows, sheetName });
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

export function getColumnStats(rows: Record<string, unknown>[], col: string) {
  const values = rows.map((r) => r[col]);
  const numericValues = values.filter(isNumeric).map(Number);
  const isNumericCol = numericValues.length > rows.length * 0.5;
  const uniqueValues = [...new Set(values.map((v) => String(v ?? "")))];

  return {
    isNumeric: isNumericCol,
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

  // Group by separate columns + multi-select columns
  const groupCols = [...separateCols, ...multiSelectCols];
  const groups: Map<string, Record<string, unknown>[]> = new Map();
  for (const row of filtered) {
    const key = groupCols
      .map((c) => `${c.col}:::${String(row[c.col] ?? "")}`)
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
