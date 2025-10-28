// makeExcel.ts
import ExcelJS from "exceljs";

/** If you already have these types elsewhere, replace with your import and delete these. */
export interface ParsedRow {
  studentNo?: string;
  sin?: string;
  gov?: string;
  cert?: string;
  name?: string;
  eosDate?: string;
  amount?: string | number;
}
export interface ParsedPage {
  sheetName?: string;
  headerLines?: string[];
  columns?: string[];            // e.g. ['STUDENT NO','SIN','GOV','CERT #','ABBREVIATED NAME','EOS DATE','AMOUNT']
  rows: ParsedRow[];
  totals?: { count?: number; amount?: number | string };
}

/* ------------------------- generic helpers ------------------------- */

function asString(v: unknown): string {
  return v == null ? "" : typeof v === "string" ? v : String(v);
}

/** Convert "2,205.00", "$2,205.00", 2205, etc. → number. */
function toNumber(v: unknown): number {
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  const n = Number(asString(v).replace(/[,$\s$]/g, "") || 0);
  return Number.isFinite(n) ? n : 0;
}

/** Pretty print 2-decimal money. */
function fmtMoney(n: number): string {
  return n.toLocaleString(undefined, {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

/** Sum page amounts robustly. */
function sumAmounts(rows: { amount?: string | number }[]): number {
  return rows.reduce((acc, r) => acc + toNumber(r.amount), 0);
}

/* ------------------------- border helpers ------------------------- */

const THIN = { style: "thin" as const };
const borderThin = { top: THIN, left: THIN, bottom: THIN, right: THIN };

/* ------------------- sheet name safety & uniqueness ------------------- */

function safeSheetName(raw: unknown, indexZeroBased: number): string {
  let s = asString(raw).replace(/[\\/*?:\[\]:]/g, "").trim();
  if (!s) s = `Page ${indexZeroBased + 1}`;
  if (s.length > 31) s = s.slice(0, 31);
  return s;
}

function uniqueSheetName(base: string, used: Set<string>): string {
  let name = base;
  let n = 2;
  while (used.has(name)) {
    // reserve space for " (n)"
    const stem = base.slice(0, 31 - 4);
    name = `${stem} (${n})`;
    n++;
  }
  used.add(name);
  return name;
}

/* ------------------------- main export ------------------------- */

export async function createWorkbookFromPages(pages: ParsedPage[]) {
  const wb = new ExcelJS.Workbook();
  const used = new Set<string>(); // track names to avoid "Worksheet name already exists"

  const defaultHeaders = [
    "STUDENT NO",
    "SIN",
    "GOV",
    "CERT #",
    "ABBREVIATED NAME",
    "EOS DATE",
    "AMOUNT",
  ];

  pages.forEach((page, idx) => {
    // Compute a safe, unique sheet name (sheetName > first header line > "Page N")
    const base = safeSheetName(page.sheetName ?? page.headerLines?.[0] ?? `Page ${idx + 1}`, idx);
    const wsName = uniqueSheetName(base, used);
    const ws = wb.addWorksheet(wsName);

    // Use parser-supplied columns or fallback to defaults
    const headers =
      Array.isArray(page.columns) && page.columns.length
        ? page.columns.map(asString)
        : defaultHeaders;

    // Reasonable column widths based on header text
    ws.columns = headers.map((h) => ({
      header: h,
      width: Math.max(12, Math.min(40, h.length + 6)),
    }));

    let rowIndex = 1;

    /* ---------- free-form header lines (exactly as parsed) ----------- */
    if (Array.isArray(page.headerLines) && page.headerLines.length) {
      for (const line of page.headerLines) {
        ws.mergeCells(rowIndex, 1, rowIndex, headers.length);
        const c = ws.getCell(rowIndex, 1);
        c.value = asString(line);
        c.alignment = { vertical: "middle", horizontal: "center" };
        rowIndex++;
      }
      rowIndex++; // blank spacer row
    }

    /* ------------------------ table header row ------------------------ */
    const headerRow = ws.getRow(rowIndex);
    headerRow.font = { bold: true };
    headerRow.alignment = { vertical: "middle", horizontal: "center" };
    headerRow.eachCell((c) => (c.border = borderThin));
    rowIndex++;

    /* --------------------------- data rows --------------------------- */

    // Maps common headers to row fields; if you pass custom headers they’ll be looked up dynamically
    const headerKeyMap = new Map<string, (r: ParsedRow) => unknown>([
      ["STUDENT NO", (r) => r.studentNo],
      ["SIN", (r) => r.sin],
      ["GOV", (r) => r.gov],
      ["CERT #", (r) => r.cert],
      ["ABBREVIATED NAME", (r) => r.name],
      ["EOS DATE", (r) => r.eosDate],
      ["AMOUNT", (r) => r.amount],
    ]);

    const rows: ParsedRow[] = Array.isArray(page.rows) ? page.rows : [];

    for (const r of rows) {
      headers.forEach((h, i) => {
        const getter = headerKeyMap.get(h.toUpperCase());
        const raw = getter ? getter(r) : (r as any)[h]; // support custom keys if your parser populates them
        const cell = ws.getCell(rowIndex, i + 1);

        if (h.toUpperCase().includes("AMOUNT")) {
          const n = toNumber(raw);
          cell.value = Number.isFinite(n) ? `$${fmtMoney(n)}` : asString(raw);
        } else {
          cell.value = asString(raw);
        }

        cell.border = borderThin;
        cell.alignment = { vertical: "middle", horizontal: "center" };
      });
      rowIndex++;
    }

    /* ---------------------------- totals ---------------------------- */

    const totalCount =
      page.totals?.count ?? (Array.isArray(page.rows) ? page.rows.length : 0);
    const totalAmountNum =
      page.totals?.amount != null ? toNumber(page.totals.amount) : sumAmounts(rows);

    rowIndex++;
    // Label "TOTAL FOR THIS DISBURSEMENT" across all but the last two columns
    ws.mergeCells(rowIndex, 1, rowIndex, Math.max(1, headers.length - 2));
    ws.getCell(rowIndex, 1).value = "TOTAL FOR THIS DISBURSEMENT";
    ws.getCell(rowIndex, headers.length - 1).value = totalCount;
    ws.getCell(rowIndex, headers.length).value = `$${fmtMoney(totalAmountNum)}`;

    for (let c = 1; c <= headers.length; c++) {
      const cell = ws.getCell(rowIndex, c);
      cell.font = { bold: true };
      cell.border = borderThin;
      cell.alignment = { vertical: "middle", horizontal: "center" };
    }
  });

  return wb;
}
