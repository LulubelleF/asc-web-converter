import Tesseract from "tesseract.js";
import type { ParsedPage, ParsedRow } from "./types";

const normalize = (s: string) =>
  s.replace(/\u00A0/g, " ").replace(/[ \t]+/g, " ").trim();

function looksLikeHeader(line: string): boolean {
  const L = line.toUpperCase();
  const tokens = ["STUDENT", "SIN", "GOV", "CERT", "NAME", "EOS", "DATE", "AMOUNT"];
  return tokens.reduce((n, t) => (L.includes(t) ? n + 1 : n), 0) >= 3;
}

function splitColumns(header: string): string[] {
  const parts = header
    .replace(/\|/g, "  ")
    .split(/ {2,}|\t+/)
    .map((s) => s.trim())
    .filter(Boolean);
  return parts.length >= 3
    ? parts
    : ["STUDENT NO", "SIN", "GOV", "CERT #", "ABBREVIATED NAME", "EOS DATE", "AMOUNT"];
}

function toNumber(v: unknown): number {
  if (typeof v === "number") return isFinite(v) ? v : 0;
  const n = Number(String(v).replace(/[,$\s$]/g, "") || 0);
  return isFinite(n) ? n : 0;
}

function parseTotals(line: string) {
  const L = line.toUpperCase();
  if (!L.includes("TOTAL")) return null;
  const mAmount = line.match(/[$]?\s*([\d,]+\.\d{2})\b/);
  const amount = mAmount ? toNumber(mAmount[1]) : undefined;
  const mCount = line.match(/\b(\d{1,4})\b(?!.*\b\d{1,4}\b)/);
  const count = mCount ? Number(mCount[1]) : undefined;
  if (amount == null && count == null) return null;
  return { count, amount };
}

function parseRow(line: string): ParsedRow | null {
  const l = normalize(line);
  const rx =
    /^(\d{3,})\s+(\d{3}-\d{3}-\d{3})\s+([A-Z]{2})?\s*(\d{4,})\s+(.+?)\s+(\d{2}\/\d{2}\/\d{2})\s+([$]?[\d,]+\.\d{2})$/;
  const m = l.match(rx);
  if (!m) return null;
  return {
    studentNo: m[1],
    sin: m[2],
    gov: (m[3] || "").trim() || undefined,
    cert: m[4],
    name: m[5],
    eosDate: m[6],
    amount: m[7],
  };
}

async function ocrCanvas(c: HTMLCanvasElement): Promise<string> {
  // These keys are valid for Tesseract itself but not declared in its TS types.
  // Cast to `any` so TypeScript doesn't complain.
  const ocrOptions: any = {
    // keep only characters we care about in these reports
    tessedit_char_whitelist:
      "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-/$:,.()# ",
    // keep spaces between words to make column splitting easier
    preserve_interword_spaces: "1",
    // You can also tweak page segmentation if needed, e.g.:
    // tessedit_pageseg_mode: "6", // equivalent to `--psm 6`
  };

  const res = await Tesseract.recognize(
    c,
    "eng",
    ocrOptions as any // type escape: the library supports it, TS typings do not
  );

  return res?.data?.text ?? "";
}

function parseFromLines(lines: string[], pageIndex: number): ParsedPage {
  const clean = lines.map(normalize).filter(Boolean);
  let headerIdx = clean.findIndex(looksLikeHeader);
  if (headerIdx < 0) headerIdx = Math.min(4, clean.length);

  const headerLines = clean.slice(0, headerIdx);
  const headerLine = clean[headerIdx] || "";
  const columns = splitColumns(headerLine);

  const rows: ParsedRow[] = [];
  let totals: { count?: number; amount?: number } | undefined;

  for (let i = headerIdx + 1; i < clean.length; i++) {
    const ln = clean[i];
    const t = parseTotals(ln);
    if (t) {
      totals = t;
      break;
    }
    const r = parseRow(ln);
    if (r) rows.push(r);
  }

  const sheetName = headerLines[0] || `Page ${pageIndex}`;
  return { sheetName, headerLines, columns, rows, totals };
}

// Public APIs used by index.tsx
export async function parseDisbursementPageFromText(
  rawText: string,
  opts: { pageIndex: number }
) {
  const lines = rawText.split(/\r?\n/);
  return parseFromLines(lines, opts.pageIndex);
}

export async function parseDisbursementPageFromCanvas(
  canvas: HTMLCanvasElement,
  opts: { pageIndex: number }
) {
  const text = await ocrCanvas(canvas);
  const lines = text.split(/\r?\n/);
  return parseFromLines(lines, opts.pageIndex);
}
