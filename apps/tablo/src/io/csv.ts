/**
 * CSV içe/dışa aktarma. Türkçe CSV ";" ayırıcı ve UTF-8 BOM ile yazılır;
 * içe aktarmada kodlama ve ayırıcı tahmin edilir.
 */
import { parseCellInput, type Scalar, MAX_COLS, MAX_ROWS } from "@officehere/formula";
import { displayText } from "../model/display";
import { WorkbookModel, type CellInput, type SheetModel } from "../model/workbook";

export interface CsvOptions {
  delimiter?: string;
  bom?: boolean;
  /** "displayed": görünen metin (Excel gibi), "raw": ham değer. */
  values?: "displayed" | "raw";
}

function field(s: string, d: string): string {
  if (s.includes(d) || s.includes('"') || s.includes("\n") || s.includes("\r") || /^\s|\s$/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

export function writeCsv(wb: WorkbookModel, sheet: SheetModel, opts: CsvOptions = {}): string {
  const d = opts.delimiter ?? ";";
  const used = wb.usedRange(sheet);
  if (!used) return opts.bom === false ? "" : "﻿";
  const lines: string[] = [];
  for (let r = 0; r <= used.r2; r++) {
    const cells: string[] = [];
    for (let c = 0; c <= used.c2; c++) {
      const v = wb.getValue(sheet, r, c);
      let s: string;
      if (opts.values === "raw") s = v === null ? "" : typeof v === "number" ? String(v).replace(".", d === ";" ? "," : ".") : typeof v === "boolean" ? (v ? "DOĞRU" : "YANLIŞ") : String(v);
      else s = displayText(wb, sheet, r, c);
      cells.push(field(s, d));
    }
    lines.push(cells.join(d));
  }
  return (opts.bom === false ? "" : "﻿") + lines.join("\r\n") + "\r\n";
}

/** Bayt dizisini metne çevirir: BOM, UTF-8, olmazsa Windows-1254. */
export function decodeText(data: ArrayBuffer | Uint8Array): string {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) return new TextDecoder("utf-8").decode(bytes.subarray(3));
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder("utf-16le").decode(bytes.subarray(2));
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder("utf-16be").decode(bytes.subarray(2));
  // BOM'suz UTF-16: ASCII ağırlıklı metinde baytların yarısı sıfırdır
  const sample = bytes.subarray(0, Math.min(bytes.length, 512)) as Uint8Array;
  if (sample.length >= 4) {
    let zeroOdd = 0;
    let zeroEven = 0;
    for (let i = 0; i + 1 < sample.length; i += 2) {
      if (sample[i] === 0) zeroEven++;
      if (sample[i + 1] === 0) zeroOdd++;
    }
    const pairs = Math.floor(sample.length / 2);
    if (zeroOdd > pairs * 0.6 && zeroEven < pairs * 0.1) return new TextDecoder("utf-16le").decode(bytes);
    if (zeroEven > pairs * 0.6 && zeroOdd < pairs * 0.1) return new TextDecoder("utf-16be").decode(bytes);
  }
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    try {
      return new TextDecoder("windows-1254").decode(bytes);
    } catch {
      return new TextDecoder("latin1").decode(bytes);
    }
  }
}

/** Ayırıcı tahmini: ilk satırlarda tutarlı en çok görülen aday. */
export function detectDelimiter(text: string): string {
  const lines = text.split(/\r?\n/).filter((l) => l.trim()).slice(0, 20);
  const cands = [";", ",", "\t", "|"];
  let best = ";";
  let bestScore = -1;
  for (const c of cands) {
    const counts = lines.map((l) => splitLine(l, c).length - 1);
    if (!counts.length || counts[0] === 0) continue;
    const consistent = counts.filter((n) => n === counts[0]).length / counts.length;
    const score = counts[0] * consistent;
    if (score > bestScore) {
      bestScore = score;
      best = c;
    }
  }
  return best;
}

function splitLine(line: string, d: string): string[] {
  const out: string[] = [];
  let cur = "";
  let q = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (q) {
      if (ch === '"') {
        if (line[i + 1] === '"') {
          cur += '"';
          i++;
        } else q = false;
      } else cur += ch;
    } else if (ch === '"' && cur === "") q = true;
    else if (ch === d) {
      out.push(cur);
      cur = "";
    } else cur += ch;
  }
  out.push(cur);
  return out;
}

/** Tırnaklı satır sonlarını destekleyen tam CSV ayrıştırıcı. */
export function parseCsv(text: string, d: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cur = "";
  let q = false;
  const t = text.replace(/^﻿/, "");
  for (let i = 0; i < t.length; i++) {
    const ch = t[i];
    if (q) {
      if (ch === '"') {
        if (t[i + 1] === '"') {
          cur += '"';
          i++;
        } else q = false;
      } else cur += ch;
      continue;
    }
    if (ch === '"' && cur === "") q = true;
    else if (ch === d) {
      row.push(cur);
      cur = "";
    } else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && t[i + 1] === "\n") i++;
      row.push(cur);
      rows.push(row);
      row = [];
      cur = "";
    } else cur += ch;
  }
  if (cur !== "" || row.length) {
    row.push(cur);
    rows.push(row);
  }
  return rows;
}

export interface CsvImportOptions {
  delimiter?: string;
  /** Sayı biçimi yerel ayarı: "tr" (1.234,5) ya da "en" (1,234.5). */
  locale?: "tr" | "en";
  /** Tüm alanları metin olarak al. */
  asText?: boolean;
}

/** Sekme/dikey çizgi ayırıcılı metinde sayıların yerel ayarı: "1.234,5" mi "1,234.5" mi? */
export function guessLocale(grid: string[][]): "tr" | "en" {
  let tr = 0;
  let en = 0;
  for (const row of grid.slice(0, 200))
    for (const f of row) {
      const s = f.trim();
      if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$|^-?\d+,\d+$/.test(s)) tr++;
      else if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$|^-?\d+\.\d+$/.test(s)) en++;
    }
  return en > tr ? "en" : "tr";
}

/** CSV metnini hücre girdilerine çevirir. */
export function csvToInputs(text: string, opts: CsvImportOptions = {}): { rows: { input: CellInput; numFmt?: string }[][]; delimiter: string } {
  const d = opts.delimiter ?? detectDelimiter(text);
  const grid = parseCsv(text, d);
  const locale = opts.locale ?? (d === "," ? "en" : d === ";" ? "tr" : guessLocale(grid));
  const rows = grid.map((r) =>
    r.map((s) => {
      if (opts.asText) return { input: s === "" ? null : s };
      const p = parseCellInput(s, { locale });
      if (p.kind === "empty") return { input: null };
      if (p.kind === "formula") return { input: { formula: p.formula! } };
      return { input: (p.value ?? null) as Scalar, numFmt: p.numFmt };
    }),
  );
  return { rows, delimiter: d };
}

/** CSV dosyasını yeni çalışma kitabı olarak açar. */
export function readCsv(data: ArrayBuffer | Uint8Array, sheetName = "Sayfa1", opts: CsvImportOptions = {}): WorkbookModel {
  const wb = new WorkbookModel(true);
  const sheet = wb.addSheet(sheetName.slice(0, 31).replace(/[\\/?*[\]:]/g, "_") || "Sayfa1");
  const { rows } = csvToInputs(decodeText(data), opts);
  // Sayfa sınırı dışındaki satır/sütunlar atlanır (16.384. sütundan sonrası hücre anahtarında
  // sonraki satıra taşar ve başka hücrelerin üzerine yazardı)
  let clipped = rows.length > MAX_ROWS;
  for (let i = 0; i < Math.min(rows.length, MAX_ROWS); i++) {
    const r = rows[i];
    if (r.length > MAX_COLS) clipped = true;
    for (let j = 0; j < Math.min(r.length, MAX_COLS); j++) {
      const c = r[j];
      if (c.input === null) continue;
      wb.setInputRaw(sheet, i, j, c.input);
      if (c.numFmt) wb.setStyleRaw(sheet, i, j, wb.styles.intern({ numFmt: c.numFmt }));
    }
  }
  if (clipped) wb.loadWarnings = ["sayfa sınırını aşan satır ve sütunlar (en çok 1.048.576 satır, 16.384 sütun)"];
  wb.engine.recalc();
  return wb;
}
