/**
 * Tablo stilleri (Tablo Olarak Biçimlendir galerisi). Renkler tema
 * vurgularından türetilir; adlar ECMA-376 yerleşik stil adlarıyla eşleşir.
 */
import { tint } from "@officehere/core";
import type { CellStyle } from "./styles";
import type { TableDef } from "./types";

/** Siyah + tema vurguları 1-6 (tema değişince setTableThemeAccents ile güncellenir). */
const BASE = ["#000000", "#156082", "#E97132", "#196B24", "#0F9ED5", "#A02B93", "#4EA72E"];

export interface TableStylePreset {
  id: string;
  label: string;
  group: "Açık" | "Orta" | "Koyu";
  color: string;
}

export const TABLE_STYLES: TableStylePreset[] = [
  ...BASE.map((c, i) => ({ id: `TableStyleLight${i + 1}`, label: `Açık Renkli ${i + 1}`, group: "Açık" as const, color: c })),
  ...BASE.map((c, i) => ({ id: `TableStyleMedium${i + 1}`, label: `Orta Renkli ${i + 1}`, group: "Orta" as const, color: c })),
  ...BASE.map((c, i) => ({ id: `TableStyleDark${i + 1}`, label: `Koyu Renkli ${i + 1}`, group: "Koyu" as const, color: c })),
];

export const DEFAULT_TABLE_STYLE = "TableStyleMedium2";

export function presetOf(id: string): TableStylePreset {
  const p = TABLE_STYLES.find((t) => t.id === id);
  if (p) return p;
  // Bilinmeyen yerleşik stil: numarasına göre yaklaşık renk
  const m = /TableStyle(Light|Medium|Dark)(\d+)/.exec(id);
  if (m) {
    const n = (parseInt(m[2], 10) - 1) % 7;
    const group = m[1] === "Light" ? "Açık" : m[1] === "Medium" ? "Orta" : "Koyu";
    return { id, label: id, group, color: BASE[n] };
  }
  return TABLE_STYLES[8];
}

/** Tablo içindeki hücrenin stil katkısı (açık dolgu yoksa uygulanır). */
export function tableCellStyle(t: TableDef, row: number, col: number): Partial<CellStyle> {
  const p = presetOf(t.style);
  const c = p.color;
  const isHeader = t.headerRow && row === t.ref.r1;
  const isTotal = t.totalRow && row === t.ref.r2;
  const dataStart = t.ref.r1 + (t.headerRow ? 1 : 0);
  const bandIdx = row - dataStart;
  const colIdx = col - t.ref.c1;
  const out: Partial<CellStyle> = {};
  const line = { style: "thin" as const, color: p.group === "Açık" ? c : tint(c, 0.4) };
  if (p.group === "Açık") {
    if (isHeader) {
      out.bold = true;
      out.bBottom = line;
      out.color = c === "#000000" ? "#000000" : tint(c, -0.25);
    } else if (isTotal) {
      out.bold = true;
      out.bTop = { style: "double", color: c };
    } else {
      if (t.bandedRows && bandIdx % 2 === 0) out.fill = tint(c, c === "#000000" ? 0.85 : 0.8);
      if (t.bandedCols && colIdx % 2 === 0) out.fill = tint(c, 0.8);
    }
    if (row === t.ref.r1) out.bTop = line;
    if (row === t.ref.r2) out.bBottom = isTotal ? line : line;
  } else if (p.group === "Orta") {
    if (isHeader) {
      out.fill = c;
      out.color = "#FFFFFF";
      out.bold = true;
    } else if (isTotal) {
      out.bold = true;
      out.bTop = { style: "double", color: c };
      out.fill = "#FFFFFF";
    } else {
      out.fill = t.bandedRows && bandIdx % 2 === 0 ? tint(c, c === "#000000" ? 0.8 : 0.8) : "#FFFFFF";
      if (t.bandedCols && colIdx % 2 === 0) out.fill = tint(c, 0.8);
      out.bTop = { style: "thin", color: tint(c, 0.6) };
      out.bBottom = { style: "thin", color: tint(c, 0.6) };
    }
  } else {
    if (isHeader) {
      out.fill = "#000000";
      out.color = "#FFFFFF";
      out.bold = true;
      out.bBottom = { style: "medium", color: "#FFFFFF" };
    } else if (isTotal) {
      out.fill = tint(c, -0.5);
      out.color = "#FFFFFF";
      out.bold = true;
    } else {
      out.fill = t.bandedRows && bandIdx % 2 === 0 ? tint(c, -0.25) : c === "#000000" ? "#404040" : c;
      out.color = "#FFFFFF";
    }
  }
  if ((t.firstColumn && col === t.ref.c1) || (t.lastColumn && col === t.ref.c2)) out.bold = true;
  return out;
}

/** Galeri önizlemesi için renkler. */
export function previewColors(p: TableStylePreset): { header: string; headerText: string; band: string; body: string; line: string } {
  const c = p.color;
  if (p.group === "Açık") return { header: "#FFFFFF", headerText: c, band: tint(c, 0.8), body: "#FFFFFF", line: c };
  if (p.group === "Orta") return { header: c, headerText: "#FFFFFF", band: tint(c, 0.8), body: "#FFFFFF", line: tint(c, 0.6) };
  return { header: "#000000", headerText: "#FFFFFF", band: tint(c, -0.25), body: c === "#000000" ? "#404040" : c, line: "#FFFFFF" };
}

/** Tablo stillerini çalışma kitabı temasının vurgu renklerine bağlar. */
export function setTableThemeAccents(accents: readonly string[]): void {
  for (let i = 0; i < 6; i++) BASE[i + 1] = accents[i];
  for (const t of TABLE_STYLES) {
    const n = (parseInt(/(\d+)$/.exec(t.id)?.[1] ?? "1", 10) - 1) % 7;
    t.color = BASE[n];
  }
}
