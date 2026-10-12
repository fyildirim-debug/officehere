/**
 * Tablo stilleri (bu projenin kendi tanımları). Dosyaya yazılırken
 * ppt/tableStyles.xml içine tam tanım olarak eklenir (ECMA-376 §20.1.4.2).
 */
import type { CellBorder, Color, Fill, SchemeColorName, TableShape } from "../model/types";

export interface BorderSpec {
  color: Color;
  width: number;
}

export interface PartStyle {
  fill?: Color | null;
  text?: Color;
  bold?: boolean;
  top?: BorderSpec | null;
  bottom?: BorderSpec | null;
  left?: BorderSpec | null;
  right?: BorderSpec | null;
  insideH?: BorderSpec | null;
  insideV?: BorderSpec | null;
}

export interface TableStyleDef {
  id: string;
  guid: string;
  name: string;
  wholeTbl: PartStyle;
  band1H?: PartStyle;
  band1V?: PartStyle;
  lastCol?: PartStyle;
  firstCol?: PartStyle;
  lastRow?: PartStyle;
  firstRow?: PartStyle;
}

const ACCENTS: (SchemeColorName | "dk1")[] = ["dk1", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
const accentName = (a: string) => (a === "dk1" ? "Koyu 1" : `Vurgu ${a.slice(6)}`);

function guid(n: number): string {
  // Bu projeye özgü sabit kimlikler.
  const hex = n.toString(16).toUpperCase().padStart(4, "0");
  return `{0FF1CE00-5A5A-4E7E-9B0B-00000000${hex}}`;
}

const W = (c: Color, width = 1): BorderSpec => ({ color: c, width });

function build(): TableStyleDef[] {
  const list: TableStyleDef[] = [
    {
      id: "none-grid",
      guid: guid(1),
      name: "Stil Yok, Tablo Kılavuzu",
      wholeTbl: { fill: null, text: { scheme: "dk1" }, top: W({ scheme: "dk1" }), bottom: W({ scheme: "dk1" }), left: W({ scheme: "dk1" }), right: W({ scheme: "dk1" }), insideH: W({ scheme: "dk1" }), insideV: W({ scheme: "dk1" }) },
      firstRow: { bold: true },
      lastRow: { bold: true },
      firstCol: { bold: true },
      lastCol: { bold: true },
    },
  ];
  let n = 16;
  for (const a of ACCENTS) {
    const c = (mods: Partial<Color> = {}): Color => ({ scheme: a as SchemeColorName, ...mods });
    list.push({
      id: `light1-${a}`,
      guid: guid(n++),
      name: `Açık Stil 1 - ${accentName(a)}`,
      wholeTbl: { fill: null, text: { scheme: "dk1" }, top: W(c()), bottom: W(c()) },
      band1H: { fill: c({ lumMod: 0.2, lumOff: 0.8 }) },
      band1V: { fill: c({ lumMod: 0.2, lumOff: 0.8 }) },
      firstRow: { bold: true, bottom: W(c()) },
      lastRow: { bold: true, top: W(c()) },
      firstCol: { bold: true },
      lastCol: { bold: true },
    });
    list.push({
      id: `medium2-${a}`,
      guid: guid(n++),
      name: `Orta Stil 2 - ${accentName(a)}`,
      wholeTbl: {
        fill: c({ lumMod: 0.2, lumOff: 0.8 }),
        text: { scheme: "dk1" },
        top: W({ scheme: "lt1" }),
        bottom: W({ scheme: "lt1" }),
        left: W({ scheme: "lt1" }),
        right: W({ scheme: "lt1" }),
        insideH: W({ scheme: "lt1" }),
        insideV: W({ scheme: "lt1" }),
      },
      band1H: { fill: c({ lumMod: 0.4, lumOff: 0.6 }) },
      band1V: { fill: c({ lumMod: 0.4, lumOff: 0.6 }) },
      lastCol: { fill: c(), text: { scheme: "lt1" }, bold: true },
      firstCol: { fill: c(), text: { scheme: "lt1" }, bold: true },
      lastRow: { fill: c(), text: { scheme: "lt1" }, bold: true, top: W({ scheme: "lt1" }, 3) },
      firstRow: { fill: c(), text: { scheme: "lt1" }, bold: true, bottom: W({ scheme: "lt1" }, 3) },
    });
    list.push({
      id: `dark1-${a}`,
      guid: guid(n++),
      name: `Koyu Stil 1 - ${accentName(a)}`,
      wholeTbl: { fill: c({ lumMod: 0.75 }), text: { scheme: "lt1" }, insideH: W({ scheme: "lt1", alpha: 0.4 }, 0.75), insideV: null },
      band1H: { fill: c({ lumMod: 0.6 }) },
      band1V: { fill: c({ lumMod: 0.6 }) },
      lastCol: { fill: c({ lumMod: 0.5 }), bold: true },
      firstCol: { fill: c({ lumMod: 0.5 }), bold: true },
      lastRow: { fill: c({ lumMod: 0.5 }), bold: true, top: W({ scheme: "lt1" }, 2) },
      firstRow: { fill: { scheme: "dk1" }, bold: true, bottom: W({ scheme: "lt1" }, 2) },
    });
  }
  return list;
}

export const TABLE_STYLES: TableStyleDef[] = build();
export const DEFAULT_TABLE_STYLE = "medium2-accent1";

export function tableStyleById(id: string): TableStyleDef {
  return TABLE_STYLES.find((s) => s.id === id) ?? TABLE_STYLES.find((s) => s.id === DEFAULT_TABLE_STYLE)!;
}

export function tableStyleByGuid(g: string): TableStyleDef | undefined {
  return TABLE_STYLES.find((s) => s.guid.toUpperCase() === g.toUpperCase());
}

export interface CellFormat {
  fill: Fill;
  text?: Color;
  bold: boolean;
  borders: { t: CellBorder | null; b: CellBorder | null; l: CellBorder | null; r: CellBorder | null };
}

/** Hücrenin tablo stilinden ve kendi biçiminden çözülmüş görünümü. */
export function cellFormat(t: TableShape, r: number, c: number): CellFormat {
  const st = tableStyleById(t.style.styleId);
  const nRows = t.rows.length;
  const nCols = t.cols.length;
  const o = t.style;
  const parts: PartStyle[] = [st.wholeTbl];
  const bodyR = r - (o.firstRow ? 1 : 0);
  const bodyC = c - (o.firstCol ? 1 : 0);
  if (o.bandCol && st.band1V && bodyC >= 0 && bodyC % 2 === 0) parts.push(st.band1V);
  if (o.bandRow && st.band1H && bodyR >= 0 && bodyR % 2 === 0) parts.push(st.band1H);
  if (o.lastCol && st.lastCol && c === nCols - 1) parts.push(st.lastCol);
  if (o.firstCol && st.firstCol && c === 0) parts.push(st.firstCol);
  if (o.lastRow && st.lastRow && r === nRows - 1) parts.push(st.lastRow);
  if (o.firstRow && st.firstRow && r === 0) parts.push(st.firstRow);
  let fill: Color | null = null;
  let text: Color | undefined;
  let bold = false;
  const borders: CellFormat["borders"] = { t: null, b: null, l: null, r: null };
  for (const p of parts) {
    if (p.fill !== undefined) fill = p.fill;
    if (p.text) text = p.text;
    if (p.bold !== undefined) bold = p.bold;
    const top = r === 0 ? p.top : p.insideH;
    const bottom = r === nRows - 1 ? p.bottom : p.insideH;
    const left = c === 0 ? p.left : p.insideV;
    const right = c === nCols - 1 ? p.right : p.insideV;
    if (top !== undefined) borders.t = top;
    if (bottom !== undefined) borders.b = bottom;
    if (left !== undefined) borders.l = left;
    if (right !== undefined) borders.r = right;
    // Bölüme özgü kenarlar komşu kenarı da etkiler (ör. başlık alt çizgisi).
    if (p !== st.wholeTbl) {
      if (p.bottom !== undefined && p === st.firstRow) borders.b = p.bottom;
      if (p.top !== undefined && p === st.lastRow) borders.t = p.top;
    }
  }
  if (o.firstRow && r === 1 && st.firstRow?.bottom) borders.t = st.firstRow.bottom;
  if (o.lastRow && r === nRows - 2 && st.lastRow?.top) borders.b = st.lastRow.top;
  const cell = t.rows[r]?.cells[c];
  const ff: Fill = cell?.fill ?? (fill ? { type: "solid", color: fill } : { type: "none" });
  if (cell?.borders) {
    for (const k of ["t", "b", "l", "r"] as const) {
      const v = cell.borders[k];
      if (v !== undefined) borders[k] = v;
    }
  }
  return { fill: ff, text, bold, borders };
}
