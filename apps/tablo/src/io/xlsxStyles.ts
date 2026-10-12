/** Hücre stili ⇄ exceljs stil nesnesi dönüşümleri, tema ve dizinli renkler. */
import type { BorderSide, BorderStyle, CellStyle } from "../model/styles";
import { DEFAULT_FONT, DEFAULT_SIZE } from "../model/styles";

/* eslint-disable @typescript-eslint/no-explicit-any */
type XColor = { argb?: string; theme?: number; tint?: number; indexed?: number } | undefined;

/** Varsayılan Office teması (theme1.xml yoksa). Sıra: lt1, dk1, lt2, dk2, accent1..6, hlink, folHlink. */
const DEFAULT_THEME_COLORS = ["FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72"];

const INDEXED = [
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
  "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
  "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
  "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
  "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
  "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
];

export class ColorResolver {
  themeColors: string[] = DEFAULT_THEME_COLORS;

  /** theme1.xml içinden renk şemasını okur. */
  loadTheme(xml: string | undefined): void {
    if (!xml) return;
    const names = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
    const found: Record<string, string> = {};
    for (const n of names) {
      const m = new RegExp(`<a:${n}>\\s*<a:(?:srgbClr val="([0-9A-Fa-f]{6})"|sysClr[^>]*lastClr="([0-9A-Fa-f]{6})")`).exec(xml);
      if (m) found[n] = (m[1] ?? m[2]).toUpperCase();
    }
    if (Object.keys(found).length >= 10) {
      // SpreadsheetML tema dizinleri: 0=lt1, 1=dk1, 2=lt2, 3=dk2
      this.themeColors = [found.lt1, found.dk1, found.lt2, found.dk2, found.accent1, found.accent2, found.accent3, found.accent4, found.accent5, found.accent6, found.hlink ?? "0563C1", found.folHlink ?? "954F72"].map((c, i) => c ?? DEFAULT_THEME_COLORS[i]);
    }
  }

  resolve(c: XColor): string | undefined {
    if (!c) return undefined;
    let hex: string | undefined;
    if (c.argb) {
      const a = c.argb.length === 8 ? c.argb.slice(2) : c.argb;
      hex = a.toUpperCase();
    } else if (c.theme !== undefined) hex = this.themeColors[c.theme] ?? "000000";
    else if (c.indexed !== undefined) {
      if (c.indexed === 64) return undefined;
      hex = INDEXED[c.indexed] ?? "000000";
    }
    if (!hex || !/^[0-9A-F]{6}$/.test(hex)) return undefined;
    if (c.tint) hex = applyTint(hex, c.tint);
    return "#" + hex;
  }
}

function applyTint(hex: string, tint: number): string {
  const rgb = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
  const out = rgb.map((v) => Math.round(tint < 0 ? v * (1 + tint) : v + (255 - v) * tint));
  return out.map((v) => Math.max(0, Math.min(255, v)).toString(16).padStart(2, "0")).join("").toUpperCase();
}

const BUILTIN_FORMAT_FIX: Record<string, string> = {
  "mm-dd-yy": "dd.mm.yyyy",
  "m/d/yy": "dd.mm.yyyy",
  "d-mmm-yy": "d.mmm.yy",
  "d-mmm": "d.mmm",
  "mmm-yy": "mmm.yy",
  "m/d/yy h:mm": "dd.mm.yyyy hh:mm",
  "h:mm AM/PM": "h:mm AM/PM",
};

function border(b: any, cr: ColorResolver): BorderSide | undefined {
  if (!b || !b.style) return undefined;
  const map: Record<string, BorderStyle> = { mediumDashDotDot: "mediumDashDot" };
  return { style: (map[b.style] ?? b.style) as BorderStyle, color: cr.resolve(b.color) ?? "#000000" };
}

/** exceljs stil nesnesi → CellStyle. */
export function fromXStyle(s: any, cr: ColorResolver): CellStyle {
  const out: CellStyle = {};
  if (!s) return out;
  const f = s.font;
  if (f) {
    if (f.name && f.name !== DEFAULT_FONT) out.font = f.name;
    if (f.size && f.size !== DEFAULT_SIZE) out.size = f.size;
    if (f.bold) out.bold = true;
    if (f.italic) out.italic = true;
    if (f.underline) out.underline = f.underline === "double" || f.underline === "doubleAccounting" ? "double" : "single";
    if (f.strike) out.strike = true;
    const col = cr.resolve(f.color);
    if (col && col !== "#000000") out.color = col;
  }
  const fill = s.fill;
  if (fill) {
    if (fill.type === "pattern" && fill.pattern && fill.pattern !== "none") {
      if (fill.pattern === "solid") {
        const c = cr.resolve(fill.fgColor) ?? cr.resolve(fill.bgColor);
        if (c) out.fill = c;
      } else {
        const fg = cr.resolve(fill.fgColor) ?? "#000000";
        const bg = cr.resolve(fill.bgColor);
        if (bg) out.fill = bg;
        out.pattern = { type: fill.pattern, color: fg };
      }
    } else if (fill.type === "gradient" && fill.stops?.length) {
      const c = cr.resolve(fill.stops[0].color);
      if (c) out.fill = c;
    }
  }
  const b = s.border;
  if (b) {
    out.bTop = border(b.top, cr);
    out.bBottom = border(b.bottom, cr);
    out.bLeft = border(b.left, cr);
    out.bRight = border(b.right, cr);
  }
  const a = s.alignment;
  if (a) {
    if (a.horizontal && a.horizontal !== "general") out.hAlign = a.horizontal;
    if (a.vertical && a.vertical !== "bottom") out.vAlign = a.vertical;
    if (a.wrapText) out.wrap = true;
    if (a.shrinkToFit) out.shrink = true;
    if (a.indent) out.indent = a.indent;
    if (a.textRotation) out.rotation = a.textRotation === "vertical" ? 255 : Number(a.textRotation);
  }
  if (s.numFmt && s.numFmt !== "General") out.numFmt = BUILTIN_FORMAT_FIX[s.numFmt] ?? s.numFmt;
  if (s.protection) {
    if (s.protection.locked === false) out.locked = false;
    if (s.protection.hidden) out.hidden = true;
  }
  for (const k of Object.keys(out) as (keyof CellStyle)[]) if (out[k] === undefined) delete out[k];
  return out;
}

function argb(hex: string): { argb: string } {
  return { argb: "FF" + hex.replace("#", "").toUpperCase() };
}

/** CellStyle → exceljs stil nesnesi. */
export function toXStyle(s: CellStyle): any {
  const out: any = {};
  out.font = { name: s.font ?? DEFAULT_FONT, size: s.size ?? DEFAULT_SIZE, family: 2 };
  if (!s.font) out.font.scheme = "minor";
  if (s.bold) out.font.bold = true;
  if (s.italic) out.font.italic = true;
  if (s.underline) out.font.underline = s.underline === "double" ? "double" : true;
  if (s.strike) out.font.strike = true;
  if (s.color) out.font.color = argb(s.color);
  if (s.pattern) out.fill = { type: "pattern", pattern: s.pattern.type, fgColor: argb(s.pattern.color), ...(s.fill ? { bgColor: argb(s.fill) } : {}) };
  else if (s.fill) out.fill = { type: "pattern", pattern: "solid", fgColor: argb(s.fill), bgColor: { indexed: 64 } };
  const bd: any = {};
  const side = (b?: BorderSide) => (b ? { style: b.style, color: argb(b.color ?? "#000000") } : undefined);
  if (s.bTop) bd.top = side(s.bTop);
  if (s.bBottom) bd.bottom = side(s.bBottom);
  if (s.bLeft) bd.left = side(s.bLeft);
  if (s.bRight) bd.right = side(s.bRight);
  if (Object.keys(bd).length) out.border = bd;
  const al: any = {};
  if (s.hAlign && s.hAlign !== "general") al.horizontal = s.hAlign;
  if (s.vAlign) al.vertical = s.vAlign;
  if (s.wrap) al.wrapText = true;
  if (s.shrink) al.shrinkToFit = true;
  if (s.indent) al.indent = s.indent;
  if (s.rotation) al.textRotation = s.rotation === 255 ? "vertical" : s.rotation;
  if (Object.keys(al).length) out.alignment = al;
  if (s.numFmt) out.numFmt = s.numFmt;
  if (s.locked === false || s.hidden) out.protection = { locked: s.locked !== false, hidden: !!s.hidden };
  return out;
}

export function isDefaultXStyle(s: CellStyle): boolean {
  return Object.keys(s).length === 0;
}
