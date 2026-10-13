/**
 * BIFF biçim kayıtları: FONT, FORMAT, XF ve PALETTE (MS-XLS §2.4.122, §2.4.126,
 * §2.4.353, §2.4.188) → CellStyle. Yerleşik sayı biçimleri 0–49 Türkçe karşılıklarıyla.
 */
import type { BorderSide, BorderStyle, CellStyle, HAlign, VAlign } from "../../model/styles";
import { ACCOUNTING_FORMAT, COMMA_FORMAT } from "../../model/styles";
import { INDEXED } from "../xlsxStyles";
import { readByteString, readString, u16, u32, u8 } from "./records";

/** Yerleşik biçim numaraları (FORMAT kaydı yoksa). */
export const BUILTIN_FORMATS: Record<number, string> = {
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  5: '#,##0 "₺";-#,##0 "₺"',
  6: '#,##0 "₺";[Red]-#,##0 "₺"',
  7: '#,##0.00 "₺";-#,##0.00 "₺"',
  8: '#,##0.00 "₺";[Red]-#,##0.00 "₺"',
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  12: "# ?/?",
  13: "# ??/??",
  14: "dd.mm.yyyy",
  15: "d.mmm.yy",
  16: "d.mmm",
  17: "mmm.yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "hh:mm",
  21: "hh:mm:ss",
  22: "dd.mm.yyyy hh:mm",
  37: "#,##0 ;(#,##0)",
  38: "#,##0 ;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)",
  40: "#,##0.00;[Red](#,##0.00)",
  41: COMMA_FORMAT.replace(/\.00/g, ""),
  42: ACCOUNTING_FORMAT.replace(/\.00/g, ""),
  43: COMMA_FORMAT,
  44: ACCOUNTING_FORMAT,
  45: "mm:ss",
  46: "[h]:mm:ss",
  47: "mm:ss.0",
  48: "##0.0E+0",
  49: "@",
};

/** ABD yerel biçimli kısa tarih kodları Türkçe karşılığına. */
const DATE_FIX: Record<string, string> = { "m/d/yy": "dd.mm.yyyy", "m/d/yyyy": "dd.mm.yyyy", "mm-dd-yy": "dd.mm.yyyy", "m/d/yy h:mm": "dd.mm.yyyy hh:mm" };

const PATTERNS: Record<number, string> = {
  2: "mediumGray", 3: "darkGray", 4: "lightGray", 5: "darkHorizontal", 6: "darkVertical", 7: "darkDown", 8: "darkUp", 9: "darkGrid",
  10: "darkTrellis", 11: "lightHorizontal", 12: "lightVertical", 13: "lightDown", 14: "lightUp", 15: "lightGrid", 16: "lightTrellis",
  17: "gray125", 18: "gray0625",
};

const BORDERS: Record<number, BorderStyle> = {
  1: "thin", 2: "medium", 3: "dashed", 4: "dotted", 5: "thick", 6: "double", 7: "hair", 8: "mediumDashed", 9: "dashDot",
  10: "mediumDashDot", 11: "dashDotDot", 12: "mediumDashDot", 13: "slantDashDot",
};

const H_ALIGN: (HAlign | undefined)[] = [undefined, "left", "center", "right", "fill", "justify", "centerContinuous", "distributed"];
const V_ALIGN: (VAlign | undefined)[] = ["top", "middle", undefined, "justify", "distributed"];

interface FontRec {
  name: string;
  size: number;
  bold: boolean;
  italic: boolean;
  strike: boolean;
  underline?: "single" | "double";
  color: number;
}

export class XlsStyles {
  fonts: FontRec[] = [];
  formats = new Map<number, string>();
  palette: string[] = INDEXED.map((c) => "#" + c);
  private xfs: Uint8Array[] = [];

  constructor(private biff: 5 | 8, private decode: (b: Uint8Array) => string) {}

  font(d: Uint8Array): void {
    const name = this.biff === 8 ? readString(d, 14, 1).text : readByteString(d, 14, 1, this.decode).text;
    const grbit = u16(d, 2);
    const uls = u8(d, 10);
    this.fonts.push({
      name,
      size: u16(d, 0) / 20,
      bold: u16(d, 6) >= 600,
      italic: !!(grbit & 0x02),
      strike: !!(grbit & 0x08),
      underline: uls === 1 || uls === 0x21 ? "single" : uls === 2 || uls === 0x22 ? "double" : undefined,
      color: u16(d, 4),
    });
  }

  format(d: Uint8Array): void {
    const ifmt = u16(d, 0);
    const code = this.biff === 8 ? readString(d, 2, 2).text : readByteString(d, 2, 1, this.decode).text;
    this.formats.set(ifmt, code);
  }

  xf(d: Uint8Array): void {
    this.xfs.push(d);
  }

  setPalette(d: Uint8Array): void {
    const n = Math.min(56, u16(d, 0));
    for (let i = 0; i < n; i++) {
      const o = 2 + i * 4;
      this.palette[8 + i] = "#" + [d[o], d[o + 1], d[o + 2]].map((x) => (x ?? 0).toString(16).padStart(2, "0")).join("").toUpperCase();
    }
  }

  color(icv: number): string | undefined {
    if (icv >= 64) return undefined;
    return this.palette[icv];
  }

  /** Font dizini: 4 numaralı font yoktur (BIFF geleneği), sonrakiler bir kayar. */
  private fontAt(ifnt: number): FontRec | undefined {
    return this.fonts[ifnt >= 4 ? ifnt - 1 : ifnt];
  }

  numFmt(ifmt: number): string | undefined {
    const code = this.formats.get(ifmt) ?? BUILTIN_FORMATS[ifmt];
    if (!code || /^general$/i.test(code)) return undefined;
    return DATE_FIX[code.toLowerCase()] ?? code;
  }

  /** Varsayılan yazı tipi (font 0). */
  defaultFont(): FontRec | undefined {
    return this.fonts[0];
  }

  /** XF dizinindeki hücre stili. */
  style(ixfe: number): CellStyle {
    const d = this.xfs[ixfe];
    const st: CellStyle = {};
    if (!d) return st;
    const f = this.fontAt(u16(d, 0));
    if (f) {
      st.font = f.name;
      st.size = f.size;
      if (f.bold) st.bold = true;
      if (f.italic) st.italic = true;
      if (f.strike) st.strike = true;
      if (f.underline) st.underline = f.underline;
      const c = f.color === 0x7fff ? undefined : this.color(f.color);
      if (c && c !== "#000000") st.color = c;
    }
    const nf = this.numFmt(u16(d, 2));
    if (nf) st.numFmt = nf;
    const flags = u16(d, 4);
    if (!(flags & 0x01)) st.locked = false;
    if (flags & 0x02) st.hidden = true;
    const al = u8(d, 6);
    const h = H_ALIGN[al & 0x07];
    if (h) st.hAlign = h;
    if (al & 0x08) st.wrap = true;
    const v = V_ALIGN[(al >> 4) & 0x07];
    if (v) st.vAlign = v;
    if (this.biff === 8) {
      const trot = u8(d, 7);
      if (trot === 255) st.rotation = 255;
      else if (trot > 0 && trot <= 90) st.rotation = trot;
      else if (trot > 90 && trot <= 180) st.rotation = -(trot - 90);
      const b8 = u8(d, 8);
      if (b8 & 0x0f) st.indent = b8 & 0x0f;
      if (b8 & 0x10) st.shrink = true;
      const bd1 = u32(d, 10);
      const bd2 = u32(d, 14);
      const side = (dg: number, icv: number): BorderSide | undefined => (BORDERS[dg] ? { style: BORDERS[dg], color: this.color(icv) ?? "#000000" } : undefined);
      st.bLeft = side(bd1 & 0x0f, (bd1 >>> 16) & 0x7f);
      st.bRight = side((bd1 >>> 4) & 0x0f, (bd1 >>> 23) & 0x7f);
      st.bTop = side((bd1 >>> 8) & 0x0f, bd2 & 0x7f);
      st.bBottom = side((bd1 >>> 12) & 0x0f, (bd2 >>> 7) & 0x7f);
      const fls = (bd2 >>> 26) & 0x3f;
      const fill = u16(d, 18);
      this.applyFill(st, fls, fill & 0x7f, (fill >> 7) & 0x7f);
    } else {
      const x = u32(d, 8);
      this.applyFill(st, (x >>> 16) & 0x3f, x & 0x7f, (x >>> 7) & 0x7f);
    }
    for (const k of Object.keys(st) as (keyof CellStyle)[]) if (st[k] === undefined) delete st[k];
    return st;
  }

  private applyFill(st: CellStyle, fls: number, fore: number, back: number): void {
    if (fls === 1) {
      const c = this.color(fore);
      if (c) st.fill = c;
    } else if (PATTERNS[fls]) {
      const bg = this.color(back);
      if (bg) st.fill = bg;
      st.pattern = { type: PATTERNS[fls], color: this.color(fore) ?? "#000000" };
    }
  }
}
