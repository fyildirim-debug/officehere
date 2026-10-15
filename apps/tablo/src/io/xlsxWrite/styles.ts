/**
 * styles.xml: hücre biçimleri (cellXfs) modelin intern stil tablosundan doğrudan, stil numarası
 * başına bir kez üretilir (hücre başına nesne yok). Yazı tipi, dolgu, kenarlık ve sayı biçimi
 * listeleri XML'leriyle tekilleştirilir. Model stili → exceljs stil nesnesi dönüşümü eski yazıcıyla
 * aynıdır (toXStyle); koşullu biçimlerin farklı biçimleri (dxf) ve onay kutulu biçim kopyaları da buradadır.
 */
import type { WorkbookModel } from "../../model/workbook";
import { toXStyle } from "../xlsxStyles";
import { attrs, esc, XML_DECL } from "./xml";

/* eslint-disable @typescript-eslint/no-explicit-any */

/** Yerleşik sayı biçimleri (kimliği sabit; styles.xml'e yazılmaz). */
const BUILTIN_FMT = new Map<string, number>([
  ["General", 0], ["0", 1], ["0.00", 2], ["#,##0", 3], ["#,##0.00", 4], ["0%", 9], ["0.00%", 10], ["0.00E+00", 11], ["# ?/?", 12],
  ["# ??/??", 13], ["mm-dd-yy", 14], ["d-mmm-yy", 15], ["d-mmm", 16], ["mmm-yy", 17], ["h:mm AM/PM", 18], ["h:mm:ss AM/PM", 19],
  ["h:mm", 20], ["h:mm:ss", 21], ['m/d/yy "h":mm', 22], ["#,##0 ;(#,##0)", 37], ["#,##0 ;[Red](#,##0)", 38], ["#,##0.00 ;(#,##0.00)", 39],
  ["#,##0.00 ;[Red](#,##0.00)", 40], ["mm:ss", 45], ["[h]:mm:ss", 46], ["mmss.0", 47], ["##0.0E+0", 48], ["@", 49],
]);
const CUSTOM_FMT_BASE = 164;

const H_ALIGN = new Set(["left", "center", "right", "fill", "centerContinuous", "distributed", "justify"]);
const V_ALIGN = new Set(["top", "middle", "bottom", "distributed", "justify"]);
const UNDERLINE: Record<string, string> = { single: "<u/>", double: '<u val="double"/>', singleAccounting: '<u val="singleAccounting"/>', doubleAccounting: '<u val="doubleAccounting"/>' };

const FPB_NS = "http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag";
const XF_EXT = `<extLst><ext uri="{C7286773-470A-42A8-94C5-96B5CB345126}" xmlns:xfpb="${FPB_NS}"><xfpb:xfComplement i="0"/></ext></extLst>`;

const validInt = (v: unknown): number => {
  const i = parseInt(String(v), 10);
  return Number.isNaN(i) ? 0 : i;
};

function colorXml(tag: string, c: any): string {
  if (!c) return "";
  if (c.argb) return `<${tag} rgb="${esc(String(c.argb))}"/>`;
  if (c.theme !== undefined) return `<${tag}${attrs([["theme", c.theme], ["tint", c.tint]])}/>`;
  if (c.indexed !== undefined) return `<${tag} indexed="${esc(String(c.indexed))}"/>`;
  return `<${tag} auto="1"/>`;
}

function fontXml(f: any): string {
  let s = "<font>";
  if (f.bold) s += "<b/>";
  if (f.italic) s += "<i/>";
  if (f.underline === true) s += "<u/>";
  else if (f.underline && UNDERLINE[f.underline]) s += UNDERLINE[f.underline];
  if (f.charset) s += `<charset val="${esc(String(f.charset))}"/>`;
  s += colorXml("color", f.color);
  if (f.family) s += `<family val="${esc(String(f.family))}"/>`;
  if (f.vertAlign !== undefined) s += `<vertAlign val="${esc(String(f.vertAlign))}"/>`;
  if (f.scheme !== undefined) s += `<scheme val="${esc(String(f.scheme))}"/>`;
  if (f.strike) s += "<strike/>";
  if (f.size) s += `<sz val="${esc(String(f.size))}"/>`;
  if (f.name !== undefined) s += `<name val="${esc(String(f.name))}"/>`;
  return s + "</font>";
}

function fillXml(f: any): string {
  if (f.type !== "pattern") return "";
  const inner = colorXml("fgColor", f.fgColor) + colorXml("bgColor", f.bgColor);
  const open = `<patternFill patternType="${esc(String(f.pattern))}"`;
  return `<fill>${inner ? `${open}>${inner}</patternFill>` : `${open}/>`}</fill>`;
}

function edgeXml(tag: string, e: any): string {
  return e?.style ? `<${tag} style="${esc(String(e.style))}">${colorXml("color", e.color)}</${tag}>` : `<${tag}/>`;
}

function borderXml(b: any): string {
  return `<border>${edgeXml("left", b.left)}${edgeXml("right", b.right)}${edgeXml("top", b.top)}${edgeXml("bottom", b.bottom)}${edgeXml("diagonal", b.diagonal)}</border>`;
}

function textRotation(v: unknown): number | undefined {
  let r: number | "vertical" | undefined;
  if (v === "vertical") r = "vertical";
  else {
    const i = validInt(v);
    r = i >= -90 && i <= 90 ? i : undefined;
  }
  if (!r) return undefined;
  if (r === "vertical") return 255;
  const t = Math.round(r);
  if (t >= 0 && t <= 90) return t;
  if (t < 0 && t >= -90) return 90 - t;
  return undefined;
}

function alignmentXml(a: any): string {
  const list: [string, string | number | undefined][] = [
    ["horizontal", H_ALIGN.has(a.horizontal) ? a.horizontal : undefined],
    ["vertical", a.vertical === "middle" ? "center" : V_ALIGN.has(a.vertical) ? a.vertical : undefined],
    ["wrapText", a.wrapText ? "1" : undefined],
    ["shrinkToFit", a.shrinkToFit ? "1" : undefined],
    ["indent", Math.max(0, validInt(a.indent)) || undefined],
    ["textRotation", textRotation(a.textRotation)],
  ];
  const s = attrs(list);
  return s ? `<alignment${s}/>` : "";
}

function protectionXml(p: any): string {
  const s = attrs([
    ["locked", (p.locked === undefined ? true : p.locked) ? undefined : "0"],
    ["hidden", (p.hidden === undefined ? false : p.hidden) ? "1" : undefined],
  ]);
  return s ? `<protection${s}/>` : "";
}

/** Bir listeye XML'iyle tekil öğe ekler; dizinini döndürür. */
class Pool {
  readonly items: string[] = [];
  private index = new Map<string, number>();
  add(xml: string): number {
    let i = this.index.get(xml);
    if (i === undefined) {
      i = this.items.length;
      this.items.push(xml);
      this.index.set(xml, i);
    }
    return i;
  }
}

export class StyleSheet {
  private fonts = new Pool();
  private fills = new Pool();
  private borders = new Pool();
  private xfs = new Pool();
  private numFmts: string[] = [];
  private numFmtIds = new Map<string, number>();
  private dxfs: string[] = [];
  /** Model stil numarası → xf (−1: henüz yok). */
  private byStyle: Int32Array;
  private merged = new Map<string, number>();
  private boxedXf = new Map<number, number>();
  private xStyles = new Map<number, any>();

  constructor(private readonly wb: WorkbookModel) {
    this.byStyle = new Int32Array(Math.max(1, wb.styles.size)).fill(-1);
    this.borders.add(borderXml({}));
    this.xfs.add('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>');
    this.fills.add(fillXml({ type: "pattern", pattern: "none" }));
    this.fills.add(fillXml({ type: "pattern", pattern: "gray125" }));
    this.fonts.add(fontXml({ size: 11, color: { theme: 1 }, name: "Calibri", family: 2, scheme: "minor" }));
  }

  /** Model stilinin exceljs biçimli nesnesi (stil başına bir kez). */
  xStyle(id: number): any {
    let x = this.xStyles.get(id);
    if (!x) this.xStyles.set(id, (x = toXStyle(this.wb.styles.get(id))));
    return x;
  }

  private numFmt(code: string): number {
    const b = BUILTIN_FMT.get(code);
    if (b !== undefined) return b;
    let id = this.numFmtIds.get(code);
    if (id === undefined) {
      id = CUSTOM_FMT_BASE + this.numFmts.length;
      this.numFmtIds.set(code, id);
      this.numFmts.push(`<numFmt numFmtId="${id}" formatCode="${esc(code)}"/>`);
    }
    return id;
  }

  /** exceljs biçimli stil nesnesinin hücre biçimi (xf) dizini. */
  xfOf(x: any): number {
    const numFmtId = x.numFmt ? this.numFmt(String(x.numFmt)) : 0;
    const fontId = x.font ? this.fonts.add(fontXml(x.font)) : 0;
    const borderId = x.border ? this.borders.add(borderXml(x.border)) : 0;
    const fillId = x.fill ? this.fills.add(fillXml(x.fill)) : 0;
    let open = `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"`;
    if (numFmtId) open += ' applyNumberFormat="1"';
    if (fontId) open += ' applyFont="1"';
    if (fillId) open += ' applyFill="1"';
    if (borderId) open += ' applyBorder="1"';
    if (x.alignment) open += ' applyAlignment="1"';
    if (x.protection) open += ' applyProtection="1"';
    const inner = (x.alignment ? alignmentXml(x.alignment) : "") + (x.protection ? protectionXml(x.protection) : "");
    return this.xfs.add(inner ? `${open}>${inner}</xf>` : `${open}/>`);
  }

  /** Model stil numarası → xf (hücre, satır ya da sütun kendi stili olarak taşıyorsa; 0 da dahil). */
  xf(id: number): number {
    if (id >= this.byStyle.length) {
      const next = new Int32Array(Math.max(id + 1, this.byStyle.length * 2)).fill(-1);
      next.set(this.byStyle);
      this.byStyle = next;
    }
    let v = this.byStyle[id];
    if (v < 0) this.byStyle[id] = v = this.xfOf(this.xStyle(id));
    return v;
  }

  /**
   * Kendi stili olmadan oluşan hücrenin biçimi: satır ve sütun stillerinin grup grup birleşimi
   * (sayı biçimi, yazı tipi, hizalama, kenarlık, dolgu, koruma: önce satırınki).
   */
  mergedXf(rowId: number | undefined, colId: number | undefined): number {
    if (!rowId && !colId) return 0;
    if (!rowId) return this.xf(colId!);
    if (!colId) return this.xf(rowId);
    const key = `${rowId}|${colId}`;
    let v = this.merged.get(key);
    if (v === undefined) {
      const r = this.xStyle(rowId);
      const c = this.xStyle(colId);
      const m: any = {};
      for (const k of ["numFmt", "font", "alignment", "border", "fill", "protection"]) {
        const val = r[k] || c[k];
        if (val) m[k] = val;
      }
      this.merged.set(key, (v = this.xfOf(m)));
    }
    return v;
  }

  /** Onay kutulu kopya (xfComplement uzantılı) biçim. */
  boxed(xf: number): number {
    let v = this.boxedXf.get(xf);
    if (v === undefined) {
      const base = this.xfs.items[xf] ?? this.xfs.items[0];
      const withExt = /\/>\s*$/.test(base) ? base.replace(/\s*\/>\s*$/, `>${XF_EXT}</xf>`) : base.replace(/<\/xf>\s*$/, `${XF_EXT}</xf>`);
      // Aynı XML'li biçim tekilleştirilmez: her kopya ayrı xf (eski yazıcı da ekliyordu)
      v = this.xfs.items.length;
      this.xfs.items.push(withExt);
      this.boxedXf.set(xf, v);
    }
    return v;
  }

  get hasBoxed(): boolean {
    return this.boxedXf.size > 0;
  }

  /** Koşullu biçimin farklı biçimi (dxf) dizini. */
  dxf(x: any): number {
    let s = "<dxf>";
    if (x.font) s += fontXml(x.font);
    if (x.numFmt) {
      const id = this.numFmt(String(x.numFmt));
      if (id) s += `<numFmt numFmtId="${id}" formatCode="${esc(String(x.numFmt))}"/>`;
    }
    if (x.fill) s += fillXml(x.fill);
    if (x.alignment) s += alignmentXml(x.alignment);
    // CT_Dxf sırası: font, numFmt, fill, alignment, border, protection
    if (x.border) s += borderXml(x.border);
    if (x.protection) s += protectionXml(x.protection);
    this.dxfs.push(s + "</dxf>");
    return this.dxfs.length - 1;
  }

  xml(): string {
    const list = (tag: string, items: string[], extra = "") => `<${tag} count="${items.length}"${extra}>${items.join("")}</${tag}>`;
    return (
      XML_DECL +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main">' +
      (this.numFmts.length ? list("numFmts", this.numFmts) : "") +
      list("fonts", this.fonts.items, ' x14ac:knownFonts="1"') +
      list("fills", this.fills.items) +
      list("borders", this.borders.items) +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      list("cellXfs", this.xfs.items) +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      (this.dxfs.length ? list("dxfs", this.dxfs) : '<dxfs count="0"/>') +
      '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>' +
      '<extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext><ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/></ext></extLst>' +
      "</styleSheet>"
    );
  }
}
