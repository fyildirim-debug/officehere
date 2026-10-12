// DOCX okuma: styles.xml, numbering.xml ve tema yazı tipleri; özellik ayrıştırma.
import { attr, attrNum, kid, kids, onOff, path, HIGHLIGHT_NAMES } from "./xml";
import type { BorderSpec } from "../../model/styles";
import { sigFromLevel, type ListSig } from "./numbering";

export interface RunProps {
  bold?: boolean;
  italic?: boolean;
  underline?: { type: string; color: string | null } | null;
  strike?: boolean;
  dstrike?: boolean;
  color?: string | null;
  highlight?: string | null;
  shd?: string | null;
  size?: number;
  font?: string;
  vertAlign?: "superscript" | "subscript" | null;
  caps?: boolean;
  smallCaps?: boolean;
  vanish?: boolean;
  outline?: boolean;
  shadow?: boolean;
  emboss?: boolean;
  imprint?: boolean;
  spacing?: number;
}

export interface ParaProps {
  spaceBefore?: number;
  spaceAfter?: number;
  lineRule?: "auto" | "exact" | "atLeast";
  lineHeight?: number;
  align?: string;
  indentLeft?: number;
  indentRight?: number;
  indentFirst?: number;
  keepNext?: boolean;
  keepLines?: boolean;
  suppressLineNumbers?: boolean;
  pageBreakBefore?: boolean;
  shading?: string | null;
  borders?: { top?: BorderSpec; bottom?: BorderSpec; left?: BorderSpec; right?: BorderSpec } | null;
  numId?: string;
  ilvl?: number;
  /** Sekme durakları (pt) ve anahat düzeyi (1-9). */
  tabs?: { pos: number; align: "left" | "center" | "right" | "decimal" | "bar"; leader: "none" | "dot" | "hyphen" | "underscore" }[];
  outlineLevel?: number;
}

export interface StyleInfo {
  id: string;
  type: string;
  name: string;
  basedOn: string | null;
  ppr: ParaProps;
  rpr: RunProps;
  tblBordersNone?: boolean;
}

export interface ThemeFonts {
  major: string;
  minor: string;
}

const UNDERLINE_MAP: Record<string, string> = {
  single: "single",
  words: "single",
  double: "double",
  thick: "thick",
  dotted: "dotted",
  dottedHeavy: "dotted",
  dash: "dashed",
  dashedHeavy: "dashed",
  dashLong: "dashed",
  dotDash: "dashed",
  dotDotDash: "dashed",
  wave: "wave",
  wavyHeavy: "wave",
  wavyDouble: "wave",
};

export function parseColor(v: string | null): string | null {
  if (!v || v === "auto") return null;
  return /^[0-9a-f]{6}$/i.test(v) ? `#${v.toUpperCase()}` : null;
}

export function parseRPr(rPr: Element | null, theme: ThemeFonts): RunProps {
  const p: RunProps = {};
  if (!rPr) return p;
  for (const c of kids(rPr)) {
    switch (c.localName) {
      case "b":
        p.bold = onOff(c) ?? undefined;
        break;
      case "i":
        p.italic = onOff(c) ?? undefined;
        break;
      case "strike":
        p.strike = onOff(c) ?? undefined;
        break;
      case "dstrike":
        p.dstrike = onOff(c) ?? undefined;
        break;
      case "outline":
      case "shadow":
      case "emboss":
      case "imprint":
        p[c.localName] = onOff(c) ?? undefined;
        break;
      case "caps":
        p.caps = onOff(c) ?? undefined;
        break;
      case "smallCaps":
        p.smallCaps = onOff(c) ?? undefined;
        break;
      case "vanish":
        p.vanish = onOff(c) ?? undefined;
        break;
      case "u": {
        const v = attr(c, "val") ?? "single";
        p.underline = v === "none" ? null : { type: UNDERLINE_MAP[v] ?? "single", color: parseColor(attr(c, "color")) };
        break;
      }
      case "color":
        p.color = parseColor(attr(c, "val"));
        break;
      case "highlight": {
        const v = attr(c, "val");
        p.highlight = v && v !== "none" ? (HIGHLIGHT_NAMES[v] ?? null) : null;
        break;
      }
      case "shd": {
        const fill = attr(c, "fill");
        p.shd = parseColor(fill);
        break;
      }
      case "sz": {
        const n = attrNum(c, "val");
        if (n) p.size = n / 2;
        break;
      }
      case "rFonts": {
        const f = attr(c, "ascii") ?? attr(c, "hAnsi");
        if (f) p.font = f;
        else {
          const t = attr(c, "asciiTheme") ?? attr(c, "hAnsiTheme");
          if (t) p.font = t.startsWith("major") ? theme.major : theme.minor;
        }
        break;
      }
      case "vertAlign": {
        const v = attr(c, "val");
        p.vertAlign = v === "superscript" || v === "subscript" ? v : null;
        break;
      }
      case "spacing": {
        const n = attrNum(c, "val");
        if (n !== null) p.spacing = n / 20;
        break;
      }
    }
  }
  return p;
}

function parseBorder(el: Element | null): BorderSpec | undefined {
  if (!el) return undefined;
  const val = attr(el, "val") ?? "single";
  if (val === "none" || val === "nil") return undefined;
  const style: BorderSpec["style"] = val === "double" ? "double" : val === "dotted" ? "dotted" : val.startsWith("dash") ? "dashed" : val === "thick" || val.startsWith("thickThin") ? "thick" : "single";
  return { style, size: (attrNum(el, "sz") ?? 4) / 8, color: parseColor(attr(el, "color")) ?? "#000000", space: attrNum(el, "space") ?? 1 };
}

export function parsePPr(pPr: Element | null): ParaProps {
  const p: ParaProps = {};
  if (!pPr) return p;
  for (const c of kids(pPr)) {
    switch (c.localName) {
      case "keepNext":
        p.keepNext = onOff(c) ?? undefined;
        break;
      case "keepLines":
        p.keepLines = onOff(c) ?? undefined;
        break;
      case "pageBreakBefore":
        p.pageBreakBefore = onOff(c) ?? undefined;
        break;
      case "suppressLineNumbers":
        p.suppressLineNumbers = onOff(c) ?? undefined;
        break;
      case "spacing": {
        const b = attrNum(c, "before");
        const a = attrNum(c, "after");
        const line = attrNum(c, "line");
        if (b !== null && attr(c, "beforeAutospacing") !== "1") p.spaceBefore = b / 20;
        if (a !== null && attr(c, "afterAutospacing") !== "1") p.spaceAfter = a / 20;
        if (line !== null) {
          const rule = attr(c, "lineRule") ?? "auto";
          if (rule === "exact" || rule === "atLeast") {
            p.lineRule = rule;
            p.lineHeight = line / 20;
          } else {
            p.lineRule = "auto";
            p.lineHeight = Math.round((line / 240) * 100) / 100;
          }
        }
        break;
      }
      case "ind": {
        const l = attrNum(c, "left") ?? attrNum(c, "start");
        const r = attrNum(c, "right") ?? attrNum(c, "end");
        const f = attrNum(c, "firstLine");
        const h = attrNum(c, "hanging");
        if (l !== null) p.indentLeft = l / 20;
        if (r !== null) p.indentRight = r / 20;
        if (h !== null && h !== 0) p.indentFirst = -h / 20;
        else if (f !== null) p.indentFirst = f / 20;
        break;
      }
      case "jc": {
        const v = attr(c, "val");
        p.align = v === "both" || v === "distribute" ? "justify" : v === "end" ? "right" : v === "start" ? "left" : (v ?? undefined);
        break;
      }
      case "shd":
        p.shading = parseColor(attr(c, "fill"));
        break;
      case "pBdr": {
        const borders = { top: parseBorder(kid(c, "top")), bottom: parseBorder(kid(c, "bottom")), left: parseBorder(kid(c, "left") ?? kid(c, "start")), right: parseBorder(kid(c, "right") ?? kid(c, "end")) };
        p.borders = borders.top || borders.bottom || borders.left || borders.right ? borders : null;
        break;
      }
      case "tabs": {
        const out: NonNullable<ParaProps["tabs"]> = [];
        for (const t of kids(c, "tab")) {
          const val = attr(t, "val") ?? "left";
          const pos = attrNum(t, "pos");
          if (pos === null || val === "clear" || val === "num") continue;
          const align = val === "center" ? "center" : val === "right" || val === "end" ? "right" : val === "decimal" ? "decimal" : val === "bar" ? "bar" : "left";
          const ld = attr(t, "leader");
          const leader = ld === "dot" || ld === "middleDot" ? "dot" : ld === "hyphen" ? "hyphen" : ld === "underscore" || ld === "heavy" ? "underscore" : "none";
          out.push({ pos: pos / 20, align, leader });
        }
        if (out.length) p.tabs = out;
        break;
      }
      case "outlineLvl": {
        const v = attrNum(c, "val");
        if (v !== null && v >= 0 && v < 9) p.outlineLevel = v + 1;
        break;
      }
      case "numPr": {
        const id = attr(kid(c, "numId"), "val");
        if (id !== null) p.numId = id;
        const l = attrNum(kid(c, "ilvl"), "val");
        if (l !== null) p.ilvl = l;
        break;
      }
    }
  }
  return p;
}

export function parseTheme(doc: Document | null): ThemeFonts {
  const out = { major: "Aptos Display", minor: "Aptos" };
  if (!doc) return out;
  const all = doc.getElementsByTagName("*");
  for (let i = 0; i < all.length; i++) {
    const el = all[i];
    if (el.localName === "majorFont") {
      const latin = kid(el, "latin");
      const tf = latin?.getAttribute("typeface");
      if (tf) out.major = tf;
    } else if (el.localName === "minorFont") {
      const latin = kid(el, "latin");
      const tf = latin?.getAttribute("typeface");
      if (tf) out.minor = tf;
    }
  }
  return out;
}

export class StyleSheet {
  styles = new Map<string, StyleInfo>();
  defaultsR: RunProps = {};
  defaultsP: ParaProps = {};
  defaultParagraph = "Normal";

  constructor(doc: Document | null, theme: ThemeFonts) {
    if (!doc) return;
    const root = doc.documentElement;
    const dd = kid(root, "docDefaults");
    this.defaultsR = parseRPr(path(dd, "rPrDefault", "rPr"), theme);
    this.defaultsP = parsePPr(path(dd, "pPrDefault", "pPr"));
    for (const s of kids(root, "style")) {
      const id = attr(s, "styleId") ?? "";
      const info: StyleInfo = {
        id,
        type: attr(s, "type") ?? "paragraph",
        name: attr(kid(s, "name"), "val") ?? id,
        basedOn: attr(kid(s, "basedOn"), "val"),
        ppr: parsePPr(kid(s, "pPr")),
        rpr: parseRPr(kid(s, "rPr"), theme),
      };
      if (info.type === "table") {
        const tb = path(s, "tblPr", "tblBorders");
        info.tblBordersNone = !!tb && kids(tb).every((b) => (attr(b, "val") ?? "none") === "none" || attr(b, "val") === "nil");
      }
      this.styles.set(id, info);
      if (attr(s, "default") === "1" && info.type === "paragraph") this.defaultParagraph = id;
    }
  }

  chain(id: string | null): StyleInfo[] {
    const out: StyleInfo[] = [];
    let cur = id ? this.styles.get(id) : undefined;
    while (cur && out.length < 12) {
      out.unshift(cur);
      cur = cur.basedOn ? this.styles.get(cur.basedOn) : undefined;
    }
    return out;
  }

  /** Belge varsayılanları + stil zinciri (paragraf stilleri için Normal dahil). */
  effective(id: string | null): { r: RunProps; p: ParaProps } {
    let chain = this.chain(id ?? this.defaultParagraph);
    if (!chain.length) chain = this.chain(this.defaultParagraph);
    const r: RunProps = { ...this.defaultsR };
    const p: ParaProps = { ...this.defaultsP };
    for (const s of chain) {
      Object.assign(r, stripUndef(s.rpr));
      Object.assign(p, stripUndef(s.ppr));
    }
    return { r, p };
  }

  charProps(id: string | null): RunProps {
    const r: RunProps = {};
    for (const s of this.chain(id)) Object.assign(r, stripUndef(s.rpr));
    return r;
  }
}

export function stripUndef<T extends object>(o: T): Partial<T> {
  const out: Partial<T> = {};
  for (const [k, v] of Object.entries(o)) if (v !== undefined) (out as Record<string, unknown>)[k] = v;
  return out;
}

interface LevelInfo {
  numFmt: string;
  lvlText: string;
  start: number;
  /** Düzeye bağlı paragraf stili (w:lvl/w:pStyle). */
  pStyle?: string;
}

export class Numbering {
  private abstracts = new Map<string, LevelInfo[]>();
  private nums = new Map<string, { abstractId: string; starts: Map<number, number> }>();

  constructor(doc: Document | null) {
    if (!doc) return;
    const root = doc.documentElement;
    for (const a of kids(root, "abstractNum")) {
      const levels: LevelInfo[] = [];
      for (const l of kids(a, "lvl")) {
        const i = attrNum(l, "ilvl") ?? levels.length;
        levels[i] = { numFmt: attr(kid(l, "numFmt"), "val") ?? "decimal", lvlText: attr(kid(l, "lvlText"), "val") ?? "%1.", start: attrNum(kid(l, "start"), "val") ?? 1, pStyle: attr(kid(l, "pStyle"), "val") ?? undefined };
      }
      this.abstracts.set(attr(a, "abstractNumId") ?? "", levels);
    }
    for (const n of kids(root, "num")) {
      const starts = new Map<number, number>();
      for (const o of kids(n, "lvlOverride")) {
        const s = attrNum(kid(o, "startOverride"), "val");
        if (s !== null) starts.set(attrNum(o, "ilvl") ?? 0, s);
      }
      this.nums.set(attr(n, "numId") ?? "", { abstractId: attr(kid(n, "abstractNumId"), "val") ?? "", starts });
    }
  }

  /** Stil üzerinden gelen numaralandırmada düzey: soyut tanımda stile bağlı düzey (w:lvl/w:pStyle). */
  styleLevel(numId: string, styleId: string | null): number | null {
    const n = this.nums.get(numId);
    if (!n || !styleId) return null;
    const i = (this.abstracts.get(n.abstractId) ?? []).findIndex((l) => l?.pStyle === styleId);
    return i >= 0 ? i : null;
  }

  /** numId + düzey → liste imzası ve başlangıç. numId 0 = liste yok. */
  resolve(numId: string, ilvl: number): { sig: ListSig; start: number } | null {
    if (numId === "0") return null;
    const n = this.nums.get(numId);
    if (!n) return null;
    const levels = this.abstracts.get(n.abstractId) ?? [];
    const lv = levels[ilvl] ?? levels[0] ?? { numFmt: "decimal", lvlText: "%1.", start: 1 };
    return { sig: sigFromLevel(lv.numFmt, lv.lvlText, ilvl, levels), start: n.starts.get(ilvl) ?? lv.start };
  }
}
