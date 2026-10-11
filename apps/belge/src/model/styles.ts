// Belge stilleri (Normal, Başlık 1… gibi). Görünüm CSS'i ve DOCX styles.xml aynı tablodan üretilir.
import { THEMES, fontStack, type ThemePalette } from "@officehere/core";

export interface BorderSpec {
  style: "single" | "double" | "dotted" | "dashed" | "thick";
  /** Kalınlık (pt). */
  size: number;
  color: string;
  /** Metinle arası (pt). */
  space?: number;
}

export interface StyleProps {
  font?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  caps?: boolean;
  /** Karakter aralığı (pt). */
  spacing?: number;
  underline?: boolean;
  spaceBefore?: number;
  spaceAfter?: number;
  /** Satır aralığı katsayısı (auto). */
  line?: number;
  align?: "left" | "center" | "right" | "justify";
  indentLeft?: number;
  indentRight?: number;
  keepNext?: boolean;
  borderTop?: BorderSpec;
  borderBottom?: BorderSpec;
  shading?: string;
}

export interface StyleDef {
  id: string;
  name: string;
  type: "paragraph" | "character";
  basedOn?: string;
  next?: string;
  /** Başlık düzeyi (Heading1 → 1). */
  level?: number;
  props: StyleProps;
}

/** Word'ün tek satır aralığı yazı tipi yüksekliğinin ~1,2 katıdır; CSS'e bu katsayıyla çevrilir. */
export const LINE_FACTOR = 1.2;

export const NORMAL_SIZE = 11;
export const NORMAL_LINE = 1.08;
export const NORMAL_AFTER = 8;

export const STYLE_ORDER = [
  "Normal",
  "NoSpacing",
  "Heading1",
  "Heading2",
  "Heading3",
  "Heading4",
  "Heading5",
  "Heading6",
  "Title",
  "Subtitle",
  "SubtleEmphasis",
  "Emphasis",
  "IntenseEmphasis",
  "Strong",
  "Quote",
  "IntenseQuote",
  "SubtleReference",
  "IntenseReference",
  "BookTitle",
  "ListParagraph",
] as const;

export type StyleId = (typeof STYLE_ORDER)[number];

interface ThemeCtx {
  major: string;
  minor: string;
  accent: string;
  heading: string;
  text1: string;
}

function darken(hex: string, f: number): string {
  const n = parseInt(hex.slice(1), 16);
  const c = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((v) => Math.round(v * f));
  return `#${c.map((v) => v.toString(16).padStart(2, "0")).join("")}`.toUpperCase();
}

export function themeByName(name: string): ThemePalette {
  return THEMES.find((t) => t.name === name) ?? THEMES[0];
}

function themeCtx(theme: ThemePalette): ThemeCtx {
  return { major: theme.majorFont, minor: theme.minorFont, accent: theme.colors[4], heading: darken(theme.colors[4], 0.75), text1: theme.colors[1] };
}

function baseStyles(t: ThemeCtx): StyleDef[] {
  const hb = (size: number, extra: StyleProps = {}): StyleProps => ({ font: t.major, size, color: t.heading, keepNext: true, spaceBefore: 8, spaceAfter: 4, ...extra });
  return [
    { id: "Normal", name: "Normal", type: "paragraph", props: { font: t.minor, size: NORMAL_SIZE, spaceBefore: 0, spaceAfter: NORMAL_AFTER, line: NORMAL_LINE } },
    { id: "NoSpacing", name: "Aralık Yok", type: "paragraph", basedOn: "Normal", props: { spaceAfter: 0, line: 1 } },
    { id: "Heading1", name: "Başlık 1", type: "paragraph", basedOn: "Normal", next: "Normal", level: 1, props: hb(20, { spaceBefore: 18 }) },
    { id: "Heading2", name: "Başlık 2", type: "paragraph", basedOn: "Normal", next: "Normal", level: 2, props: hb(16) },
    { id: "Heading3", name: "Başlık 3", type: "paragraph", basedOn: "Normal", next: "Normal", level: 3, props: hb(14, { font: t.minor }) },
    { id: "Heading4", name: "Başlık 4", type: "paragraph", basedOn: "Normal", next: "Normal", level: 4, props: hb(11, { font: t.minor, italic: true, spaceBefore: 4, spaceAfter: 2 }) },
    { id: "Heading5", name: "Başlık 5", type: "paragraph", basedOn: "Normal", next: "Normal", level: 5, props: hb(11, { font: t.minor, spaceBefore: 4, spaceAfter: 2 }) },
    { id: "Heading6", name: "Başlık 6", type: "paragraph", basedOn: "Normal", next: "Normal", level: 6, props: hb(11, { font: t.minor, italic: true, color: "#595959", spaceBefore: 2, spaceAfter: 0 }) },
    { id: "Title", name: "Konu Başlığı", type: "paragraph", basedOn: "Normal", next: "Normal", props: { font: t.major, size: 28, spacing: -0.5, spaceAfter: 4, line: 1 } },
    { id: "Subtitle", name: "Alt Konu Başlığı", type: "paragraph", basedOn: "Normal", next: "Normal", props: { font: t.minor, size: 14, color: "#595959", spacing: 0.75 } },
    { id: "SubtleEmphasis", name: "Hafif Vurgu", type: "character", props: { italic: true, color: "#404040" } },
    { id: "Emphasis", name: "Vurgu", type: "character", props: { italic: true } },
    { id: "IntenseEmphasis", name: "Belirgin Vurgu", type: "character", props: { italic: true, color: t.heading } },
    { id: "Strong", name: "Güçlü", type: "character", props: { bold: true } },
    { id: "Quote", name: "Alıntı", type: "paragraph", basedOn: "Normal", next: "Normal", props: { italic: true, color: "#404040", align: "center", spaceBefore: 8, spaceAfter: 8 } },
    {
      id: "IntenseQuote",
      name: "Belirgin Alıntı",
      type: "paragraph",
      basedOn: "Normal",
      next: "Normal",
      props: {
        italic: true,
        color: t.heading,
        align: "center",
        spaceBefore: 18,
        spaceAfter: 18,
        indentLeft: 43.2,
        indentRight: 43.2,
        borderTop: { style: "single", size: 0.5, color: t.heading, space: 10 },
        borderBottom: { style: "single", size: 0.5, color: t.heading, space: 10 },
      },
    },
    { id: "SubtleReference", name: "Hafif Başvuru", type: "character", props: { caps: true, color: "#404040" } },
    { id: "IntenseReference", name: "Belirgin Başvuru", type: "character", props: { bold: true, caps: true, color: t.heading, spacing: 0.25 } },
    { id: "BookTitle", name: "Kitap Başlığı", type: "character", props: { bold: true, italic: true, spacing: 0.25 } },
    { id: "ListParagraph", name: "Liste Paragrafı", type: "paragraph", basedOn: "Normal", props: { indentLeft: 36 } },
  ];
}

/* ------------------------------------------------------------------ */
/* Stil kümeleri (Tasarım > Belge Biçimlendirme)                        */
/* ------------------------------------------------------------------ */

export interface StyleSetDef {
  id: string;
  label: string;
  apply: (s: Record<string, StyleDef>, t: ThemeCtx) => void;
}

const patch = (s: Record<string, StyleDef>, id: string, p: StyleProps) => {
  s[id] = { ...s[id], props: { ...s[id].props, ...p } };
};

export const STYLE_SETS: StyleSetDef[] = [
  { id: "default", label: "Varsayılan", apply: () => {} },
  {
    id: "basic",
    label: "Temel (Basit)",
    apply: (s, t) => {
      patch(s, "Heading1", { size: 18, color: t.text1, bold: true, spaceBefore: 12, spaceAfter: 6 });
      patch(s, "Heading2", { size: 14, color: t.text1, bold: true, spaceBefore: 10, spaceAfter: 4 });
      patch(s, "Heading3", { size: 12, color: t.text1, bold: true });
      patch(s, "Title", { size: 26, bold: true });
    },
  },
  {
    id: "classic",
    label: "Siyah Beyaz (Klasik)",
    apply: (s) => {
      patch(s, "Normal", { font: "Times New Roman", size: 12, line: 1.15 });
      for (const id of ["Heading1", "Heading2", "Heading3", "Heading4", "Heading5", "Heading6", "Title"]) patch(s, id, { font: "Times New Roman", color: "#000000" });
      patch(s, "Heading1", { size: 18, bold: true, align: "center" });
      patch(s, "Heading2", { size: 14, bold: true });
      patch(s, "Heading3", { size: 12, bold: true, italic: true });
      patch(s, "Title", { size: 26, align: "center", caps: true, spacing: 1 });
    },
  },
  {
    id: "lines",
    label: "Çizgili",
    apply: (s, t) => {
      patch(s, "Heading1", { size: 18, borderBottom: { style: "single", size: 1, color: t.accent, space: 2 }, spaceBefore: 16, spaceAfter: 6 });
      patch(s, "Heading2", { size: 14, borderBottom: { style: "single", size: 0.5, color: t.accent, space: 1 } });
      patch(s, "Title", { borderBottom: { style: "single", size: 1.5, color: t.accent, space: 4 }, spaceAfter: 12 });
    },
  },
  {
    id: "shaded",
    label: "Gölgelendirme",
    apply: (s, t) => {
      patch(s, "Heading1", { size: 16, color: "#FFFFFF", shading: t.heading, spaceBefore: 12, spaceAfter: 6 });
      patch(s, "Heading2", { size: 13, color: t.heading, shading: "#DCE6F2" });
      patch(s, "Title", { color: t.heading, caps: true });
    },
  },
  {
    id: "centered",
    label: "Ortalanmış",
    apply: (s) => {
      for (const id of ["Heading1", "Heading2", "Heading3", "Title", "Subtitle"]) patch(s, id, { align: "center" });
      patch(s, "Heading1", { caps: true, spacing: 1 });
    },
  },
  {
    id: "elegant",
    label: "Şık",
    apply: (s, t) => {
      patch(s, "Normal", { font: "Georgia", size: 11, line: 1.2, spaceAfter: 10 });
      for (const id of ["Heading1", "Heading2", "Heading3", "Title", "Subtitle"]) patch(s, id, { font: "Georgia" });
      patch(s, "Heading1", { size: 20, italic: true, color: t.accent });
      patch(s, "Heading2", { size: 15, italic: true, color: t.accent });
      patch(s, "Title", { size: 32, italic: true, color: t.heading });
    },
  },
];

/** Tema + stil kümesine göre çözülmüş stil tablosu. */
export function resolveStyles(styleSet: string, themeName: string): Record<string, StyleDef> {
  const t = themeCtx(themeByName(themeName));
  const out: Record<string, StyleDef> = {};
  for (const s of baseStyles(t)) out[s.id] = s;
  (STYLE_SETS.find((x) => x.id === styleSet) ?? STYLE_SETS[0]).apply(out, t);
  return out;
}

/** basedOn zinciriyle birleşmiş özellikler. */
export function effectiveProps(styles: Record<string, StyleDef>, id: string): StyleProps {
  const chain: StyleDef[] = [];
  let cur: StyleDef | undefined = styles[id];
  while (cur && chain.length < 10) {
    chain.unshift(cur);
    cur = cur.basedOn ? styles[cur.basedOn] : undefined;
  }
  if (chain[0]?.type === "paragraph" && chain[0].id !== "Normal" && styles.Normal) chain.unshift(styles.Normal);
  return Object.assign({}, ...chain.map((c) => c.props)) as StyleProps;
}

/** Paragraf düğümünden stil kimliği. */
export function styleIdOf(nodeType: string, attrs: Record<string, unknown>): string {
  if (nodeType === "heading") return `Heading${attrs.level ?? 1}`;
  return (attrs.styleId as string | null) ?? "Normal";
}

/* ------------------------------------------------------------------ */
/* CSS üretimi                                                          */
/* ------------------------------------------------------------------ */

export function borderCss(b: BorderSpec | undefined): string | undefined {
  if (!b) return undefined;
  const style = b.style === "thick" ? "solid" : b.style === "single" ? "solid" : b.style;
  const width = b.style === "thick" ? Math.max(b.size, 2.25) : b.style === "double" ? Math.max(b.size * 3, 2.25) : b.size;
  return `${Math.max(width, 0.5)}pt ${style} ${b.color}`;
}

function propsCss(p: StyleProps, paragraph: boolean): string {
  const d: string[] = [];
  if (p.font) d.push(`font-family:${fontStack(p.font)}`);
  if (p.size) d.push(`font-size:${p.size}pt`);
  if (p.bold !== undefined) d.push(`font-weight:${p.bold ? 700 : 400}`);
  if (p.italic !== undefined) d.push(`font-style:${p.italic ? "italic" : "normal"}`);
  if (p.color) d.push(`color:${p.color}`);
  if (p.caps) d.push("text-transform:uppercase");
  if (p.spacing !== undefined) d.push(`letter-spacing:${p.spacing}pt`);
  if (p.underline) d.push("text-decoration:underline");
  if (paragraph) {
    if (p.spaceBefore !== undefined) d.push(`margin-top:${p.spaceBefore}pt`);
    if (p.spaceAfter !== undefined) d.push(`margin-bottom:${p.spaceAfter}pt`);
    if (p.line !== undefined) d.push(`line-height:${+(p.line * LINE_FACTOR).toFixed(4)}`);
    if (p.align) d.push(`text-align:${p.align}`);
    if (p.indentLeft !== undefined) d.push(`padding-left:${p.indentLeft}pt`);
    if (p.indentRight !== undefined) d.push(`padding-right:${p.indentRight}pt`);
    const bt = borderCss(p.borderTop);
    const bb = borderCss(p.borderBottom);
    if (bt) d.push(`border-top:${bt}`, `padding-top:${p.borderTop?.space ?? 1}pt`);
    if (bb) d.push(`border-bottom:${bb}`, `padding-bottom:${p.borderBottom?.space ?? 1}pt`);
    if (p.shading) d.push(`background-color:${p.shading}`);
  }
  return d.join(";");
}

/** Belge alanı için stil CSS'i (scope: .bl-doc). */
export function styleSheetCss(styles: Record<string, StyleDef>): string {
  const rules: string[] = [];
  const normal = styles.Normal.props;
  rules.push(`.bl-doc{${propsCss({ font: normal.font, size: normal.size }, false)};line-height:${+((normal.line ?? 1) * LINE_FACTOR).toFixed(4)}}`);
  rules.push(`.bl-doc p,.bl-doc h1,.bl-doc h2,.bl-doc h3,.bl-doc h4,.bl-doc h5,.bl-doc h6{${propsCss({ ...normal, font: undefined, size: undefined }, true)}}`);
  for (const id of STYLE_ORDER) {
    const s = styles[id];
    if (!s || id === "Normal") continue;
    const props = s.props;
    if (s.type === "character") {
      rules.push(`.bl-doc [data-cstyle="${id}"]{${propsCss(props, false)}}`);
    } else if (s.level) {
      rules.push(`.bl-doc h${s.level}{font-weight:400;${propsCss(props, true)}}`);
    } else {
      rules.push(`.bl-doc p[data-style="${id}"]{${propsCss(props, true)}}`);
    }
  }
  return rules.join("\n");
}

/** Galeri önizlemesinde kullanılacak satır içi stil. */
export function previewStyle(styles: Record<string, StyleDef>, id: string): React.CSSProperties {
  const p = id in styles && styles[id].type === "paragraph" ? effectiveProps(styles, id) : { ...styles.Normal.props, ...(styles[id]?.props ?? {}) };
  return {
    fontFamily: p.font ? fontStack(p.font) : undefined,
    fontSize: `${Math.min(15, Math.max(9, (p.size ?? 11) * 0.72))}pt`,
    fontWeight: p.bold ? 700 : 400,
    fontStyle: p.italic ? "italic" : "normal",
    color: p.color && p.color.toUpperCase() !== "#FFFFFF" ? p.color : p.shading ? p.shading : "#000",
    textTransform: p.caps ? "uppercase" : undefined,
    letterSpacing: p.spacing ? `${p.spacing}pt` : undefined,
    borderBottom: p.borderBottom ? `1px solid ${p.borderBottom.color}` : undefined,
  };
}
