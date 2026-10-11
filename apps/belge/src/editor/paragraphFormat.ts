// Paragraf biçimi: stil, girintiler, aralıklar, satır aralığı, gölgelendirme, kenarlıklar, sayfa sonu denetimleri.
import { Extension } from "@tiptap/core";
import { LINE_FACTOR, borderCss, type BorderSpec } from "../model/styles";

export interface ParaBorders {
  top?: BorderSpec;
  bottom?: BorderSpec;
  left?: BorderSpec;
  right?: BorderSpec;
}

export type LineRule = "auto" | "exact" | "atLeast";

export interface ParagraphAttrs {
  styleId: string | null;
  indentLeft: number | null;
  indentRight: number | null;
  /** İlk satır girintisi (pt); negatifse asılı girinti. */
  indentFirst: number | null;
  spaceBefore: number | null;
  spaceAfter: number | null;
  lineRule: LineRule | null;
  /** auto: katsayı, exact/atLeast: pt. */
  lineHeight: number | null;
  shading: string | null;
  borders: ParaBorders | null;
  keepNext: boolean;
  keepLines: boolean;
  pageBreakBefore: boolean;
}

export const PARA_TYPES = ["paragraph", "heading"];

const num = (v: string | null | undefined): number | null => {
  if (!v) return null;
  const m = /^(-?[\d.]+)(pt|px|cm|in)?$/.exec(v.trim());
  if (!m) return null;
  const n = parseFloat(m[1]);
  switch (m[2]) {
    case "px":
      return +(n * 0.75).toFixed(2);
    case "cm":
      return +((n / 2.54) * 72).toFixed(2);
    case "in":
      return n * 72;
    default:
      return n;
  }
};

/** Paragraf özniteliklerinden satır içi CSS. */
export function paragraphCss(a: Partial<ParagraphAttrs>): string {
  const d: string[] = [];
  if (a.indentLeft != null) d.push(`padding-left:${a.indentLeft}pt`);
  if (a.indentRight != null) d.push(`padding-right:${a.indentRight}pt`);
  if (a.indentFirst != null) d.push(`text-indent:${a.indentFirst}pt`);
  if (a.spaceBefore != null) d.push(`margin-top:${a.spaceBefore}pt`);
  if (a.spaceAfter != null) d.push(`margin-bottom:${a.spaceAfter}pt`);
  if (a.lineHeight != null) {
    if (!a.lineRule || a.lineRule === "auto") d.push(`line-height:${+(a.lineHeight * LINE_FACTOR).toFixed(4)}`);
    else if (a.lineRule === "exact") d.push(`line-height:${a.lineHeight}pt`);
    else d.push(`line-height:max(${a.lineHeight}pt, ${LINE_FACTOR}em)`);
  }
  if (a.shading) d.push(`background-color:${a.shading}`);
  if (a.borders) {
    for (const side of ["top", "bottom", "left", "right"] as const) {
      const b = a.borders[side];
      const css = borderCss(b);
      if (css) {
        d.push(`border-${side}:${css}`);
        if (side === "top" || side === "bottom") d.push(`padding-${side}:${b?.space ?? 1}pt`);
        else if (a[side === "left" ? "indentLeft" : "indentRight"] == null) d.push(`padding-${side}:${b?.space ?? 4}pt`);
      }
    }
  }
  if (a.pageBreakBefore) d.push("break-before:page");
  return d.join(";");
}

export const ParagraphFormat = Extension.create({
  name: "paragraphFormat",

  addGlobalAttributes() {
    return [
      {
        types: PARA_TYPES,
        attributes: {
          styleId: {
            default: null,
            parseHTML: (el) => el.getAttribute("data-style"),
            renderHTML: (a) => (a.styleId ? { "data-style": a.styleId } : {}),
          },
          indentLeft: { default: null, parseHTML: (el) => num(el.style.paddingLeft) ?? num(el.style.marginLeft), renderHTML: () => ({}) },
          indentRight: { default: null, parseHTML: (el) => num(el.style.paddingRight) ?? num(el.style.marginRight), renderHTML: () => ({}) },
          indentFirst: { default: null, parseHTML: (el) => num(el.style.textIndent), renderHTML: () => ({}) },
          spaceBefore: { default: null, parseHTML: (el) => num(el.style.marginTop), renderHTML: () => ({}) },
          spaceAfter: { default: null, parseHTML: (el) => num(el.style.marginBottom), renderHTML: () => ({}) },
          lineRule: { default: null, parseHTML: (el) => el.getAttribute("data-line-rule"), renderHTML: (a) => (a.lineRule && a.lineRule !== "auto" ? { "data-line-rule": a.lineRule } : {}) },
          lineHeight: {
            default: null,
            parseHTML: (el) => {
              const v = el.style.lineHeight;
              if (!v) return null;
              if (/^[\d.]+$/.test(v)) return +(parseFloat(v) / LINE_FACTOR).toFixed(2);
              if (v.endsWith("%")) return +(parseFloat(v) / 100 / LINE_FACTOR).toFixed(2);
              return num(v);
            },
            renderHTML: () => ({}),
          },
          shading: { default: null, parseHTML: (el) => el.style.backgroundColor || null, renderHTML: () => ({}) },
          borders: {
            default: null,
            parseHTML: (el) => {
              const raw = el.getAttribute("data-borders");
              if (!raw) return null;
              try {
                return JSON.parse(raw) as ParaBorders;
              } catch {
                return null;
              }
            },
            renderHTML: (a) => (a.borders ? { "data-borders": JSON.stringify(a.borders) } : {}),
          },
          keepNext: { default: false, parseHTML: (el) => el.hasAttribute("data-keep-next"), renderHTML: (a) => (a.keepNext ? { "data-keep-next": "" } : {}) },
          keepLines: { default: false, parseHTML: (el) => el.hasAttribute("data-keep-lines"), renderHTML: (a) => (a.keepLines ? { "data-keep-lines": "" } : {}) },
          pageBreakBefore: {
            default: false,
            parseHTML: (el) => el.style.breakBefore === "page" || el.style.pageBreakBefore === "always",
            // Tüm stil özellikleri tek bir style özniteliğinde birleştirilir.
            renderHTML: (a) => {
              const css = paragraphCss(a as Partial<ParagraphAttrs>);
              return css ? { style: css } : {};
            },
          },
        },
      },
    ];
  },
});

/** Seçimdeki tüm paragraf/başlık düğümlerinin öznitelikleri ortaksa döner. */
export const PARA_DEFAULTS: ParagraphAttrs = {
  styleId: null,
  indentLeft: null,
  indentRight: null,
  indentFirst: null,
  spaceBefore: null,
  spaceAfter: null,
  lineRule: null,
  lineHeight: null,
  shading: null,
  borders: null,
  keepNext: false,
  keepLines: false,
  pageBreakBefore: false,
};
