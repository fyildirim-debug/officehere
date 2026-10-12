/**
 * Metin gövdesini HTML dizgesine çevirir. Aynı çıktı statik çizimde,
 * küçük resimlerde, gösteride ve düzenleyicide (contentEditable) kullanılır;
 * böylece düzenleme sırasında görünüm birebir aynı kalır.
 */
import { fontStack } from "@officehere/core";
import { cssColor } from "../model/color";
import { effectiveRun, resolveFont } from "../model/textStyle";
import type { AutoNumScheme, LevelStyle, Paragraph, ParaProps, RunProps, TextBody, Theme } from "../model/types";
import { PT } from "./style";

export interface TextRenderCtx {
  theme: Theme;
  lvl: (n: number) => LevelStyle;
  fontScale?: number;
  lnReduction?: number;
  slideNumber?: number;
  editing?: boolean;
  /** Boş metinde gösterilecek istem metni. */
  prompt?: string;
}

const ESC: Record<string, string> = { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" };
export const escHtml = (s: string) => s.replace(/[&<>"]/g, (c) => ESC[c]);

const ROMAN: [number, string][] = [
  [1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"], [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"],
];

function roman(n: number): string {
  let s = "";
  for (const [v, r] of ROMAN) while (n >= v) (s += r), (n -= v);
  return s;
}

function alpha(n: number): string {
  let s = "";
  while (n > 0) {
    n -= 1;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

/** Numara etiketinin yaklaşık genişliği (px); sekme durağı kararı için. */
function numberWidthPx(label: string, sizePt: number): number {
  let em = 0;
  for (const ch of label) em += /[.,:;()[\]|!iIl]/.test(ch) ? 0.27 : /[MWmw]/.test(ch) ? 0.85 : 0.52;
  return (em * sizePt * 96) / 72;
}

export function autoNumber(scheme: AutoNumScheme, n: number): string {
  switch (scheme) {
    case "arabicPeriod":
      return `${n}.`;
    case "arabicParenR":
      return `${n})`;
    case "romanUcPeriod":
      return `${roman(n)}.`;
    case "romanLcPeriod":
      return `${roman(n).toLowerCase()}.`;
    case "alphaUcPeriod":
      return `${alpha(n)}.`;
    case "alphaLcPeriod":
      return `${alpha(n).toLowerCase()}.`;
    case "alphaLcParenR":
      return `${alpha(n).toLowerCase()})`;
    case "alphaUcParenR":
      return `${alpha(n)})`;
  }
}

export function runCss(p: RunProps, ctx: TextRenderCtx): string {
  const scale = ctx.fontScale ?? 1;
  const size = (p.size ?? 18) * scale;
  const parts: string[] = [`font-family:${fontStack(resolveFont(p.font, ctx.theme)).replace(/"/g, "'")}`];
  let fs = size;
  if (p.baseline) fs = size * 0.68;
  parts.push(`font-size:${fs.toFixed(2)}pt`);
  if (p.bold) parts.push("font-weight:700");
  if (p.italic) parts.push("font-style:italic");
  const deco: string[] = [];
  if (p.underline) deco.push("underline");
  if (p.strike) deco.push("line-through");
  if (deco.length) parts.push(`text-decoration:${deco.join(" ")}`);
  parts.push(`color:${cssColor(p.color, ctx.theme)}`);
  if (p.highlight) parts.push(`background:${cssColor(p.highlight, ctx.theme)}`);
  const shadows: string[] = [];
  if (p.shadow) shadows.push("1.5px 1.5px 2px rgba(0,0,0,0.45)");
  if (p.glow) {
    const g = cssColor(p.glow, ctx.theme);
    shadows.push(`0 0 6px ${g}`, `0 0 10px ${g}`);
  }
  if (shadows.length) parts.push(`text-shadow:${shadows.join(",")}`);
  if (p.outline) parts.push(`-webkit-text-stroke:${(size / 40).toFixed(2)}pt ${cssColor(p.outline, ctx.theme)};paint-order:stroke fill`);
  if (p.caps === "all") parts.push("text-transform:uppercase");
  if (p.caps === "small") parts.push("font-variant:small-caps");
  if (p.spacing) parts.push(`letter-spacing:${p.spacing}pt`);
  if (p.baseline) parts.push(`vertical-align:${p.baseline > 0 ? "super" : "sub"}`);
  return parts.join(";");
}

function fieldText(p: RunProps, text: string, ctx: TextRenderCtx): string {
  if (p.field === "slidenum") return ctx.slideNumber !== undefined ? String(ctx.slideNumber) : text;
  if (p.field === "datetime") return new Intl.DateTimeFormat("tr-TR", { day: "2-digit", month: "2-digit", year: "numeric" }).format(new Date());
  return text;
}

function textToHtml(t: string): string {
  return t
    .split("\v")
    .map((part) => escHtml(part))
    .join('<br data-lb="1">');
}

export function paraCss(ep: ParaProps, ctx: TextRenderCtx, first: boolean, refSize: number): string {
  const scale = ctx.fontScale ?? 1;
  const parts: string[] = [];
  const align = ep.align === "ctr" ? "center" : ep.align === "r" ? "right" : ep.align === "just" ? "justify" : "left";
  parts.push(`text-align:${align}`);
  if (ep.marL) parts.push(`margin-left:${ep.marL.toFixed(2)}px`);
  if (ep.indent) parts.push(`text-indent:${ep.indent.toFixed(2)}px`);
  if (ep.lineSpacingPt) parts.push(`line-height:${(ep.lineSpacingPt * scale).toFixed(2)}pt`);
  else {
    const ls = Math.max(0.3, (ep.lineSpacing ?? 100) / 100 - (ctx.lnReduction ?? 0));
    parts.push(`line-height:${(1.2 * ls).toFixed(3)}`);
  }
  if (!first && ep.spaceBefore) parts.push(`margin-top:${(ep.spaceBefore * scale).toFixed(2)}pt`);
  if (ep.spaceAfter) parts.push(`margin-bottom:${(ep.spaceAfter * scale).toFixed(2)}pt`);
  void refSize;
  return parts.join(";");
}

/** Paragrafların HTML'i (düzenleyici kapsayıcısının içeriği). */
export function paragraphsHtml(body: TextBody, ctx: TextRenderCtx): string {
  const counters: number[] = [];
  const out: string[] = [];
  const empty = body.paragraphs.every((p) => p.runs.every((r) => !r.text));
  if (empty && ctx.prompt && !ctx.editing) {
    const p0 = body.paragraphs[0];
    const st = ctx.lvl(p0?.level ?? 0);
    const rp = effectiveRun(p0?.endProps ?? {}, p0 ?? { runs: [], level: 0, props: {} }, ctx.lvl);
    const ep = { ...st.para, ...(p0?.props ?? {}) };
    const bu = ep.bullet;
    let attr = "";
    let extra = "";
    if (bu && bu.type === "char") {
      attr = ` data-bullet="${escHtml(bu.char)}"`;
      const bfont = bu.font ? fontStack(bu.font) : fontStack(resolveFont(rp.font, ctx.theme));
      extra = `;--bu-font:${bfont.replace(/"/g, "'")};--bu-size:${(((rp.size ?? 18) * (bu.size ?? 100)) / 100).toFixed(2)}pt;--bu-color:${cssColor(bu.color ?? rp.color, ctx.theme)};--bu-w:${Math.max(0, -(ep.indent ?? 0)).toFixed(2)}px`;
    }
    return `<div class="tb-p is-prompt"${attr} style="${paraCss({ ...ep, bullet: undefined }, ctx, true, rp.size ?? 18)}${extra}"><span style="${runCss(rp, ctx)}">${escHtml(ctx.prompt)}</span></div>`;
  }
  body.paragraphs.forEach((p, i) => {
    const st = ctx.lvl(p.level);
    const ep: ParaProps = { ...st.para };
    for (const [k, v] of Object.entries(p.props)) if (v !== undefined) (ep as Record<string, unknown>)[k] = v;
    counters.length = Math.min(counters.length, p.level + 1);
    const firstRun = p.runs.find((r) => r.text) ?? null;
    const refProps = effectiveRun(firstRun?.props ?? p.endProps ?? {}, p, ctx.lvl);
    let bulletAttr = "";
    let bulletStyle = "";
    const isEmpty = !p.runs.some((r) => r.text);
    const bu = ep.bullet;
    if (bu && bu.type === "auto") {
      counters[p.level] = (counters[p.level] ?? (bu.startAt ?? 1) - 1) + 1;
    } else {
      counters[p.level] = undefined as unknown as number;
    }
    if (bu && bu.type !== "none" && (!isEmpty || ctx.editing)) {
      const label = bu.type === "char" ? bu.char : autoNumber(bu.scheme, counters[p.level] ?? 1);
      const bsize = ((refProps.size ?? 18) * (ctx.fontScale ?? 1) * (bu.size ?? 100)) / 100;
      const bfont = bu.type === "char" && bu.font ? fontStack(bu.font) : fontStack(resolveFont(refProps.font, ctx.theme));
      const bcolor = cssColor(bu.color ?? refProps.color, ctx.theme);
      let minW = Math.max(0, -(ep.indent ?? 0));
      if (bu.type === "auto") {
        // Numara asılı girintiye sığmazsa metin bir sonraki varsayılan sekme durağından (1 inç) başlar.
        const est = numberWidthPx(label, bsize);
        if (est > minW - 2) {
          const start = Math.max(0, (ep.marL ?? 0) + (ep.indent ?? 0));
          minW = Math.ceil((start + est + 4) / 96) * 96 - start;
        }
      }
      bulletAttr = ` data-bullet="${escHtml(label)}"`;
      bulletStyle = `;--bu-font:${bfont.replace(/"/g, "'")};--bu-size:${bsize.toFixed(2)}pt;--bu-color:${bcolor};--bu-w:${minW.toFixed(2)}px`;
      if (isEmpty && ctx.editing) bulletStyle += ";--bu-op:0.45";
    }
    const css = paraCss(ep, ctx, i === 0, refProps.size ?? 18) + bulletStyle;
    let inner = "";
    if (!p.runs.length || isEmpty) {
      inner = `<span data-r="1" style="${runCss(effectiveRun(p.endProps ?? {}, p, ctx.lvl), ctx)}"><br data-ph="1"></span>`;
    } else {
      for (const r of p.runs) {
        if (!r.text) continue;
        const rp = effectiveRun(r.props, p, ctx.lvl);
        if (r.props.field) {
          inner += `<span data-r="1" data-field="1" data-len="${r.text.length}" contenteditable="false" style="${runCss(rp, ctx)}">${escHtml(fieldText(r.props, r.text, ctx))}</span>`;
        } else {
          inner += `<span data-r="1" style="${runCss(rp, ctx)}">${textToHtml(r.text)}</span>`;
        }
      }
      const lastText = p.runs[p.runs.length - 1]?.text ?? "";
      if (lastText.endsWith("\v")) inner += '<br data-ph="1">';
    }
    out.push(`<div class="tb-p" data-p="${i}"${bulletAttr} style="${css}">${inner}</div>`);
  });
  return out.join("");
}

export function isParagraphEmpty(p: Paragraph): boolean {
  return !p.runs.some((r) => r.text);
}

export { PT };
