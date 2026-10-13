/** ODF metin okuyucu: paragraflar, listeler (düzeyler, madde işaretleri), aralıklar, alanlar. */
import { normalizeRuns } from "../../model/text";
import type { Align, AutoNumScheme, BulletDef, Paragraph, ParaProps, Run, RunProps } from "../../model/types";
import { attr, kids } from "../xml";
import { lengthPt, lengthPx, odfColor, percent } from "./common";
import type { Props, StyleSheet } from "./styles";

const MAX_PARAS = 5000;
const MAX_RUNS = 20000;

export interface TextReadCtx {
  sheet: StyleSheet;
  /** Yer tutucu anahat metni: liste dışı paragraflar madde işaretsizdir. */
  outline?: boolean;
  /** Köprü çevirici (iç slayt bağlantıları için). */
  link?: (href: string) => string | undefined;
}

/** Metin özellikleri → çalıştırma biçimi. */
export function runPropsOf(tp: Props, sheet: StyleSheet): RunProps {
  const p: RunProps = {};
  const sz = tp["font-size"];
  if (sz && !sz.endsWith("%")) {
    const v = lengthPt(sz);
    if (v !== undefined && v > 0 && v < 4000) p.size = Math.round(v * 100) / 100;
  }
  const fw = tp["font-weight"];
  if (fw) p.bold = fw === "bold" || (/^\d+$/.test(fw) && Number(fw) >= 600);
  const fs = tp["font-style"];
  if (fs) p.italic = fs === "italic" || fs === "oblique";
  const u = tp["text-underline-style"];
  if (u) p.underline = u !== "none";
  else if (tp["text-underline-type"]) p.underline = tp["text-underline-type"] !== "none";
  const st = tp["text-line-through-style"];
  if (st) p.strike = st !== "none";
  if (tp["use-window-font-color"] !== "true") {
    const c = odfColor(tp["color"]);
    if (c) p.color = c;
  }
  const font = sheet.fontOf(tp);
  if (font) p.font = font;
  if (tp["text-transform"] === "uppercase") p.caps = "all";
  else if (tp["font-variant"] === "small-caps") p.caps = "small";
  else if (tp["text-transform"] === "none" || tp["font-variant"] === "normal") p.caps = "none";
  const pos = tp["text-position"];
  if (pos) {
    const first = pos.trim().split(/\s+/)[0];
    if (first === "super") p.baseline = 30;
    else if (first === "sub") p.baseline = -25;
    else {
      const v = percent(first);
      if (v) p.baseline = Math.round(v * 100);
    }
  }
  const bg = tp["background-color"];
  if (bg && bg !== "transparent") {
    const c = odfColor(bg);
    if (c) p.highlight = c;
  }
  const sh = tp["text-shadow"];
  if (sh) p.shadow = sh !== "none";
  const ls = tp["letter-spacing"];
  if (ls && ls !== "normal") {
    const v = lengthPt(ls);
    if (v) p.spacing = Math.round(v * 100) / 100;
  }
  return p;
}

const ALIGN: Record<string, Align> = { start: "l", left: "l", center: "ctr", end: "r", right: "r", justify: "just" };

/** Paragraf özellikleri → paragraf biçimi. */
export function paraPropsOf(pp: Props): ParaProps {
  const p: ParaProps = {};
  const a = ALIGN[pp["text-align"] ?? ""];
  if (a) p.align = a;
  const mt = lengthPt(pp["margin-top"]);
  if (mt !== undefined) p.spaceBefore = Math.max(0, Math.round(mt * 100) / 100);
  const mb = lengthPt(pp["margin-bottom"]);
  if (mb !== undefined) p.spaceAfter = Math.max(0, Math.round(mb * 100) / 100);
  const lh = pp["line-height"];
  if (lh && lh !== "normal") {
    const pc = percent(lh);
    if (pc !== undefined) p.lineSpacing = Math.round(pc * 100);
    else {
      const v = lengthPt(lh);
      if (v) p.lineSpacingPt = Math.round(v * 100) / 100;
    }
  }
  const ml = lengthPx(pp["margin-left"]);
  if (ml !== undefined) p.marL = ml;
  const ti = lengthPx(pp["text-indent"]);
  if (ti !== undefined) p.indent = ti;
  return p;
}

const NUM_SCHEME: Record<string, [AutoNumScheme, AutoNumScheme]> = {
  "1": ["arabicPeriod", "arabicParenR"],
  a: ["alphaLcPeriod", "alphaLcParenR"],
  A: ["alphaUcPeriod", "alphaUcParenR"],
  i: ["romanLcPeriod", "romanLcPeriod"],
  I: ["romanUcPeriod", "romanUcPeriod"],
};

/** Simge yazı tiplerindeki özel kullanım alanı karakterleri yerine düz madde işareti. */
function bulletChar(ch: string | undefined, font: string | undefined): string {
  if (!ch) return "•";
  const code = ch.codePointAt(0) ?? 0;
  if (code >= 0xe000 && code <= 0xf8ff) return code === 0xe00c || code === 0xf0a7 ? "▪" : code === 0xf0d8 || code === 0xe00a ? "➢" : "•";
  if (font && /wingdings|symbol|starbats|dingbats/i.test(font) && code < 0x100) return "•";
  return ch;
}

export interface ListLevel {
  bullet: BulletDef;
  marL?: number;
  indent?: number;
}

/** Liste stilinin bir düzeyi (0 tabanlı). */
export function listLevel(list: Element | undefined, level: number, sheet: StyleSheet): ListLevel | undefined {
  if (!list) return undefined;
  const want = String(level + 1);
  const lv = kids(list).find((k) => attr(k, "level") === want);
  if (!lv) return undefined;
  const lp = kids(lv, "list-level-properties")[0];
  let marL: number | undefined;
  let indent: number | undefined;
  if (lp) {
    if (attr(lp, "list-level-position-and-space-mode") === "label-alignment") {
      const la = kids(lp, "list-level-label-alignment")[0];
      marL = lengthPx(attr(la, "margin-left"));
      indent = lengthPx(attr(la, "text-indent"));
    } else {
      const sb = lengthPx(attr(lp, "space-before")) ?? 0;
      const w = lengthPx(attr(lp, "min-label-width")) ?? 0;
      marL = sb + w;
      indent = -w;
    }
  }
  const tp: Props = {};
  const tpe = kids(lv, "text-properties")[0];
  if (tpe) for (const a of Array.from(tpe.attributes)) tp[a.localName] = a.value;
  const color = tp["use-window-font-color"] === "true" ? undefined : odfColor(tp["color"]);
  const pc = percent(tp["font-size"]);
  const size = pc !== undefined ? Math.round(pc * 100) : undefined;
  let bullet: BulletDef;
  if (lv.localName === "list-level-style-number") {
    const fmt = attr(lv, "num-format") ?? "";
    if (!fmt) bullet = { type: "none" };
    else {
      const pair = NUM_SCHEME[fmt] ?? NUM_SCHEME["1"];
      const start = Number(attr(lv, "start-value") ?? "1");
      bullet = { type: "auto", scheme: (attr(lv, "num-suffix") ?? ".").includes(")") ? pair[1] : pair[0], startAt: Number.isFinite(start) && start > 1 ? start : undefined, color, size };
    }
  } else {
    const font = sheet.fontOf(tp);
    const raw = attr(lv, "bullet-char");
    const ch = bulletChar(raw, font);
    bullet = { type: "char", char: ch, font: ch === raw && font && !/opensymbol|starsymbol|starbats/i.test(font) ? font : undefined, color, size };
  }
  return { bullet, marL, indent };
}

interface Inline {
  props: RunProps;
  link?: string;
}

/** Satır içi düğümleri çalıştırmalara çevirir. */
function collectRuns(el: Element, inh: Inline, ctx: TextReadCtx, out: Run[]) {
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (out.length > MAX_RUNS) return;
    if (n.nodeType === 3) {
      // ODF: ardışık boşluklar tek boşluk sayılır.
      const t = (n.nodeValue ?? "").replace(/[\t\n\r ]+/g, " ");
      if (t) out.push({ text: t, props: withLink(inh) });
      continue;
    }
    if (n.nodeType !== 1) continue;
    const c = n as Element;
    switch (c.localName) {
      case "span": {
        const sp = ctx.sheet.resolve("text", attr(c, "style-name"));
        collectRuns(c, { ...inh, props: { ...inh.props, ...runPropsOf(sp.text, ctx.sheet) } }, ctx, out);
        break;
      }
      case "a": {
        const href = attr(c, "href");
        collectRuns(c, { ...inh, link: href ? (ctx.link ? ctx.link(href) : href) : inh.link }, ctx, out);
        break;
      }
      case "s": {
        const n2 = Math.max(1, Math.min(1000, Number(attr(c, "c") ?? "1") || 1));
        out.push({ text: " ".repeat(n2), props: withLink(inh) });
        break;
      }
      case "tab":
        out.push({ text: "\t", props: withLink(inh) });
        break;
      case "line-break":
        out.push({ text: "\v", props: withLink(inh) });
        break;
      case "page-number":
        out.push({ text: "‹#›", props: { ...withLink(inh), field: "slidenum" } });
        break;
      case "date":
      case "time":
      case "date-time":
        out.push({ text: c.textContent?.trim() || "gg.aa.yyyy", props: { ...withLink(inh), field: "datetime" } });
        break;
      case "footer":
      case "header":
      case "soft-page-break":
      case "bookmark":
      case "bookmark-start":
      case "bookmark-end":
      case "annotation":
      case "annotation-end":
      case "note":
        break;
      default:
        collectRuns(c, inh, ctx, out);
    }
  }
}

function withLink(inh: Inline): RunProps {
  return inh.link ? { ...inh.props, link: inh.link } : { ...inh.props };
}

/** Tek paragraf (text:p / text:h). */
function readPara(p: Element, level: number, list: ListLevel | undefined, inList: boolean, ctx: TextReadCtx): Paragraph {
  const ps = ctx.sheet.resolve("paragraph", attr(p, "style-name"));
  const base = runPropsOf(ps.text, ctx.sheet);
  const runs: Run[] = [];
  collectRuns(p, { props: base }, ctx, runs);
  // Paragraf başı/sonu boşlukları ODF'de görünmez.
  if (runs.length) {
    runs[0] = { ...runs[0], text: runs[0].text.replace(/^ +/, "") };
    const li = runs.length - 1;
    runs[li] = { ...runs[li], text: runs[li].text.replace(/ +$/, "") };
  }
  const props = paraPropsOf(ps.paragraph);
  if (inList && list) {
    // Listede girinti liste düzeyinden gelir; paragraf sol boşluğu üstüne eklenir.
    if (list.bullet) props.bullet = list.bullet;
    if (list.marL !== undefined) props.marL = list.marL + (props.marL ?? 0);
    if (list.indent !== undefined) props.indent = list.indent;
  } else if (!inList && ctx.outline) {
    props.bullet = { type: "none" };
    props.marL ??= 0;
    props.indent ??= 0;
  } else if (inList && !list) {
    props.bullet = { type: "char", char: "•" };
  }
  const para: Paragraph = { runs: normalizeRuns(runs.filter((r) => r.text)), level: Math.max(0, Math.min(8, level)), props };
  if (Object.keys(base).length) para.endProps = base;
  return para;
}

/** Bir kaptaki (draw:text-box ya da şekil) paragrafları okur. */
export function readParagraphs(container: Element, ctx: TextReadCtx): Paragraph[] {
  const out: Paragraph[] = [];
  const walkList = (list: Element, depth: number, inheritedStyle: string | undefined) => {
    if (depth > 9) return;
    const styleName = attr(list, "style-name") ?? inheritedStyle;
    const listEl = styleName ? ctx.sheet.lists.get(styleName) : undefined;
    for (const item of kids(list)) {
      if (item.localName !== "list-item" && item.localName !== "list-header") continue;
      const header = item.localName === "list-header";
      for (const c of kids(item)) {
        if (out.length >= MAX_PARAS) return;
        if (c.localName === "p" || c.localName === "h") {
          const lv = header ? undefined : listLevel(listEl, depth, ctx.sheet);
          const para = readPara(c, depth, lv, !header, ctx);
          if (header) para.props.bullet = { type: "none" };
          out.push(para);
        } else if (c.localName === "list") walkList(c, depth + 1, styleName);
      }
    }
  };
  for (const c of kids(container)) {
    if (out.length >= MAX_PARAS) break;
    if (c.localName === "p" || c.localName === "h") out.push(readPara(c, 0, undefined, false, ctx));
    else if (c.localName === "list") walkList(c, 0, undefined);
  }
  return out;
}

/** Paragraflar; boşsa tek boş paragraf. */
export function paragraphsOf(container: Element, ctx: TextReadCtx): Paragraph[] {
  const paras = readParagraphs(container, ctx);
  return paras.length ? paras : [{ runs: [], level: 0, props: {} }];
}

/** Düz metin (notlar için). */
export function paragraphsText(paras: Paragraph[]): string {
  return paras
    .map((p) => p.runs.map((r) => r.text.replace(/\v/g, "\n")).join(""))
    .join("\n")
    .replace(/\s+$/, "");
}
