/** DrawingML metin gövdesi, paragraf ve karakter özellikleri okuyucuları. */
import { normalizeRuns } from "../model/text";
import type {
  AutoNumScheme,
  BulletDef,
  LevelStyle,
  Paragraph,
  ParaProps,
  Run,
  RunProps,
  TextBody,
} from "../model/types";
import { attr, fromEmu, kid, kidPath, kids, numAttr, relId } from "./xml";

import { type ReadCtx, readColor, readFill } from "./readStyle";

function fontName(tf: string | undefined): string | undefined {
  if (!tf) return undefined;
  if (tf.startsWith("+mj")) return "+mj";
  if (tf.startsWith("+mn")) return "+mn";
  return tf;
}

export function readRunProps(rPr: Element | undefined, ctx: ReadCtx): RunProps {
  const p: RunProps = {};
  if (!rPr) return p;
  const sz = numAttr(rPr, "sz");
  if (sz !== undefined) p.size = sz / 100;
  const b = attr(rPr, "b");
  if (b !== undefined) p.bold = b === "1" || b === "true";
  const i = attr(rPr, "i");
  if (i !== undefined) p.italic = i === "1" || i === "true";
  const u = attr(rPr, "u");
  if (u !== undefined) p.underline = u !== "none";
  const st = attr(rPr, "strike");
  if (st !== undefined) p.strike = st !== "noStrike";
  const cap = attr(rPr, "cap");
  if (cap === "all" || cap === "small" || cap === "none") p.caps = cap;
  const spc = numAttr(rPr, "spc");
  if (spc !== undefined && spc !== 0) p.spacing = spc / 100;
  const bl = numAttr(rPr, "baseline");
  if (bl) p.baseline = bl / 1000;
  const fill = readFill(rPr, ctx);
  if (fill?.type === "solid") p.color = fill.color;
  else if (fill?.type === "gradient") p.color = fill.stops[0].color;
  const ln = kid(rPr, "ln");
  if (ln) {
    const lf = readFill(ln, ctx);
    if (lf?.type === "solid") p.outline = lf.color;
  }
  const hl = kid(rPr, "highlight");
  if (hl) p.highlight = readColor(hl, ctx);
  const fx = kid(rPr, "effectLst");
  if (fx) {
    if (kid(fx, "outerShdw")) p.shadow = true;
    const g = kid(fx, "glow");
    if (g) p.glow = readColor(g, ctx);
  }
  const latin = fontName(attr(kid(rPr, "latin"), "typeface"));
  if (latin) p.font = latin;
  const link = relId(kid(rPr, "hlinkClick"));
  if (link) {
    const rel = ctx.rels.get(link);
    if (rel?.external) p.link = rel.target;
  }
  return p;
}

const AUTONUM: Set<string> = new Set(["arabicPeriod", "arabicParenR", "romanUcPeriod", "romanLcPeriod", "alphaUcPeriod", "alphaLcPeriod", "alphaLcParenR", "alphaUcParenR"]);

export function readParaProps(pPr: Element | undefined, ctx: ReadCtx): { props: ParaProps; level: number; defRPr?: RunProps } {
  const props: ParaProps = {};
  if (!pPr) return { props, level: 0 };
  const algn = attr(pPr, "algn");
  if (algn) props.align = algn === "ctr" ? "ctr" : algn === "r" ? "r" : algn === "just" || algn === "dist" || algn === "justLow" ? "just" : "l";
  const marL = numAttr(pPr, "marL");
  if (marL !== undefined) props.marL = fromEmu(marL);
  const ind = numAttr(pPr, "indent");
  if (ind !== undefined) props.indent = fromEmu(ind);
  const ls = kid(pPr, "lnSpc");
  if (ls) {
    const pctv = numAttr(kid(ls, "spcPct"), "val");
    const pts = numAttr(kid(ls, "spcPts"), "val");
    if (pctv !== undefined) props.lineSpacing = pctv / 1000;
    else if (pts !== undefined) props.lineSpacingPt = pts / 100;
  }
  const sb = numAttr(kidPath(pPr, "spcBef", "spcPts"), "val");
  if (sb !== undefined) props.spaceBefore = sb / 100;
  else {
    const sbp = numAttr(kidPath(pPr, "spcBef", "spcPct"), "val");
    if (sbp !== undefined) props.spaceBefore = (sbp / 100000) * 18;
  }
  const sa = numAttr(kidPath(pPr, "spcAft", "spcPts"), "val");
  if (sa !== undefined) props.spaceAfter = sa / 100;
  else {
    const sap = numAttr(kidPath(pPr, "spcAft", "spcPct"), "val");
    if (sap !== undefined) props.spaceAfter = (sap / 100000) * 18;
  }
  let bullet: BulletDef | undefined;
  const buClr = readColor(kid(pPr, "buClr"), ctx);
  const buSz = numAttr(kid(pPr, "buSzPct"), "val");
  const buFont = attr(kid(pPr, "buFont"), "typeface");
  if (kid(pPr, "buNone")) bullet = { type: "none" };
  const ch = kid(pPr, "buChar");
  if (ch) bullet = { type: "char", char: attr(ch, "char") ?? "•", font: buFont && !buFont.startsWith("+") ? buFont : undefined, color: buClr, size: buSz !== undefined ? buSz / 1000 : undefined };
  const an = kid(pPr, "buAutoNum");
  if (an) {
    const t = attr(an, "type") ?? "arabicPeriod";
    bullet = { type: "auto", scheme: (AUTONUM.has(t) ? t : t.startsWith("roman") ? "romanUcPeriod" : t.startsWith("alpha") ? "alphaLcPeriod" : "arabicPeriod") as AutoNumScheme, startAt: numAttr(an, "startAt"), color: buClr, size: buSz !== undefined ? buSz / 1000 : undefined };
  }
  if (bullet) props.bullet = bullet;
  const defR = kid(pPr, "defRPr");
  return { props, level: numAttr(pPr, "lvl") ?? 0, defRPr: defR ? readRunProps(defR, ctx) : undefined };
}

export function readLevels(lst: Element | undefined, ctx: ReadCtx): Partial<LevelStyle>[] | undefined {
  if (!lst) return undefined;
  const out: Partial<LevelStyle>[] = [];
  let any = false;
  for (let i = 1; i <= 9; i++) {
    const el = kid(lst, `lvl${i}pPr`);
    if (!el) continue;
    const r = readParaProps(el, ctx);
    out[i - 1] = { para: r.props, run: r.defRPr ?? {} };
    any = true;
  }
  return any ? out : undefined;
}

export function readParagraph(p: Element, ctx: ReadCtx): Paragraph {
  const pp = readParaProps(kid(p, "pPr"), ctx);
  const runs: Run[] = [];
  for (const el of kids(p)) {
    if (el.localName === "r") {
      runs.push({ text: kid(el, "t")?.textContent ?? "", props: readRunProps(kid(el, "rPr"), ctx) });
    } else if (el.localName === "br") {
      runs.push({ text: "\v", props: readRunProps(kid(el, "rPr"), ctx) });
    } else if (el.localName === "fld") {
      const type = attr(el, "type") ?? "";
      const props = readRunProps(kid(el, "rPr"), ctx);
      const text = kid(el, "t")?.textContent ?? "";
      if (type === "slidenum") props.field = "slidenum";
      else if (type.startsWith("datetime")) props.field = "datetime";
      runs.push({ text: text || (props.field === "slidenum" ? "‹#›" : " "), props });
    }
  }
  const endEl = kid(p, "endParaRPr");
  return { runs: normalizeRuns(runs), level: pp.level, props: pp.props, endProps: endEl ? readRunProps(endEl, ctx) : undefined };
}

export function readTextBody(tx: Element | undefined, ctx: ReadCtx): TextBody | undefined {
  if (!tx) return undefined;
  const body: TextBody = { paragraphs: kids(tx, "p").map((p) => readParagraph(p, ctx)) };
  if (!body.paragraphs.length) body.paragraphs.push({ runs: [], level: 0, props: {} });
  const bp = kid(tx, "bodyPr");
  if (bp) {
    const wrap = attr(bp, "wrap");
    if (wrap) body.wrap = wrap !== "none";
    const anchor = attr(bp, "anchor");
    if (anchor === "t" || anchor === "ctr" || anchor === "b") body.anchor = anchor;
    else if (anchor === "just" || anchor === "dist") body.anchor = "ctr";
    const vert = attr(bp, "vert");
    if (vert === "vert" || vert === "eaVert" || vert === "mongolianVert") body.vert = "vert";
    else if (vert === "vert270") body.vert = "vert270";
    else if (vert === "wordArtVert" || vert === "wordArtVertRtl") body.vert = "stacked";
    else if (vert === "horz") body.vert = "horz";
    const ins = ["lIns", "tIns", "rIns", "bIns"].map((k) => numAttr(bp, k));
    if (ins.some((v) => v !== undefined)) {
      body.insets = { l: fromEmu(ins[0], 9.6), t: fromEmu(ins[1], 4.8), r: fromEmu(ins[2], 9.6), b: fromEmu(ins[3], 4.8) };
    }
    const na = kid(bp, "normAutofit");
    if (na) {
      body.autofit = "shrink";
      const fs = numAttr(na, "fontScale");
      if (fs !== undefined) body.fontScale = fs / 100000;
      const lr = numAttr(na, "lnSpcReduction");
      if (lr !== undefined) body.lnSpcReduction = lr / 100000;
    } else if (kid(bp, "spAutoFit")) body.autofit = "shape";
    else if (kid(bp, "noAutofit")) body.autofit = "none";
  }
  const levels = readLevels(kid(tx, "lstStyle"), ctx);
  if (levels) body.levels = levels;
  return body;
}

/* ------------------------------------------------------------------ */
/* Geometri                                                             */
/* ------------------------------------------------------------------ */
