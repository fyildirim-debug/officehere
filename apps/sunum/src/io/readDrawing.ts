/** DrawingML/PresentationML şekil ve metin okuyucusu. */
import { uid } from "../model/factory";
import { normalizeRuns } from "../model/text";
import type {
  ArrowType,
  AutoNumScheme,
  BulletDef,
  CellBorder,
  Color,
  CustomPath,
  DashType,
  Effects,
  Fill,
  GroupShape,
  LevelStyle,
  Line,
  Paragraph,
  ParaProps,
  PathCmd,
  PicShape,
  Placeholder,
  Run,
  RunProps,
  SchemeColorName,
  Shape,
  SpShape,
  TableCell,
  TableShape,
  TextBody,
} from "../model/types";
import { isLineGeom } from "../render/geometry";
import { tableStyleByGuid, DEFAULT_TABLE_STYLE } from "../render/tableStyles";
import { attr, boolAttr, descendants, fromEmu, kid, kidPath, kids, numAttr, relId, type Rel } from "./xml";

export interface ReadCtx {
  rels: Map<string, Rel>;
  media: Map<string, string>;
  /** schemeClr adı → tema yuvası (clrMap). */
  clrMap: Record<string, SchemeColorName>;
  /** cNvPr kimliği → model kimliği. */
  idMap: Map<number, string>;
  /** Yer tutucu kalıtımı (düzen/asıl). */
  inherit?: (ph: Placeholder) => SpShape | undefined;
  /** Tema satır kalınlıkları (lnStyleLst, pt). */
  lnWidths: number[];
  phClr?: Color;
}

export const DEFAULT_CLRMAP: Record<string, SchemeColorName> = {
  bg1: "lt1",
  tx1: "dk1",
  bg2: "lt2",
  tx2: "dk2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hlink: "hlink",
  folHlink: "folHlink",
  dk1: "dk1",
  lt1: "lt1",
  dk2: "dk2",
  lt2: "lt2",
};

const PRESET_COLORS: Record<string, string> = {
  black: "#000000",
  white: "#FFFFFF",
  red: "#FF0000",
  green: "#008000",
  lime: "#00FF00",
  blue: "#0000FF",
  yellow: "#FFFF00",
  gray: "#808080",
  grey: "#808080",
  silver: "#C0C0C0",
  orange: "#FFA500",
  purple: "#800080",
  navy: "#000080",
  maroon: "#800000",
  teal: "#008080",
  cyan: "#00FFFF",
  magenta: "#FF00FF",
  ltGray: "#D3D3D3",
  dkGray: "#A9A9A9",
};

const COLOR_TAGS = new Set(["srgbClr", "schemeClr", "sysClr", "prstClr", "scrgbClr", "hslClr"]);

function hslToHex(h: number, s: number, l: number): string {
  const f = (n: number) => {
    const k = (n + h / 30) % 12;
    const a = s * Math.min(l, 1 - l);
    return Math.round(255 * (l - a * Math.max(-1, Math.min(k - 3, 9 - k, 1))));
  };
  return `#${[f(0), f(8), f(4)].map((v) => v.toString(16).padStart(2, "0")).join("")}`.toUpperCase();
}

/** Renk öğesini (srgbClr vb.) okur. */
export function readColorEl(el: Element, ctx: ReadCtx): Color | undefined {
  let c: Color | undefined;
  switch (el.localName) {
    case "srgbClr":
      c = { rgb: `#${(attr(el, "val") ?? "000000").toUpperCase()}` };
      break;
    case "schemeClr": {
      const v = attr(el, "val") ?? "tx1";
      if (v === "phClr") {
        c = ctx.phClr ? { ...ctx.phClr } : { scheme: "accent1" };
        break;
      }
      c = { scheme: ctx.clrMap[v] ?? DEFAULT_CLRMAP[v] ?? "dk1" };
      break;
    }
    case "sysClr":
      c = { rgb: `#${(attr(el, "lastClr") ?? (attr(el, "val") === "window" ? "FFFFFF" : "000000")).toUpperCase()}` };
      break;
    case "prstClr":
      c = { rgb: PRESET_COLORS[attr(el, "val") ?? "black"] ?? "#000000" };
      break;
    case "scrgbClr": {
      const t = (k: string) => Math.round(Math.max(0, Math.min(1, (numAttr(el, k) ?? 0) / 100000)) ** (1 / 2.2) * 255);
      c = { rgb: `#${[t("r"), t("g"), t("b")].map((v) => v.toString(16).padStart(2, "0")).join("")}`.toUpperCase() };
      break;
    }
    case "hslClr":
      c = { rgb: hslToHex((numAttr(el, "hue") ?? 0) / 60000, (numAttr(el, "sat") ?? 0) / 100000, (numAttr(el, "lum") ?? 0) / 100000) };
      break;
    default:
      return undefined;
  }
  for (const m of kids(el)) {
    const v = (numAttr(m, "val") ?? 0) / 100000;
    switch (m.localName) {
      case "lumMod":
        c.lumMod = v;
        break;
      case "lumOff":
        c.lumOff = v;
        break;
      case "tint":
        c.tint = v;
        break;
      case "shade":
        c.shade = v;
        break;
      case "alpha":
        c.alpha = v;
        break;
      default:
        break;
    }
  }
  return c;
}

/** Bir kapsayıcıdaki ilk renk öğesi. */
export function readColor(parent: Element | undefined, ctx: ReadCtx): Color | undefined {
  if (!parent) return undefined;
  for (const k of kids(parent)) if (COLOR_TAGS.has(k.localName)) return readColorEl(k, ctx);
  return undefined;
}

export function mediaFor(rid: string | undefined, ctx: ReadCtx): string | undefined {
  if (!rid) return undefined;
  const rel = ctx.rels.get(rid);
  if (!rel || rel.external) return undefined;
  return ctx.media.get(rel.target);
}

/** Dolgu öğesi (noFill/solidFill/…); bulunamazsa undefined. */
export function readFill(parent: Element | undefined, ctx: ReadCtx): Fill | undefined {
  if (!parent) return undefined;
  for (const el of kids(parent)) {
    switch (el.localName) {
      case "noFill":
        return { type: "none" };
      case "solidFill": {
        const c = readColor(el, ctx);
        return c ? { type: "solid", color: c } : { type: "none" };
      }
      case "gradFill": {
        const stops = kids(kid(el, "gsLst"), "gs")
          .map((gs) => ({ pos: (numAttr(gs, "pos") ?? 0) / 100000, color: readColor(gs, ctx) ?? { rgb: "#000000" } }))
          .sort((a, b) => a.pos - b.pos);
        if (!stops.length) return { type: "none" };
        const lin = kid(el, "lin");
        const path = kid(el, "path");
        return { type: "gradient", kind: path ? "radial" : "linear", angle: lin ? (numAttr(lin, "ang") ?? 0) / 60000 : 90, stops };
      }
      case "blipFill": {
        const src = mediaFor(attr(kid(el, "blip"), "embed"), ctx);
        if (!src) return { type: "none" };
        const amt = numAttr(kidPath(el, "blip", "alphaModFix"), "amt");
        return { type: "image", src, mode: kid(el, "tile") ? "tile" : "stretch", alpha: amt !== undefined ? amt / 100000 : undefined };
      }
      case "pattFill": {
        const c = readColor(kid(el, "fgClr"), ctx);
        return c ? { type: "solid", color: c } : { type: "none" };
      }
      case "grpFill":
        return { type: "none" };
      default:
        break;
    }
  }
  return undefined;
}

const DASHES = new Set(["solid", "sysDot", "sysDash", "dash", "dashDot", "lgDash", "lgDashDot", "lgDashDotDot"]);

export function readLine(ln: Element | undefined, ctx: ReadCtx, base?: Line): Line | undefined {
  if (!ln) return base;
  const w = numAttr(ln, "w");
  const f = readFill(ln, ctx);
  const dashV = attr(kid(ln, "prstDash"), "val");
  const head = attr(kid(ln, "headEnd"), "type") as ArrowType | undefined;
  const tail = attr(kid(ln, "tailEnd"), "type") as ArrowType | undefined;
  const line: Line = {
    fill: f ? (f.type === "solid" ? f : f.type === "gradient" ? { type: "solid", color: f.stops[0].color } : { type: "none" }) : (base?.fill ?? { type: "none" }),
    width: w !== undefined ? w / 12700 : (base?.width ?? 0.75),
    dash: dashV && DASHES.has(dashV) ? (dashV as DashType) : dashV ? "dash" : (base?.dash ?? "solid"),
    head: head ?? base?.head,
    tail: tail ?? base?.tail,
  };
  if (line.head === "none") line.head = undefined;
  if (line.tail === "none") line.tail = undefined;
  return line;
}

export function readEffects(parent: Element | undefined, ctx: ReadCtx): Effects | undefined {
  const lst = kid(parent, "effectLst");
  if (!lst) return undefined;
  const fx: Effects = {};
  const sh = kid(lst, "outerShdw");
  if (sh) fx.shadow = { color: readColor(sh, ctx) ?? { rgb: "#000000", alpha: 0.4 }, blur: fromEmu(numAttr(sh, "blurRad")), dist: fromEmu(numAttr(sh, "dist")), dir: (numAttr(sh, "dir") ?? 0) / 60000 };
  const glow = kid(lst, "glow");
  if (glow) fx.glow = { color: readColor(glow, ctx) ?? { scheme: "accent1" }, radius: fromEmu(numAttr(glow, "rad")) };
  if (kid(lst, "reflection")) fx.reflection = true;
  const se = kid(lst, "softEdge");
  if (se) fx.softEdge = fromEmu(numAttr(se, "rad"));
  return fx.shadow || fx.glow || fx.reflection || fx.softEdge ? fx : undefined;
}

/* ------------------------------------------------------------------ */
/* Metin                                                                */
/* ------------------------------------------------------------------ */

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

interface Xf {
  x: number;
  y: number;
  w: number;
  h: number;
  rot: number;
  flipH?: boolean;
  flipV?: boolean;
  chX?: number;
  chY?: number;
  chW?: number;
  chH?: number;
}

export function readXfrm(el: Element | undefined): Xf | undefined {
  if (!el) return undefined;
  const off = kid(el, "off");
  const ext = kid(el, "ext");
  const chOff = kid(el, "chOff");
  const chExt = kid(el, "chExt");
  return {
    x: fromEmu(numAttr(off, "x")),
    y: fromEmu(numAttr(off, "y")),
    w: fromEmu(numAttr(ext, "cx")),
    h: fromEmu(numAttr(ext, "cy")),
    rot: (numAttr(el, "rot") ?? 0) / 60000,
    flipH: boolAttr(el, "flipH") || undefined,
    flipV: boolAttr(el, "flipV") || undefined,
    chX: chOff ? fromEmu(numAttr(chOff, "x")) : undefined,
    chY: chOff ? fromEmu(numAttr(chOff, "y")) : undefined,
    chW: chExt ? fromEmu(numAttr(chExt, "cx")) : undefined,
    chH: chExt ? fromEmu(numAttr(chExt, "cy")) : undefined,
  };
}

function arcToBeziers(cx: number, cy: number, rx: number, ry: number, start: number, sweep: number): number[][] {
  const segs = Math.max(1, Math.ceil(Math.abs(sweep) / (Math.PI / 2)));
  const d = sweep / segs;
  const k = (4 / 3) * Math.tan(d / 4);
  const out: number[][] = [];
  let a = start;
  for (let i = 0; i < segs; i++) {
    const a2 = a + d;
    const x1 = cx + rx * Math.cos(a);
    const y1 = cy + ry * Math.sin(a);
    const x2 = cx + rx * Math.cos(a2);
    const y2 = cy + ry * Math.sin(a2);
    out.push([x1 - k * rx * Math.sin(a), y1 + k * ry * Math.cos(a), x2 + k * rx * Math.sin(a2), y2 - k * ry * Math.cos(a2), x2, y2]);
    a = a2;
  }
  return out;
}

export function readCustGeom(cg: Element): CustomPath[] {
  const paths: CustomPath[] = [];
  for (const p of kids(kid(cg, "pathLst"), "path")) {
    const w = numAttr(p, "w") ?? 0;
    const h = numAttr(p, "h") ?? 0;
    const cmds: PathCmd[] = [];
    let cur: [number, number] = [0, 0];
    const pt = (e: Element): [number, number] => [numAttr(e, "x") ?? 0, numAttr(e, "y") ?? 0];
    for (const c of kids(p)) {
      const pts = kids(c, "pt").map(pt);
      switch (c.localName) {
        case "moveTo":
          cur = pts[0] ?? cur;
          cmds.push({ c: "M", pts: [...cur] });
          break;
        case "lnTo":
          cur = pts[0] ?? cur;
          cmds.push({ c: "L", pts: [...cur] });
          break;
        case "cubicBezTo":
          if (pts.length >= 3) {
            cmds.push({ c: "C", pts: pts.flat() });
            cur = pts[2];
          }
          break;
        case "quadBezTo":
          if (pts.length >= 2) {
            cmds.push({ c: "Q", pts: pts.flat() });
            cur = pts[1];
          }
          break;
        case "arcTo": {
          const wR = numAttr(c, "wR") ?? 0;
          const hR = numAttr(c, "hR") ?? 0;
          const st = (((numAttr(c, "stAng") ?? 0) / 60000) * Math.PI) / 180;
          const sw = (((numAttr(c, "swAng") ?? 0) / 60000) * Math.PI) / 180;
          const cx = cur[0] - wR * Math.cos(st);
          const cy = cur[1] - hR * Math.sin(st);
          for (const b of arcToBeziers(cx, cy, wR, hR, st, sw)) {
            cmds.push({ c: "C", pts: b });
            cur = [b[4], b[5]];
          }
          break;
        }
        case "close":
          cmds.push({ c: "Z", pts: [] });
          break;
        default:
          break;
      }
    }
    const maxX = Math.max(1, ...cmds.flatMap((c) => c.pts.filter((_, i) => i % 2 === 0)));
    const maxY = Math.max(1, ...cmds.flatMap((c) => c.pts.filter((_, i) => i % 2 === 1)));
    paths.push({ w: w || maxX, h: h || maxY, cmds, fill: attr(p, "fill") === "none" ? false : undefined, stroke: attr(p, "stroke") === "0" || attr(p, "stroke") === "false" ? false : undefined });
  }
  return paths;
}

function readGeom(spPr: Element | undefined): { geom: string; adj?: Record<string, number>; paths?: CustomPath[] } {
  const pg = kid(spPr, "prstGeom");
  if (pg) {
    const adj: Record<string, number> = {};
    for (const gd of kids(kid(pg, "avLst"), "gd")) {
      const m = /val\s+(-?\d+)/.exec(attr(gd, "fmla") ?? "");
      if (m) adj[attr(gd, "name") ?? "adj"] = Number(m[1]);
    }
    return { geom: attr(pg, "prst") ?? "rect", adj: Object.keys(adj).length ? adj : undefined };
  }
  const cg = kid(spPr, "custGeom");
  if (cg) {
    const paths = readCustGeom(cg);
    return paths.length ? { geom: "custom", paths } : { geom: "rect" };
  }
  return { geom: "rect" };
}

/* ------------------------------------------------------------------ */
/* Şekiller                                                             */
/* ------------------------------------------------------------------ */

function nvProps(nv: Element | undefined, ctx: ReadCtx): { id: string; name: string; descr?: string; hidden?: boolean; ph?: Placeholder; txBox?: boolean } {
  const c = kid(nv, "cNvPr");
  const nid = numAttr(c, "id");
  const id = uid();
  if (nid !== undefined) ctx.idMap.set(nid, id);
  const phEl = kidPath(nv, "nvPr", "ph");
  let ph: Placeholder | undefined;
  if (phEl) {
    const type = (attr(phEl, "type") ?? "obj") as Placeholder["type"];
    ph = { type, idx: numAttr(phEl, "idx"), sz: attr(phEl, "sz"), orient: attr(phEl, "orient") };
    if (ph.idx === undefined) delete ph.idx;
    if (!ph.sz) delete ph.sz;
    if (!ph.orient) delete ph.orient;
  }
  const cNvSpPr = kids(nv).find((k) => k.localName === "cNvSpPr");
  return {
    id,
    name: attr(c, "name") ?? "",
    descr: attr(c, "descr") || undefined,
    hidden: boolAttr(c, "hidden") || undefined,
    ph,
    txBox: boolAttr(cNvSpPr, "txBox") || undefined,
  };
}

function styleRefs(style: Element | undefined, ctx: ReadCtx) {
  if (!style) return {};
  const fillRef = kid(style, "fillRef");
  const lnRef = kid(style, "lnRef");
  const fontRef = kid(style, "fontRef");
  const fx = kid(style, "effectRef");
  const fillIdx = numAttr(fillRef, "idx") ?? 0;
  const lnIdx = numAttr(lnRef, "idx") ?? 0;
  const fill: Fill | undefined = fillRef ? (fillIdx > 0 ? { type: "solid", color: readColor(fillRef, ctx) ?? { scheme: "accent1" } } : { type: "none" }) : undefined;
  const lnColor = readColor(lnRef, ctx);
  const line: Line | undefined = lnRef ? (lnIdx > 0 && lnColor ? { fill: { type: "solid", color: lnColor }, width: ctx.lnWidths[lnIdx - 1] ?? 1, dash: "solid" } : { fill: { type: "none" }, width: 1, dash: "solid" }) : undefined;
  const fontColor = readColor(fontRef, ctx);
  const fontIdx = attr(fontRef, "idx");
  const effects: Effects | undefined = (numAttr(fx, "idx") ?? 0) >= 3 ? { shadow: { color: { rgb: "#000000", alpha: 0.4 }, blur: 4, dist: 2, dir: 90 } } : undefined;
  return { fill, line, fontColor, font: fontIdx === "major" ? "+mj" : fontIdx === "minor" ? "+mn" : undefined, effects };
}

function readSp(el: Element, ctx: ReadCtx, connector: boolean): SpShape {
  const nv = kid(el, connector ? "nvCxnSpPr" : "nvSpPr");
  const info = nvProps(nv, ctx);
  const spPr = kid(el, "spPr");
  const st = styleRefs(kid(el, "style"), ctx);
  const xf = readXfrm(kid(spPr, "xfrm"));
  const inh = info.ph && ctx.inherit ? ctx.inherit(info.ph) : undefined;
  const g = readGeom(spPr);
  const geomEl = kid(spPr, "prstGeom") || kid(spPr, "custGeom");
  let fill = readFill(spPr, ctx) ?? st.fill ?? inh?.fill ?? { type: "none" };
  let line = readLine(kid(spPr, "ln"), ctx, st.line ?? inh?.line) ?? st.line ?? inh?.line ?? { fill: { type: "none" }, width: 1, dash: "solid" };
  if (connector && fill.type !== "none" && !readFill(spPr, ctx)) fill = { type: "none" };
  if (line.fill.type === "solid" && line.width === undefined) line = { ...line, width: 1 };
  let text = readTextBody(kid(el, "txBody"), ctx);
  if ((st.fontColor || st.font) && text) {
    const lv = text.levels ? [...text.levels] : [];
    const l0 = lv[0] ?? {};
    lv[0] = { ...l0, run: { ...(st.fontColor ? { color: st.fontColor } : {}), ...(st.font && st.font !== "+mn" ? { font: st.font } : {}), ...(l0.run ?? {}) } };
    text = { ...text, levels: lv };
  }
  const sp: SpShape = {
    kind: "sp",
    id: info.id,
    name: info.name,
    descr: info.descr,
    hidden: info.hidden,
    x: xf?.x ?? inh?.x ?? 0,
    y: xf?.y ?? inh?.y ?? 0,
    w: xf?.w ?? inh?.w ?? 0,
    h: xf?.h ?? inh?.h ?? 0,
    rot: xf ? xf.rot : (inh?.rot ?? 0),
    flipH: xf?.flipH,
    flipV: xf?.flipV,
    geom: geomEl ? g.geom : (inh?.geom ?? "rect"),
    adj: geomEl ? g.adj : inh?.adj,
    paths: geomEl ? g.paths : inh?.paths,
    fill,
    line,
    effects: readEffects(spPr, ctx) ?? st.effects,
    text,
    ph: info.ph,
    txBox: info.txBox,
    connector: connector || undefined,
  };
  if (!sp.text && !connector && !isLineGeom(sp.geom) && !info.ph) sp.text = undefined;
  return sp;
}

function readPic(el: Element, ctx: ReadCtx): PicShape {
  const info = nvProps(kid(el, "nvPicPr"), ctx);
  const spPr = kid(el, "spPr");
  const xf = readXfrm(kid(spPr, "xfrm"));
  const inh = info.ph && ctx.inherit ? ctx.inherit(info.ph) : undefined;
  const bf = kid(el, "blipFill");
  const blip = kid(bf, "blip");
  const src = mediaFor(attr(blip, "embed"), ctx) ?? "";
  const sr = kid(bf, "srcRect");
  const crop = sr ? { l: (numAttr(sr, "l") ?? 0) / 100000, t: (numAttr(sr, "t") ?? 0) / 100000, r: (numAttr(sr, "r") ?? 0) / 100000, b: (numAttr(sr, "b") ?? 0) / 100000 } : undefined;
  const lum = kid(blip, "lum");
  const amt = numAttr(kid(blip, "alphaModFix"), "amt");
  const adjust = {
    brightness: lum ? (numAttr(lum, "bright") ?? 0) / 100000 || undefined : undefined,
    contrast: lum ? (numAttr(lum, "contrast") ?? 0) / 100000 || undefined : undefined,
    grayscale: kid(blip, "grayscl") ? true : undefined,
    alpha: amt !== undefined ? amt / 100000 : undefined,
  };
  const g = readGeom(spPr);
  return {
    kind: "pic",
    id: info.id,
    name: info.name,
    descr: info.descr,
    hidden: info.hidden,
    x: xf?.x ?? inh?.x ?? 0,
    y: xf?.y ?? inh?.y ?? 0,
    w: xf?.w ?? inh?.w ?? 0,
    h: xf?.h ?? inh?.h ?? 0,
    rot: xf?.rot ?? 0,
    flipH: xf?.flipH,
    flipV: xf?.flipV,
    src,
    crop: crop && (crop.l || crop.t || crop.r || crop.b) ? crop : undefined,
    geom: g.geom !== "rect" && g.geom !== "custom" ? g.geom : undefined,
    line: readLine(kid(spPr, "ln"), ctx) ?? { fill: { type: "none" }, width: 1, dash: "solid" },
    effects: readEffects(spPr, ctx),
    adjust: adjust.brightness || adjust.contrast || adjust.grayscale || adjust.alpha !== undefined ? adjust : undefined,
    ph: info.ph,
  };
}

function readBorder(ln: Element | undefined, ctx: ReadCtx): CellBorder | null | undefined {
  if (!ln) return undefined;
  const f = readFill(ln, ctx);
  if (!f || f.type === "none") return f ? null : undefined;
  const color = f.type === "solid" ? f.color : f.type === "gradient" ? f.stops[0].color : { rgb: "#000000" };
  return { color, width: (numAttr(ln, "w") ?? 12700) / 12700 };
}

function readTable(gf: Element, tbl: Element, ctx: ReadCtx): TableShape {
  const info = nvProps(kid(gf, "nvGraphicFramePr"), ctx);
  const xf = readXfrm(kid(gf, "xfrm"));
  const tblPr = kid(tbl, "tblPr");
  const guid = kid(tblPr, "tableStyleId")?.textContent?.trim();
  const style = guid ? tableStyleByGuid(guid) : undefined;
  const cols = kids(kid(tbl, "tblGrid"), "gridCol").map((c) => fromEmu(numAttr(c, "w")));
  const rows = kids(tbl, "tr").map((tr) => ({
    h: fromEmu(numAttr(tr, "h")),
    cells: kids(tr, "tc").map((tc): TableCell => {
      const tcPr = kid(tc, "tcPr");
      const fill = readFill(tcPr, ctx);
      const borders = { l: readBorder(kid(tcPr, "lnL"), ctx), r: readBorder(kid(tcPr, "lnR"), ctx), t: readBorder(kid(tcPr, "lnT"), ctx), b: readBorder(kid(tcPr, "lnB"), ctx) };
      const hasB = Object.values(borders).some((v) => v !== undefined);
      const anchor = attr(tcPr, "anchor");
      return {
        text: readTextBody(kid(tc, "txBody"), ctx) ?? { paragraphs: [{ runs: [], level: 0, props: {} }] },
        fill,
        gridSpan: numAttr(tc, "gridSpan"),
        rowSpan: numAttr(tc, "rowSpan"),
        hMerge: boolAttr(tc, "hMerge") || undefined,
        vMerge: boolAttr(tc, "vMerge") || undefined,
        borders: hasB ? borders : undefined,
        anchor: anchor === "ctr" || anchor === "b" || anchor === "t" ? anchor : undefined,
      };
    }),
  }));
  // Eksik hücreleri tamamla.
  for (const r of rows) while (r.cells.length < cols.length) r.cells.push({ text: { paragraphs: [{ runs: [], level: 0, props: {} }] } });
  const total = rows.reduce((t, r) => t + r.h, 0);
  return {
    kind: "table",
    id: info.id,
    name: info.name,
    x: xf?.x ?? 0,
    y: xf?.y ?? 0,
    w: cols.reduce((t, c) => t + c, 0) || (xf?.w ?? 0),
    h: Math.max(total, xf?.h ?? 0),
    rot: 0,
    cols,
    rows,
    style: {
      styleId: style?.id ?? DEFAULT_TABLE_STYLE,
      firstRow: boolAttr(tblPr, "firstRow") ?? false,
      lastRow: boolAttr(tblPr, "lastRow") ?? false,
      firstCol: boolAttr(tblPr, "firstCol") ?? false,
      lastCol: boolAttr(tblPr, "lastCol") ?? false,
      bandRow: boolAttr(tblPr, "bandRow") ?? false,
      bandCol: boolAttr(tblPr, "bandCol") ?? false,
    },
  };
}

/** Alt şekli grup koordinatlarından üst koordinatlara taşır. */
function transformShape(s: Shape, f: (x: number, y: number) => [number, number], sx: number, sy: number): Shape {
  const [nx, ny] = f(s.x, s.y);
  const base = { ...s, x: nx, y: ny, w: s.w * sx, h: s.h * sy };
  if (s.kind === "group") return { ...base, children: s.children.map((c) => transformShape(c, f, sx, sy)) } as GroupShape;
  if (s.kind === "table") return { ...base, cols: s.cols.map((c) => c * sx), rows: s.rows.map((r) => ({ ...r, h: r.h * sy })) } as TableShape;
  return base as Shape;
}

function readGroup(el: Element, ctx: ReadCtx): GroupShape | undefined {
  const info = nvProps(kid(el, "nvGrpSpPr"), ctx);
  const xf = readXfrm(kidPath(el, "grpSpPr", "xfrm"));
  const children = readShapeList(el, ctx);
  if (!children.length) return undefined;
  let mapped = children;
  if (xf) {
    const chW = xf.chW || xf.w || 1;
    const chH = xf.chH || xf.h || 1;
    const sx = xf.w && chW ? xf.w / chW : 1;
    const sy = xf.h && chH ? xf.h / chH : 1;
    const ox = xf.chX ?? 0;
    const oy = xf.chY ?? 0;
    mapped = children.map((c) => transformShape(c, (x, y) => [xf.x + (x - ox) * sx, xf.y + (y - oy) * sy], sx, sy));
  }
  return {
    kind: "group",
    id: info.id,
    name: info.name || "Grup",
    x: xf?.x ?? 0,
    y: xf?.y ?? 0,
    w: xf?.w ?? 0,
    h: xf?.h ?? 0,
    rot: xf?.rot ?? 0,
    flipH: xf?.flipH,
    flipV: xf?.flipV,
    hidden: info.hidden,
    children: mapped,
  };
}

function readShapeEl(el: Element, ctx: ReadCtx): Shape | undefined {
  switch (el.localName) {
    case "sp":
      return readSp(el, ctx, false);
    case "cxnSp":
      return readSp(el, ctx, true);
    case "pic":
      return readPic(el, ctx);
    case "grpSp":
      return readGroup(el, ctx);
    case "graphicFrame": {
      const tbl = descendants(el, "tbl")[0];
      if (tbl) return readTable(el, tbl, ctx);
      // Grafik/diyagram: yedek resim varsa onu kullan.
      const pic = descendants(el, "pic")[0];
      if (pic) {
        const p = readPic(pic, ctx);
        const xf = readXfrm(kid(el, "xfrm"));
        if (xf) Object.assign(p, { x: xf.x, y: xf.y, w: xf.w, h: xf.h });
        return p.src ? p : undefined;
      }
      return undefined;
    }
    case "AlternateContent": {
      const fb = kid(el, "Fallback") ?? kid(el, "Choice");
      if (!fb) return undefined;
      const list = kids(fb).map((c) => readShapeEl(c, ctx)).filter((x): x is Shape => !!x);
      return list[0];
    }
    default:
      return undefined;
  }
}

/** spTree / grpSp altındaki şekiller. */
export function readShapeList(tree: Element | undefined, ctx: ReadCtx): Shape[] {
  const out: Shape[] = [];
  for (const el of kids(tree)) {
    if (el.localName === "nvGrpSpPr" || el.localName === "grpSpPr") continue;
    const s = readShapeEl(el, ctx);
    if (s) out.push(s);
  }
  return out;
}
