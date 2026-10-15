/**
 * ODF çizim nesneleri okuyucusu: çerçeveler (metin kutusu, resim, tablo),
 * hazır/özel şekiller, çizgiler, gruplar; dolgu, çizgi, döndürme.
 */
import { uid } from "../../model/factory";
import type { CustomPath, Fill, GroupShape, Line, Paragraph, PathCmd, PicShape, Placeholder, Shape, SpShape, TextBody } from "../../model/types";
import { svgPathToCmds } from "../../render/svgPath";
import { attr, kids } from "../xml";
import { angleDeg, lengthPt, lengthPx, odfColor, percent } from "./common";
import { EMPTY, type Props, type Resolved, type StyleSheet } from "./styles";
import { readTable } from "./table";
import { paragraphsOf, paraPropsOf, runPropsOf } from "./text";

const PRES_NS = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const DRAW_NS = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";

const MAX_SHAPES = 3000;

export interface ShapeReadCtx {
  sheet: StyleSheet;
  /** Paket içi yol ya da gömülü veri → data: URL (gösterilemiyorsa undefined). */
  image: (href: string | undefined, mime: string | undefined, inline: string | undefined) => string | undefined;
  link?: (href: string) => string | undefined;
  /** Yer tutucunun ana slayttaki stil düzeyleri (sınıf → düzeyler). */
  placeholderLevels?: (cls: string, own: Resolved) => TextBody["levels"];
  count: { n: number };
  /** Gömülü nesne çerçevesi (draw:object/draw:object-ole) → özgün XML'i taşıyan şekil (geçiş katmanı; odp/reader.ts). */
  object?: (el: Element, base: ShapeBaseProps) => Shape | undefined;
}

/** ODF şekil türü → model hazır şekil adı. */
const ODF_TYPES: Record<string, string> = {
  rectangle: "rect",
  "round-rectangle": "roundRect",
  ellipse: "ellipse",
  circle: "ellipse",
  "isosceles-triangle": "triangle",
  "right-triangle": "rtTriangle",
  parallelogram: "parallelogram",
  trapezoid: "trapezoid",
  diamond: "diamond",
  pentagon: "pentagon",
  hexagon: "hexagon",
  octagon: "octagon",
  cross: "plus",
  ring: "donut",
  can: "can",
  cube: "cube",
  paper: "foldedCorner",
  smiley: "smileyFace",
  heart: "heart",
  sun: "sun",
  moon: "moon",
  lightning: "lightningBolt",
  cloud: "cloud",
  forbidden: "noSmoking",
  frame: "frame",
  star4: "star4",
  star5: "star5",
  star6: "star6",
  star8: "star8",
  star12: "star12",
  star24: "star12",
  "right-arrow": "rightArrow",
  "left-arrow": "leftArrow",
  "up-arrow": "upArrow",
  "down-arrow": "downArrow",
  "left-right-arrow": "leftRightArrow",
  "up-down-arrow": "upDownArrow",
  "quad-arrow": "quadArrow",
  "pentagon-right": "homePlate",
  chevron: "chevron",
  "flowchart-process": "flowChartProcess",
  "flowchart-decision": "flowChartDecision",
  "flowchart-data": "flowChartInputOutput",
  "flowchart-predefined-process": "flowChartPredefinedProcess",
  "flowchart-document": "flowChartDocument",
  "flowchart-terminator": "flowChartTerminator",
  "flowchart-connector": "flowChartConnector",
  "rectangular-callout": "wedgeRectCallout",
  "round-rectangular-callout": "wedgeRoundRectCallout",
  "round-callout": "wedgeEllipseCallout",
  "cloud-callout": "cloudCallout",
  "octagon-bevel": "bevel",
  "quad-bevel": "bevel",
  "horizontal-scroll": "rect",
  "mso-spt100": "rect",
};

/** Desteklenen hazır şekiller (ooxml-* türleri için). */
const KNOWN = new Set([
  ...Object.values(ODF_TYPES),
  "snip1Rect", "round2SameRect", "heptagon", "decagon", "dodecagon", "pie", "teardrop", "corner", "wave", "star10", "mathPlus", "mathMinus", "mathMultiply", "mathEqual", "irregularSeal1",
  "line", "straightConnector1", "bentConnector3", "curvedConnector3", "arc",
]);

export function geomFromOdfType(t: string | undefined): string | undefined {
  if (!t) return undefined;
  if (t.startsWith("ooxml-")) {
    const g = t.slice(6);
    return KNOWN.has(g) || g.startsWith("actionButton") ? g : undefined;
  }
  return ODF_TYPES[t];
}

/* ------------------------------------------------------------------ */
/* Dolgu ve çizgi                                                       */
/* ------------------------------------------------------------------ */

export function readFillProps(g: Props, sheet: StyleSheet, ctx: ShapeReadCtx, def: "none" | "solid"): Fill {
  const kind = g["fill"] ?? def;
  const opacity = percent(g["opacity"]);
  const alpha = opacity !== undefined && opacity < 1 ? Math.max(0, opacity) : undefined;
  switch (kind) {
    case "none":
      return { type: "none" };
    case "gradient": {
      const gr = sheet.named.get(`gradient:${g["fill-gradient-name"] ?? ""}`);
      if (!gr) break;
      const a = odfColor(attr(gr, "start-color")) ?? { rgb: "#FFFFFF" };
      const b = odfColor(attr(gr, "end-color")) ?? { rgb: "#000000" };
      const style = attr(gr, "style") ?? "linear";
      const ang = angleDeg(attr(gr, "angle")) ?? 0;
      const angle = (((90 - ang) % 360) + 360) % 360;
      if (style === "axial") return { type: "gradient", kind: "linear", angle, stops: [{ pos: 0, color: b }, { pos: 0.5, color: a }, { pos: 1, color: b }] };
      return { type: "gradient", kind: style === "linear" ? "linear" : "radial", angle, stops: [{ pos: 0, color: a }, { pos: 1, color: b }] };
    }
    case "bitmap": {
      const fi = sheet.named.get(`fill-image:${g["fill-image-name"] ?? ""}`);
      const inline = fi ? kids(fi, "binary-data")[0]?.textContent ?? undefined : undefined;
      const src = fi ? ctx.image(attr(fi, "href"), undefined, inline) : undefined;
      if (!src) break;
      return { type: "image", src, mode: g["repeat"] === "repeat" ? "tile" : "stretch", alpha };
    }
    case "hatch": {
      const h = sheet.named.get(`hatch:${g["fill-hatch-name"] ?? ""}`);
      const c = g["fill-hatch-solid"] === "true" ? odfColor(g["fill-color"]) : odfColor(attr(h, "color"));
      return c ? { type: "solid", color: alpha !== undefined ? { ...c, alpha } : c } : { type: "none" };
    }
  }
  if (kind === "none") return { type: "none" };
  const c = odfColor(g["fill-color"]) ?? { rgb: "#729FCF" };
  return { type: "solid", color: alpha !== undefined ? { ...c, alpha } : c };
}

export function readLineProps(g: Props, sheet: StyleSheet, def: "none" | "solid"): Line {
  const kind = g["stroke"] ?? def;
  if (kind === "none") return { fill: { type: "none" }, width: 1, dash: "solid" };
  const c = odfColor(g["stroke-color"]) ?? { rgb: "#3465A4" };
  const op = percent(g["stroke-opacity"]);
  const w = lengthPt(g["stroke-width"]);
  const line: Line = { fill: { type: "solid", color: op !== undefined && op < 1 ? { ...c, alpha: op } : c }, width: w && w > 0 ? Math.round(w * 100) / 100 : 0.75, dash: "solid" };
  if (kind === "dash") {
    const d = sheet.named.get(`stroke-dash:${g["stroke-dash"] ?? ""}`);
    const dots = Number(attr(d, "dots1") ?? "1") + Number(attr(d, "dots2") ?? "0");
    const len = lengthPx(attr(d, "dots1-length")) ?? percent(attr(d, "dots1-length")) ?? 0;
    line.dash = len && len < 3 ? "sysDot" : dots > 1 ? "dashDot" : "dash";
  }
  if (g["marker-start"]) line.head = "triangle";
  if (g["marker-end"]) line.tail = "triangle";
  const cap = g["stroke-linecap"];
  if (cap === "round") line.cap = "rnd";
  else if (cap === "square") line.cap = "sq";
  return line;
}

/* ------------------------------------------------------------------ */
/* Konum                                                                */
/* ------------------------------------------------------------------ */

type Mat = [number, number, number, number, number, number];
const mul = (m: Mat, n: Mat): Mat => [m[0] * n[0] + m[2] * n[1], m[1] * n[0] + m[3] * n[1], m[0] * n[2] + m[2] * n[3], m[1] * n[2] + m[3] * n[3], m[0] * n[4] + m[2] * n[5] + m[4], m[1] * n[4] + m[3] * n[5] + m[5]];

/** draw:transform → matris (işlemler yazıldığı sırayla uygulanır). */
export function parseTransform(v: string | undefined): Mat | undefined {
  if (!v) return undefined;
  let m: Mat = [1, 0, 0, 1, 0, 0];
  const re = /(rotate|translate|scale|skewX|skewY|matrix)\s*\(([^)]*)\)/gi;
  let hit: RegExpExecArray | null;
  let any = false;
  while ((hit = re.exec(v))) {
    any = true;
    const args = hit[2].trim().split(/[\s,]+/).filter(Boolean);
    let op: Mat;
    switch (hit[1].toLowerCase()) {
      case "rotate": {
        const a = Number(args[0]) || 0;
        // ODF: pozitif açı saat yönünün tersine.
        op = [Math.cos(a), -Math.sin(a), Math.sin(a), Math.cos(a), 0, 0];
        break;
      }
      case "translate":
        op = [1, 0, 0, 1, lengthPx(args[0]) ?? 0, lengthPx(args[1] ?? "0") ?? 0];
        break;
      case "scale": {
        const sx = Number(args[0]) || 1;
        op = [sx, 0, 0, args[1] !== undefined ? Number(args[1]) || 1 : sx, 0, 0];
        break;
      }
      case "matrix": {
        const n = args.map((x, i) => (i >= 4 ? (lengthPx(x) ?? 0) : Number(x) || 0));
        op = [n[0] ?? 1, n[1] ?? 0, n[2] ?? 0, n[3] ?? 1, n[4] ?? 0, n[5] ?? 0];
        break;
      }
      default:
        continue;
    }
    m = mul(op, m);
  }
  return any ? m : undefined;
}

interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
  rot: number;
}

export function readBox(el: Element): Box {
  const w = Math.max(0, lengthPx(attr(el, "width")) ?? 0);
  const h = Math.max(0, lengthPx(attr(el, "height")) ?? 0);
  const x0 = lengthPx(attr(el, "x")) ?? 0;
  const y0 = lengthPx(attr(el, "y")) ?? 0;
  const m = parseTransform(attr(el, "transform"));
  if (!m) return { x: x0, y: y0, w, h, rot: 0 };
  const full = mul(m, [1, 0, 0, 1, x0, y0]);
  const sx = Math.hypot(full[0], full[1]) || 1;
  const sy = Math.hypot(full[2], full[3]) || 1;
  const W = w * sx;
  const H = h * sy;
  const cx = full[0] * (w / 2) + full[2] * (h / 2) + full[4];
  const cy = full[1] * (w / 2) + full[3] * (h / 2) + full[5];
  const ccw = (Math.atan2(-full[1], full[0]) * 180) / Math.PI;
  let rot = Math.round(((-ccw % 360) + 360) * 1000) / 1000 % 360;
  if (Math.abs(rot - 360) < 1e-6) rot = 0;
  return { x: cx - W / 2, y: cy - H / 2, w: W, h: H, rot };
}

/* ------------------------------------------------------------------ */
/* Metin gövdesi                                                        */
/* ------------------------------------------------------------------ */

function bodyProps(g: Props, isFrame: boolean): Omit<TextBody, "paragraphs"> {
  const b: Omit<TextBody, "paragraphs"> = {};
  const va = g["textarea-vertical-align"];
  b.anchor = va === "middle" ? "ctr" : va === "bottom" ? "b" : va === "top" ? "t" : isFrame ? "t" : "ctr";
  // Varsayılan stildeki fo:wrap-option güvenilmez; genişliğe göre büyüyen çerçeve satır kaydırmaz.
  if (isFrame && g["auto-grow-width"] === "true") b.wrap = false;
  if (g["shrink-to-fit"] === "true" || g["fit-to-size"] === "shrink-to-fit") b.autofit = "shrink";
  else if (g["auto-grow-height"] === "true" && isFrame) b.autofit = "shape";
  else b.autofit = "none";
  const pads = ["padding-left", "padding-top", "padding-right", "padding-bottom"].map((k) => lengthPx(g[k] ?? g["padding"]));
  if (pads.some((v) => v !== undefined)) b.insets = { l: pads[0] ?? 9.6, t: pads[1] ?? 4.8, r: pads[2] ?? 9.6, b: pads[3] ?? 4.8 };
  const wm = g["writing-mode"];
  if (wm === "tb-rl" || wm === "tb") b.vert = "vert";
  return b;
}

/** Çerçevenin/şeklin stil zincirinden gelen metin düzeyleri (yer tutucu olmayan şekiller). */
function baseLevels(r: Resolved, sheet: StyleSheet): TextBody["levels"] {
  const run = readRunBase(r, sheet);
  const para = readParaBase(r);
  if (!Object.keys(run).length && !Object.keys(para).length) return undefined;
  return Array.from({ length: 9 }, () => ({ para: { ...para }, run: { ...run } }));
}

const readRunBase = (r: Resolved, sheet: StyleSheet) => runPropsOf(r.text, sheet);
const readParaBase = (r: Resolved) => {
  const p = paraPropsOf(r.paragraph);
  delete p.marL;
  delete p.indent;
  return p;
};

function textOf(container: Element, ctx: ShapeReadCtx, outline: boolean): Paragraph[] {
  return paragraphsOf(container, { sheet: ctx.sheet, outline, link: ctx.link });
}

const PH_CLASS: Record<string, Placeholder["type"]> = {
  title: "title",
  subtitle: "subTitle",
  outline: "body",
  text: "body",
  graphic: "pic",
  object: "obj",
  chart: "obj",
  table: "tbl",
  "date-time": "dt",
  footer: "ftr",
  "page-number": "sldNum",
};

/* ------------------------------------------------------------------ */
/* Şekiller                                                             */
/* ------------------------------------------------------------------ */

export type ShapeBaseProps = Omit<SpShape, "kind" | "geom" | "fill" | "line">;

function baseShape(el: Element, box: Box): ShapeBaseProps {
  const name = attr(el, "name") ?? "";
  const desc = kids(el, "desc")[0]?.textContent ?? kids(el, "title")[0]?.textContent ?? undefined;
  return { id: uid(), name: name || "Şekil", x: box.x, y: box.y, w: box.w, h: box.h, rot: box.rot, ...(desc ? { descr: desc } : {}) };
}

/** Bir sayfadaki (ya da gruptaki) çizim nesneleri. */
export function readShapes(parent: Element, ctx: ShapeReadCtx): Shape[] {
  const out: Shape[] = [];
  for (const el of kids(parent)) {
    if (ctx.count.n >= MAX_SHAPES) break;
    const s = readShape(el, ctx);
    if (s) {
      ctx.count.n++;
      out.push(s);
    }
  }
  return out;
}

function readShape(el: Element, ctx: ShapeReadCtx): Shape | undefined {
  const s = readShapeInner(el, ctx);
  if (s && attr(el, "display") === "none") s.hidden = true;
  return s;
}

function readShapeInner(el: Element, ctx: ShapeReadCtx): Shape | undefined {
  switch (el.localName) {
    case "frame":
      return readFrame(el, ctx);
    case "rect":
    case "ellipse":
    case "circle":
    case "custom-shape":
    case "polygon":
    case "polyline":
    case "path":
    case "caption":
      return readDrawShape(el, ctx);
    case "line":
    case "connector":
    case "measure":
      return readLine(el, ctx);
    case "g":
      return readGroup(el, ctx);
    default:
      return undefined;
  }
}

function styleOf(el: Element, ctx: ShapeReadCtx): Resolved {
  const pres = el.getAttributeNS(PRES_NS, "style-name");
  if (pres && ctx.sheet.has("presentation", pres)) return ctx.sheet.resolve("presentation", pres);
  return ctx.sheet.resolve("graphic", el.getAttributeNS(DRAW_NS, "style-name") || undefined, { withDefault: true });
}

function readFrame(el: Element, ctx: ShapeReadCtx): Shape | undefined {
  const box = readBox(el);
  const cls = attr(el, "class");
  const st = styleOf(el, ctx);
  const tb = kids(el, "text-box")[0];
  const images = kids(el, "image");
  const table = kids(el, "table")[0];
  const base = baseShape(el, box);
  if (table) return readTable(table, base, ctx);
  if (ctx.object && (kids(el, "object")[0] || kids(el, "object-ole")[0])) {
    const o = ctx.object(el, base);
    if (o) return o;
  }
  if (images.length && !tb) {
    for (const im of images) {
      const src = ctx.image(attr(im, "href"), attr(im, "mime-type"), kids(im, "binary-data")[0]?.textContent ?? undefined);
      if (!src) continue;
      const pic: PicShape = { ...base, kind: "pic", src, line: readLineProps(st.graphic, ctx.sheet, "none"), name: base.name === "Şekil" ? "Resim" : base.name };
      const mirror = st.graphic["mirror"] ?? "";
      if (mirror.includes("horizontal")) pic.flipH = true;
      if (mirror.includes("vertical")) pic.flipV = true;
      if (cls === "graphic") pic.ph = { type: "pic" };
      return pic;
    }
    return undefined;
  }
  if (!tb) return undefined;
  const phType = cls ? PH_CLASS[cls] : undefined;
  if (cls && !phType) return undefined;
  const isPh = !!phType;
  const own = isPh ? ctx.sheet.resolve("presentation", el.getAttributeNS(PRES_NS, "style-name") || undefined, { stopAtCommon: true }) : EMPTY;
  const text: TextBody = { ...bodyProps(st.graphic, true), paragraphs: textOf(tb, ctx, cls === "outline") };
  const levels = isPh && ctx.placeholderLevels ? ctx.placeholderLevels(cls!, own) : baseLevels(st, ctx.sheet);
  if (levels) text.levels = levels;
  const sp: SpShape = {
    ...base,
    kind: "sp",
    geom: "rect",
    fill: readFillProps(st.graphic, ctx.sheet, ctx, "none"),
    line: readLineProps(st.graphic, ctx.sheet, "none"),
    text,
  };
  if (isPh) {
    sp.ph = { type: phType! };
    if (base.name === "Şekil") sp.name = cls === "title" ? "Başlık" : cls === "subtitle" ? "Alt Başlık" : "İçerik Yer Tutucusu";
  } else {
    sp.txBox = true;
    if (base.name === "Şekil") sp.name = "Metin Kutusu";
  }
  return sp;
}

/** draw:enhanced-path (yalnız sayısal) → yollar; formül içeriyorsa undefined. */
function enhancedPath(d: string | undefined, vb: number[]): CustomPath[] | undefined {
  if (!d || /[?$]/.test(d)) return undefined;
  const tk = d.match(/[A-Za-z]|[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?/g);
  if (!tk) return undefined;
  const paths: CustomPath[] = [];
  let cur: PathCmd[] = [];
  let fill = true;
  let stroke = true;
  let i = 0;
  let cmd = "";
  const num = () => Number(tk[i++]);
  const isNum = () => i < tk.length && !/^[A-Za-z]$/.test(tk[i]);
  const flush = () => {
    if (cur.length) paths.push({ w: vb[2] || 1, h: vb[3] || 1, cmds: cur.map((c) => ({ ...c, pts: c.pts.map((v, k) => v - (k % 2 === 0 ? vb[0] : vb[1])) })), fill, stroke });
    cur = [];
    fill = true;
    stroke = true;
  };
  let guard = 0;
  while (i < tk.length && guard++ < 100000) {
    if (/^[A-Za-z]$/.test(tk[i])) cmd = tk[i++];
    switch (cmd) {
      case "M":
        cur.push({ c: "M", pts: [num(), num()] });
        cmd = "L";
        break;
      case "L":
        cur.push({ c: "L", pts: [num(), num()] });
        break;
      case "C":
        cur.push({ c: "C", pts: [num(), num(), num(), num(), num(), num()] });
        break;
      case "Q":
        cur.push({ c: "Q", pts: [num(), num(), num(), num()] });
        break;
      case "Z":
        cur.push({ c: "Z", pts: [] });
        while (isNum()) i++;
        cmd = "";
        break;
      case "N":
        flush();
        cmd = "";
        break;
      case "F":
        fill = false;
        cmd = "";
        break;
      case "S":
        stroke = false;
        cmd = "";
        break;
      default:
        // Yay ve açı komutları desteklenmiyor: hazır şekle düş.
        return undefined;
    }
    if (cur.some((c) => c.pts.some((v) => !Number.isFinite(v)))) return undefined;
  }
  flush();
  return paths.length ? paths : undefined;
}

function readDrawShape(el: Element, ctx: ShapeReadCtx): Shape | undefined {
  const box = readBox(el);
  const st = styleOf(el, ctx);
  const base = baseShape(el, box);
  const g = st.graphic;
  let geom = "rect";
  let paths: CustomPath[] | undefined;
  let adj: Record<string, number> | undefined;
  let flipH = false;
  let flipV = false;
  const ln = el.localName;
  if (ln === "ellipse" || ln === "circle") geom = "ellipse";
  else if (ln === "rect") geom = lengthPx(attr(el, "corner-radius")) ? "roundRect" : "rect";
  else if (ln === "custom-shape") {
    const eg = kids(el, "enhanced-geometry")[0];
    const t = attr(eg, "type");
    const known = geomFromOdfType(t);
    flipH = attr(eg, "mirror-horizontal") === "true";
    flipV = attr(eg, "mirror-vertical") === "true";
    if (known) {
      geom = known;
      if (t?.startsWith("ooxml-")) {
        const mods = (attr(eg, "modifiers") ?? "").trim().split(/\s+/).filter(Boolean).map(Number);
        if (mods.length && mods.every(Number.isFinite)) {
          adj = {};
          const keys = mods.length === 1 ? ["adj"] : mods.map((_, k) => `adj${k + 1}`);
          keys.forEach((k, idx) => (adj![k] = mods[idx]));
        }
      }
    } else {
      let vb = (attr(eg, "viewBox") ?? "0 0 21600 21600").split(/[\s,]+/).map(Number);
      if (vb.length !== 4 || !vb.every(Number.isFinite) || !(vb[2] > 0 && vb[3] > 0)) {
        // LibreOffice OOXML şekilleri: yol birimi "sub-view-size" ile verilir.
        const sv = (attr(eg, "sub-view-size") ?? "").split(/[\s,]+/).map(Number);
        vb = sv.length >= 2 && sv[0] > 0 && sv[1] > 0 ? [0, 0, sv[0], sv[1]] : [0, 0, 21600, 21600];
      }
      paths = enhancedPath(attr(eg, "enhanced-path"), vb);
      geom = paths ? "custom" : "rect";
    }
  } else if (ln === "polygon" || ln === "polyline") {
    const vb = (attr(el, "viewBox") ?? `0 0 ${box.w} ${box.h}`).split(/[\s,]+/).map(Number);
    const nums = (attr(el, "points") ?? "").split(/[\s,]+/).filter(Boolean).map(Number);
    if (nums.length >= 4 && nums.every(Number.isFinite) && vb.every(Number.isFinite)) {
      const cmds: PathCmd[] = [];
      for (let k = 0; k + 1 < nums.length && k < 20000; k += 2) cmds.push({ c: k ? "L" : "M", pts: [nums[k] - vb[0], nums[k + 1] - vb[1]] });
      if (ln === "polygon") cmds.push({ c: "Z", pts: [] });
      paths = [{ w: vb[2] || 1, h: vb[3] || 1, cmds, fill: ln === "polygon", stroke: true }];
      geom = "custom";
    }
  } else if (ln === "path") {
    const vb = (attr(el, "viewBox") ?? `0 0 ${box.w} ${box.h}`).split(/[\s,]+/).map(Number);
    const d = attr(el, "d");
    if (d && d.length < 200000 && vb.every(Number.isFinite)) {
      try {
        const cmds = svgPathToCmds(d).map((c) => ({ ...c, pts: c.pts.map((v, k) => v - (k % 2 === 0 ? vb[0] : vb[1])) }));
        if (cmds.length) {
          paths = [{ w: vb[2] || 1, h: vb[3] || 1, cmds, fill: cmds.some((c) => c.c === "Z"), stroke: true }];
          geom = "custom";
        }
      } catch {
        // Okunamayan yol dikdörtgen olur.
      }
    }
  }
  const open = paths?.every((p) => p.fill === false) ?? false;
  const paras = textOf(el, ctx, false);
  const hasText = paras.some((p) => p.runs.some((r) => r.text.trim()));
  const sp: SpShape = {
    ...base,
    kind: "sp",
    geom,
    fill: open ? { type: "none" } : readFillProps(g, ctx.sheet, ctx, "solid"),
    line: readLineProps(g, ctx.sheet, "solid"),
  };
  if (adj) sp.adj = adj;
  if (paths) sp.paths = paths;
  if (flipH) sp.flipH = true;
  if (flipV) sp.flipV = true;
  if (hasText) {
    sp.text = { ...bodyProps(g, false), paragraphs: paras };
    let levels = baseLevels(st, ctx.sheet);
    // Metin alanı yatayda ortalı/sağa yaslı ise paragraf hizası oradan gelir.
    const ha = g["textarea-horizontal-align"];
    const align = ha === "center" ? "ctr" : ha === "right" ? "r" : undefined;
    if (align) levels = (levels ?? Array.from({ length: 9 }, () => ({ para: {}, run: {} }))).map((l) => ({ ...l, para: { align, ...l.para } }));
    if (levels) sp.text.levels = levels;
  }
  readShadow(g, sp);
  return sp;
}

function readShadow(g: Props, sp: SpShape) {
  if (g["shadow"] !== "visible") return;
  const dx = lengthPx(g["shadow-offset-x"]) ?? 3;
  const dy = lengthPx(g["shadow-offset-y"]) ?? 3;
  const color = odfColor(g["shadow-color"]) ?? { rgb: "#808080" };
  const op = percent(g["shadow-opacity"]);
  sp.effects = { ...(sp.effects ?? {}), shadow: { color: op !== undefined ? { ...color, alpha: op } : color, blur: lengthPx(g["shadow-blur"]) ?? 0, dist: Math.hypot(dx, dy), dir: (Math.atan2(dy, dx) * 180) / Math.PI } };
}

function readLine(el: Element, ctx: ShapeReadCtx): Shape | undefined {
  const st = styleOf(el, ctx);
  let x1 = lengthPx(attr(el, "x1")) ?? 0;
  let y1 = lengthPx(attr(el, "y1")) ?? 0;
  let x2 = lengthPx(attr(el, "x2")) ?? 0;
  let y2 = lengthPx(attr(el, "y2")) ?? 0;
  const m = parseTransform(attr(el, "transform"));
  if (m) {
    [x1, y1] = [m[0] * x1 + m[2] * y1 + m[4], m[1] * x1 + m[3] * y1 + m[5]];
    [x2, y2] = [m[0] * x2 + m[2] * y2 + m[4], m[1] * x2 + m[3] * y2 + m[5]];
  }
  const base = baseShape(el, { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1), rot: 0 });
  const sp: SpShape = {
    ...base,
    kind: "sp",
    name: base.name === "Şekil" ? "Düz Bağlayıcı" : base.name,
    geom: el.localName === "connector" ? "straightConnector1" : "line",
    fill: { type: "none" },
    line: readLineProps(st.graphic, ctx.sheet, "solid"),
  };
  if (x2 < x1) sp.flipH = true;
  if (y2 < y1) sp.flipV = true;
  if (el.localName === "connector") sp.connector = true;
  return sp;
}

function readGroup(el: Element, ctx: ShapeReadCtx): Shape | undefined {
  const children = readShapes(el, ctx);
  if (!children.length) return undefined;
  if (children.length === 1) return children[0];
  const x = Math.min(...children.map((c) => c.x));
  const y = Math.min(...children.map((c) => c.y));
  const r = Math.max(...children.map((c) => c.x + c.w));
  const b = Math.max(...children.map((c) => c.y + c.h));
  const g: GroupShape = { id: uid(), kind: "group", name: attr(el, "name") || "Grup", x, y, w: r - x, h: b - y, rot: 0, children };
  return g;
}
