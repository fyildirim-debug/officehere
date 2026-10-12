/** DrawingML/PresentationML şekil okuyucusu (geometri, şekil, resim, tablo, grup). */
import { uid } from "../model/factory";
import type {
  CellBorder,
  ChartShape,
  MediaInfo,
  CustomPath,
  Effects,
  Fill,
  GroupShape,
  Line,
  PathCmd,
  PicShape,
  Placeholder,
  Shape,
  SpShape,
  TableCell,
  TableShape,
} from "../model/types";
import { isLineGeom } from "../render/geometry";
import { tableStyleByGuid, DEFAULT_TABLE_STYLE } from "../render/tableStyles";
import { attr, boolAttr, descendants, fromEmu, kid, kidPath, kids, numAttr, relId } from "./xml";

import { type ReadCtx, readColor, mediaFor, readFill, readLine, readEffects } from "./readStyle";
import { readHlink, readTextBody } from "./readText";

export * from "./readStyle";
export * from "./readText";

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

function nvProps(nv: Element | undefined, ctx: ReadCtx): { id: string; name: string; descr?: string; hidden?: boolean; ph?: Placeholder; txBox?: boolean; link?: string } {
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
    link: readHlink(kid(c, "hlinkClick"), ctx),
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
  // effectRef: tema efekt stili (effectStyleLst); tema okunamadıysa 3. stil gölge sayılır.
  const fxIdx = numAttr(fx, "idx") ?? 0;
  const themeFx = ctx.effectStyles ? ctx.effectStyles[fxIdx - 1] : fxIdx >= 3 ? { shadow: { color: { rgb: "#000000", alpha: 0.4 }, blur: 4, dist: 2, dir: 90 } } : undefined;
  const effects: Effects | undefined = fxIdx > 0 && themeFx ? structuredClone(themeFx) : undefined;
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
    // Açık (boş da olsa) effectLst stil başvurusunu geçersiz kılar.
    effects: kid(spPr, "effectLst") ? readEffects(spPr, ctx) : (st.effects ?? inh?.effects),
    text,
    ph: info.ph,
    txBox: info.txBox,
    connector: connector || undefined,
    link: info.link,
  };
  if (!sp.text && !connector && !isLineGeom(sp.geom) && !info.ph) sp.text = undefined;
  // Çizim sekmesiyle oluşturulan mürekkep (silgiyle silinebilmesi için).
  if (sp.geom === "custom" && /^Mürekkep \d+$/.test(sp.name ?? "") && sp.fill.type === "none") sp.ink = true;
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
  // Video/ses (a:videoFile / a:audioFile, p14:media).
  const nvPr = kidPath(el, "nvPicPr", "nvPr");
  const vf = kid(nvPr, "videoFile") ?? kid(nvPr, "audioFile");
  let media: MediaInfo | undefined;
  if (vf) {
    const embed = descendants(nvPr, "media")[0];
    const rel = ctx.rels.get(relId(embed) ?? "") ?? ctx.rels.get(attr(vf, "link") ?? "");
    const data = rel && !rel.external ? ctx.media.get(rel.target) : undefined;
    if (data) {
      media = { kind: vf.localName === "videoFile" ? "video" : "audio", src: data, mime: /^data:([^;,]+)/.exec(data)?.[1] ?? "" };
      const opt = descendants(nvPr, "media").find((m) => m !== embed);
      if (opt && boolAttr(opt, "autoPlay")) media.autoPlay = true;
      if (opt && boolAttr(opt, "loop")) media.loop = true;
    }
  }
  return {
    kind: "pic",
    media,
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
    link: info.link,
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
    link: info.link,
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
      const chartRef = descendants(el, "chart")[0];
      const chartRel = chartRef ? ctx.rels.get(relId(chartRef) ?? "") : undefined;
      const chart = chartRel && !chartRel.external ? ctx.charts?.get(chartRel.target) : undefined;
      if (chart) {
        const info = nvProps(kid(el, "nvGraphicFramePr"), ctx);
        const xf = readXfrm(kid(el, "xfrm"));
        return { kind: "chart", id: info.id, name: info.name || "Grafik", descr: info.descr, hidden: info.hidden, link: info.link, x: xf?.x ?? 0, y: xf?.y ?? 0, w: xf?.w ?? 0, h: xf?.h ?? 0, rot: xf?.rot ?? 0, chart: structuredClone(chart) } satisfies ChartShape;
      }
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
