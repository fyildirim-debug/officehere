/**
 * MS-ODRAW çizim okuyucusu (PPT slaytlarındaki PPDrawing): şekil kapları,
 * şekil türü (MSOSPT), konum (ClientAnchor 576/inç; grup içinde ChildAnchor),
 * özellik tablosu (dolgu, çizgi, döndürme, resim), yer tutucular ve metin.
 */
import { uid } from "../../model/factory";
import type { Color, CustomPath, Fill, GroupShape, Line, PathCmd, PicShape, Placeholder, Shape, SpShape, TextBody } from "../../model/types";
import { RT, type Rec, type Stream } from "./records";
import { hex, itemParagraphs, textItems, type TextEnv, type TextItem } from "./text";

const MAX_SHAPES = 3000;

/** MSOSPT → model hazır şekli. */
const SPT: Record<number, string> = {
  1: "rect",
  2: "roundRect",
  3: "ellipse",
  4: "diamond",
  5: "triangle",
  6: "rtTriangle",
  7: "parallelogram",
  8: "trapezoid",
  9: "hexagon",
  10: "octagon",
  11: "plus",
  12: "star5",
  13: "rightArrow",
  15: "homePlate",
  16: "cube",
  19: "arc",
  20: "line",
  22: "can",
  23: "donut",
  32: "straightConnector1",
  33: "bentConnector3",
  34: "bentConnector3",
  35: "bentConnector3",
  36: "bentConnector3",
  37: "curvedConnector3",
  38: "curvedConnector3",
  39: "curvedConnector3",
  40: "curvedConnector3",
  55: "chevron",
  56: "pentagon",
  57: "noSmoking",
  58: "star8",
  59: "star12",
  60: "star12",
  61: "wedgeRectCallout",
  62: "wedgeRoundRectCallout",
  63: "wedgeEllipseCallout",
  64: "wave",
  65: "foldedCorner",
  66: "leftArrow",
  67: "downArrow",
  68: "upArrow",
  69: "leftRightArrow",
  70: "upDownArrow",
  71: "irregularSeal1",
  73: "lightningBolt",
  74: "heart",
  76: "quadArrow",
  84: "bevel",
  92: "star12",
  96: "smileyFace",
  106: "cloudCallout",
  109: "flowChartProcess",
  110: "flowChartDecision",
  111: "flowChartInputOutput",
  112: "flowChartPredefinedProcess",
  114: "flowChartDocument",
  116: "flowChartTerminator",
  120: "flowChartConnector",
  176: "roundRect",
  183: "sun",
  184: "moon",
  187: "star4",
  189: "actionButtonBlank",
  190: "actionButtonHome",
  191: "actionButtonHelp",
  192: "actionButtonInformation",
  193: "actionButtonForwardNext",
  194: "actionButtonBackPrevious",
  195: "actionButtonEnd",
  196: "actionButtonBeginning",
  197: "actionButtonReturn",
  198: "actionButtonDocument",
  202: "rect",
};

const LINE_TYPES = new Set([20, 32, 33, 34, 35, 36, 37, 38, 39, 40]);

/** PlaceholderEnum → model yer tutucusu. */
const PH: Record<number, Placeholder["type"]> = {
  1: "title",
  2: "body",
  3: "ctrTitle",
  4: "subTitle",
  7: "dt",
  8: "sldNum",
  9: "ftr",
  13: "title",
  14: "body",
  15: "ctrTitle",
  16: "subTitle",
  17: "title",
  18: "body",
  19: "obj",
  20: "obj",
  21: "tbl",
  22: "pic",
  23: "obj",
  24: "obj",
  25: "obj",
  26: "pic",
};

export interface Props {
  get(pid: number): number | undefined;
  complex(pid: number): Uint8Array | undefined;
}

/** OfficeArtFOPT / TertiaryFOPT özellik tabloları. */
export function readProps(s: Stream, recs: (Rec | undefined)[], base?: Map<number, number>): Props {
  const vals = new Map<number, number>(base);
  const cx = new Map<number, Uint8Array>();
  for (const r of recs) {
    if (!r) continue;
    const n = Math.min(r.inst, Math.floor(r.len / 6));
    let co = r.start + n * 6;
    const end = r.start + r.len;
    for (let i = 0; i < n; i++) {
      const o = r.start + i * 6;
      const opid = s.u16(o);
      const pid = opid & 0x3fff;
      const v = s.u32(o + 2);
      vals.set(pid, v);
      if (opid & 0x8000) {
        const len = Math.min(v, Math.max(0, end - co));
        cx.set(pid, s.bytes.subarray(co, co + len));
        co += len;
      }
    }
  }
  return { get: (p) => vals.get(p), complex: (p) => cx.get(p) };
}

export interface DrawCtx extends TextEnv {
  /** Blip dizini (1 tabanlı) → data: URL. */
  blip: (pib: number) => string | undefined;
  /** Belge varsayılan şekil özellikleri (DggContainer FOPT). */
  defaults: Map<number, number>;
  /** OutlineTextRefAtom için slaydın anahat metinleri. */
  outline: TextItem[];
  /** Metin türüne göre düzey stilleri (asıl slayt). */
  levels: (textType: number) => TextBody["levels"];
  count: { n: number };
}

/** OfficeArtCOLORREF → renk. */
function colorRef(v: number | undefined, scheme: string[]): Color | undefined {
  if (v === undefined) return undefined;
  const r = v & 0xff;
  const g = (v >>> 8) & 0xff;
  const b = (v >>> 16) & 0xff;
  const f = v >>> 24;
  if (f & 0x08) return scheme[r] ? { rgb: scheme[r] } : undefined;
  if (f & 0x10) return undefined;
  return { rgb: hex(r, g, b) };
}

const fixed = (v: number | undefined) => (v === undefined ? undefined : (v | 0) / 65536);

function readFill(p: Props, ctx: DrawCtx, explicitOnly: boolean): Fill {
  const bools = p.get(0x1bf);
  let filled = !explicitOnly;
  if (bools !== undefined && bools & 0x100000) filled = !!(bools & 0x10);
  if (!filled) return { type: "none" };
  const type = p.get(0x180) ?? 0;
  const op = fixed(p.get(0x182));
  const alpha = op !== undefined && op < 1 ? Math.max(0, op) : undefined;
  const c = colorRef(p.get(0x181), ctx.scheme) ?? { rgb: "#FFFFFF" };
  if (type === 3 || type === 2) {
    const src = ctx.blip(p.get(0x186) ?? 0);
    if (src) return { type: "image", src, mode: type === 2 ? "tile" : "stretch" };
  }
  if (type >= 4 && type <= 8) {
    const c2 = colorRef(p.get(0x183), ctx.scheme) ?? { rgb: "#FFFFFF" };
    const ang = fixed(p.get(0x18b)) ?? 0;
    return { type: "gradient", kind: type === 4 ? "linear" : "radial", angle: (((90 - ang) % 360) + 360) % 360, stops: [{ pos: 0, color: c }, { pos: 1, color: c2 }] };
  }
  if (type === 9) return { type: "none" };
  return { type: "solid", color: alpha !== undefined ? { ...c, alpha } : c };
}

const DASH: Line["dash"][] = ["solid", "dash", "sysDot", "dashDot", "lgDashDotDot", "sysDot", "dash", "lgDash", "dashDot", "lgDashDot", "lgDashDotDot"];
const ARROW: Line["head"][] = [undefined, "triangle", "stealth", "diamond", "oval", "arrow"];

function readLine(p: Props, ctx: DrawCtx, explicitOnly: boolean): Line {
  const bools = p.get(0x1ff);
  let on = !explicitOnly;
  if (bools !== undefined && bools & 0x80000) on = !!(bools & 0x8);
  if (!on) return { fill: { type: "none" }, width: 1, dash: "solid" };
  const c = colorRef(p.get(0x1c0), ctx.scheme) ?? { rgb: "#000000" };
  const op = fixed(p.get(0x1c1));
  const w = p.get(0x1cb);
  const line: Line = { fill: { type: "solid", color: op !== undefined && op < 1 ? { ...c, alpha: op } : c }, width: w !== undefined ? Math.max(0.25, Math.min(200, w / 12700)) : 0.75, dash: DASH[p.get(0x1ce) ?? 0] ?? "solid" };
  const h = ARROW[p.get(0x1d0) ?? 0];
  const t = ARROW[p.get(0x1d1) ?? 0];
  if (h) line.head = h;
  if (t) line.tail = t;
  return line;
}

function utf16z(b: Uint8Array | undefined): string | undefined {
  if (!b || b.length < 2) return undefined;
  let s = "";
  for (let i = 0; i + 1 < b.length; i += 2) {
    const c = b[i] | (b[i + 1] << 8);
    if (!c) break;
    s += String.fromCharCode(c);
  }
  return s.trim() || undefined;
}

/** IMsoArray (karmaşık özellik) öğeleri. */
function msoArray(b: Uint8Array | undefined): { n: number; size: number; dv: DataView } | undefined {
  if (!b || b.length < 6) return undefined;
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  const n = dv.getUint16(0, true);
  let size = dv.getUint16(4, true);
  if (size === 0xfff0) size = 4;
  if (!size || 6 + n * size > b.length) return undefined;
  return { n: Math.min(n, 20000), size, dv };
}

/** Serbest biçim (pVertices + pSegmentInfo) → özel yol. */
function freeform(p: Props): CustomPath[] | undefined {
  const v = msoArray(p.complex(0x145));
  if (!v) return undefined;
  const pts: [number, number][] = [];
  for (let i = 0; i < v.n; i++) {
    const o = 6 + i * v.size;
    pts.push(v.size >= 8 ? [v.dv.getInt32(o, true), v.dv.getInt32(o + 4, true)] : [v.dv.getInt16(o, true), v.dv.getInt16(o + 2, true)]);
  }
  // Formül (kılavuz) başvurulu köşeler desteklenmiyor: hazır şekle düşülür.
  const bad = v.size >= 8 ? (n: number) => Math.abs(n) > 1e7 : (n: number) => n <= -32513;
  if (pts.some(([x, y]) => bad(x) || bad(y))) return undefined;
  const l = p.get(0x140) ?? 0;
  const t = p.get(0x141) ?? 0;
  const w = (p.get(0x142) ?? 21600) - l || 21600;
  const h = (p.get(0x143) ?? 21600) - t || 21600;
  const at = (k: number) => [(pts[k]?.[0] ?? 0) - l, (pts[k]?.[1] ?? 0) - t];
  const seg = msoArray(p.complex(0x146));
  const cmds: PathCmd[] = [];
  let k = 0;
  if (seg && seg.size === 2) {
    for (let i = 0; i < seg.n && k <= pts.length; i++) {
      const info = seg.dv.getUint16(6 + i * 2, true);
      const type = info >> 13;
      const count = info & 0x1fff;
      if (type === 0) for (let c = 0; c < Math.max(1, count) && k < pts.length; c++) cmds.push({ c: "L", pts: at(k++) });
      else if (type === 1) for (let c = 0; c < Math.max(1, count) && k + 2 < pts.length; c++) cmds.push({ c: "C", pts: [...at(k++), ...at(k++), ...at(k++)] });
      else if (type === 2) cmds.push({ c: "M", pts: at(k++) });
      else if (type === 3) cmds.push({ c: "Z", pts: [] });
      else if (type === 4) continue;
      else k += info & 0xff;
    }
  } else {
    pts.forEach((_, i) => cmds.push({ c: i ? "L" : "M", pts: at(i) }));
  }
  if (!cmds.length || cmds[0].c !== "M") cmds.unshift({ c: "M", pts: at(0) });
  return [{ w, h, cmds, fill: cmds.some((c) => c.c === "Z"), stroke: true }];
}

interface Xf {
  /** Grup koordinatı → slayt pikseli. */
  map(x: number, y: number): [number, number];
  sx: number;
  sy: number;
}

const IDENT: Xf = { map: (x, y) => [x / 6, y / 6], sx: 1 / 6, sy: 1 / 6 };

function anchorOf(s: Stream, sp: Rec, xf: Xf, top: boolean): { x: number; y: number; w: number; h: number } | undefined {
  const ca = s.kid(sp, RT.ClientAnchor);
  if (ca && top) {
    const small = ca.len < 16;
    const t = small ? s.i16(ca.start) : s.i32(ca.start);
    const l = small ? s.i16(ca.start + 2) : s.i32(ca.start + 4);
    const r = small ? s.i16(ca.start + 4) : s.i32(ca.start + 8);
    const b = small ? s.i16(ca.start + 6) : s.i32(ca.start + 12);
    return { x: l / 6, y: t / 6, w: Math.max(0, r - l) / 6, h: Math.max(0, b - t) / 6 };
  }
  const ch = s.kid(sp, RT.ChildAnchor);
  if (ch) {
    const [x1, y1] = xf.map(s.i32(ch.start), s.i32(ch.start + 4));
    const [x2, y2] = xf.map(s.i32(ch.start + 8), s.i32(ch.start + 12));
    return { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1) };
  }
  return undefined;
}

export interface ReadShape {
  shape?: Shape;
  /** Arka plan şekli (fBackground). */
  background?: Fill;
  /** Not metni ya da yer tutucu sırası için. */
  ph?: number;
}

/** Tek şekil kabı (OfficeArtSpContainer). */
function readSp(s: Stream, sp: Rec, ctx: DrawCtx, xf: Xf, top: boolean): ReadShape {
  const fsp = s.kid(sp, RT.FSP);
  if (!fsp) return {};
  const spt = fsp.inst;
  const flags = s.u32(fsp.start + 4);
  if (flags & 0x8) return {};
  const p = readProps(s, [s.kid(sp, RT.FOPT), s.kid(sp, RT.TertiaryFOPT)], ctx.defaults);
  if (flags & 0x400) return { background: readFill(p, ctx, false) };
  if (flags & 0x5 && !(flags & 0x200)) return {};
  const box = anchorOf(s, sp, xf, top);
  // Yer tutucu ve metin.
  const cd = s.kid(sp, RT.ClientData);
  const pa = cd ? s.kid(cd, RT.PlaceholderAtom) : undefined;
  const phId = pa ? s.u8(pa.start + 4) : 0;
  const tbx = s.kid(sp, RT.ClientTextbox);
  let item: TextItem | undefined;
  if (tbx) {
    const kids = s.kids(tbx);
    const ref = kids.find((k) => k.type === RT.OutlineTextRefAtom);
    item = ref ? ctx.outline[s.i32(ref.start)] : textItems(s, kids)[0];
  }
  if (!box) return { ph: phId };
  let { x, y, w, h } = box;
  // Döndürme (16.16 derece, saat yönü); 45–135 ve 225–315 aralığında çapa 90° döndürülmüş saklanır.
  let rot = (((fixed(p.get(0x4)) ?? 0) % 360) + 360) % 360;
  if ((rot >= 45 && rot < 135) || (rot >= 225 && rot < 315)) {
    const cx = x + w / 2;
    const cy = y + h / 2;
    [w, h] = [h, w];
    x = cx - w / 2;
    y = cy - h / 2;
  }
  rot = Math.round(rot * 1000) / 1000;
  const base = { id: uid(), name: utf16z(p.complex(0x380)) ?? "", x, y, w, h, rot, flipH: !!(flags & 0x40) || undefined, flipV: !!(flags & 0x80) || undefined };
  const descr = utf16z(p.complex(0x381));
  const pib = p.get(0x104);
  if (spt === 75 || (pib && !tbx)) {
    const src = pib ? ctx.blip(pib) : undefined;
    if (src) {
      const pic: PicShape = { ...base, kind: "pic", name: base.name || "Resim", src, line: readLine(p, ctx, true), ...(descr ? { descr } : {}) };
      const crop = { t: fixed(p.get(0x100)) ?? 0, b: fixed(p.get(0x101)) ?? 0, l: fixed(p.get(0x102)) ?? 0, r: fixed(p.get(0x103)) ?? 0 };
      if (crop.t || crop.b || crop.l || crop.r) pic.crop = crop;
      const ph = PH[phId];
      if (ph === "pic") pic.ph = { type: "pic" };
      return { shape: pic, ph: phId };
    }
    if (spt === 75) return { ph: phId };
  }
  const textBox = spt === 202 || !!PH[phId];
  const paths = spt === 0 || spt === 0xfff ? freeform(p) : undefined;
  const geom = paths ? "custom" : (SPT[spt] ?? "rect");
  const sp2: SpShape = {
    ...base,
    kind: "sp",
    name: base.name || (spt === 202 ? "Metin Kutusu" : "Şekil"),
    geom,
    fill: LINE_TYPES.has(spt) ? { type: "none" } : readFill(p, ctx, textBox),
    line: readLine(p, ctx, textBox && !LINE_TYPES.has(spt)),
  };
  if (descr) sp2.descr = descr;
  if (paths) {
    sp2.paths = paths;
    if (!paths[0].fill) sp2.fill = { type: "none" };
  }
  if (LINE_TYPES.has(spt) && spt !== 20) sp2.connector = true;
  const adj = p.get(0x147);
  if (adj !== undefined && geom === "roundRect") sp2.adj = { adj: Math.round(((adj | 0) / 21600) * 100000 * 2) / 2 };
  if (item) {
    const anchor = p.get(0x87) ?? (textBox ? 0 : 1);
    const ml = p.get(0x81);
    const mt = p.get(0x82);
    const mr = p.get(0x83);
    const mb = p.get(0x84);
    const text: TextBody = {
      paragraphs: itemParagraphs(item, ctx),
      anchor: anchor % 3 === 1 ? "ctr" : anchor % 3 === 2 ? "b" : "t",
      wrap: p.get(0x85) !== 2,
      autofit: "none",
      insets: { l: (ml ?? 91440) / 9525, t: (mt ?? 45720) / 9525, r: (mr ?? 91440) / 9525, b: (mb ?? 45720) / 9525 },
    };
    const levels = ctx.levels(item.type);
    if (levels) text.levels = levels;
    sp2.text = text;
  }
  const ph = PH[phId];
  if (ph) sp2.ph = { type: ph };
  else if (spt === 202) sp2.txBox = true;
  return { shape: sp2, ph: phId };
}

/** Grup kabı (OfficeArtSpgrContainer): ilk şekil grubun kendisidir. */
function readGroup(s: Stream, gr: Rec, ctx: DrawCtx, xf: Xf, top: boolean, out: ReadShape[]): Shape | undefined {
  const kids = s.kids(gr);
  const head = kids[0];
  if (!head || head.type !== RT.SpContainer) return undefined;
  const fspgr = s.kid(head, RT.FSPGR);
  const fsp = s.kid(head, RT.FSP);
  const patriarch = !!fsp && !!(s.u32(fsp.start + 4) & 0x4);
  let inner = xf;
  let box: ReturnType<typeof anchorOf>;
  if (!patriarch) {
    box = anchorOf(s, head, xf, top);
    if (fspgr && box) {
      const l = s.i32(fspgr.start);
      const t = s.i32(fspgr.start + 4);
      const r = s.i32(fspgr.start + 8);
      const b = s.i32(fspgr.start + 12);
      const sx = r !== l ? box.w / (r - l) : 0;
      const sy = b !== t ? box.h / (b - t) : 0;
      const bx = box;
      inner = { map: (x, y) => [bx.x + (x - l) * sx, bx.y + (y - t) * sy], sx, sy };
    }
  }
  const children: Shape[] = [];
  for (const k of kids.slice(1)) {
    if (ctx.count.n >= MAX_SHAPES) break;
    if (k.type === RT.SpContainer) {
      const r = readSp(s, k, ctx, inner, patriarch);
      out.push(r);
      if (r.shape) {
        ctx.count.n++;
        children.push(r.shape);
      }
    } else if (k.type === RT.SpgrContainer) {
      const g = readGroup(s, k, ctx, inner, patriarch, out);
      if (g) children.push(g);
    }
  }
  if (patriarch) {
    // Kök grup: şekiller doğrudan listeye.
    const g: GroupShape = { id: "", kind: "group", name: "", x: 0, y: 0, w: 0, h: 0, rot: 0, children };
    return g;
  }
  if (!children.length) return undefined;
  if (children.length === 1) return children[0];
  const x = Math.min(...children.map((c) => c.x));
  const y = Math.min(...children.map((c) => c.y));
  const r = Math.max(...children.map((c) => c.x + c.w));
  const b = Math.max(...children.map((c) => c.y + c.h));
  return { id: uid(), kind: "group", name: "Grup", x, y, w: r - x, h: b - y, rot: 0, children };
}

export interface DrawingResult {
  shapes: Shape[];
  background?: Fill;
  /** Okunan tüm şekiller (yer tutucu numaralarıyla). */
  all: ReadShape[];
}

/** PPDrawing (RT_Drawing) kaydındaki şekiller. */
export function readDrawing(s: Stream, drawing: Rec | undefined, ctx: DrawCtx): DrawingResult {
  const res: DrawingResult = { shapes: [], all: [] };
  const dg = s.kid(drawing, RT.DgContainer);
  if (!dg) return res;
  for (const k of s.kids(dg)) {
    if (k.type === RT.SpgrContainer) {
      const root = readGroup(s, k, ctx, IDENT, true, res.all);
      if (root?.kind === "group" && !root.id) res.shapes.push(...root.children);
      else if (root) res.shapes.push(root);
    } else if (k.type === RT.SpContainer) {
      const r = readSp(s, k, ctx, IDENT, true);
      if (r.background) res.background = r.background;
    }
  }
  for (const r of res.all) if (r.background) res.background = r.background;
  return res;
}
