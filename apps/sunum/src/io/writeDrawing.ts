/** DrawingML/PresentationML şekil yazıcısı (sp, cxnSp, pic, graphicFrame/tbl, grpSp, txBody). */
import type {
  BulletDef,
  CellBorder,
  ChartData,
  ChartShape,
  Color,
  CustomPath,
  Effects,
  Fill,
  GroupShape,
  LevelStyle,
  Line,
  Paragraph,
  ParaProps,
  PicShape,
  RunProps,
  Shape,
  SpShape,
  TableShape,
  TextBody,
} from "../model/types";
import { isLineGeom } from "../render/geometry";
import { tableStyleById } from "../render/tableStyles";
import { attrs, emu, esc, MEDIA_EXT_URI, NS, OH_EXT_URI, REL, type RelsBuilder } from "./xml";

export interface MediaStore {
  /** data URL → paket yolu ("ppt/media/image1.png"). */
  add(dataUrl: string): string | null;
  /** Video/ses dosyası → paket yolu ("ppt/media/media1.mp4"). */
  addMedia?(dataUrl: string): string | null;
}

export interface ChartStore {
  /** Grafik parçasını kaydeder, paket yolunu döndürür ("ppt/charts/chart1.xml"). */
  add(chart: ChartData): string;
}

export interface WriteCtx {
  rels: RelsBuilder;
  media: MediaStore;
  /** Bu parçanın yolu ("ppt/slides/slide1.xml"). */
  part: string;
  nextId: () => number;
  idMap: Map<string, number>;
  /** Slayt içeriği mi (tx1/bg1 adlandırması). */
  slideContent: boolean;
  /** Slayt kimliği → slayt numarası (1 tabanlı), sunu içi köprüler için. */
  slideNumbers?: Map<string, number>;
  charts?: ChartStore;
}

const JUMP_ACTIONS: Record<string, string> = {
  "#next": "nextslide",
  "#prev": "previousslide",
  "#first": "firstslide",
  "#last": "lastslide",
};

/** a:hlinkClick (ECMA-376 §21.1.2.3.5); sunu içi hedefler ppaction ile yazılır. */
export function hlinkXml(link: string | undefined, ctx: WriteCtx, tag = "a:hlinkClick"): string {
  if (!link) return "";
  const jump = JUMP_ACTIONS[link];
  if (jump) return `<${tag} r:id="" action="ppaction://hlinkshowjump?jump=${jump}"/>`;
  if (link.startsWith("#slide:")) {
    const n = ctx.slideNumbers?.get(link.slice(7));
    if (!n) return "";
    const rid = ctx.rels.add(REL.slide, ctx.part.startsWith("ppt/slides/") ? `slide${n}.xml` : `../slides/slide${n}.xml`);
    return `<${tag} r:id="${rid}" action="ppaction://hlinksldjump"/>`;
  }
  if (link.startsWith("#")) return "";
  const rid = ctx.rels.add(REL.hyperlink, link, true);
  return `<${tag} r:id="${rid}"/>`;
}

const SCHEME_OUT: Record<string, string> = { dk1: "tx1", lt1: "bg1", dk2: "tx2", lt2: "bg2" };

const pct = (v: number) => Math.round(v * 100000);

export function colorXml(c: Color, ctx?: { slideContent: boolean }): string {
  const mods =
    (c.tint !== undefined ? `<a:tint val="${pct(c.tint)}"/>` : "") +
    (c.shade !== undefined ? `<a:shade val="${pct(c.shade)}"/>` : "") +
    (c.alpha !== undefined && c.alpha < 1 ? `<a:alpha val="${pct(c.alpha)}"/>` : "") +
    (c.lumMod !== undefined ? `<a:lumMod val="${pct(c.lumMod)}"/>` : "") +
    (c.lumOff !== undefined ? `<a:lumOff val="${pct(c.lumOff)}"/>` : "");
  if (c.scheme) {
    const name = ctx?.slideContent === false ? c.scheme : (SCHEME_OUT[c.scheme] ?? c.scheme);
    return mods ? `<a:schemeClr val="${name}">${mods}</a:schemeClr>` : `<a:schemeClr val="${name}"/>`;
  }
  const hex = (c.rgb ?? "#000000").replace("#", "").toUpperCase().padEnd(6, "0").slice(0, 6);
  return mods ? `<a:srgbClr val="${hex}">${mods}</a:srgbClr>` : `<a:srgbClr val="${hex}"/>`;
}

export function relativeTarget(fromPart: string, target: string): string {
  const from = fromPart.split("/");
  from.pop();
  const to = target.split("/");
  let i = 0;
  while (i < from.length && i < to.length && from[i] === to[i]) i++;
  return [...Array(from.length - i).fill(".."), ...to.slice(i)].join("/");
}

export function imageRel(ctx: WriteCtx, src: string): string | null {
  const path = ctx.media.add(src);
  if (!path) return null;
  return ctx.rels.add(REL.image, relativeTarget(ctx.part, path));
}

export function fillXml(f: Fill | undefined, ctx: WriteCtx): string {
  if (!f || f.type === "none") return "<a:noFill/>";
  if (f.type === "solid") return `<a:solidFill>${colorXml(f.color, ctx)}</a:solidFill>`;
  if (f.type === "gradient") {
    const gs = f.stops.map((s) => `<a:gs pos="${pct(s.pos)}">${colorXml(s.color, ctx)}</a:gs>`).join("");
    const shade = f.kind === "radial" ? '<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>' : `<a:lin ang="${Math.round((((f.angle % 360) + 360) % 360) * 60000)}" scaled="0"/>`;
    return `<a:gradFill rotWithShape="1"><a:gsLst>${gs}</a:gsLst>${shade}</a:gradFill>`;
  }
  const rid = imageRel(ctx, f.src);
  if (!rid) return "<a:noFill/>";
  const alpha = f.alpha !== undefined && f.alpha < 1 ? `<a:alphaModFix amt="${pct(f.alpha)}"/>` : "";
  const mode = f.mode === "tile" ? '<a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/>' : "<a:stretch><a:fillRect/></a:stretch>";
  return `<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="${rid}">${alpha}</a:blip><a:srcRect/>${mode}</a:blipFill>`;
}

export function lineXml(l: Line | undefined, ctx: WriteCtx): string {
  if (!l) return "";
  if (l.fill.type === "none") return "<a:ln><a:noFill/></a:ln>";
  const dash = l.dash && l.dash !== "solid" ? `<a:prstDash val="${l.dash}"/>` : "";
  const head = l.head && l.head !== "none" ? `<a:headEnd type="${l.head}"/>` : "";
  const tail = l.tail && l.tail !== "none" ? `<a:tailEnd type="${l.tail}"/>` : "";
  return `<a:ln w="${Math.round(l.width * 12700)}"${l.cap ? ` cap="${l.cap}"` : ""}><a:solidFill>${colorXml(l.fill.color, ctx)}</a:solidFill>${dash}<a:round/>${head}${tail}</a:ln>`;
}

export function effectsXml(fx: Effects | undefined, ctx: WriteCtx): string {
  if (!fx) return "";
  let s = "";
  if (fx.glow) s += `<a:glow rad="${emu(fx.glow.radius)}">${colorXml(fx.glow.color, ctx)}</a:glow>`;
  if (fx.shadow) {
    const sh = fx.shadow;
    const dir = Math.round((((sh.dir % 360) + 360) % 360) * 60000);
    s += `<a:outerShdw blurRad="${emu(sh.blur)}" dist="${emu(sh.dist)}" dir="${dir}" algn="tl" rotWithShape="0">${colorXml(sh.color, ctx)}</a:outerShdw>`;
  }
  if (fx.reflection) s += '<a:reflection blurRad="6350" stA="50000" endA="300" endPos="55000" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/>';
  if (fx.softEdge) s += `<a:softEdge rad="${emu(fx.softEdge)}"/>`;
  return s ? `<a:effectLst>${s}</a:effectLst>` : "";
}

export function xfrmXml(s: { x: number; y: number; w: number; h: number; rot?: number; flipH?: boolean; flipV?: boolean }, tag = "a:xfrm", child?: { x: number; y: number; w: number; h: number }): string {
  const rot = s.rot ? Math.round((((s.rot % 360) + 360) % 360) * 60000) : undefined;
  const ch = child ? `<a:chOff x="${emu(child.x)}" y="${emu(child.y)}"/><a:chExt cx="${emu(Math.max(0, child.w))}" cy="${emu(Math.max(0, child.h))}"/>` : "";
  return `<${tag}${attrs({ rot, flipH: s.flipH, flipV: s.flipV })}><a:off x="${emu(s.x)}" y="${emu(s.y)}"/><a:ext cx="${emu(Math.max(0, s.w))}" cy="${emu(Math.max(0, s.h))}"/>${ch}</${tag}>`;
}

function customGeomXml(paths: CustomPath[]): string {
  const k = 100;
  const pl = paths
    .map((p) => {
      let body = "";
      for (const c of p.cmds) {
        const pts: string[] = [];
        for (let i = 0; i + 1 < c.pts.length; i += 2) pts.push(`<a:pt x="${Math.round(c.pts[i] * k)}" y="${Math.round(c.pts[i + 1] * k)}"/>`);
        if (c.c === "M") body += `<a:moveTo>${pts[0] ?? '<a:pt x="0" y="0"/>'}</a:moveTo>`;
        else if (c.c === "L") body += `<a:lnTo>${pts[0]}</a:lnTo>`;
        else if (c.c === "C") body += `<a:cubicBezTo>${pts.join("")}</a:cubicBezTo>`;
        else if (c.c === "Q") body += `<a:quadBezTo>${pts.join("")}</a:quadBezTo>`;
        else if (c.c === "Z") body += "<a:close/>";
      }
      return `<a:path${attrs({ w: Math.round(p.w * k), h: Math.round(p.h * k), fill: p.fill === false ? "none" : undefined, stroke: p.stroke === false ? "0" : undefined })}>${body}</a:path>`;
    })
    .join("");
  return `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst>${pl}</a:pathLst></a:custGeom>`;
}

function geomXml(geom: string, adj: Record<string, number> | undefined, paths?: CustomPath[]): string {
  if (geom === "custom" && paths?.length) return customGeomXml(paths);
  const gd = adj ? Object.entries(adj).map(([k, v]) => `<a:gd name="${k}" fmla="val ${Math.round(v)}"/>`).join("") : "";
  return `<a:prstGeom prst="${esc(geom === "custom" ? "rect" : geom)}"><a:avLst>${gd}</a:avLst></a:prstGeom>`;
}

/* ------------------------------------------------------------------ */
/* Metin                                                                */
/* ------------------------------------------------------------------ */

function fontXml(font: string | undefined): string {
  if (!font) return "";
  if (font === "+mj") return '<a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/>';
  if (font === "+mn") return '<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/>';
  return `<a:latin typeface="${esc(font)}"/><a:cs typeface="${esc(font)}"/>`;
}

export function rPrXml(tag: string, p: RunProps, ctx: WriteCtx, lang = true): string {
  const a = attrs({
    lang: lang ? "tr-TR" : undefined,
    sz: p.size !== undefined ? Math.round(p.size * 100) : undefined,
    b: p.bold === undefined ? undefined : p.bold ? "1" : "0",
    i: p.italic === undefined ? undefined : p.italic ? "1" : "0",
    u: p.underline === undefined ? undefined : p.underline ? "sng" : "none",
    strike: p.strike === undefined ? undefined : p.strike ? "sngStrike" : "noStrike",
    cap: p.caps,
    spc: p.spacing !== undefined ? Math.round(p.spacing * 100) : undefined,
    baseline: p.baseline ? Math.round(p.baseline * 1000) : undefined,
    dirty: lang ? "0" : undefined,
  });
  let inner = "";
  if (p.outline) inner += `<a:ln w="9525"><a:solidFill>${colorXml(p.outline, ctx)}</a:solidFill></a:ln>`;
  if (p.color) inner += `<a:solidFill>${colorXml(p.color, ctx)}</a:solidFill>`;
  if (p.shadow || p.glow) {
    inner += "<a:effectLst>";
    if (p.glow) inner += `<a:glow rad="101600">${colorXml({ ...p.glow, alpha: p.glow.alpha ?? 0.6 }, ctx)}</a:glow>`;
    if (p.shadow) inner += '<a:outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl"><a:srgbClr val="000000"><a:alpha val="43137"/></a:srgbClr></a:outerShdw>';
    inner += "</a:effectLst>";
  }
  if (p.highlight) inner += `<a:highlight>${colorXml(p.highlight, ctx)}</a:highlight>`;
  inner += fontXml(p.font);
  inner += hlinkXml(p.link, ctx);
  return inner ? `<${tag}${a}>${inner}</${tag}>` : `<${tag}${a}/>`;
}

function bulletXml(b: BulletDef | undefined, ctx: WriteCtx): string {
  if (!b) return "";
  if (b.type === "none") return "<a:buNone/>";
  const clr = b.color ? `<a:buClr>${colorXml(b.color, ctx)}</a:buClr>` : "";
  const sz = b.size ? `<a:buSzPct val="${Math.round(b.size * 1000)}"/>` : "";
  if (b.type === "char") return `${clr}${sz}${b.font ? `<a:buFont typeface="${esc(b.font)}"/>` : ""}<a:buChar char="${esc(b.char)}"/>`;
  return `${clr}${sz}<a:buFont typeface="+mj-lt"/><a:buAutoNum type="${b.scheme}"${b.startAt && b.startAt !== 1 ? ` startAt="${b.startAt}"` : ""}/>`;
}

export function pPrXml(tag: string, p: ParaProps, ctx: WriteCtx, extra: { lvl?: number; defRPr?: RunProps } = {}): string {
  const a = attrs({
    marL: p.marL !== undefined ? emu(p.marL) : undefined,
    lvl: extra.lvl ? extra.lvl : undefined,
    indent: p.indent !== undefined ? emu(p.indent) : undefined,
    algn: p.align,
  });
  let inner = "";
  if (p.lineSpacingPt !== undefined) inner += `<a:lnSpc><a:spcPts val="${Math.round(p.lineSpacingPt * 100)}"/></a:lnSpc>`;
  else if (p.lineSpacing !== undefined) inner += `<a:lnSpc><a:spcPct val="${Math.round(p.lineSpacing * 1000)}"/></a:lnSpc>`;
  if (p.spaceBefore !== undefined) inner += `<a:spcBef><a:spcPts val="${Math.round(p.spaceBefore * 100)}"/></a:spcBef>`;
  if (p.spaceAfter !== undefined) inner += `<a:spcAft><a:spcPts val="${Math.round(p.spaceAfter * 100)}"/></a:spcAft>`;
  inner += bulletXml(p.bullet, ctx);
  if (extra.defRPr && Object.keys(extra.defRPr).length) inner += rPrXml("a:defRPr", extra.defRPr, ctx, false);
  return inner ? `<${tag}${a}>${inner}</${tag}>` : a ? `<${tag}${a}/>` : "";
}

export function lstStyleXml(levels: Partial<LevelStyle>[] | undefined, ctx: WriteCtx): string {
  if (!levels?.length) return "<a:lstStyle/>";
  const parts = levels
    .slice(0, 9)
    .map((l, i) => (l ? pPrXml(`a:lvl${i + 1}pPr`, l.para ?? {}, ctx, { defRPr: l.run }) : ""))
    .join("");
  return `<a:lstStyle>${parts}</a:lstStyle>`;
}

let fldSeq = 0;
function fldId(): string {
  fldSeq += 1;
  const h = (fldSeq * 2654435761).toString(16).toUpperCase().padStart(8, "0").slice(-8);
  return `{${h}-0F1C-4E7E-9B0B-${(fldSeq + 0x100000000).toString(16).toUpperCase().slice(-12).padStart(12, "0")}}`;
}

export function paragraphXml(p: Paragraph, ctx: WriteCtx): string {
  const pPr = pPrXml("a:pPr", p.props, ctx, { lvl: p.level });
  let runs = "";
  for (const r of p.runs) {
    if (!r.text) continue;
    if (r.props.field) {
      const type = r.props.field === "slidenum" ? "slidenum" : "datetime1";
      const { field: _f, ...rp } = r.props;
      void _f;
      runs += `<a:fld id="${fldId()}" type="${type}">${rPrXml("a:rPr", rp, ctx)}<a:t>${esc(r.text)}</a:t></a:fld>`;
      continue;
    }
    const parts = r.text.split("\v");
    parts.forEach((part, i) => {
      if (i > 0) runs += `<a:br>${rPrXml("a:rPr", r.props, ctx)}</a:br>`;
      if (part) runs += `<a:r>${rPrXml("a:rPr", r.props, ctx)}<a:t>${esc(part)}</a:t></a:r>`;
    });
  }
  const end = rPrXml("a:endParaRPr", p.endProps ?? {}, ctx);
  return `<a:p>${pPr}${runs}${end}</a:p>`;
}

export function bodyPrXml(b: TextBody | undefined, forceDefaults = false): string {
  const ins = b?.insets;
  const a = attrs({
    vert: b?.vert && b.vert !== "horz" ? (b.vert === "stacked" ? "wordArtVert" : b.vert) : undefined,
    wrap: b?.wrap === undefined ? undefined : b.wrap ? "square" : "none",
    lIns: ins ? emu(ins.l) : undefined,
    tIns: ins ? emu(ins.t) : undefined,
    rIns: ins ? emu(ins.r) : undefined,
    bIns: ins ? emu(ins.b) : undefined,
    rtlCol: forceDefaults ? "0" : undefined,
    anchor: b?.anchor,
  });
  let fit = "";
  if (b?.autofit === "shrink") fit = `<a:normAutofit${attrs({ fontScale: b.fontScale && b.fontScale < 1 ? Math.round(b.fontScale * 100000) : undefined, lnSpcReduction: b.lnSpcReduction ? Math.round(b.lnSpcReduction * 100000) : undefined })}/>`;
  else if (b?.autofit === "shape") fit = "<a:spAutoFit/>";
  else if (b?.autofit === "none") fit = "<a:noAutofit/>";
  return fit ? `<a:bodyPr${a}>${fit}</a:bodyPr>` : `<a:bodyPr${a}/>`;
}

export function txBodyXml(tag: string, b: TextBody, ctx: WriteCtx, promptText?: string): string {
  let paras = b.paragraphs;
  if (promptText !== undefined && paras.every((p) => !p.runs.some((r) => r.text))) {
    paras = [{ runs: [{ text: promptText, props: {} }], level: 0, props: {} }];
  }
  const ps = (paras.length ? paras : [{ runs: [], level: 0, props: {} }]).map((p) => paragraphXml(p, ctx)).join("");
  return `<${tag}>${bodyPrXml(b, true)}${lstStyleXml(b.levels, ctx)}${ps}</${tag}>`;
}

/* ------------------------------------------------------------------ */
/* Şekiller                                                             */
/* ------------------------------------------------------------------ */

function cNvPr(s: Shape, ctx: WriteCtx): string {
  const id = ctx.nextId();
  ctx.idMap.set(s.id, id);
  const a = attrs({ id, name: s.name || `Şekil ${id}`, descr: s.descr, hidden: s.hidden ? "1" : undefined });
  const link = hlinkXml(s.link, ctx);
  return link ? `<p:cNvPr${a}>${link}</p:cNvPr>` : `<p:cNvPr${a}/>`;
}

function phXml(s: { ph?: { type: string; idx?: number; sz?: string; orient?: string } }): string {
  if (!s.ph) return "<p:nvPr/>";
  const type = s.ph.type === "obj" ? undefined : s.ph.type;
  return `<p:nvPr><p:ph${attrs({ type, orient: s.ph.orient, sz: s.ph.sz, idx: s.ph.idx })}/></p:nvPr>`;
}

export function spXml(s: SpShape, ctx: WriteCtx, opts: { promptAsText?: string } = {}): string {
  const connector = s.connector || (isLineGeom(s.geom) && !s.text);
  const spPr = `<p:spPr>${xfrmXml(s)}${geomXml(s.geom, s.adj, s.paths)}${fillXml(s.fill, ctx)}${lineXml(s.line, ctx)}${effectsXml(s.effects, ctx)}</p:spPr>`;
  if (connector) {
    return `<p:cxnSp><p:nvCxnSpPr>${cNvPr(s, ctx)}<p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>${spPr}</p:cxnSp>`;
  }
  const locks = s.ph ? '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' : s.txBox ? '<p:cNvSpPr txBox="1"/>' : "<p:cNvSpPr/>";
  const tx = s.text ? txBodyXml("p:txBody", s.text, ctx, opts.promptAsText) : "";
  return `<p:sp><p:nvSpPr>${cNvPr(s, ctx)}${locks}${phXml(s)}</p:nvSpPr>${spPr}${tx}</p:sp>`;
}

export function picXml(s: PicShape, ctx: WriteCtx): string {
  const rid = imageRel(ctx, s.src);
  const adj = s.adjust;
  let blipInner = "";
  if (adj?.alpha !== undefined && adj.alpha < 1) blipInner += `<a:alphaModFix amt="${pct(adj.alpha)}"/>`;
  if (adj?.grayscale) blipInner += "<a:grayscl/>";
  if (adj?.brightness || adj?.contrast) blipInner += `<a:lum${attrs({ bright: adj.brightness ? pct(adj.brightness) : undefined, contrast: adj.contrast ? pct(adj.contrast) : undefined })}/>`;
  const c = s.crop;
  const src = c && (c.l || c.t || c.r || c.b) ? `<a:srcRect${attrs({ l: c.l ? pct(c.l) : undefined, t: c.t ? pct(c.t) : undefined, r: c.r ? pct(c.r) : undefined, b: c.b ? pct(c.b) : undefined })}/>` : "";
  const blip = rid ? `<a:blip r:embed="${rid}">${blipInner}</a:blip>` : "<a:blip/>";
  let nvPr = phXml(s);
  let cnv = cNvPr(s, ctx);
  const mediaPath = s.media && ctx.media.addMedia ? ctx.media.addMedia(s.media.src) : null;
  if (s.media && mediaPath) {
    // Video/ses: a:videoFile/a:audioFile bağlantısı + p14:media gömülü ortam (ECMA-376 §20.1.3, PowerPoint 2010 uzantısı).
    const target = relativeTarget(ctx.part, mediaPath);
    const linkRid = ctx.rels.add(s.media.kind === "video" ? REL.video : REL.audio, target);
    const embedRid = ctx.rels.add(REL.media, target);
    const tag = s.media.kind === "video" ? "a:videoFile" : "a:audioFile";
    const opts = s.media.autoPlay || s.media.loop ? `<p:ext uri="${OH_EXT_URI}"><oh:media xmlns:oh="${NS.oh}"${attrs({ autoPlay: s.media.autoPlay, loop: s.media.loop })}/></p:ext>` : "";
    const ph = s.ph ? `<p:ph${attrs({ type: s.ph.type === "obj" ? undefined : s.ph.type, idx: s.ph.idx })}/>` : "";
    nvPr = `<p:nvPr>${ph}<${tag} r:link="${linkRid}"/><p:extLst><p:ext uri="${MEDIA_EXT_URI}"><p14:media xmlns:p14="${NS.p14}" r:embed="${embedRid}"/></p:ext>${opts}</p:extLst></p:nvPr>`;
    if (!s.link) cnv = cnv.replace(/\/>$/, '><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr>');
  }
  return (
    `<p:pic><p:nvPicPr>${cnv}<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>${nvPr}</p:nvPicPr>` +
    `<p:blipFill>${blip}${src}<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr>${xfrmXml(s)}${geomXml(s.geom ?? "rect", undefined)}${lineXml(s.line, ctx)}${effectsXml(s.effects, ctx)}</p:spPr></p:pic>`
  );
}

function borderXml(tag: string, b: CellBorder | null | undefined, ctx: WriteCtx): string {
  if (b === undefined) return "";
  if (b === null) return `<a:${tag} w="12700"><a:noFill/></a:${tag}>`;
  return `<a:${tag} w="${Math.round(b.width * 12700)}"><a:solidFill>${colorXml(b.color, ctx)}</a:solidFill></a:${tag}>`;
}

export function tableXml(t: TableShape, ctx: WriteCtx): string {
  const st = tableStyleById(t.style.styleId);
  const o = t.style;
  const tblPr = `<a:tblPr${attrs({ firstRow: o.firstRow, firstCol: o.firstCol, lastRow: o.lastRow, lastCol: o.lastCol, bandRow: o.bandRow, bandCol: o.bandCol })}><a:tableStyleId>${st.guid}</a:tableStyleId></a:tblPr>`;
  const grid = `<a:tblGrid>${t.cols.map((w) => `<a:gridCol w="${emu(w)}"/>`).join("")}</a:tblGrid>`;
  const rows = t.rows
    .map((r) => {
      const cells = r.cells
        .map((c) => {
          const b = c.borders ?? {};
          const tcPr = `<a:tcPr${attrs({ anchor: c.anchor })}>${borderXml("lnL", b.l, ctx)}${borderXml("lnR", b.r, ctx)}${borderXml("lnT", b.t, ctx)}${borderXml("lnB", b.b, ctx)}${c.fill ? fillXml(c.fill, ctx) : ""}</a:tcPr>`;
          return `<a:tc${attrs({ gridSpan: c.gridSpan && c.gridSpan > 1 ? c.gridSpan : undefined, rowSpan: c.rowSpan && c.rowSpan > 1 ? c.rowSpan : undefined, hMerge: c.hMerge, vMerge: c.vMerge })}>${txBodyXml("a:txBody", c.text, ctx)}${tcPr}</a:tc>`;
        })
        .join("");
      return `<a:tr h="${emu(r.h)}">${cells}</a:tr>`;
    })
    .join("");
  return (
    `<p:graphicFrame><p:nvGraphicFramePr>${cNvPr(t, ctx)}<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>` +
    `${xfrmXml(t, "p:xfrm")}<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>${tblPr}${grid}${rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`
  );
}

export function groupXml(g: GroupShape, ctx: WriteCtx): string {
  const nv = `<p:nvGrpSpPr>${cNvPr(g, ctx)}<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>`;
  const pr = `<p:grpSpPr>${xfrmXml(g, "a:xfrm", g)}</p:grpSpPr>`;
  return `<p:grpSp>${nv}${pr}${g.children.map((c) => shapeXml(c, ctx)).join("")}</p:grpSp>`;
}

export function shapeXml(s: Shape, ctx: WriteCtx, opts: { promptAsText?: (s: SpShape) => string | undefined } = {}): string {
  switch (s.kind) {
    case "sp":
      return spXml(s, ctx, { promptAsText: opts.promptAsText?.(s) });
    case "pic":
      return picXml(s, ctx);
    case "table":
      return tableXml(s, ctx);
    case "group":
      return groupXml(s, ctx);
    case "chart":
      return chartFrameXml(s, ctx);
  }
}

/** Grafik çerçevesi (p:graphicFrame + c:chart başvurusu). */
export function chartFrameXml(s: ChartShape, ctx: WriteCtx): string {
  if (!ctx.charts) return "";
  const path = ctx.charts.add(s.chart);
  const rid = ctx.rels.add(REL.chart, relativeTarget(ctx.part, path));
  return (
    `<p:graphicFrame><p:nvGraphicFramePr>${cNvPr(s, ctx)}<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>` +
    `${xfrmXml(s, "p:xfrm")}<a:graphic><a:graphicData uri="${NS.c}"><c:chart xmlns:c="${NS.c}" r:id="${rid}"/></a:graphicData></a:graphic></p:graphicFrame>`
  );
}

export function spTreeXml(shapes: Shape[], ctx: WriteCtx, opts: { promptAsText?: (s: SpShape) => string | undefined } = {}): string {
  const head =
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
  return `<p:spTree>${head}${shapes.map((s) => shapeXml(s, ctx, opts)).join("")}</p:spTree>`;
}

export function bgXml(fill: Fill | undefined, ctx: WriteCtx): string {
  if (!fill) return "";
  return `<p:bg><p:bgPr>${fillXml(fill, ctx)}<a:effectLst/></p:bgPr></p:bg>`;
}
