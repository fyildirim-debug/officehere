/** ODP şekil yazıcısı: metin çerçeveleri, yer tutucular, özel şekiller, çizgiler, resimler, tablolar, gruplar. */
import { resolveHex } from "../../model/color";
import { newPresentation } from "../../model/presentation";
import { effectiveBodyProps, levelResolver, type TextContext } from "../../model/textStyle";
import type { Layout, PicShape, Presentation, RunProps, Shape, SpShape, TableShape, TextBody } from "../../model/types";
import { customPathD, isLineGeom, presetGeometry } from "../../render/geometry";
import { svgPathToCmds } from "../../render/svgPath";
import { cellFormat } from "../../render/tableStyles";
import { esc } from "../xml";
import { cm, odfAttrs } from "./common";
import { fillAttrs, strokeAttrs, textProps, type Attrs, type AutoStyles, type Shared } from "./writeStyles";
import { ownRun, textBodyXml } from "./writeText";

export interface ShapeWriteCtx {
  sh: Shared;
  auto: AutoStyles;
  pres: Presentation;
  layout?: Layout;
  /** Yer tutucu stillerinin ana sayfa öneki ("M1" → "M1-title"). */
  master: string;
  link: (l: string) => string | undefined;
  /** Ana sayfa (styles.xml) içinde mi. */
  inMaster?: boolean;
  /**
   * Yer tutucu metninin ana sayfa sunu stillerinden (başlık, alt başlık, anahat
   * düzeyleri) devralacağı çalıştırma biçimi; okuyucu bu stilleri düzey olarak okur.
   */
  phInherited?: (cls: "title" | "subtitle" | "outline", level: number) => RunProps;
}

let tableDefaults: RunProps | undefined;
/**
 * Biçimsiz tablo hücresi metninin yeniden açılışta alacağı biçim: ODP okuyucusu
 * hücre metnine düzey yazmaz; sunu varsayılan temanın "diğer" metin stilini kullanır.
 */
function tableTextDefaults(): RunProps {
  if (!tableDefaults) {
    const p = newPresentation("varsayilan", 0);
    const run = p.master.textStyles.other[0]?.run ?? {};
    tableDefaults = { ...run, color: run.color ? { rgb: resolveHex(run.color, p.theme, "#000000") } : undefined };
  }
  return tableDefaults;
}

const r2 = (n: number) => Math.round(n * 100) / 100;

/** Konum öznitelikleri; döndürülmüş şekilde draw:transform (ODF: pozitif açı saat yönünün tersi). */
function boxAttrs(s: Shape): Attrs {
  const w = Math.max(0, s.w);
  const h = Math.max(0, s.h);
  const rot = ((s.rot % 360) + 360) % 360;
  if (rot < 0.01 || rot > 359.99) return { "svg:x": cm(s.x), "svg:y": cm(s.y), "svg:width": cm(w), "svg:height": cm(h) };
  const t = (-rot * Math.PI) / 180;
  const cx = s.x + w / 2;
  const cy = s.y + h / 2;
  // Yerel merkez R(t) ile dönünce merkeze otursun diye öteleme.
  const rx = (w / 2) * Math.cos(t) + (h / 2) * Math.sin(t);
  const ry = -(w / 2) * Math.sin(t) + (h / 2) * Math.cos(t);
  return { "svg:width": cm(w), "svg:height": cm(h), "draw:transform": `rotate (${Math.round(t * 1e9) / 1e9}) translate (${cm(cx - rx)} ${cm(cy - ry)})` };
}

function textCtx(ctx: ShapeWriteCtx): TextContext {
  return { master: ctx.pres.master, layout: ctx.layout };
}

/** Metin gövdesi grafik özellikleri. */
function bodyAttrs(s: SpShape, ctx: ShapeWriteCtx, frame: boolean): Attrs {
  const b = effectiveBodyProps(s, textCtx(ctx));
  return {
    "draw:textarea-vertical-align": b.anchor === "ctr" ? "middle" : b.anchor === "b" ? "bottom" : "top",
    "draw:textarea-horizontal-align": frame ? undefined : "justify",
    "fo:padding-left": cm(b.insets.l),
    "fo:padding-top": cm(b.insets.t),
    "fo:padding-right": cm(b.insets.r),
    "fo:padding-bottom": cm(b.insets.b),
    "draw:auto-grow-height": b.autofit === "shape" ? "true" : "false",
    "draw:auto-grow-width": frame && !b.wrap ? "true" : "false",
    "fo:wrap-option": b.wrap ? "wrap" : "no-wrap",
    "style:shrink-to-fit": b.autofit === "shrink" ? "true" : undefined,
    "fo:min-height": "0cm",
    "style:writing-mode": b.vert === "vert" || b.vert === "vert270" ? "tb-rl" : undefined,
  };
}

function shadowAttrs(s: SpShape | PicShape, ctx: ShapeWriteCtx): Attrs {
  const sh = s.effects?.shadow;
  if (!sh) return {};
  const rad = (sh.dir * Math.PI) / 180;
  return {
    "draw:shadow": "visible",
    "draw:shadow-offset-x": cm(Math.cos(rad) * sh.dist),
    "draw:shadow-offset-y": cm(Math.sin(rad) * sh.dist),
    "draw:shadow-color": ctx.sh.color(sh.color, "#808080"),
    "draw:shadow-opacity": sh.color.alpha !== undefined ? `${Math.round(sh.color.alpha * 100)}%` : undefined,
    "draw:shadow-blur": sh.blur ? cm(sh.blur) : undefined,
  };
}

function graphicStyle(ctx: ShapeWriteCtx, a: Attrs, family = "graphic", parent: string | undefined = "standard", kind = "gr", extra = ""): string {
  return ctx.auto.style(family, kind, `<style:graphic-properties${odfAttrs(a)}/>${extra}`, parent);
}

/**
 * Metin. Yer tutucuda devralınan biçim ana sayfa sunu stillerinden, diğer
 * şekillerde şeklin grafik stiline yazılan düzey biçiminden gelir; çalıştırmalara
 * yalnız açık ya da farklı özellikler yazılır.
 */
function textXml(s: SpShape, body: TextBody | undefined, ctx: ShapeWriteCtx, phCls?: "title" | "subtitle" | "outline", phOwn: RunProps = {}): string {
  const lvl = levelResolver(s, textCtx(ctx));
  const inherited = phCls && ctx.phInherited ? (level: number) => ({ ...ctx.phInherited!(phCls, level), ...phOwn }) : () => lvl(0).run;
  return textBodyXml(body, { sh: ctx.sh, auto: ctx.auto, lvl, link: ctx.link, inherited });
}

/**
 * Yer tutucunun kendi düzey biçiminin ana sayfa stilinden farkı: şeklin sunu
 * stiline yazılır (okuyucu tüm düzeylere uygular). Anahatta yalnız kullanılan
 * düzeylerin farkı aynıysa; değilse fark çalıştırmalara kalır.
 */
function placeholderOwnRun(s: SpShape, cls: "title" | "subtitle" | "outline", ctx: ShapeWriteCtx): RunProps {
  if (!ctx.phInherited || !s.text) return {};
  const lvl = levelResolver(s, textCtx(ctx));
  const levels = cls === "outline" ? [...new Set(s.text.paragraphs.map((p) => p.level))] : [0];
  const diffs = levels.map((l) => ownRun({}, lvl(l).run, ctx.phInherited!(cls, l), ctx.sh));
  const first = textProps(diffs[0] ?? {}, ctx.sh);
  return diffs.every((d) => textProps(d, ctx.sh) === first) ? { ...diffs[0], link: undefined, field: undefined } : {};
}

/** Yer tutucu olmayan metinli şeklin düzey biçimi grafik stiline yazılır (okuyucu şekil düzeyi olarak okur). */
function shapeTextProps(s: SpShape, ctx: ShapeWriteCtx): string {
  if (!s.text) return "";
  const run = { ...levelResolver(s, textCtx(ctx))(0).run, link: undefined, field: undefined };
  return textProps(run, ctx.sh);
}

const PH_CLASS: Record<string, [string, string]> = {
  title: ["title", "title"],
  ctrTitle: ["title", "title"],
  subTitle: ["subtitle", "subtitle"],
  body: ["outline", "outline1"],
  obj: ["outline", "outline1"],
};

function isEmpty(t: TextBody | undefined): boolean {
  return !t || t.paragraphs.every((p) => p.runs.every((r) => !r.text));
}

function spXml(s: SpShape, ctx: ShapeWriteCtx): string {
  const name = odfAttrs({ "draw:name": s.name || undefined });
  const ph = s.ph ? PH_CLASS[s.ph.type] : undefined;
  if (ph && (s.text || ctx.inMaster)) {
    const cls = ph[0] === "title" ? "title" : ph[0] === "subtitle" ? "subtitle" : "outline";
    const empty = isEmpty(s.text);
    const own = empty ? {} : placeholderOwnRun(s, cls, ctx);
    const st = graphicStyle(ctx, { ...fillAttrs(s.fill, ctx.sh), ...strokeAttrs(s.line, ctx.sh), ...bodyAttrs(s, ctx, true) }, "presentation", `${ctx.master}-${ph[1]}`, "pr", textProps(own, ctx.sh));
    const body = empty ? "" : textXml(s, s.text, ctx, cls, own);
    return `<draw:frame${name}${odfAttrs({ "presentation:style-name": st, "draw:layer": ctx.inMaster ? "backgroundobjects" : "layout", ...boxAttrs(s), "presentation:class": ph[0], "presentation:placeholder": empty ? "true" : undefined })}><draw:text-box>${body}</draw:text-box></draw:frame>`;
  }
  const g = s.geom === "custom" ? "custom" : s.geom;
  if (isLineGeom(g) && (g === "line" || g === "straightConnector1")) return lineXml(s, ctx);
  const frame = !!s.txBox || (g === "rect" && !!s.text && s.fill.type === "none" && s.line.fill.type === "none" && !s.ph);
  const st = graphicStyle(ctx, { ...fillAttrs(isLineGeom(g) ? { type: "none" } : s.fill, ctx.sh), ...strokeAttrs(s.line, ctx.sh), ...(s.text ? bodyAttrs(s, ctx, frame) : {}), ...shadowAttrs(s, ctx) }, "graphic", "standard", "gr", shapeTextProps(s, ctx));
  const layer = odfAttrs({ "draw:layer": ctx.inMaster ? "backgroundobjects" : "layout" });
  if (frame) return `<draw:frame${name}${odfAttrs({ "draw:style-name": st })}${layer}${odfAttrs(boxAttrs(s))}><draw:text-box>${textXml(s, s.text, ctx)}</draw:text-box></draw:frame>`;
  const text = s.text && !isEmpty(s.text) ? textXml(s, s.text, ctx) : "";
  return `<draw:custom-shape${name}${odfAttrs({ "draw:style-name": st })}${layer}${odfAttrs(boxAttrs(s))}>${text}${geometryXml(s)}</draw:custom-shape>`;
}

/** Hazır/özel geometri → draw:enhanced-geometry (sayısal yol; viewBox = piksel × 100). */
function geometryXml(s: SpShape): string {
  const w = Math.max(1, s.w);
  const h = Math.max(1, s.h);
  const S = 100;
  const n = (v: number) => Math.round(v * S);
  let paths: { d: string; fill: boolean; stroke: boolean }[];
  let text = { x: 0, y: 0, w, h };
  if (s.geom === "custom" && s.paths?.length) paths = s.paths.map((p) => ({ d: customPathD(p, w, h), fill: p.fill !== false, stroke: p.stroke !== false }));
  else {
    const g = presetGeometry(s.geom, w, h, s.adj);
    paths = g.paths.map((p) => ({ d: p.d, fill: !g.open && p.fill !== "none", stroke: p.stroke }));
    text = g.text;
  }
  let ep = "";
  for (const p of paths) {
    let cmds;
    try {
      cmds = svgPathToCmds(p.d);
    } catch {
      continue;
    }
    for (const c of cmds) {
      const pts = c.pts.map(n).join(" ");
      ep += `${c.c}${pts ? ` ${pts}` : ""} `;
    }
    if (!p.fill) ep += "F ";
    if (!p.stroke) ep += "S ";
    ep += "N ";
  }
  const isPreset = s.geom !== "custom";
  // Hazır şekiller "ooxml-<ad>" türüyle yazılır (yol açıkça verilir; ayarlar draw:modifiers'ta).
  const odfType = s.geom === "rect" ? "rectangle" : s.geom === "ellipse" ? "ellipse" : isPreset ? `ooxml-${s.geom}` : "non-primitive";
  const adj = isPreset && s.adj ? Object.keys(s.adj).sort().map((k) => Math.round(s.adj![k])) : [];
  return `<draw:enhanced-geometry${odfAttrs({
    "svg:viewBox": `0 0 ${n(w)} ${n(h)}`,
    "draw:type": odfType,
    "draw:modifiers": adj.length ? adj.join(" ") : undefined,
    // LibreOffice hazır (ooxml-*) türlerde metin alanını kendisi hesaplar; açık değer metni gizleyebiliyor.
    "draw:text-areas": isPreset ? undefined : `${n(text.x)} ${n(text.y)} ${n(text.x + text.w)} ${n(text.y + text.h)}`,
    "draw:mirror-horizontal": s.flipH ? "true" : undefined,
    "draw:mirror-vertical": s.flipV ? "true" : undefined,
    "draw:enhanced-path": ep.trim() || `M 0 0 L ${n(w)} 0 ${n(w)} ${n(h)} 0 ${n(h)} Z N`,
  })}/>`;
}

function lineXml(s: SpShape, ctx: ShapeWriteCtx): string {
  let x1 = s.flipH ? s.x + s.w : s.x;
  let y1 = s.flipV ? s.y + s.h : s.y;
  let x2 = s.flipH ? s.x : s.x + s.w;
  let y2 = s.flipV ? s.y : s.y + s.h;
  if (s.rot) {
    const t = (s.rot * Math.PI) / 180;
    const cx = s.x + s.w / 2;
    const cy = s.y + s.h / 2;
    const rot = (x: number, y: number): [number, number] => [cx + (x - cx) * Math.cos(t) - (y - cy) * Math.sin(t), cy + (x - cx) * Math.sin(t) + (y - cy) * Math.cos(t)];
    [x1, y1] = rot(x1, y1);
    [x2, y2] = rot(x2, y2);
  }
  const st = graphicStyle(ctx, { "draw:fill": "none", ...strokeAttrs(s.line, ctx.sh) });
  const tag = s.connector || s.geom === "straightConnector1" ? "draw:connector" : "draw:line";
  return `<${tag}${odfAttrs({ "draw:name": s.name || undefined, "draw:style-name": st, "draw:layer": ctx.inMaster ? "backgroundobjects" : "layout", "draw:type": tag === "draw:connector" ? "line" : undefined, "svg:x1": cm(x1), "svg:y1": cm(y1), "svg:x2": cm(x2), "svg:y2": cm(y2) })}><text:p/></${tag}>`;
}

function picXml(s: PicShape, ctx: ShapeWriteCtx): string {
  const path = ctx.sh.picture(s.src);
  if (!path) return "";
  const mime = ctx.sh.files.find((f) => f.path === path)?.mime;
  const mirror = s.flipH && s.flipV ? "horizontal vertical" : s.flipH ? "horizontal" : s.flipV ? "vertical" : undefined;
  const st = graphicStyle(ctx, { "draw:fill": "none", ...strokeAttrs(s.line, ctx.sh), "style:mirror": mirror, ...shadowAttrs(s, ctx) });
  const desc = s.descr ? `<svg:desc>${esc(s.descr)}</svg:desc>` : "";
  return `<draw:frame${odfAttrs({ "draw:name": s.name || undefined, "draw:style-name": st, "draw:layer": ctx.inMaster ? "backgroundobjects" : "layout", ...boxAttrs(s) })}><draw:image${odfAttrs({ "xlink:href": path, "xlink:type": "simple", "xlink:show": "embed", "xlink:actuate": "onLoad", "draw:mime-type": mime })}><text:p/></draw:image>${desc}</draw:frame>`;
}

function tableXml(t: TableShape, ctx: ShapeWriteCtx): string {
  const cols = t.cols.map((w) => `<table:table-column table:style-name="${ctx.auto.style("table-column", "co", `<style:table-column-properties style:column-width="${cm(w)}"/>`)}"/>`).join("");
  const rows = t.rows
    .map((row, ri) => {
      const rs = ctx.auto.style("table-row", "ro", `<style:table-row-properties style:row-height="${cm(row.h)}"/>`);
      const cells = row.cells
        .map((cell, ci) => {
          if (cell.hMerge || cell.vMerge) return "<table:covered-table-cell/>";
          const f = cellFormat(t, ri, ci);
          const border = (b: { color: Parameters<Shared["color"]>[0]; width: number } | null) => (b ? `${r2(b.width)}pt solid ${ctx.sh.color(b.color)}` : "none");
          const cs = ctx.auto.style(
            "table-cell",
            "ce",
            `<style:graphic-properties${odfAttrs({ ...fillAttrs(f.fill, ctx.sh), "draw:textarea-vertical-align": cell.anchor === "ctr" ? "middle" : cell.anchor === "b" ? "bottom" : "top" })}/><style:paragraph-properties${odfAttrs({ "fo:border-top": border(f.borders.t), "fo:border-bottom": border(f.borders.b), "fo:border-left": border(f.borders.l), "fo:border-right": border(f.borders.r) })}/>`,
          );
          const pseudo = { kind: "sp", text: cell.text } as unknown as SpShape;
          const body = textBodyXml(cell.text, { sh: ctx.sh, auto: ctx.auto, lvl: levelResolver(pseudo, textCtx(ctx), "other"), link: ctx.link, cell: { color: f.text, bold: f.bold || undefined }, inherited: tableTextDefaults });
          return `<table:table-cell${odfAttrs({ "table:style-name": cs, "table:number-columns-spanned": cell.gridSpan && cell.gridSpan > 1 ? cell.gridSpan : undefined, "table:number-rows-spanned": cell.rowSpan && cell.rowSpan > 1 ? cell.rowSpan : undefined })}>${body || "<text:p/>"}</table:table-cell>`;
        })
        .join("");
      return `<table:table-row table:style-name="${rs}">${cells}</table:table-row>`;
    })
    .join("");
  const st = graphicStyle(ctx, { "draw:fill": "none", "draw:stroke": "none" });
  return `<draw:frame${odfAttrs({ "draw:name": t.name || undefined, "draw:style-name": st, "draw:layer": "layout", ...boxAttrs(t) })}><table:table>${cols}${rows}</table:table></draw:frame>`;
}

/** Şekil listesini yazar (desteklenmeyen grafik nesneleri atlanır). */
export function shapesXml(shapes: Shape[], ctx: ShapeWriteCtx): string {
  let out = "";
  for (const s of shapes) {
    const before = out.length;
    switch (s.kind) {
      case "sp":
        out += spXml(s, ctx);
        break;
      case "pic":
        out += picXml(s, ctx);
        break;
      case "table":
        out += tableXml(s, ctx);
        break;
      case "group": {
        const inner = shapesXml(s.children, ctx);
        if (inner) out += `<draw:g${odfAttrs({ "draw:name": s.name || undefined })}>${inner}</draw:g>`;
        break;
      }
      case "chart":
        break;
    }
    // Gizli şekil: LibreOffice uzantı özniteliğiyle saklanır (görünmez ama korunur).
    if (s.hidden && out.length > before) out = out.slice(0, before) + out.slice(before).replace(/^<([\w:-]+)/, '<$1 drawooo:display="none"');
  }
  return out;
}
