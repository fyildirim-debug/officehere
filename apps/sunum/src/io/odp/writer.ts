/**
 * OpenDocument sunusu (.odp) yazıcısı. Paket: mimetype (ilk, sıkıştırılmamış),
 * META-INF/manifest.xml, content.xml, styles.xml, meta.xml, Pictures/.
 * Kullanılan her slayt düzeni bir ana sayfa olur (arka plan + süsler).
 */
import JSZip from "jszip";
import { findPlaceholder, levelResolver } from "../../model/textStyle";
import type { Layout, LevelStyle, Presentation, Shape, SpShape } from "../../model/types";
import { esc, XML_HEAD } from "../xml";
import { cm, ODF_NS, ODF_NS_DECL, ODP_MIME, odfAttrs } from "./common";
import { AutoStyles, fillAttrs, paraProps, Shared, textProps } from "./writeStyles";
import { shapesXml, type ShapeWriteCtx } from "./writeShapes";
import { listStyleXml, textBodyXml } from "./writeText";

const EXTRA_NS = 'xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"';

function levelBody(lv: LevelStyle, sh: Shared, list?: string): string {
  const g = `<style:graphic-properties draw:stroke="none" draw:fill="none">${list ?? ""}</style:graphic-properties>`;
  const p = { ...lv.para };
  delete p.marL;
  delete p.indent;
  return g + paraProps(p, true) + textProps(lv.run, sh);
}

/** Bir ana sayfanın sunu stilleri (başlık, alt başlık, anahat 1–9, notlar). */
function masterPresStyles(name: string, pres: Presentation, layout: Layout, sh: Shared): string {
  const ctx = { master: pres.master, layout };
  const title = findPlaceholder(layout.shapes, { type: "title" }) ?? findPlaceholder(pres.master.shapes, { type: "title" });
  const body = layout.shapes.find((s): s is SpShape => s.kind === "sp" && (s.ph?.type === "obj" || s.ph?.type === "body")) ?? findPlaceholder(pres.master.shapes, { type: "body" });
  const sub = findPlaceholder(layout.shapes, { type: "subTitle" });
  const tl = levelResolver(title ?? { ph: { type: "title" } }, ctx, "title");
  const bl = levelResolver(body ?? { ph: { type: "body" } }, ctx, "body");
  const sl = sub ? levelResolver(sub, ctx) : bl;
  const style = (suffix: string, inner: string, parent?: string) => `<style:style${odfAttrs({ "style:name": `${name}-${suffix}`, "style:family": "presentation", "style:parent-style-name": parent })}>${inner}</style:style>`;
  const levels = new Map(Array.from({ length: 9 }, (_, k) => {
    const lv = bl(k);
    return [k, { bullet: lv.para.bullet ?? { type: "char", char: "•" }, marL: lv.para.marL ?? 24 + k * 48, indent: lv.para.indent ?? -24 }] as const;
  }));
  const list = `<text:list-style style:name="${name}-outline1">${listStyleXml(levels, sh)}</text:list-style>`;
  let out = style("title", levelBody(tl(0), sh).replace('draw:fill="none">', 'draw:fill="none" draw:textarea-vertical-align="middle">'));
  out += style("subtitle", levelBody({ ...sl(0), para: { ...sl(0).para, bullet: undefined } }, sh).replace('draw:fill="none">', 'draw:fill="none" draw:textarea-vertical-align="middle">'));
  for (let k = 1; k <= 9; k++) out += style(`outline${k}`, levelBody(bl(k - 1), sh, k === 1 ? list : undefined), k > 1 ? `${name}-outline${k - 1}` : undefined);
  out += style("notes", `<style:graphic-properties draw:stroke="none" draw:fill="none"/>${textProps({ size: 12, font: "+mn" }, sh)}`);
  out += style("background", `<style:graphic-properties draw:stroke="none" draw:fill="none"/>`);
  out += style("backgroundobjects", `<style:graphic-properties draw:shadow="hidden"/>`);
  return out;
}

/** Sunuyu ODP bayt dizisine yazar. */
export async function writeOdp(pres: Presentation): Promise<Uint8Array> {
  const sh = new Shared(pres.theme);
  const auto = new AutoStyles("");
  const mauto = new AutoStyles("M");

  // Kullanılan düzenler → ana sayfalar.
  const used: Layout[] = [];
  for (const s of pres.slides) {
    const l = pres.layouts.find((x) => x.id === s.layoutId) ?? pres.layouts[0];
    if (l && !used.includes(l)) used.push(l);
  }
  if (!used.length && pres.layouts[0]) used.push(pres.layouts[0]);
  const masterName = (l: Layout | undefined) => `M${Math.max(0, used.indexOf(l!)) + 1}`;

  // Sayfa adları (iç köprüler için benzersiz).
  const names: string[] = [];
  for (const [i, s] of pres.slides.entries()) {
    let n = (s.name ?? "").trim() || `Slayt ${i + 1}`;
    while (names.includes(n)) n += " ";
    names.push(n);
  }
  const byId = new Map(pres.slides.map((s, i) => [s.id, names[i]]));
  const link = (l: string) => (l.startsWith("#slide:") && byId.has(l.slice(7)) ? `#${byId.get(l.slice(7))}` : undefined);

  // Slaytlar.
  let pages = "";
  pres.slides.forEach((slide, i) => {
    const layout = pres.layouts.find((x) => x.id === slide.layoutId) ?? used[0];
    const mname = masterName(layout);
    const dp = auto.style(
      "drawing-page",
      "dp",
      `<style:drawing-page-properties${odfAttrs({
        ...(slide.background ? fillAttrs(slide.background, sh) : {}),
        "draw:background-size": slide.background ? "full" : undefined,
        "presentation:background-visible": "true",
        "presentation:background-objects-visible": slide.hideMasterShapes ? "false" : "true",
        "presentation:visibility": slide.hidden ? "hidden" : undefined,
        "presentation:display-footer": "true",
        "presentation:display-page-number": "false",
        "presentation:display-date-time": "true",
      })}/>`,
    );
    const ctx: ShapeWriteCtx = { sh, auto, pres, layout, master: mname, link };
    let notes = "";
    if (slide.notes.trim()) {
      const body = { paragraphs: slide.notes.split("\n").map((t) => ({ runs: t ? [{ text: t, props: {} }] : [], level: 0, props: {} })) };
      const lvl = () => ({ para: {}, run: {} }) as LevelStyle;
      const ns = auto.style("presentation", "pr", `<style:graphic-properties draw:fill="none" draw:stroke="none"/>`, `${mname}-notes`);
      notes = `<presentation:notes><draw:page-thumbnail${odfAttrs({ "draw:layer": "layout", "svg:x": "2cm", "svg:y": "2cm", "svg:width": "17cm", "svg:height": cm(((17 * 96) / 2.54) * (pres.height / pres.width)), "draw:page-number": i + 1, "presentation:class": "page" })}/><draw:frame${odfAttrs({ "presentation:style-name": ns, "draw:layer": "layout", "svg:x": "2cm", "svg:y": "13.5cm", "svg:width": "17cm", "svg:height": "12cm", "presentation:class": "notes" })}><draw:text-box>${textBodyXml(body, { sh, auto, lvl, link })}</draw:text-box></draw:frame></presentation:notes>`;
    }
    pages += `<draw:page${odfAttrs({ "draw:name": names[i], "draw:style-name": dp, "draw:master-page-name": mname })}>${shapesXml(slide.shapes, ctx)}${notes}</draw:page>`;
  });

  // Ana sayfalar.
  let masters = "";
  let presStyles = "";
  used.forEach((layout, i) => {
    const name = `M${i + 1}`;
    presStyles += masterPresStyles(name, pres, layout, sh);
    const bg = layout.background ?? pres.master.background;
    const mdp = mauto.style("drawing-page", "dp", `<style:drawing-page-properties${odfAttrs({ "draw:background-size": "full", ...fillAttrs(bg, sh) })}/>`);
    const deco: Shape[] = [...(layout.showMasterShapes ? pres.master.shapes.filter((s) => !(s.kind === "sp" && s.ph)) : []), ...layout.shapes.filter((s) => !(s.kind === "sp" && s.ph))];
    const ctx: ShapeWriteCtx = { sh, auto: mauto, pres, layout, master: name, link, inMaster: true };
    const title = findPlaceholder(layout.shapes, { type: "title" }) ?? findPlaceholder(pres.master.shapes, { type: "title" });
    const body = layout.shapes.find((s): s is SpShape => s.kind === "sp" && (s.ph?.type === "obj" || s.ph?.type === "body")) ?? findPlaceholder(pres.master.shapes, { type: "body" });
    const phs = [title, body].filter((s): s is SpShape => !!s).map((s) => ({ ...s, text: undefined, ph: { type: s.ph!.type === "ctrTitle" ? ("title" as const) : s.ph!.type === "obj" ? ("body" as const) : s.ph!.type } }));
    masters += `<style:master-page${odfAttrs({ "style:name": name, "style:display-name": layout.name, "style:page-layout-name": "PM1", "draw:style-name": mdp })}>${shapesXml(deco, ctx)}${shapesXml(phs, ctx)}</style:master-page>`;
  });

  const pageLayout = `<style:page-layout style:name="PM1"><style:page-layout-properties${odfAttrs({ "fo:margin-top": "0cm", "fo:margin-bottom": "0cm", "fo:margin-left": "0cm", "fo:margin-right": "0cm", "fo:page-width": cm(pres.width), "fo:page-height": cm(pres.height), "style:print-orientation": pres.width >= pres.height ? "landscape" : "portrait" })}/></style:page-layout>`;
  const minor = sh.font(pres.theme.fonts.minor);
  const common =
    `<style:default-style style:family="graphic"><style:graphic-properties svg:stroke-color="#3465a4" draw:fill-color="#729fcf"/>${textProps({ size: 18, font: minor }, sh)}</style:default-style>` +
    `<style:style style:name="standard" style:family="graphic"><style:graphic-properties draw:stroke="none" draw:fill="none"/>${textProps({ size: 18, font: minor }, sh)}</style:style>`;

  const content = `${XML_HEAD}<office:document-content ${ODF_NS_DECL} ${EXTRA_NS} office:version="1.3">${sh.fontDecls()}<office:automatic-styles>${auto.xml()}</office:automatic-styles><office:body><office:presentation>${pages}</office:presentation></office:body></office:document-content>`;
  const styles = `${XML_HEAD}<office:document-styles ${ODF_NS_DECL} ${EXTRA_NS} office:version="1.3">${sh.fontDecls()}<office:styles>${sh.named.join("")}${common}${presStyles}</office:styles><office:automatic-styles>${pageLayout}${mauto.xml()}</office:automatic-styles><office:master-styles>${masters}</office:master-styles></office:document-styles>`;

  const now = new Date().toISOString().replace(/\.\d+Z$/, "");
  const p = pres.props;
  const metaItems = [
    `<meta:generator>OfficeHere Sunum</meta:generator>`,
    p.title ? `<dc:title>${esc(p.title)}</dc:title>` : "",
    p.subject ? `<dc:subject>${esc(p.subject)}</dc:subject>` : "",
    p.creator ? `<meta:initial-creator>${esc(p.creator)}</meta:initial-creator><dc:creator>${esc(p.creator)}</dc:creator>` : "",
    `<meta:creation-date>${esc((p.created ?? new Date().toISOString()).replace(/\.\d+Z$|Z$/, ""))}</meta:creation-date>`,
    `<dc:date>${now}</dc:date>`,
  ].join("");
  const meta = `${XML_HEAD}<office:document-meta xmlns:office="${ODF_NS.office}" xmlns:meta="${ODF_NS.meta}" xmlns:dc="${ODF_NS.dc}" office:version="1.3"><office:meta>${metaItems}</office:meta></office:document-meta>`;

  const entries = [
    `<manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="${ODP_MIME}"/>`,
    `<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>`,
    `<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>`,
    `<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>`,
    ...sh.files.map((f) => `<manifest:file-entry manifest:full-path="${esc(f.path)}" manifest:media-type="${esc(f.mime)}"/>`),
  ].join("");
  const manifest = `${XML_HEAD}<manifest:manifest xmlns:manifest="${ODF_NS.manifest}" manifest:version="1.3">${entries}</manifest:manifest>`;

  const zip = new JSZip();
  // mimetype ilk girdi ve sıkıştırılmamış olmalı (ODF paket kuralı).
  zip.file("mimetype", ODP_MIME, { compression: "STORE" });
  zip.file("META-INF/manifest.xml", manifest);
  zip.file("content.xml", content);
  zip.file("styles.xml", styles);
  zip.file("meta.xml", meta);
  for (const f of sh.files) zip.file(f.path, f.bytes, { compression: "STORE" });
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 }, mimeType: ODP_MIME });
}
