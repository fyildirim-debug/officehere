/**
 * PPTX yazıcı (ECMA-376 PresentationML, JSZip ile). Asıl slayt, düzenler,
 * tema, slaytlar, notlar, ortam dosyaları, geçişler ve animasyonlar yazılır.
 */
import JSZip from "jszip";
import { flatShapes } from "../model/ops";
import type { Layout, Presentation, Shape, SpShape } from "../model/types";
import { bgXml, spTreeXml, type MediaStore, type WriteCtx } from "./writeDrawing";
import { appXml, CLR_MAP, coreXml, notesMasterXml, notesSlideXml, NS_DECL, presentationXml, PRES_PROPS, samplePrompt, tableStylesXml, themeXml, txStylesXml, VIEW_PROPS } from "./writeParts";
import { slideExtXml, timingXml, transitionXml } from "./writeTiming";
import { attrs, esc, NS, REL, RelsBuilder, XML_HEAD } from "./xml";

const CT = {
  presentation: "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
  slide: "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
  layout: "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
  master: "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
  theme: "application/vnd.openxmlformats-officedocument.theme+xml",
  notesMaster: "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml",
  notesSlide: "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
  presProps: "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml",
  viewProps: "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml",
  tableStyles: "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml",
  comments: "application/vnd.openxmlformats-officedocument.presentationml.comments+xml",
  commentAuthors: "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml",
  core: "application/vnd.openxmlformats-package.core-properties+xml",
  app: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
};

const MIME_EXT: Record<string, string> = {
  "image/png": "png",
  "image/jpeg": "jpeg",
  "image/jpg": "jpeg",
  "image/gif": "gif",
  "image/bmp": "bmp",
  "image/svg+xml": "svg",
  "image/webp": "webp",
  "image/tiff": "tiff",
  "image/x-emf": "emf",
  "image/x-wmf": "wmf",
};

export const EXT_MIME: Record<string, string> = {
  png: "image/png",
  jpeg: "image/jpeg",
  jpg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  svg: "image/svg+xml",
  webp: "image/webp",
  tiff: "image/tiff",
  tif: "image/tiff",
  emf: "image/x-emf",
  wmf: "image/x-wmf",
};

export function dataUrlToBytes(url: string): { bytes: Uint8Array; mime: string } | null {
  const m = /^data:([^;,]+)?(;base64)?,(.*)$/s.exec(url);
  if (!m) return null;
  const mime = (m[1] ?? "application/octet-stream").toLowerCase();
  const data = m[3];
  if (m[2]) {
    const bin = atob(data);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return { bytes, mime };
  }
  return { bytes: new TextEncoder().encode(decodeURIComponent(data)), mime };
}

class Media implements MediaStore {
  map = new Map<string, string>();
  files: { path: string; bytes: Uint8Array; ext: string }[] = [];
  add(src: string): string | null {
    const hit = this.map.get(src);
    if (hit) return hit;
    const d = dataUrlToBytes(src);
    if (!d) return null;
    const ext = MIME_EXT[d.mime] ?? "png";
    const path = `ppt/media/image${this.files.length + 1}.${ext}`;
    this.files.push({ path, bytes: d.bytes, ext });
    this.map.set(src, path);
    return path;
  }
}

function makeCtx(part: string, media: Media, slideContent: boolean): WriteCtx {
  let id = 1;
  return { rels: new RelsBuilder(), media, part, nextId: () => ++id, idMap: new Map(), slideContent };
}

function tableStyleIds(shapes: Shape[], out: Set<string>) {
  for (const s of flatShapes(shapes)) if (s.kind === "table") out.add(s.style.styleId);
}

/** Sunuyu PPTX bayt dizisine yazar. */
export async function writePptx(pres: Presentation): Promise<Uint8Array> {
  const zip = new JSZip();
  // OPC okuyucularının çoğu içerik türleri parçasını ilk sırada bekler.
  zip.file("[Content_Types].xml", "");
  const media = new Media();
  const overrides: [string, string][] = [];
  const add = (path: string, xml: string, ct?: string) => {
    zip.file(path, XML_HEAD + xml);
    if (ct) overrides.push([`/${path}`, ct]);
  };

  // Tema
  add("ppt/theme/theme1.xml", themeXml(pres.theme), CT.theme);
  add("ppt/theme/theme2.xml", themeXml(pres.theme, "Not Teması"), CT.theme);

  // Asıl slayt
  const masterPath = "ppt/slideMasters/slideMaster1.xml";
  const mctx = makeCtx(masterPath, media, false);
  const layoutPaths = pres.layouts.map((_, i) => `ppt/slideLayouts/slideLayout${i + 1}.xml`);
  const layoutRids = layoutPaths.map((p) => mctx.rels.add(REL.slideLayout, `../slideLayouts/${p.split("/").pop()}`));
  const masterTree = spTreeXml(withFooterFields(pres.master.shapes), mctx, { promptAsText: (s) => samplePrompt(s, true) });
  const layoutIdLst = layoutRids.map((rid, i) => `<p:sldLayoutId id="${2147483649 + i}" r:id="${rid}"/>`).join("");
  const masterTx = txStylesXml(pres.master.textStyles, mctx);
  mctx.rels.add(REL.theme, "../theme/theme1.xml");
  add(masterPath, `<p:sldMaster ${NS_DECL}><p:cSld>${bgXml(pres.master.background, mctx)}${masterTree}</p:cSld>${CLR_MAP}<p:sldLayoutIdLst>${layoutIdLst}</p:sldLayoutIdLst>${masterTx}</p:sldMaster>`, CT.master);
  zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", mctx.rels.xml());

  // Düzenler
  pres.layouts.forEach((l: Layout, i) => {
    const path = layoutPaths[i];
    const ctx = makeCtx(path, media, false);
    ctx.rels.add(REL.slideMaster, "../slideMasters/slideMaster1.xml");
    const tree = spTreeXml(withFooterFields(l.shapes), ctx, { promptAsText: (s) => samplePrompt(s, false) });
    const type = l.type === "cust" ? undefined : l.type;
    add(
      path,
      `<p:sldLayout ${NS_DECL}${attrs({ type, preserve: "1", showMasterSp: l.showMasterShapes ? undefined : "0" })}><p:cSld name="${esc(l.name)}">${bgXml(l.background, ctx)}${tree}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
      CT.layout,
    );
    zip.file(`ppt/slideLayouts/_rels/slideLayout${i + 1}.xml.rels`, ctx.rels.xml());
  });

  // Not asıl sayfası
  const hasNotes = pres.slides.some((s) => s.notes.trim());
  if (hasNotes) {
    add("ppt/notesMasters/notesMaster1.xml", notesMasterXml(), CT.notesMaster);
    const r = new RelsBuilder();
    r.add(REL.theme, "../theme/theme2.xml");
    zip.file("ppt/notesMasters/_rels/notesMaster1.xml.rels", r.xml());
  }

  // Açıklama yazarları (ECMA-376 §19.4: cmAuthorLst / cmLst; konum birimi 1/576 inç).
  const authors: string[] = [];
  const authorIdx = new Map<string, number>();
  for (const sl of pres.slides) for (const c of sl.comments ?? []) if (!authorIdx.has(c.author)) authorIdx.set(c.author, authors.push(c.author) - 1);
  const authorLast = authors.map(() => 0);

  // Slaytlar
  const tableStyles = new Set<string>();
  pres.slides.forEach((slide, i) => {
    const path = `ppt/slides/slide${i + 1}.xml`;
    const ctx = makeCtx(path, media, true);
    const li = Math.max(0, pres.layouts.findIndex((l) => l.id === slide.layoutId));
    ctx.rels.add(REL.slideLayout, `../slideLayouts/slideLayout${li + 1}.xml`);
    tableStyleIds(slide.shapes, tableStyles);
    const tree = spTreeXml(slide.shapes, ctx);
    const trans = transitionXml(slide.transition);
    const timing = timingXml(slide, ctx.idMap);
    const ext = slideExtXml(slide, ctx.idMap);
    if (slide.notes.trim()) {
      const np = `ppt/notesSlides/notesSlide${i + 1}.xml`;
      add(np, notesSlideXml(slide.notes), CT.notesSlide);
      const nr = new RelsBuilder();
      nr.add(REL.notesMaster, "../notesMasters/notesMaster1.xml");
      nr.add(REL.slide, `../slides/slide${i + 1}.xml`);
      zip.file(`ppt/notesSlides/_rels/notesSlide${i + 1}.xml.rels`, nr.xml());
      ctx.rels.add(REL.notesSlide, `../notesSlides/notesSlide${i + 1}.xml`);
    }
    if (slide.comments?.length) {
      const cp = `ppt/comments/comment${i + 1}.xml`;
      const items = slide.comments
        .map((c) => {
          const a = authorIdx.get(c.author) ?? 0;
          const idx = ++authorLast[a];
          const dt = new Date(c.date || Date.now()).toISOString();
          return `<p:cm authorId="${a}" dt="${dt}" idx="${idx}"><p:pos x="${Math.round(c.x * 6)}" y="${Math.round(c.y * 6)}"/><p:text>${esc(c.text)}</p:text></p:cm>`;
        })
        .join("");
      add(cp, `<p:cmLst ${NS_DECL}>${items}</p:cmLst>`, CT.comments);
      ctx.rels.add(REL.comments, `../comments/comment${i + 1}.xml`);
    }
    add(
      path,
      `<p:sld ${NS_DECL}${attrs({ show: slide.hidden ? "0" : undefined, showMasterSp: slide.hideMasterShapes ? "0" : undefined })}><p:cSld>${bgXml(slide.background, ctx)}${tree}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${trans}${timing}${ext}</p:sld>`,
      CT.slide,
    );
    zip.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`, ctx.rels.xml());
  });

  // Sunu
  const pctx = makeCtx("ppt/presentation.xml", media, false);
  const masterRid = pctx.rels.add(REL.slideMaster, "slideMasters/slideMaster1.xml");
  const slideRids = pres.slides.map((_, i) => pctx.rels.add(REL.slide, `slides/slide${i + 1}.xml`));
  const notesRid = hasNotes ? pctx.rels.add(REL.notesMaster, "notesMasters/notesMaster1.xml") : null;
  pctx.rels.add(REL.presProps, "presProps.xml");
  pctx.rels.add(REL.viewProps, "viewProps.xml");
  pctx.rels.add(REL.theme, "theme/theme1.xml");
  pctx.rels.add(REL.tableStyles, "tableStyles.xml");
  if (authors.length) {
    pctx.rels.add(REL.commentAuthors, "commentAuthors.xml");
    const list = authors
      .map((name, id) => {
        const initials = name.split(/\s+/).filter(Boolean).map((w) => w[0]!.toLocaleUpperCase("tr-TR")).join("").slice(0, 3) || "?";
        return `<p:cmAuthor id="${id}" name="${esc(name)}" initials="${esc(initials)}" lastIdx="${authorLast[id]}" clrIdx="${id % 8}"/>`;
      })
      .join("");
    add("ppt/commentAuthors.xml", `<p:cmAuthorLst ${NS_DECL}>${list}</p:cmAuthorLst>`, CT.commentAuthors);
  }
  add("ppt/presentation.xml", presentationXml(pres, masterRid, notesRid, slideRids, pctx), CT.presentation);
  zip.file("ppt/_rels/presentation.xml.rels", pctx.rels.xml());
  add("ppt/presProps.xml", PRES_PROPS, CT.presProps);
  add("ppt/viewProps.xml", VIEW_PROPS, CT.viewProps);
  add("ppt/tableStyles.xml", tableStylesXml(tableStyles, pctx), CT.tableStyles);

  // Belge özellikleri
  add("docProps/core.xml", coreXml(pres), CT.core);
  add("docProps/app.xml", appXml(pres), CT.app);
  const root = new RelsBuilder();
  root.add(REL.officeDocument, "ppt/presentation.xml");
  root.add(REL.coreProps, "docProps/core.xml");
  root.add(REL.extProps, "docProps/app.xml");
  zip.file("_rels/.rels", root.xml());

  // Ortam
  const exts = new Set<string>();
  for (const f of media.files) {
    zip.file(f.path, f.bytes);
    exts.add(f.ext);
  }

  const defaults = [
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`,
    `<Default Extension="xml" ContentType="application/xml"/>`,
    ...[...exts].map((e) => `<Default Extension="${e}" ContentType="${EXT_MIME[e] ?? "application/octet-stream"}"/>`),
  ].join("");
  const ov = overrides.map(([p, c]) => `<Override PartName="${p}" ContentType="${c}"/>`).join("");
  zip.file("[Content_Types].xml", `${XML_HEAD}<Types xmlns="${NS.ct}">${defaults}${ov}</Types>`);

  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 }, mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
}

/** Asıl/düzen alt bilgi yer tutucularına alan (slayt numarası, tarih) ekler. */
function withFooterFields(shapes: Shape[]): Shape[] {
  return shapes.map((s) => {
    if (s.kind !== "sp" || !s.ph || (s.ph.type !== "sldNum" && s.ph.type !== "dt")) return s;
    const empty = !s.text || s.text.paragraphs.every((p) => !p.runs.some((r) => r.text));
    if (!empty) return s;
    const run = s.ph.type === "sldNum" ? { text: "‹#›", props: { field: "slidenum" as const } } : { text: "gg.aa.yyyy", props: { field: "datetime" as const } };
    return { ...s, text: { ...(s.text ?? { paragraphs: [] }), paragraphs: [{ runs: [run], level: 0, props: {} }] } } as SpShape;
  });
}
