/**
 * PPTX yazıcı (ECMA-376 PresentationML, JSZip ile). Asıl slayt, düzenler,
 * tema, slaytlar, notlar, ortam dosyaları, geçişler ve animasyonlar yazılır.
 *
 * Geçiş katmanı: sunu açılan bir paketten geliyorsa (`pres.kaynak`) modelin tanımadığı
 * öğeler özgün XML'leriyle (ham şekiller, slayt extLst/denetimleri, presentation.xml yaması)
 * ve parçalar kaynak paketten taşınır (core `carryParts`; slaytlar kaynak parçalarına eşlenir,
 * silinen slaytın parçaları taşınmaz).
 */
import JSZip from "jszip";
import { carryParts, type CarryReport } from "@officehere/core/opc";
import { flatShapes } from "../model/ops";
import type { ChartData, Layout, Master, Presentation, Shape, SpShape, Theme } from "../model/types";
import { kaynakPaket, patchPresentationXml, RawWriter } from "./rawPassthrough";
import { bgXml, spTreeXml, type ChartStore, type MediaStore, type WriteCtx } from "./writeDrawing";
import { chartWorkbook, chartXml } from "./writeChart";
import { appXml, CLR_MAP, coreXml, notesMasterXml, notesSlideXml, NS_DECL, presentationXml, samplePrompt, tableStylesXml, themeXml, txStylesXml, VIEW_PROPS } from "./writeParts";
import { presPropsXml } from "./presProps";
import { slideExtXml, timingXml, transitionXml } from "./writeTiming";
import { attrs, esc, NS, REL, RelsBuilder, relsPathFor, XML_HEAD } from "./xml";

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
  chart: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
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

/** Video/ses uzantıları ve içerik türleri. */
export const MEDIA_MIME: Record<string, string> = {
  mp4: "video/mp4",
  m4v: "video/mp4",
  webm: "video/webm",
  ogv: "video/ogg",
  mov: "video/quicktime",
  mp3: "audio/mpeg",
  m4a: "audio/mp4",
  wav: "audio/wav",
  ogg: "audio/ogg",
  oga: "audio/ogg",
  aac: "audio/aac",
  flac: "audio/flac",
};

const MEDIA_EXT: Record<string, string> = {
  "video/mp4": "mp4",
  "video/webm": "webm",
  "video/ogg": "ogv",
  "video/quicktime": "mov",
  "audio/mpeg": "mp3",
  "audio/mp3": "mp3",
  "audio/mp4": "m4a",
  "audio/x-m4a": "m4a",
  "audio/wav": "wav",
  "audio/x-wav": "wav",
  "audio/wave": "wav",
  "audio/ogg": "ogg",
  "audio/aac": "aac",
  "audio/flac": "flac",
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
  private mediaCount = 0;
  addMedia(src: string): string | null {
    const hit = this.map.get(src);
    if (hit) return hit;
    const d = dataUrlToBytes(src);
    if (!d) return null;
    const ext = MEDIA_EXT[d.mime] ?? (d.mime.startsWith("audio/") ? "mp3" : "mp4");
    const path = `ppt/media/media${++this.mediaCount}.${ext}`;
    this.files.push({ path, bytes: d.bytes, ext });
    this.map.set(src, path);
    return path;
  }
}

class Charts implements ChartStore {
  list: ChartData[] = [];
  /** Kaynak grafik parçası → yazılan grafik (biçem/renk parçaları taşınsın). */
  sources: { src: string; part: string; out: string }[] = [];
  add(chart: ChartData, src?: { src: string; part: string }): string {
    this.list.push(chart);
    const out = `ppt/charts/chart${this.list.length}.xml`;
    if (src) this.sources.push({ ...src, out });
    return out;
  }
}

function makeCtx(part: string, media: Media, slideContent: boolean, slideNumbers?: Map<string, number>, charts?: Charts, raw?: RawWriter): WriteCtx {
  let id = 1;
  return { rels: new RelsBuilder(), media, part, nextId: () => ++id, idMap: new Map(), slideContent, slideNumbers, charts, raw };
}

function tableStyleIds(shapes: Shape[], out: Set<string>) {
  for (const s of flatShapes(shapes)) if (s.kind === "table") out.add(s.style.styleId);
}

/** Ana bölüm içerik türleri: sunu, slayt gösterisi (.ppsx), şablon (.potx). */
export const MAIN_CT = {
  presentation: CT.presentation,
  slideshow: "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml",
  template: "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
} as const;

const PACKAGE_MIME = {
  presentation: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  slideshow: "application/vnd.openxmlformats-officedocument.presentationml.slideshow",
  template: "application/vnd.openxmlformats-officedocument.presentationml.template",
} as const;

export type PptxKind = keyof typeof MAIN_CT;

/** Makrolu ana bölüm içerik türleri (.pptm / .ppsm / .potm). */
const MACRO_CT: Record<PptxKind, string> = {
  presentation: "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
  slideshow: "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml",
  template: "application/vnd.ms-powerpoint.template.macroEnabled.main+xml",
};

export interface PptxWriteOptions {
  /** .pptx / .ppsx / .potx paketi. */
  kind?: PptxKind;
  /** Makrolu hedef (.pptm): kaynaktaki VBA projesi korunur, ana bölüm türü makrolu olur. */
  macro?: boolean;
  /** Kaydetme uyarıları (Türkçe): düz gruba çevrilen SmartArt, kaldırılan imza/makro… */
  warnings?: string[];
}

/** Paket düzeyi taşıma raporundan kullanıcı uyarıları. */
const DROP_WARN: Partial<Record<CarryReport["dropped"][number]["reason"], string>> = {
  signature: "Dijital imza kaldırıldı; kaydedilen dosya imzasız.",
  macro: "Makrolar bu biçimde saklanamadığı için kaldırıldı (.pptm olarak kaydedin).",
  comments: "Yeni tür açıklamalar (yanıt zinciri) bu sürümde korunamadı ve kaldırıldı.",
};

/** Sunuyu PPTX bayt dizisine yazar (kind: .pptx / .ppsx / .potx paketi). */
export async function writePptx(pres: Presentation, opts: PptxWriteOptions = {}): Promise<Uint8Array> {
  const kind = opts.kind ?? "presentation";
  const raw = new RawWriter();
  const source = kaynakPaket(pres.kaynak);
  const sourceNotesMaster = source?.mainPart ? (await source.rels(source.mainPart)).some((r) => r.type.endsWith("/notesMaster")) : false;
  // Ham XML'deki slayt başvuruları yeni slayt parçalarına bağlanır (silinen slayta başvuran ham öğe yedek görünümle yazılır).
  pres.slides.forEach((s, i) => {
    const key = s.raw ? `${s.raw.src}|${s.raw.part}` : "";
    if (key && !raw.slideParts.has(key)) raw.slideParts.set(key, `ppt/slides/slide${i + 1}.xml`);
  });
  const zip = new JSZip();
  // OPC okuyucularının çoğu içerik türleri parçasını ilk sırada bekler.
  zip.file("[Content_Types].xml", "");
  const media = new Media();
  const charts = new Charts();
  const overrides: [string, string][] = [];
  const add = (path: string, xml: string, ct?: string) => {
    zip.file(path, XML_HEAD + xml);
    if (ct) overrides.push([`/${path}`, ct]);
  };

  // Tema: içe aktarılan tema değişmediyse özgün tema parçası (efekt/dolgu stilleri ve resimleriyle).
  const writeTheme = async (path: string, t: Theme, generate: () => string) => {
    const kept = await raw.theme(t, path);
    add(path, kept ? kept.xml : generate(), CT.theme);
    if (kept?.rels) zip.file(relsPathFor(path), kept.rels);
  };
  await writeTheme("ppt/theme/theme1.xml", pres.theme, () => themeXml(pres.theme, pres.theme.name, !!pres.masterEdited));
  add("ppt/theme/theme2.xml", themeXml(pres.theme, "Not Teması"), CT.theme);

  // Asıl slaytlar: sununun asıl slaydı ve seçili slaytlara uygulanmış her tema için ikincil asıl slayt
  // (arka planı ve süsleri düzenlere taşınmış; kendi tema parçasıyla). Kaynakta ayrı asıl slayda bağlı
  // düzenler (Layout.master) yine ayrı asıl slayt altında yazılır.
  const designs: { key: string; theme: Theme; master: Master; layouts: Layout[] }[] = [{ key: "", theme: pres.theme, master: pres.master, layouts: pres.layouts.filter((l) => !l.theme && !l.master) }];
  for (const l of pres.layouts) {
    if (!l.theme && !l.master) continue;
    const key = l.master ?? `${l.theme!.name}|${l.theme!.id ?? ""}`;
    const d = designs.find((x, i) => i > 0 && x.key === key);
    if (d) d.layouts.push(l);
    else designs.push({ key, theme: l.theme ?? pres.theme, master: { background: l.background ?? pres.master.background, shapes: [], textStyles: pres.master.textStyles }, layouts: [l] });
  }
  const themeFileOf = (mi: number) => (mi === 0 ? "theme1.xml" : `theme${mi + 2}.xml`);
  for (const [mi, d] of designs.entries()) if (mi > 0) await writeTheme(`ppt/theme/${themeFileOf(mi)}`, d.theme, () => themeXml(d.theme, d.theme.name));
  const layoutOrder: Layout[] = designs.flatMap((d) => d.layouts);
  const layoutFile = (l: Layout) => `slideLayout${layoutOrder.indexOf(l) + 1}.xml`;
  let sldId = 2147483648;
  const masterIds: number[] = [];
  designs.forEach((d, mi) => {
    const n = mi + 1;
    const themeFile = themeFileOf(mi);
    const masterPath = `ppt/slideMasters/slideMaster${n}.xml`;
    const mctx = makeCtx(masterPath, media, false, undefined, charts, raw);
    masterIds.push(sldId++);
    const layoutIdLst = d.layouts.map((l) => `<p:sldLayoutId id="${sldId++}" r:id="${mctx.rels.add(REL.slideLayout, `../slideLayouts/${layoutFile(l)}`)}"/>`).join("");
    const masterTree = spTreeXml(withFooterFields(d.master.shapes), mctx, { promptAsText: (s) => samplePrompt(s, true) });
    const masterTx = txStylesXml(d.master.textStyles, mctx);
    mctx.rels.add(REL.theme, `../theme/${themeFile}`);
    add(masterPath, `<p:sldMaster ${NS_DECL}><p:cSld>${bgXml(d.master.background, mctx)}${masterTree}</p:cSld>${CLR_MAP}<p:sldLayoutIdLst>${layoutIdLst}</p:sldLayoutIdLst>${masterTx}</p:sldMaster>`, CT.master);
    zip.file(`ppt/slideMasters/_rels/slideMaster${n}.xml.rels`, mctx.rels.xml());

    // Düzenler
    for (const l of d.layouts) {
      const file = layoutFile(l);
      const path = `ppt/slideLayouts/${file}`;
      const ctx = makeCtx(path, media, false, undefined, charts, raw);
      ctx.rels.add(REL.slideMaster, `../slideMasters/slideMaster${n}.xml`);
      const tree = spTreeXml(withFooterFields(l.shapes), ctx, { promptAsText: (s) => samplePrompt(s, false) });
      const type = l.type === "cust" ? undefined : l.type;
      add(
        path,
        `<p:sldLayout ${NS_DECL}${attrs({ type, preserve: "1", showMasterSp: l.showMasterShapes ? undefined : "0" })}><p:cSld name="${esc(l.name)}">${bgXml(l.background, ctx)}${tree}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
        CT.layout,
      );
      zip.file(`ppt/slideLayouts/_rels/${file}.rels`, ctx.rels.xml());
    }
  });

  // Not asıl sayfası
  // Not sayfası/dinleyici notu üst-alt bilgisi ayarlıysa not asıl sayfası notsuz sunuda da yazılır.
  // Kaynakta not asıl sayfası / not sayfası olan slaytlar notsuz da yazılır (parçalar kaybolmasın).
  const notesPage = (s: Presentation["slides"][number]) => !!s.notes.trim() || !!s.raw?.notesPage;
  const hasNotes = pres.slides.some(notesPage) || !!pres.notesHeaderFooter || sourceNotesMaster;
  if (hasNotes) {
    add("ppt/notesMasters/notesMaster1.xml", notesMasterXml(pres.notesHeaderFooter), CT.notesMaster);
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
  const slideNumbers = new Map(pres.slides.map((s, i) => [s.id, i + 1]));
  pres.slides.forEach((slide, i) => {
    const path = `ppt/slides/slide${i + 1}.xml`;
    const ctx = makeCtx(path, media, true, slideNumbers, charts, raw);
    const sl = layoutOrder.find((l) => l.id === slide.layoutId) ?? layoutOrder[0];
    ctx.rels.add(REL.slideLayout, `../slideLayouts/${layoutFile(sl)}`);
    tableStyleIds(slide.shapes, tableStyles);
    const tree = spTreeXml(slide.shapes, ctx);
    const trans = transitionXml(slide.transition);
    const timing = timingXml(slide, ctx.idMap);
    const kept = raw.slide(slide, ctx);
    const ext = mergeExtLst(kept.ext, slideExtXml(slide, ctx.idMap));
    if (notesPage(slide)) {
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
      `<p:sld ${NS_DECL}${attrs({ show: slide.hidden ? "0" : undefined, showMasterSp: slide.hideMasterShapes ? "0" : undefined })}><p:cSld${attrs({ name: slide.name })}>${bgXml(slide.background, ctx)}${tree}${kept.cSld}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${trans}${timing}${ext}</p:sld>`,
      CT.slide,
    );
    zip.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`, ctx.rels.xml());
  });

  // Sunu
  const pctx = makeCtx("ppt/presentation.xml", media, false);
  const masterRids = designs.map((_, i) => pctx.rels.add(REL.slideMaster, `slideMasters/slideMaster${i + 1}.xml`));
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
  // Kaynaktaki dinleyici notu asıl sayfası, gömülü yazı tipleri, kılavuzlar (p15:sldGuideLst)… eklenir.
  const presXml = await patchPresentationXml(presentationXml(pres, masterRids.map((rid, i) => [masterIds[i], rid] as [number, string]), notesRid, slideRids, pctx), pres.kaynak, raw, pctx);
  add("ppt/presentation.xml", presXml, opts.macro ? MACRO_CT[kind] : MAIN_CT[kind]);
  zip.file("ppt/_rels/presentation.xml.rels", pctx.rels.xml());
  add("ppt/presProps.xml", presPropsXml(pres), CT.presProps);
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

  // Grafikler: chartN.xml + gömülü çalışma kitabı.
  let hasXlsx = false;
  for (let i = 0; i < charts.list.length; i++) {
    const cp = `ppt/charts/chart${i + 1}.xml`;
    const xp = `ppt/embeddings/Calisma_Sayfasi${i + 1}.xlsx`;
    const rels = new RelsBuilder();
    const rid = rels.add(REL.package, `../embeddings/Calisma_Sayfasi${i + 1}.xlsx`);
    zip.file(cp, chartXml(charts.list[i], rid));
    overrides.push([`/${cp}`, CT.chart]);
    zip.file(`ppt/charts/_rels/chart${i + 1}.xml.rels`, rels.xml());
    zip.file(xp, await chartWorkbook(charts.list[i]));
    hasXlsx = true;
  }

  // Ortam
  const exts = new Set<string>();
  for (const f of media.files) {
    zip.file(f.path, f.bytes);
    exts.add(f.ext);
  }

  const defaults = [
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`,
    `<Default Extension="xml" ContentType="application/xml"/>`,
    ...[...exts].map((e) => `<Default Extension="${e}" ContentType="${EXT_MIME[e] ?? MEDIA_MIME[e] ?? "application/octet-stream"}"/>`),
    ...(hasXlsx ? ['<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'] : []),
  ].join("");
  const ov = overrides.map(([p, c]) => `<Override PartName="${p}" ContentType="${c}"/>`).join("");
  zip.file("[Content_Types].xml", `${XML_HEAD}<Types xmlns="${NS.ct}">${defaults}${ov}</Types>`);

  const out = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 }, mimeType: PACKAGE_MIME[kind] });
  const bytes = await carryFromSource(pres, out, raw, charts, !!opts.macro);
  opts.warnings?.push(...raw.warnings);
  return bytes;
}

/**
 * Paket düzeyi taşıma: modelin tanımadığı parçalar (makro, özel XML, özel özellikler, yazıcı
 * ayarları, VML çizimleri…) ve ham öğelerin kopyaları kaynak paketten eklenir. Slaytlar ve
 * grafikler kaynak parçalarına eşlenir; eşlenmeyen (silinen) slaydın parçaları taşınmaz.
 */
async function carryFromSource(pres: Presentation, out: Uint8Array, raw: RawWriter, charts: Charts, macro: boolean): Promise<Uint8Array> {
  const source = kaynakPaket(pres.kaynak);
  if (pres.kaynak && !source) raw.warnings.add("Özgün dosya bellekte bulunamadığı için modelin tanımadığı bölümler (makrolar, özel veriler…) taşınamadı.");
  if (!source && !raw.carry.size) return out;
  const partMap: [string, string][] = [];
  pres.slides.forEach((s, i) => {
    if (s.raw && s.raw.src === pres.kaynak) partMap.push([s.raw.part, `ppt/slides/slide${i + 1}.xml`]);
  });
  for (const c of charts.sources) if (c.src === pres.kaynak) partMap.push([c.part, c.out]);
  // Açıklama yazarları: yazıcı yazmadıysa (açıklamasız sunu) kaynaktaki liste korunur.
  const { bytes, report } = await carryParts({ app: "sunum", source: source ?? null, output: out, macro, partMap, carry: raw.carry, overrides: { carry: ["commentAuthors"] } });
  for (const d of report.dropped) {
    const w = DROP_WARN[d.reason];
    if (w) raw.warnings.add(w);
  }
  return bytes;
}

/** Slaydın p:extLst'i: kaynaktaki uzantılar + uygulamanınki. */
function mergeExtLst(kept: string[], ours: string): string {
  if (!kept.length) return ours;
  const inner = ours.replace(/^<p:extLst>/, "").replace(/<\/p:extLst>$/, "");
  return `<p:extLst>${kept.join("")}${inner}</p:extLst>`;
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
