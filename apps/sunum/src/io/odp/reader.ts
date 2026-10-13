/**
 * OpenDocument sunusu (.odp / .otp) okuyucu. Sayfa boyutu, ana sayfalar
 * (arka plan, süs şekilleri, başlık/anahat stilleri), slaytlar, yer tutucular,
 * metin kutuları, şekiller, resimler, tablolar ve konuşmacı notları okunur.
 */
import JSZip from "jszip";
import { uid } from "../../model/factory";
import { newPresentation } from "../../model/presentation";
import type { Fill, LevelStyle, Presentation, Slide, SpShape, TextBody } from "../../model/types";
import { applyPrimaryMaster, chooseLayout, layoutResolver, type ImportMaster } from "../importBuild";
import { attr, kid, kids, parseXml } from "../xml";
import { DISPLAYABLE, IMAGE_MIME, lengthPx, OdpError, toBase64 } from "./common";
import { readFillProps, readShapes, type ShapeReadCtx } from "./shapes";
import { type Resolved, StyleSheet } from "./styles";
import { readOdpTransition, transitionDur } from "./transitions";
import { listLevel, paragraphsText, paraPropsOf, readParagraphs, runPropsOf } from "./text";

const MAX_SLIDES = 2000;
const MAX_IMAGE = 40 * 1024 * 1024;

function parse(xml: string, what: string): Element {
  const doc = parseXml(xml);
  const root = doc.documentElement;
  if (!root || root.localName === "parsererror" || doc.getElementsByTagName("parsererror").length) throw new OdpError(`${what} okunamadı (bozuk XML).`);
  return root;
}

/** Ana sayfa stilinden düzey stili. */
function levelFrom(r: Resolved, sheet: StyleSheet): LevelStyle {
  const para = paraPropsOf(r.paragraph);
  const run = runPropsOf(r.text, sheet);
  return { para, run };
}

function mergeLevel(a: LevelStyle, b: Partial<LevelStyle>): LevelStyle {
  const strip = <T extends object>(o: T | undefined) => Object.fromEntries(Object.entries(o ?? {}).filter(([, v]) => v !== undefined)) as Partial<T>;
  return { para: { ...a.para, ...strip(b.para) }, run: { ...a.run, ...strip(b.run) } };
}

export async function readOdp(data: ArrayBuffer | Uint8Array): Promise<Presentation> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(data);
  } catch {
    throw new OdpError("Dosya geçerli bir OpenDocument (ZIP) paketi değil.");
  }
  const text = async (p: string) => {
    const f = zip.file(p);
    return f ? f.async("string") : undefined;
  };
  const mime = (await text("mimetype"))?.trim();
  if (mime && !mime.startsWith("application/vnd.oasis.opendocument.presentation")) throw new OdpError("Bu dosya bir OpenDocument sunusu değil.");
  const manifest = await text("META-INF/manifest.xml");
  if (manifest && /encryption-data/.test(manifest)) throw new OdpError("Bu sunu parola ile korunuyor (şifreli); açılamıyor.");
  const contentXml = await text("content.xml");
  if (!contentXml) throw new OdpError("Sunu içeriği (content.xml) bulunamadı.");
  const content = parse(contentXml, "Sunu içeriği");
  const stylesXml = await text("styles.xml");
  const styles = stylesXml ? parse(stylesXml, "Stil bölümü") : undefined;

  // Resimler (paket içi): eşzamanlı erişim için önceden okunur.
  const images = new Map<string, string>();
  await Promise.all(
    Object.keys(zip.files)
      .filter((p) => !zip.files[p].dir && /\.(png|jpe?g|gif|bmp|svg|webp)$/i.test(p))
      .map(async (p) => {
        const bytes = await zip.file(p)!.async("uint8array");
        if (bytes.length > MAX_IMAGE) return;
        const ext = p.split(".").pop()!.toLowerCase();
        images.set(p, `data:${IMAGE_MIME[ext]};base64,${toBase64(bytes)}`);
      }),
  );
  const image: ShapeReadCtx["image"] = (href, mimeType, inline) => {
    if (inline) {
      const m = mimeType ?? "image/png";
      return DISPLAYABLE.has(m) ? `data:${m};base64,${inline.replace(/\s+/g, "")}` : undefined;
    }
    if (!href) return undefined;
    const path = href.replace(/^\.\//, "");
    if (mimeType && !DISPLAYABLE.has(mimeType)) return undefined;
    return images.get(path);
  };

  const stylesCommon = kid(styles, "styles");
  const stylesAuto = kid(styles, "automatic-styles");
  const fontsS = kid(styles, "font-face-decls");
  const fontsC = kid(content, "font-face-decls");
  const masterSheet = new StyleSheet([stylesAuto, stylesCommon], [fontsS, fontsC], stylesCommon);
  const contentSheet = new StyleSheet([kid(content, "automatic-styles"), stylesCommon], [fontsC, fontsS], stylesCommon);

  const body = kid(kid(content, "body"), "presentation");
  if (!body) throw new OdpError("Dosyada sunu gövdesi yok.");
  const pages = kids(body, "page").slice(0, MAX_SLIDES);
  const masterPages = kids(kid(styles, "master-styles"), "master-page");

  // Sayfa boyutu: ilk slaydın ana sayfasının sayfa düzeni.
  const firstMaster = masterPages.find((m) => attr(m, "name") === attr(pages[0], "master-page-name")) ?? masterPages[0];
  const pl = kids(stylesAuto).find((s) => s.localName === "page-layout" && attr(s, "name") === attr(firstMaster, "page-layout-name"));
  const plp = kid(pl, "page-layout-properties");
  const clampDim = (v: number | undefined, def: number) => (v && v >= 96 && v <= 96 * 56 ? v : def);
  const width = clampDim(lengthPx(attr(plp, "page-width")), 1058.27);
  const height = clampDim(lengthPx(attr(plp, "page-height")), 595.28);

  const pres = newPresentation("varsayilan", 0, { width, height });
  pres.slides = [];
  pres.props = {};

  // Slayt adları → kimlikler (iç köprüler için).
  const slideIds = pages.map(() => uid("sl"));
  const byName = new Map<string, string>();
  pages.forEach((p, i) => {
    const n = attr(p, "name");
    if (n && !byName.has(n)) byName.set(n, slideIds[i]);
  });
  const link = (href: string) => (href.startsWith("#") ? (byName.has(decodeURIComponent(href.slice(1))) ? `#slide:${byName.get(decodeURIComponent(href.slice(1)))}` : undefined) : href);
  const count = { n: 0 };

  // Ana sayfa stilleri: "<ana sayfa>-title", "-subtitle", "-outline1..9".
  const presStyle = (master: string, suffix: string) => masterSheet.resolve("presentation", `${master}-${suffix}`);
  const placeholderLevels = (master: string) => (cls: string, own: Resolved): TextBody["levels"] => {
    const ownLv = levelFrom(own, contentSheet);
    if (cls === "outline" || cls === "text") {
      return Array.from({ length: 9 }, (_, i) => {
        const st = presStyle(master, `outline${i + 1}`);
        let lv = levelFrom(st, masterSheet);
        const ll = listLevel(st.listStyle, i, masterSheet);
        if (ll) lv = { para: { ...lv.para, bullet: ll.bullet, marL: ll.marL, indent: ll.indent }, run: lv.run };
        return mergeLevel(lv, ownLv);
      });
    }
    const suffix = cls === "subtitle" ? "subtitle" : cls === "title" ? "title" : cls;
    const st = presStyle(master, suffix);
    const lv = mergeLevel(levelFrom(st, masterSheet), ownLv);
    if (cls === "subtitle" || cls === "title") lv.para = { ...lv.para, bullet: { type: "none" }, marL: lv.para.marL ?? 0, indent: lv.para.indent ?? 0 };
    return Array.from({ length: 9 }, () => structuredClone(lv));
  };

  const pageFill = (r: Resolved, sheet: StyleSheet): Fill | undefined => {
    const f = r.page["fill"];
    if (!f || f === "none") return undefined;
    const ctx = { sheet, image, count, link } as ShapeReadCtx;
    const fill = readFillProps(r.page, sheet, ctx, "none");
    return fill.type === "none" ? undefined : fill;
  };

  // Ana sayfalar.
  const masters = new Map<string, ImportMaster>();
  for (const mp of masterPages) {
    const name = attr(mp, "name") ?? "";
    const bgStyle = masterSheet.resolve("drawing-page", attr(mp, "style-name"));
    const bg = pageFill(bgStyle, masterSheet) ?? { type: "solid", color: { rgb: "#FFFFFF" } };
    const shapes = readShapes(mp, { sheet: masterSheet, image, link, count: { n: 0 }, placeholderLevels: placeholderLevels(name) });
    masters.set(name, {
      key: name,
      display: attr(mp, "display-name") ?? name,
      bg,
      deco: shapes.filter((s) => !(s.kind === "sp" && s.ph)),
      title: shapes.find((s): s is SpShape => s.kind === "sp" && s.ph?.type === "title"),
      body: shapes.find((s): s is SpShape => s.kind === "sp" && s.ph?.type === "body"),
    });
  }
  const primaryName = attr(pages[0], "master-page-name") ?? attr(masterPages[0], "name") ?? "";
  const primary = masters.get(primaryName);

  // Asıl slayt: arka plan, süsler, yer tutucu konumları ve metin stilleri.
  if (primary) {
    applyPrimaryMaster(pres, primary);
    const ts = pres.master.textStyles;
    const t = presStyle(primaryName, "title");
    if (Object.keys(t.text).length) ts.title = ts.title.map((l, i) => (i === 0 ? mergeLevel(l, levelFrom(t, masterSheet)) : l));
    ts.body = ts.body.map((l, i) => {
      const st = presStyle(primaryName, `outline${i + 1}`);
      if (!Object.keys(st.text).length) return l;
      let lv = mergeLevel(l, levelFrom(st, masterSheet));
      const ll = listLevel(st.listStyle, i, masterSheet);
      if (ll) lv = { para: { ...lv.para, bullet: ll.bullet, marL: ll.marL ?? lv.para.marL, indent: ll.indent ?? lv.para.indent }, run: lv.run };
      return lv;
    });
    const tf = masterSheet.fontOf(t.text);
    const bf = masterSheet.fontOf(presStyle(primaryName, "outline1").text);
    if (tf) pres.theme = { ...pres.theme, fonts: { ...pres.theme.fonts, major: tf } };
    if (bf) pres.theme = { ...pres.theme, fonts: { ...pres.theme.fonts, minor: bf } };
    pres.theme = { ...pres.theme, name: primary.display || pres.theme.name, id: undefined, variant: undefined };
  }
  const layoutFor = layoutResolver(pres, masters, primary);

  // Slaytlar.
  pages.forEach((page, i) => {
    const masterName = attr(page, "master-page-name");
    const ps = contentSheet.resolve("drawing-page", attr(page, "style-name"));
    const shapes = readShapes(page, { sheet: contentSheet, image, link, count: { n: 0 }, placeholderLevels: placeholderLevels(masterName ?? primaryName) });
    const layout = layoutFor(masterName, chooseLayout(shapes));
    const notesEl = kid(page, "notes");
    let notes = "";
    if (notesEl) {
      const frames = kids(notesEl, "frame");
      const nf = frames.find((f) => attr(f, "class") === "notes") ?? frames.find((f) => kid(f, "text-box"));
      const tb = kid(nf, "text-box");
      if (tb) notes = paragraphsText(readParagraphs(tb, { sheet: contentSheet }));
    }
    const name = attr(page, "name");
    const slide: Slide = {
      id: slideIds[i],
      layoutId: layout.id,
      shapes,
      notes,
      animations: [],
    };
    if (name && !/^(page|slide|slayt|sayfa)\s*\d+$/i.test(name)) slide.name = name;
    const bg = pageFill(ps, contentSheet);
    if (bg) slide.background = bg;
    if (ps.page["visibility"] === "hidden") slide.hidden = true;
    const tr = readOdpTransition(ps.page, transitionDur(page));
    if (tr) slide.transition = tr;
    if (ps.page["background-objects-visible"] === "false") slide.hideMasterShapes = true;
    pres.slides.push(slide);
  });
  if (!pres.slides.length) throw new OdpError("Sunuda slayt bulunamadı.");

  // Belge özellikleri.
  const metaXml = await text("meta.xml");
  if (metaXml) {
    try {
      const meta = kid(parse(metaXml, "Özellikler"), "meta");
      const val = (n: string) => kid(meta, n)?.textContent?.trim() || undefined;
      // Kesirli saniyeler 3 haneye indirilir (Date.parse uyumu).
      const date = (n: string) => val(n)?.replace(/(\.\d{3})\d+/, "$1");
      pres.props = { title: val("title"), subject: val("subject"), creator: val("initial-creator") ?? val("creator"), created: date("creation-date"), modified: date("date") };
    } catch {
      // Özellikler okunamazsa boş kalır.
    }
  }
  return pres;
}
