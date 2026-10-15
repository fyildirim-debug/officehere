/**
 * OpenDocument sunusu (.odp / .otp) okuyucu. Sayfa boyutu, ana sayfalar
 * (arka plan, süs şekilleri, başlık/anahat stilleri), slaytlar, yer tutucular,
 * metin kutuları, şekiller, resimler, tablolar ve konuşmacı notları okunur.
 */
import { uid } from "../../model/factory";
import { newPresentation } from "../../model/presentation";
import type { Fill, LevelStyle, Presentation, Slide, SpShape, TextBody } from "../../model/types";
import { applyPrimaryMaster, chooseLayout, layoutResolver, type ImportMaster } from "../importBuild";
import { flatShapes } from "../../model/ops";
import { resolveHex } from "../../model/color";
import { levelResolver, resolveFont } from "../../model/textStyle";
import type { Color, LevelStyle as LS, SchemeColorName, Theme } from "../../model/types";
import { attr, kid, kids, parseXml } from "../xml";
import { DISPLAYABLE, IMAGE_MIME, lengthPx, OdpError, toBase64 } from "./common";
import { readFillProps, readShapes, type ShapeReadCtx } from "./shapes";
import { type Resolved, StyleSheet } from "./styles";
import { readOdpTransition, transitionDur } from "./transitions";
import { readOdpSettings } from "./showSettings";
import { listLevel, paragraphsText, paraPropsOf, readParagraphs, runPropsOf } from "./text";
import { LoadWarnings } from "../loadWarnings";
import { formatMB, openZip, PackageError, PartError, xmlHealth, ZIP_LIMITS, type SafeZip } from "../zipSafe";
import { OpcPackage } from "@officehere/core/opc";
import { objectPreviews, odfInfo } from "@officehere/core/odf";
import { kaynakBirak, kaynakKaydet } from "../rawPassthrough";
import { objectShape } from "./objects";

const MAX_SLIDES = 2000;
const MAX_IMAGE = 40 * 1024 * 1024;

function parse(xml: string, what: string): Element {
  const doc = parseXml(xml);
  const root = doc.documentElement;
  if (!root || root.localName === "parsererror" || doc.getElementsByTagName("parsererror").length) throw new OdpError(`${what} okunamadı (bozuk XML).`);
  return root;
}

const SCHEME: SchemeColorName[] = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];

/** OfficeHere'nin meta.xml'e yazdığı tema (yoksa ya da geçersizse undefined). */
function themeFromMeta(meta: Element | undefined): Theme | undefined {
  const el = kids(meta).find((e) => e.localName === "user-defined" && attr(e, "name") === "OfficeHere.Tema");
  if (!el) return undefined;
  try {
    const t = JSON.parse(el.textContent ?? "") as Partial<Theme>;
    const colors = {} as Theme["colors"];
    for (const k of SCHEME) {
      const v = t.colors?.[k];
      if (typeof v !== "string" || !/^#[0-9a-f]{6}$/i.test(v)) return undefined;
      colors[k] = v.toUpperCase();
    }
    const major = t.fonts?.major;
    const minor = t.fonts?.minor;
    if (typeof major !== "string" || typeof minor !== "string" || !major || !minor) return undefined;
    return { name: typeof t.name === "string" && t.name ? t.name.slice(0, 200) : "Tema", ...(typeof t.id === "string" ? { id: t.id } : {}), ...(typeof t.variant === "number" ? { variant: t.variant } : {}), colors, fonts: { major, minor } };
  } catch {
    return undefined;
  }
}

/**
 * Şekil düzeyi metin stillerinden, zincirin (asıl slayt + düzen) zaten verdiği
 * değerler atılır: ODF'de tema olmadığından okunan renk/yazı tipi mutlak gelir;
 * temanın çözdüğü değere eşit olanlar devralınan sayılır (tema değişince izlesin).
 */
function pruneInheritedLevels(pres: Presentation) {
  const theme = pres.theme;
  const same = (k: string, a: unknown, b: unknown): boolean => {
    if (a === undefined) return true;
    if (b === undefined) return false;
    if (k === "color") {
      const x = a as Color;
      const y = b as Color;
      return resolveHex(x, theme) === resolveHex(y, theme) && (x.alpha ?? 1) === (y.alpha ?? 1);
    }
    if (k === "font") return resolveFont(a as string, theme) === resolveFont(b as string, theme);
    return JSON.stringify(a) === JSON.stringify(b);
  };
  const prune = <T extends object>(own: T | undefined, base: T): T => {
    const out = {} as T;
    for (const [k, v] of Object.entries(own ?? {})) if (v !== undefined && !same(k, v, (base as Record<string, unknown>)[k])) (out as Record<string, unknown>)[k] = v;
    return out;
  };
  for (const slide of pres.slides) {
    const layout = pres.layouts.find((l) => l.id === slide.layoutId);
    for (const sh of flatShapes(slide.shapes)) {
      if (sh.kind !== "sp" || !sh.text?.levels?.length) continue;
      const chain = levelResolver({ ...sh, text: { ...sh.text, levels: undefined } }, { master: pres.master, layout });
      const levels: LS[] = sh.text.levels.map((lv, i) => ({ para: prune(lv?.para, chain(i).para), run: prune(lv?.run, chain(i).run) }));
      sh.text.levels = levels.some((l) => Object.keys(l.para).length || Object.keys(l.run).length) ? levels : undefined;
    }
  }
}

export interface OdpReadOptions {
  /** Açılış uyarıları (eksik/bozuk parçalar); verilmezse yok sayılır. */
  warnings?: LoadWarnings;
  /**
   * Özgün paket geçiş katmanı için kaydedilsin mi (varsayılan evet; tek kullanımlık doğrulama okumalarında hayır).
   * Kayıtlıysa gömülü nesneler özgün çerçeveleriyle taşınır, .odp'ye kaydederken ayarlar ve RDF üst verisi korunur.
   */
  kaynak?: boolean;
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

export async function readOdp(data: ArrayBuffer | Uint8Array, opts: OdpReadOptions = {}): Promise<Presentation> {
  const W = opts.warnings ?? new LoadWarnings();
  let pkg: SafeZip;
  try {
    pkg = await openZip(data, "Dosya geçerli bir OpenDocument (ZIP) paketi değil.");
  } catch (e) {
    if (e instanceof PackageError) throw new OdpError(e.message, e.detail);
    throw e;
  }
  // Özgün paket (geçiş katmanı): açılamazsa ya da ODF değilse taşıma olmadan okunur.
  const opc = opts.kaynak === false ? undefined : await OpcPackage.open(data).catch(() => undefined);
  const info = opc ? await odfInfo(opc).catch(() => null) : null;
  const kayit = opc && info ? await kaynakKaydet(opc) : undefined;
  try {
    const pres = await readOdpPackage(pkg, W, kayit?.id, kayit ? await objectPreviews(opc!).catch(() => new Map<string, string>()) : new Map<string, string>());
    if (kayit) pres.kaynak = kayit.id;
    return pres;
  } catch (e) {
    if (kayit?.yeni) kaynakBirak(kayit.id);
    throw e;
  }
}

async function readOdpPackage(pkg: SafeZip, W: LoadWarnings, src: string | undefined, previews: Map<string, string>): Promise<Presentation> {
  /** Metin parçası; bozuk ya da sınırı aşan parça atlanır (verilirse uyarı yazılır). */
  const text = async (p: string, label?: string) => {
    try {
      return await pkg.text(p);
    } catch (e) {
      if (!(e instanceof PartError)) throw e;
      if (label) W.add(label, `${p}: ${e.detail ?? e.message}`);
      else W.detail(`${p}: ${e.detail ?? e.message}`);
      return undefined;
    }
  };
  const mime = (await text("mimetype"))?.trim();
  if (mime && !mime.startsWith("application/vnd.oasis.opendocument.presentation")) throw new OdpError("Bu dosya bir OpenDocument sunusu değil.");
  const manifest = await text("META-INF/manifest.xml");
  if (manifest && /encryption-data/.test(manifest)) throw new OdpError("Bu sunu parola ile korunuyor (şifreli); açılamıyor.");
  const listed = (p: string) => !!manifest && new RegExp(`full-path="${p.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}"`).test(manifest);
  let contentXml: string | undefined;
  try {
    contentXml = await pkg.text("content.xml");
  } catch (e) {
    if (!(e instanceof PartError)) throw e;
    if (e.kind === "tooLarge") throw new OdpError(`Sunu içeriği (content.xml) olağandışı büyük${pkg.size("content.xml") ? ` (${formatMB(pkg.size("content.xml")!)})` : ""}; dosya bozuk ya da zararlı olabilir.`, e.detail);
    throw new OdpError("Sunu içeriği (content.xml) bozuk; dosya açılamadı.", e.detail);
  }
  if (!contentXml) throw new OdpError("Sunu içeriği (content.xml) bulunamadı.");
  // Kısmen ayrıştırılabilen içerikte (tarayıcılar okunabilen kısmı korur) okunabilen slaytlar açılır.
  const contentDoc = parseXml(contentXml);
  const contentHealth = xmlHealth(contentDoc, "document-content");
  if (contentHealth === "bad") throw new OdpError("Sunu içeriği okunamadı (bozuk XML).");
  const content = contentDoc.documentElement;
  if (contentHealth === "partial") W.add("Sunu içeriği bozuk; yalnız okunabilen slaytlar açıldı.", "content.xml: XML yalnız kısmen ayrıştırılabildi");
  const stylesXml = await text("styles.xml", "Stil bölümü (styles.xml) bozuk; asıl sayfa biçimleri varsayılanlarla açıldı.");
  let styles: Element | undefined;
  if (stylesXml) {
    const sd = parseXml(stylesXml);
    const h = xmlHealth(sd, "document-styles");
    if (h === "bad") W.add("Stil bölümü (styles.xml) bozuk; asıl sayfa biçimleri varsayılanlarla açıldı.", "styles.xml: XML ayrıştırılamadı");
    else {
      styles = sd.documentElement;
      if (h === "partial") W.add("Stil bölümü (styles.xml) bozuk; yalnız okunabilen kısmı kullanıldı.", "styles.xml: XML yalnız kısmen ayrıştırılabildi");
    }
  } else if (!pkg.has("styles.xml") && listed("styles.xml")) W.add("Stil bölümü (styles.xml) dosyada bulunamadı; asıl sayfa biçimleri varsayılanlarla açıldı.");

  // Resimler (paket içi): eşzamanlı erişim için önceden okunur.
  const images = new Map<string, string>();
  const badImages = new Set<string>();
  await Promise.all(
    pkg.paths
      .filter((p) => /\.(png|jpe?g|gif|bmp|svg|webp)$/i.test(p))
      .map(async (p) => {
        try {
          const bytes = await pkg.bytes(p, Math.min(MAX_IMAGE, ZIP_LIMITS.media));
          if (!bytes) return;
          const ext = p.split(".").pop()!.toLowerCase();
          images.set(p, `data:${IMAGE_MIME[ext]};base64,${toBase64(bytes)}`);
        } catch (e) {
          if (!(e instanceof PartError)) throw e;
          badImages.add(p);
          W.detail(`${p}: ${e.detail ?? e.message}`);
        }
      }),
  );
  if (badImages.size) W.add(`${badImages.size} resim bozuk ya da çok büyük olduğu için yüklenemedi.`);
  const missingImages = new Set<string>();
  const image: ShapeReadCtx["image"] = (href, mimeType, inline) => {
    if (inline) {
      const m = mimeType ?? "image/png";
      return DISPLAYABLE.has(m) ? `data:${m};base64,${inline.replace(/\s+/g, "")}` : undefined;
    }
    if (!href) return undefined;
    const path = href.replace(/^\.\//, "");
    if (mimeType && !DISPLAYABLE.has(mimeType)) return undefined;
    const got = images.get(path);
    if (got === undefined && !/^[a-z][a-z0-9+.-]*:/i.test(path) && !pkg.has(path) && !badImages.has(path)) missingImages.add(path);
    return got;
  };

  const stylesCommon = kid(styles, "styles");
  const stylesAuto = kid(styles, "automatic-styles");
  const fontsS = kid(styles, "font-face-decls");
  const fontsC = kid(content, "font-face-decls");
  const masterSheet = new StyleSheet([stylesAuto, stylesCommon], [fontsS, fontsC], stylesCommon);
  const contentSheet = new StyleSheet([kid(content, "automatic-styles"), stylesCommon], [fontsC, fontsS], stylesCommon);
  // Gömülü nesneler (kaynak kayıtlıysa): özgün çerçeve ve bölümünün otomatik stilleriyle
  const objectHook = (auto: Element | undefined): ShapeReadCtx["object"] => (src ? (el, base) => objectShape(el, base, src, auto, previews) : undefined);
  const masterObjects = objectHook(stylesAuto);
  const contentObjects = objectHook(kid(content, "automatic-styles"));

  const body = kid(kid(content, "body"), "presentation");
  if (!body) throw new OdpError("Dosyada sunu gövdesi yok.");
  const allPages = kids(body, "page");
  if (allPages.length > MAX_SLIDES) W.add(`Sunuda ${allPages.length.toLocaleString("tr-TR")} slayt var; yalnız ilk ${MAX_SLIDES.toLocaleString("tr-TR")} slayt açıldı.`);
  const pages = allPages.slice(0, MAX_SLIDES);
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
    const shapes = readShapes(mp, { sheet: masterSheet, image, link, count: { n: 0 }, placeholderLevels: placeholderLevels(name), object: masterObjects });
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
    try {
      pres.slides.push(readPage(page, i));
    } catch (e) {
      W.slide("corrupt", i + 1, `draw:page ${i + 1}: ${(e as Error)?.message ?? e}`);
      pres.slides.push({ id: slideIds[i], layoutId: layoutFor(primaryName, "blank").id, shapes: [], notes: "", animations: [] });
    }
  });
  function readPage(page: Element, i: number): Slide {
    const masterName = attr(page, "master-page-name");
    const ps = contentSheet.resolve("drawing-page", attr(page, "style-name"));
    const shapes = readShapes(page, { sheet: contentSheet, image, link, count: { n: 0 }, placeholderLevels: placeholderLevels(masterName ?? primaryName), object: contentObjects });
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
    return slide;
  }
  if (!pres.slides.length) throw new OdpError("Sunuda slayt bulunamadı.");
  // Uygulamanın kendi yazdığı tema (meta.xml) geri yüklenir; devralınan değerler temaya bağlanır.
  const metaXml = await text("meta.xml", "Belge özellikleri (başlık, yazar) okunamadı.");
  let metaEl: Element | undefined;
  if (metaXml) {
    try {
      metaEl = kid(parse(metaXml, "Özellikler"), "meta");
    } catch {
      metaEl = undefined;
    }
  }
  const saved = themeFromMeta(metaEl);
  if (saved) pres.theme = saved;
  pruneInheritedLevels(pres);
  if (missingImages.size) W.add(`${missingImages.size} resim pakette bulunamadı; bu resimler slaytlarda görünmüyor.`, `eksik resimler: ${[...missingImages].slice(0, 20).join(", ")}`);
  // Özel gösteriler ve gösteri ayarları.
  Object.assign(pres, readOdpSettings(body, byName, slideIds));

  // Belge özellikleri.
  if (metaEl) {
    try {
      const meta = metaEl;
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
