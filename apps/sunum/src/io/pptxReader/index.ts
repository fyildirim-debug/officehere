/**
 * PPTX okuyucu: slayt boyutu, sıra, asıl slayt/düzen kalıtımı, tema renkleri
 * ve yazı tipleri, şekiller, resimler, tablolar, gruplar, arka planlar,
 * notlar, gizli slaytlar, geçişler ve animasyonlar.
 *
 * Okuma parça başına modüllere bölünmüştür; bu dosya yalnız sırayı kurar:
 * paket → sunu → asıl slayt + tema → ortam/grafik → asıl slaytlar ve düzenler
 * → slaytlar (açıklama, not) → bölümler, köprüler, özel gösteriler → sunu
 * düzeyi ayarlar. Uyarıların ve kimliklerin sırası bu sıraya bağlıdır.
 * Geçiş katmanı: özgün paket (OpcPackage) kimlikle kaydedilir (`Presentation.kaynak`);
 * tanınmayan parçalar (customXml, vbaProject, gömülü yazı tipleri…) kaydederken oradan
 * taşınır, tanınmayan çizim öğeleri `raw` alanlarıyla okunur (rawPassthrough.ts).
 */
import { OpcPackage } from "@officehere/core/opc";
import type { Presentation } from "../../model/types";
import { SIGNED_NOTICE } from "../formatInfo";
import { readEffects, type ReadCtx } from "../readDrawing";
import { LoadWarnings } from "../loadWarnings";
import { kaynakBirak, kaynakKaydet, themeSig } from "../rawPassthrough";
import { kid, type Rel } from "../xml";
import { commentAuthors } from "./comments";
import { readClrMap, type MkCtx } from "./common";
import { readMasters, resolveMaster } from "./master";
import { readCharts, readDiagrams, readMedia } from "./media";
import { readNotesMasterHf } from "./notes";
import { openPptxZip, pptxPackage, readContentTypes } from "./package";
import { headerFooterOf, openPresentation, readCoreProps, readCustomShows, readShowSettings, slideSize } from "./presentation";
import { applySections } from "./sections";
import { linkSlides, readSlides, resolveSlides } from "./slide";
import { readTheme, resolveTheme } from "./theme";

export { PptxError } from "./package";
export { SIGNED_NOTICE };

export interface PptxReadOptions {
  /** Açılış uyarıları (eksik/bozuk parçalar); verilmezse yok sayılır. */
  warnings?: LoadWarnings;
  /** Hasar olmayan bilgi iletileri (ör. dijital imza kaydedince kalkar). */
  notices?: string[];
  /** Özgün paket geçiş katmanı için kaydedilsin mi (varsayılan evet; tek kullanımlık doğrulama okumalarında hayır). */
  kaynak?: boolean;
}

/** PPTX baytlarını sunu modeline okur. Eksik/bozuk parçalar `opts.warnings`'e yazılır. */
export async function readPptx(data: ArrayBuffer | Uint8Array, opts: PptxReadOptions = {}): Promise<Presentation> {
  const W = opts.warnings ?? new LoadWarnings();
  const pk = pptxPackage(await openPptxZip(data), W);
  // Özgün paket (merkez dizin; baytlar kopyalanmaz). Açılamazsa geçiş katmanı olmadan okunur.
  const opc = opts.kaynak === false ? undefined : await OpcPackage.open(data).catch(() => undefined);
  const kayit = opc ? await kaynakKaydet(opc) : undefined;
  const rawSrc = kayit?.id;
  try {
    const pres = await readParts(pk, W, rawSrc);
    if (rawSrc) pres.kaynak = rawSrc;
    if (opc?.signed) opts.notices?.push(SIGNED_NOTICE);
    return pres;
  } catch (e) {
    if (kayit?.yeni) kaynakBirak(rawSrc);
    throw e;
  }
}

async function readParts(pk: ReturnType<typeof pptxPackage>, W: LoadWarnings, rawSrc: string | undefined): Promise<Presentation> {
  const contentTypes = await readContentTypes(pk);
  const { media, badMedia } = await readMedia(pk);
  const { diagrams, diagramRels } = rawSrc ? await readDiagrams(pk, contentTypes) : { diagrams: undefined, diagramRels: undefined };
  const missingMedia = new Set<string>();
  const pres = await openPresentation(pk);
  const { width, height } = slideSize(pres.doc);

  // Asıl slayt, tema ve tema efekt stilleri (ana asıl slaydın renk eşlemesiyle).
  const main = await resolveMaster(pres);
  const themeRef = await resolveTheme(pres, main);
  const { theme, lnWidths, edited: masterEdited, effectEls } = readTheme(themeRef.part.root ? themeRef.part.xml : undefined);
  if (rawSrc && themeRef.path && themeRef.part.root && !theme.id) theme.src = { src: rawSrc, part: themeRef.path, sig: themeSig(theme) };
  const clrMap = readClrMap(kid(main.doc, "clrMap"));
  const effectStyles = effectEls.map((e) => readEffects(e, { rels: new Map(), media: new Map(), clrMap, idMap: new Map(), lnWidths }));
  const charts = await readCharts(pk, media, clrMap, lnWidths);
  const mkCtx: MkCtx = (rels: Map<string, Rel>, map = clrMap, inherit?: ReadCtx["inherit"]): ReadCtx => ({ rels, media, clrMap: map, idMap: new Map(), inherit, lnWidths, charts, effectStyles, missingMedia, rawSrc, diagrams, diagramRels });
  const authors = commentAuthors(pres);

  const { master, set } = await readMasters({ pres, main, mkCtx, clrMap, theme, themePath: themeRef.path, rawSrc });
  const list = resolveSlides(pres, contentTypes);
  const { slides, slideAt } = await readSlides({ pk, mkCtx, clrMap, set, masterShapes: master.shapes, authors, rawSrc, media }, list);
  applySections(pres.doc, list.sldIds, slides);
  const idByPath = linkSlides(list, slideAt, slides);
  const { customShows, customShowIds } = readCustomShows(pres, idByPath);

  const notesHeaderFooter = await readNotesMasterHf(pres);
  const presProps = await readShowSettings(pres, customShowIds);
  const props = await readCoreProps(pres);

  // Pakette bulunamayan (bozuk olanlar dışında) resim/ortam başvuruları.
  const lost = [...missingMedia].filter((p) => !badMedia.has(p));
  if (lost.length) W.add(`${lost.length} resim ya da ortam dosyası pakette bulunamadı; yerlerinde boş çerçeve görünür.`, `eksik ortam: ${lost.slice(0, 20).join(", ")}`);
  const recoveredFromPackage = main.recovered || themeRef.recovered || list.recovered;
  if (recoveredFromPackage && !W.list().some((w) => w.startsWith("Sununun ilişki bölümü bozuk"))) W.add("Sununun ilişki bölümü bozuk; asıl slayt ve tema paket içinden bulunarak kullanıldı.");
  const headerFooter = headerFooterOf(slides);

  return { width, height, theme, master, layouts: set.layouts, slides, headerFooter, props, masterEdited, ...(customShows.length ? { customShows } : {}), ...presProps, ...(notesHeaderFooter ? { notesHeaderFooter } : {}) };
}
