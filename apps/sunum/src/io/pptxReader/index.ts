/**
 * PPTX okuyucu: slayt boyutu, sıra, asıl slayt/düzen kalıtımı, tema renkleri
 * ve yazı tipleri, şekiller, resimler, tablolar, gruplar, arka planlar,
 * notlar, gizli slaytlar, geçişler ve animasyonlar.
 *
 * Okuma parça başına modüllere bölünmüştür; bu dosya yalnız sırayı kurar:
 * paket → sunu → asıl slayt + tema → ortam/grafik → asıl slaytlar ve düzenler
 * → slaytlar (açıklama, not) → bölümler, köprüler, özel gösteriler → sunu
 * düzeyi ayarlar. Uyarıların ve kimliklerin sırası bu sıraya bağlıdır.
 * Geçiş kancası (dalga 3): okunan parçalar dışındaki (tanınmayan) parçalar
 * (customXml, vbaProject, gömülü yazı tipleri…) burada pakete bakılarak toplanacak.
 */
import type { Presentation } from "../../model/types";
import { readEffects, type ReadCtx } from "../readDrawing";
import { LoadWarnings } from "../loadWarnings";
import { kid, type Rel } from "../xml";
import { commentAuthors } from "./comments";
import { readClrMap, type MkCtx } from "./common";
import { readMasters, resolveMaster } from "./master";
import { readCharts, readMedia } from "./media";
import { readNotesMasterHf } from "./notes";
import { openPptxZip, pptxPackage, readContentTypes } from "./package";
import { headerFooterOf, openPresentation, readCoreProps, readCustomShows, readShowSettings, slideSize } from "./presentation";
import { applySections } from "./sections";
import { linkSlides, readSlides, resolveSlides } from "./slide";
import { readTheme, resolveTheme } from "./theme";

export { PptxError } from "./package";

export interface PptxReadOptions {
  /** Açılış uyarıları (eksik/bozuk parçalar); verilmezse yok sayılır. */
  warnings?: LoadWarnings;
}

/** PPTX baytlarını sunu modeline okur. Eksik/bozuk parçalar `opts.warnings`'e yazılır. */
export async function readPptx(data: ArrayBuffer | Uint8Array, opts: PptxReadOptions = {}): Promise<Presentation> {
  const W = opts.warnings ?? new LoadWarnings();
  const pk = pptxPackage(await openPptxZip(data), W);
  const contentTypes = await readContentTypes(pk);
  const { media, badMedia } = await readMedia(pk);
  const missingMedia = new Set<string>();
  const pres = await openPresentation(pk);
  const { width, height } = slideSize(pres.doc);

  // Asıl slayt, tema ve tema efekt stilleri (ana asıl slaydın renk eşlemesiyle).
  const main = await resolveMaster(pres);
  const themeRef = await resolveTheme(pres, main);
  const { theme, lnWidths, edited: masterEdited, effectEls } = readTheme(themeRef.part.root ? themeRef.part.xml : undefined);
  const clrMap = readClrMap(kid(main.doc, "clrMap"));
  const effectStyles = effectEls.map((e) => readEffects(e, { rels: new Map(), media: new Map(), clrMap, idMap: new Map(), lnWidths }));
  const charts = await readCharts(pk, media, clrMap, lnWidths);
  const mkCtx: MkCtx = (rels: Map<string, Rel>, map = clrMap, inherit?: ReadCtx["inherit"]): ReadCtx => ({ rels, media, clrMap: map, idMap: new Map(), inherit, lnWidths, charts, effectStyles, missingMedia });
  const authors = commentAuthors(pres);

  const { master, set } = await readMasters({ pres, main, mkCtx, clrMap, theme, themePath: themeRef.path });
  const list = resolveSlides(pres, contentTypes);
  const { slides, slideAt } = await readSlides({ pk, mkCtx, clrMap, set, masterShapes: master.shapes, authors }, list);
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
