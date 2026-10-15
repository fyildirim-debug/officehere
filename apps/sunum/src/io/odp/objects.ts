/**
 * ODF geçiş katmanı (Sunum): .odp'deki gömülü nesneler (draw:object ile grafik/formül alt belgeleri, OLE) modelde
 * yedek görünümleriyle (yedek görüntüsü çizilebiliyorsa o, değilse yer tutucu resim) gösterilir; şeklin `odf`
 * alanında özgün çerçeve, andığı otomatik stiller ve görünüm imzası durur. .odp'ye kaydederken görünüm
 * değişmemişse özgün çerçeve (konum ve boyut güncel) yazılır, alt belge kaynak paketten kopyalanır
 * (@officehere/core/odf). .pptx'e yedek görünümü yazılır.
 */
import { odfObjectFromElement, replacementHref } from "@officehere/core/odf";
import type { PicShape, Shape } from "../../model/types";
import { placeholderPic, rawSig } from "../rawPassthrough";
import type { ShapeBaseProps } from "./shapes";

/** .odp'ye özgü gömülü nesneler başka biçimde yedek görünümleriyle yazıldı. */
export const ODF_OBJECTS_FALLBACK = "Gömülü grafikler ve nesneler bu biçimde saklanamadığı için yedek görünümleriyle (resim) kaydedildi; korumak için .odp olarak kaydedin.";
/** Görünümü düzenlenen gömülü nesne resim olarak yazıldı. */
export const ODF_OBJECTS_EDITED = "Düzenlenen gömülü nesneler yedek görünümleriyle (resim) kaydedildi; özgün nesne kaldırıldı.";
/** Gömülü nesnenin özgün dosyası bellekte yok. */
export const ODF_OBJECTS_SOURCE = "Gömülü nesnelerin özgün dosyası bulunamadığı için yedek görünümleriyle (resim) kaydedildi.";

/** Slaytlarda (gruplar dahil) .odp'den gelen gömülü nesne var mı? */
export function hasOdfShapes(holders: { shapes: Shape[] }[]): boolean {
  const any = (shapes: Shape[]): boolean => shapes.some((s) => !!s.odf || (s.kind === "group" && any(s.children)));
  return holders.some((h) => any(h.shapes));
}

/** Görünüm imzası: `odf` alanı dışındaki her şey (konum ve boyut hariç; bkz. rawSig). */
export function odfSig(s: Shape): string {
  return rawSig({ ...s, odf: undefined } as Shape);
}

/** Şeklin görünümü okunduğu andakiyle aynı mı (kayan nokta farkı yok sayılır)? */
export function odfSigMatches(s: Shape): boolean {
  const sig = s.odf?.sig;
  if (!sig) return false;
  const [h1, g1 = ""] = odfSig(s).split("|");
  const [h2, g2 = ""] = sig.split("|");
  if (h1 !== h2) return false;
  const a = g1 ? g1.split(",").map(Number) : [];
  const b = g2 ? g2.split(",").map(Number) : [];
  return a.length === b.length && a.every((v, i) => Math.abs(v - b[i]) < 2e-3);
}

/**
 * Gömülü nesne çerçevesi → yedek görünümlü resim şekli (`odf` ile). `autoStyles`: çerçevenin bölümünün otomatik
 * stilleri; `previews`: çizilebilen yedek görüntüler (yol → data: adresi).
 */
export function objectShape(el: Element, base: ShapeBaseProps, src: string, autoStyles: Element | null | undefined, previews: Map<string, string>): PicShape | undefined {
  let odf;
  try {
    odf = odfObjectFromElement(el, autoStyles, src);
  } catch {
    return undefined;
  }
  const rep = replacementHref(odf.xml);
  const pic = placeholderPic("ole", base);
  pic.rot = base.rot;
  const preview = rep ? previews.get(rep) : undefined;
  if (preview) pic.src = preview;
  return { ...pic, odf: { src, xml: odf.xml, styles: odf.styles, sig: odfSig(pic) } };
}
