/**
 * Bölümler (presentation.xml extLst → p14:sectionLst): her bölümün ilk
 * slaydına bölüm işareti konur, ardından bölümler normalleştirilir.
 */
import { uid } from "../../model/factory";
import { normalizeSections } from "../../model/sections";
import type { Slide } from "../../model/types";
import { attr, kid, kids } from "../xml";

/** `slides` yerinde güncellenir; `sldIds`: sunudaki p:sldId öğeleri (slaytlarla aynı sırada). */
export function applySections(presDoc: Element, sldIds: Element[], slides: Slide[]): void {
  const sldIdNum = new Map<string, Slide>();
  sldIds.forEach((sid, i) => {
    const sl = slides[i];
    const num = attr(sid, "id");
    if (sl && num) sldIdNum.set(num, sl);
  });
  for (const ext of kids(kid(presDoc, "extLst"), "ext")) {
    const lst = kid(ext, "sectionLst");
    // Geçiş kancası (dalga 3): bölüm dışı sunu uzantıları (p15:sldGuideLst…) burada atlanıyor.
    if (!lst) continue;
    for (const sec of kids(lst, "section")) markSection(sec, sldIdNum, slides);
  }
  const sectioned = normalizeSections(slides);
  if (sectioned !== slides) slides.splice(0, slides.length, ...sectioned);
}

/** Bölümün ilk (bulunan) slaydına bölüm işareti; o slayt zaten bir bölüm başıysa dokunulmaz. */
function markSection(sec: Element, sldIdNum: Map<string, Slide>, slides: Slide[]): void {
  const first = kids(kid(sec, "sldIdLst"), "sldId")
    .map((e) => sldIdNum.get(attr(e, "id") ?? ""))
    .find((x): x is Slide => !!x);
  if (!first || first.section) return;
  const k = slides.indexOf(first);
  if (k >= 0) slides[k] = { ...first, section: { id: uid("sec"), name: attr(sec, "name") ?? "Bölüm" } };
}
