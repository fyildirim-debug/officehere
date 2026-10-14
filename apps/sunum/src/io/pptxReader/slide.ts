/**
 * Slaytlar: sldIdLst'teki her girdinin parçası çözülür (ilişkisi bozuk girdiler
 * paket içinden kurtarılır), her slayt okunur (şekiller, arka plan, geçiş,
 * animasyonlar, açıklamalar, notlar); okunamayan slaytın yerine boş slayt konur.
 */
import { uid } from "../../model/factory";
import { mapSlideLinks } from "../../model/links";
import { findPlaceholder } from "../../model/textStyle";
import type { Layout, Placeholder, SchemeColorName, Shape, Slide } from "../../model/types";
import { readShapeList } from "../readDrawing";
import { readAnimations, readTransition } from "../readTiming";
import { attr, kid, kidPath, kids, parseRels, relId, relsPathFor, REL, type Rel } from "../xml";
import { readSlideComments, type CommentAuthors } from "./comments";
import { readBg, readClrMap, type MkCtx } from "./common";
import type { LayoutSet } from "./layout";
import { readSlideNotes } from "./notes";
import { byNumber, PptxError, type Part, type PptxPackage } from "./package";
import type { PresPart } from "./presentation";

/** Bir sunuda okunacak en çok slayt (bozuk sldIdLst'e karşı). */
const MAX_SLIDES = 5000;
const SLIDE_CT = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";

/** Sunudaki slayt girdileri ve parçaları. */
export interface SlideList {
  sldIds: Element[];
  /** Her girdinin slayt parçası (sldIds ile aynı sırada); bulunamayanlar tanımsız. */
  paths: (string | undefined)[];
  /** Girdinin sunu ilişkisindeki hedefi (bozuk olabilir). */
  relTarget(sid: Element): string | undefined;
  /** En az bir slayt paket içinden bulundu. */
  recovered: boolean;
}

/**
 * Slayt girdilerinin parçaları. İlişkisi bozuk (hedefi yok ya da slayt olmayan bir
 * parçaya işaret eden) girdiler, paket içinde hiçbir sağlam girdinin kullanmadığı
 * slayt parçalarıyla sırasıyla eşlenir (kurtarma).
 */
export function resolveSlides(pres: PresPart, contentTypes: Map<string, string>): SlideList {
  const { zip, W } = pres.pk;
  const isSlidePart = (p: string) => (contentTypes.has(p) ? contentTypes.get(p) === SLIDE_CT : /^ppt\/slides\/[^/]+\.xml$/i.test(p));
  let sldIds = kids(kid(pres.doc, "sldIdLst"), "sldId");
  if (sldIds.length > MAX_SLIDES) {
    W.add(`Sunuda ${sldIds.length.toLocaleString("tr-TR")} slayt girdisi var; yalnız ilk ${MAX_SLIDES.toLocaleString("tr-TR")} slayt açıldı.`);
    sldIds = sldIds.slice(0, MAX_SLIDES);
  }
  const relTarget = (sid: Element) => {
    const rid = relId(sid);
    return rid ? pres.rels.get(rid)?.target : undefined;
  };
  const used = new Set<string>();
  const paths: (string | undefined)[] = sldIds.map((sid) => {
    const t = relTarget(sid);
    if (!t || used.has(t) || !zip.has(t) || !isSlidePart(t)) return undefined;
    used.add(t);
    return t;
  });
  const list: SlideList = { sldIds, paths, relTarget, recovered: false };
  const failedIdx = paths.flatMap((p, i) => (p ? [] : [i]));
  if (failedIdx.length) {
    const orphans = zip.paths.filter((p) => isSlidePart(p) && !used.has(p) && /\.xml$/i.test(p) && !/\/_rels\//.test(p)).sort(byNumber);
    list.recovered = recoverSlidePaths(pres.pk, list, failedIdx, orphans);
  }
  return list;
}

/** Bulunamayan girdilere sahipsiz slayt parçalarını sırasıyla atar; en az biri atandıysa doğru. */
function recoverSlidePaths(pk: PptxPackage, list: SlideList, failedIdx: number[], orphans: string[]): boolean {
  let k = 0;
  for (const i of failedIdx) {
    if (k >= orphans.length) break;
    pk.W.detail(`${i + 1}. slayt: ilişki hedefi ${list.relTarget(list.sldIds[i]) ?? "(yok)"} yerine ${orphans[k]} kullanıldı`);
    list.paths[i] = orphans[k++];
  }
  if (k) pk.W.add(`Sununun ilişki bölümü bozuk; ${k} slayt paket içinden bulunarak açıldı. Slayt sırasını denetleyin.`);
  return k > 0;
}

/** Slayt okuma ortamı. */
export interface SlideEnv {
  pk: PptxPackage;
  mkCtx: MkCtx;
  /** Ana asıl slaydın renk eşlemesi. */
  clrMap: Record<string, SchemeColorName>;
  set: LayoutSet;
  /** Ana asıl slaydın şekilleri (düzenin asıl slaydı bilinmezse kalıtım için). */
  masterShapes: Shape[];
  authors: CommentAuthors;
}

/**
 * Tüm slaytları okur. `slideAt[i]`: okunabilen slayt ya da (yerine boş slayt
 * konduysa) tanımsız. Hiçbir slayt okunamazsa PptxError.
 */
export async function readSlides(env: SlideEnv, list: SlideList): Promise<{ slides: Slide[]; slideAt: (Slide | undefined)[] }> {
  const layouts = env.set.layouts;
  const blankLayout = layouts.find((l) => l.type === "blank") ?? layouts[0];
  const blankSlide = (): Slide => ({ id: uid("sl"), layoutId: blankLayout.id, shapes: [], notes: "", animations: [] });
  const slides: Slide[] = [];
  const slideAt: (Slide | undefined)[] = [];
  for (const [i, sp] of list.paths.entries()) {
    const slide = await readSlideAt(env, list, i, sp);
    slides.push(slide ?? blankSlide());
    slideAt.push(slide);
  }
  if (list.sldIds.length && !slideAt.some(Boolean)) throw new PptxError("Dosya bozuk: sunudaki slaytların hiçbiri okunamadı.", env.pk.W.details.join("\n"));
  return { slides, slideAt };
}

/** `i`. girdinin slaydı; parça yoksa ya da okunamazsa (uyarıyla) tanımsız. */
async function readSlideAt(env: SlideEnv, list: SlideList, i: number, sp: string | undefined): Promise<Slide | undefined> {
  const { W } = env.pk;
  const n = i + 1;
  const sp0 = sp ?? list.relTarget(list.sldIds[i]);
  const sPart = sp ? await env.pk.part(sp, "sld") : ({ problem: "missing", detail: `${sp0 ?? "(ilişki yok)"}: slayt parçası yok` } as Part);
  if (!sPart.root) {
    W.slide(sPart.problem === "tooLarge" ? "tooLarge" : sPart.problem === "missing" ? "missing" : "corrupt", n, sPart.detail);
    return undefined;
  }
  if (sPart.problem === "partial") W.slide("partial", n, sPart.detail);
  try {
    return await readSlide(env, sp!, sPart.root, n);
  } catch (e) {
    W.slide("corrupt", n, `${sp}: ${(e as Error)?.message ?? e}`);
    return undefined;
  }
}

/**
 * Tek slayt. Geçiş kancası (dalga 3): slaydın extLst'i, custDataLst ve
 * tanınmayan spTree öğeleri burada (readShapeList çevresinde) toplanacak.
 */
async function readSlide(env: SlideEnv, sp: string, doc: Element, n: number): Promise<Slide> {
  const { pk } = env;
  const relsXml = await pk.text(relsPathFor(sp));
  if (relsXml === undefined) pk.W.slide("rels", n, `${relsPathFor(sp)}: pakette yok ya da okunamadı`);
  const rels = parseRels(relsXml, sp);
  const layout = layoutOf(env, rels);
  const sctx = env.mkCtx(rels, readClrMap(kidPath(doc, "clrMapOvr", "overrideClrMapping"), env.clrMap), (ph: Placeholder) => findPlaceholder(layout.shapes, ph) ?? findPlaceholder(env.set.masterShapes.get(layout) ?? env.masterShapes, { type: ph.type }));
  const cSld = kid(doc, "cSld");
  const shapes = readShapeList(kid(cSld, "spTree"), sctx);
  const slide: Slide = {
    id: uid("sl"),
    ...(attr(cSld, "name") ? { name: attr(cSld, "name") } : {}),
    layoutId: layout.id,
    shapes,
    background: readBg(cSld, sctx),
    hideMasterShapes: attr(doc, "showMasterSp") === "0" || attr(doc, "showMasterSp") === "false" || undefined,
    hidden: attr(doc, "show") === "0" || attr(doc, "show") === "false" || undefined,
    notes: "",
    transition: readTransition(doc),
    animations: readAnimations(doc, sctx.idMap),
  };
  const comments = await readSlideComments(pk, rels, n, env.authors);
  if (comments) slide.comments = comments;
  slide.notes = await readSlideNotes(pk, env.mkCtx, rels, n);
  return slide;
}

/** Slaydın düzeni (ilişkisi bozuksa ilk düzen). */
function layoutOf(env: SlideEnv, rels: Map<string, Rel>): Layout {
  const lrel = [...rels.values()].find((r) => r.type === REL.slideLayout);
  return (lrel && env.set.byPath.get(lrel.target)) || env.set.layouts[0];
}

/** Slayt atlama köprüleri (parça yolu) slayt kimliklerine çevrilir; döner: parça yolu → slayt kimliği. */
export function linkSlides(list: SlideList, slideAt: (Slide | undefined)[], slides: Slide[]): Map<string, string> {
  const idByPath = new Map<string, string>();
  list.paths.forEach((sp, i) => {
    if (sp && slideAt[i]) idByPath.set(sp, slides[i].id);
  });
  for (let k = 0; k < slides.length; k++) {
    slides[k] = mapSlideLinks(slides[k], (l) => (l.startsWith("#sldpath:") ? (idByPath.has(l.slice(9)) ? `#slide:${idByPath.get(l.slice(9))}` : undefined) : l));
  }
  return idByPath;
}
