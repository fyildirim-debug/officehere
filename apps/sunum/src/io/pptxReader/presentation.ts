/**
 * Sunu bölümü (presentation.xml) ve sunu düzeyi parçalar: slayt boyutu, özel
 * gösteriler, gösteri/yazdırma ayarları (presProps.xml), belge özellikleri
 * (docProps/core.xml) ve slaytlardan çıkarılan üst/alt bilgi durumu.
 */
import { DEFAULT_HF } from "../../model/presentation";
import type { CustomShow, Presentation, Slide, SpShape } from "../../model/types";
import { uid } from "../../model/factory";
import { DEFAULT_PRINT_SETTINGS, notesOrientation, readPresProps } from "../presProps";
import { attr, fromEmu, kid, kids, numAttr, parseRels, parseXml, relId, relsPathFor, REL, type Rel } from "../xml";
import { PptxError, type Part, type PptxPackage } from "./package";

/** Sunu bölümü, ilişkileri ve paket kökü ilişkileri. */
export interface PresPart {
  pk: PptxPackage;
  path: string;
  doc: Element;
  rels: Map<string, Rel>;
  /** Paket kökü ilişkileri (_rels/.rels). */
  rootRels: Map<string, Rel>;
  /** İlişki hedefi (pakette varsa); yoksa ilişkiler bozukken alışılmış parça yolu. */
  partOf(type: string, fallback: RegExp): string | undefined;
}

/**
 * Sunu bölümünü ve ilişkilerini açar; okunamazsa anlaşılır PptxError.
 * Geçiş katmanı: tanınmayan sunu öğeleri (embeddedFontLst, custDataLst, extLst'teki diğer
 * uzantılar…) kaydederken kaynak paketten yamalanır (rawPassthrough.patchPresentationXml).
 */
export async function openPresentation(pk: PptxPackage): Promise<PresPart> {
  const { zip, W } = pk;
  const rootRels = parseRels(await pk.text("_rels/.rels"), "");
  let path = [...rootRels.values()].find((r) => r.type === REL.officeDocument)?.target ?? "ppt/presentation.xml";
  if (!zip.has(path) && zip.has("ppt/presentation.xml")) path = "ppt/presentation.xml";
  const presPart = await pk.part(path, "presentation");
  if (!presPart.root) throw presentationError(presPart);
  if (presPart.problem === "partial") W.add("Sunu bölümü (presentation.xml) bozuk; yalnız okunabilen kısmı kullanıldı.", presPart.detail);
  const relsXml = await pk.text(relsPathFor(path));
  if (relsXml === undefined) W.detail(`${relsPathFor(path)}: pakette yok ya da okunamadı`);
  const rels = parseRels(relsXml, path);
  const partOf = (type: string, fallback: RegExp): string | undefined => {
    const t = [...rels.values()].find((r) => r.type === type)?.target;
    if (t && zip.has(t)) return t;
    const alt = pk.parts(fallback)[0];
    if (alt) W.detail(`${type.split("/").pop()}: ${t ?? "(ilişki yok)"} yerine ${alt} kullanıldı`);
    return alt ?? t;
  };
  return { pk, path, doc: presPart.root, rels, rootRels, partOf };
}

function presentationError(p: Part): PptxError {
  if (p.problem === "missing") return new PptxError("Sunu bölümü (presentation.xml) bulunamadı.", p.detail);
  return new PptxError(p.problem === "tooLarge" ? "Sunu bölümü (presentation.xml) olağandışı büyük; dosya bozuk ya da zararlı olabilir." : "Sunu bölümü (presentation.xml) bozuk; dosya açılamadı.", p.detail);
}

/** Slayt boyutu (piksel); yoksa 1280×720. */
export function slideSize(doc: Element): { width: number; height: number } {
  const sz = kid(doc, "sldSz");
  return { width: fromEmu(numAttr(sz, "cx"), 1280) || 1280, height: fromEmu(numAttr(sz, "cy"), 720) || 720 };
}

/** Özel slayt gösterileri; `idByPath`: slayt parça yolu → model kimliği. */
export function readCustomShows(pres: PresPart, idByPath: Map<string, string>): { customShows: CustomShow[]; customShowIds: Map<string, string> } {
  const customShows: CustomShow[] = [];
  const customShowIds = new Map<string, string>();
  for (const cs of kids(kid(pres.doc, "custShowLst"), "custShow")) {
    const ids = kids(kid(cs, "sldLst"), "sld")
      .map((e) => {
        const rid = relId(e);
        const sp = rid ? pres.rels.get(rid)?.target : undefined;
        return sp ? idByPath.get(sp) : undefined;
      })
      .filter((x): x is string => !!x);
    const show = { id: uid("cs"), name: attr(cs, "name") ?? `Özel Gösteri ${customShows.length + 1}`, slideIds: ids };
    if (attr(cs, "id") !== undefined) customShowIds.set(attr(cs, "id")!, show.id);
    customShows.push(show);
  }
  return { customShows, customShowIds };
}

/** Gösteri ve yazdırma ayarları (presProps.xml); not sayfası yönü p:notesSz'den. */
export async function readShowSettings(pres: PresPart, customShowIds: Map<string, string>): Promise<ReturnType<typeof readPresProps>> {
  const propsPath = pres.partOf(REL.presProps, /^ppt\/presProps\.xml$/i);
  const presProps = readPresProps(propsPath ? await pres.pk.text(propsPath, "Gösteri ve yazdırma ayarları okunamadı; varsayılanlar kullanıldı.") : undefined, customShowIds);
  const notesSz = kid(pres.doc, "notesSz");
  if (notesOrientation(numAttr(notesSz, "cx"), numAttr(notesSz, "cy")) === "landscape") presProps.printSettings = { ...(presProps.printSettings ?? DEFAULT_PRINT_SETTINGS), orientation: "landscape" };
  return presProps;
}

/** Belge özellikleri (başlık, konu, yazar, tarihler). */
export async function readCoreProps(pres: PresPart): Promise<Presentation["props"]> {
  const { zip } = pres.pk;
  const props: Presentation["props"] = {};
  const coreRel = [...pres.rootRels.values()].find((r) => r.type === REL.coreProps)?.target;
  const corePath = coreRel && zip.has(coreRel) ? coreRel : zip.has("docProps/core.xml") ? "docProps/core.xml" : coreRel;
  const core = corePath ? await pres.pk.text(corePath, "Belge özellikleri (başlık, yazar) okunamadı.") : undefined;
  if (core) {
    const d = parseXml(core).documentElement;
    props.title = kid(d, "title")?.textContent || undefined;
    props.subject = kid(d, "subject")?.textContent || undefined;
    props.creator = kid(d, "creator")?.textContent || undefined;
    props.created = kid(d, "created")?.textContent || undefined;
    props.modified = kid(d, "modified")?.textContent || undefined;
  }
  return props;
}

/** Üst bilgi / alt bilgi durumu slaytlardaki yer tutuculardan çıkarılır. */
export function headerFooterOf(slides: Slide[]): Presentation["headerFooter"] {
  const hasPh = (t: string) => slides.some((s) => s.shapes.some((x) => x.kind === "sp" && x.ph?.type === t));
  const ftr = slides.flatMap((s) => s.shapes).find((x): x is SpShape => x.kind === "sp" && x.ph?.type === "ftr");
  return {
    ...DEFAULT_HF,
    dateTime: hasPh("dt"),
    slideNumber: hasPh("sldNum"),
    footer: hasPh("ftr"),
    footerText: ftr?.text ? ftr.text.paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join(" ") : "",
  };
}
