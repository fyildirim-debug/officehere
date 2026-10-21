// DOCX parça dönüştürücüsü: gövde / üst bilgi / alt bilgi / not içeriğini (paragraflar, koşular,
// alanlar, çizimler, metin kutuları, tablolar, İçindekiler) ProseMirror JSON'a çevirir.
//
// Geçiş katmanı, öğe düzeyi: modelin karşılamadığı öğeler (gömülü nesne, denetim, anlaşılmayan çizim ve
// mc:AlternateContent, alt belge, içe aktarılan içerik, bilinmeyen blok/satır içi/koşu çocukları) opak düğüm
// olur (editor/opaque.ts): özgün XML'i, ilişkileri ve kaynak paket kimliğiyle; pPr/rPr'nin okunmayan çocukları
// paragrafta ve textStyle işaretinde `ooxmlExtra` olarak taşınır. Yazımı partWriter.ts'te.
import { explicitRelIds } from "@officehere/core/opc";
import { toLinear } from "../../editor/math/ast";
import { fromOmml } from "../../editor/math/render";
import { indexNode, xeNode } from "./indexXml";
import type { JSONContent } from "@tiptap/core";
import type { DocSettings } from "../../model/document";
import { fieldPicture, type TocEntry } from "../../editor/nodes";
import { pictureOfInstr } from "../../editor/dateFormat";
import { normalizeDropCap } from "../../editor/dropCap";
import type { OpaqueAttrs, OpaqueKind, OpaqueRef } from "../../editor/opaque";
import { attr, attrNum, descendants, kid, kids, path, NS } from "./xml";
import { parsePPr, parseRPr, stripUndef, type ParaProps, type RunProps } from "./readerStyles";
import { tableStyleFromId } from "./styleXml";
import type { ListSig } from "./numbering";
import type { DocxReader } from "./reader";
import { readDrawing, readPict } from "./readerDrawing";
import { CHART_NS } from "./chartXml";
import { DOC_PREFIXES, OUTPUT_NAMESPACES } from "./partWriter";
import { extendCommentToWord } from "../common";
import type { LoadWarnings } from "../warnings";

export type Mark = { type: string; attrs?: Record<string, unknown> };

export interface Rel {
  type: string;
  target: string;
  external: boolean;
}

interface FieldState {
  instr: string;
  phase: "instr" | "result";
  kind: "link" | "page" | "numpages" | "other";
  href?: string;
  emitted?: boolean;
}

/** Bir parçanın (gövde / üst bilgi / alt bilgi) dönüştürücüsü. */
/** Sayfa/tarih alanı olarak okunan alan anahtar sözcükleri. */
const PAGE_FIELDS = new Set(["PAGE", "NUMPAGES", "DATE", "TIME"]);

/** Büyük ilk harf paragrafının harfi: çerçeve boyutu (yazı tipi boyutu) atılır, yazı tipi ayrıca döner. */
function dropLetter(n: JSONContent): { node: JSONContent; font: string | null } {
  let font: string | null = null;
  const marks = (n.marks ?? [])
    .map((m) => {
      if (m.type !== "textStyle") return m;
      font = (m.attrs?.fontFamily as string | null) ?? null;
      const { fontSize: _s, fontFamily: _f, ...rest } = m.attrs ?? {};
      return { ...m, attrs: rest };
    })
    .filter((m) => m.type !== "textStyle" || Object.values(m.attrs ?? {}).some((v) => v !== null && v !== undefined && v !== false));
  return { node: { ...n, ...(marks.length ? { marks } : { marks: undefined }) }, font };
}

/* ------------------------------------------------------------------ */
/* Opak öğeler ve ek özellikler                                          */
/* ------------------------------------------------------------------ */

const XMLNS = "http://www.w3.org/2000/xmlns/";
const PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture";
/** Yer imleri, izinler, yazım denetimi işaretleri ve özellik öğeleri: içerik değildir, opak olmaz. */
const MARKERS = new Set([
  "bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd", "permStart", "permEnd", "proofErr",
  "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd",
  "customXmlInsRangeStart", "customXmlInsRangeEnd", "customXmlDelRangeStart", "customXmlDelRangeEnd",
  "customXmlMoveFromRangeStart", "customXmlMoveFromRangeEnd", "customXmlMoveToRangeStart", "customXmlMoveToRangeEnd",
  "pPr", "rPr", "sectPr", "tcPr", "trPr", "tblPr", "tblPrEx", "sdtPr", "sdtEndPr", "customXmlPr", "smartTagPr", "fldData",
]);
/** Koşunun okunmayan ama içerik taşımayan çocukları. */
const RUN_SKIP = new Set(["rPr", "lastRenderedPageBreak", "annotationRef", "separator", "continuationSeparator", "delInstrText", "proofErr"]);
/** Modelin okuduğu (yazıcının yeniden ürettiği) paragraf özellikleri; geri kalan W öğeleri ek özellik olarak taşınır. */
const PPR_MODELED = new Set(["pStyle", "keepNext", "keepLines", "pageBreakBefore", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "spacing", "ind", "jc", "outlineLvl", "sectPr", "pPrChange"]);
/** Modelin okuduğu koşu özellikleri (+ izlenen değişiklik öğeleri). */
const RPR_MODELED = new Set(["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "vanish", "color", "spacing", "sz", "szCs", "highlight", "u", "shd", "vertAlign", "rPrChange", "ins", "del", "moveFrom", "moveTo"]);
/** Paragraf işaretinin koşu özelliklerinden atılanlar (izlenen değişiklik; kimlik ve yazar taşır). */
const MARK_REVISIONS = new Set(["ins", "del", "moveFrom", "moveTo", "rPrChange"]);
/** Bağlam ad alanı bildirimi gerektiren MCE öznitelikleri (değerlerinde ön ek geçer). */
const MCE_ATTRS = new Set(["Requires", "Ignorable", "ProcessContent", "MustUnderstand", "PreserveElements", "PreserveAttributes"]);
/** Bozuk belgede bildirilmemiş ön ekler için okuyucunun yer tutucu ad alanı (xml.ts): opak olarak yazılmaz. */
const PLACEHOLDER_NS = "urn:officehere:bilinmeyen:";
/** Tarayıcıda çizilebilen yedek görüntü türleri. */
const PREVIEWABLE = /^data:image\/(png|jpeg|gif|bmp|webp|svg\+xml)[;,]/;

const escAttr = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;").replace(/\t/g, "&#9;").replace(/\n/g, "&#10;").replace(/\r/g, "&#13;");
const escText = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

let serializer: XMLSerializer | null = null;
function serialize(el: Element): string {
  serializer ??= new XMLSerializer();
  return serializer.serializeToString(el);
}

/**
 * Parça kökünde (DOC_NS) aynı önekle bildirilen ad alanlarındaki öğenin hızlı yazımı (ad alanı bildirimsiz).
 * Başka ön ek, öneksiz öğe ya da beklenmeyen düğümde null (XMLSerializer'a bırakılır).
 */
function quickXml(el: Element): string | null {
  const p = el.prefix;
  if (!p || DOC_PREFIXES.get(p) !== el.namespaceURI) return null;
  let s = `<${p}:${el.localName}`;
  for (const a of Array.from(el.attributes)) {
    if (a.prefix === "xmlns" || a.name === "xmlns") continue;
    if (a.prefix && a.prefix !== "xml" && DOC_PREFIXES.get(a.prefix) !== a.namespaceURI) return null;
    s += ` ${a.name}="${escAttr(a.value)}"`;
  }
  let inner = "";
  for (let c = el.firstChild; c; c = c.nextSibling) {
    if (c.nodeType === 1) {
      const x = quickXml(c as Element);
      if (x === null) return null;
      inner += x;
    } else if (c.nodeType === 3 || c.nodeType === 4) inner += escText(c.nodeValue ?? "");
  }
  return inner ? `${s}>${inner}</${p}:${el.localName}>` : `${s}/>`;
}

/** Ek özellik öğesinin XML'i: kökteki önekle aynı bildirimler atılır (yazıldığı parçanın kökünde var). */
function extraXml(el: Element): string {
  const quick = quickXml(el);
  if (quick !== null) return quick;
  return serialize(el).replace(/\sxmlns:([\w.-]+)="([^"]*)"/g, (m, p: string, uri: string) => (DOC_PREFIXES.get(p) === uri ? "" : m));
}

/** pPr/rPr'nin modelde karşılığı olmayan çocukları (W ve w14 ad alanı; sıra korunur). */
function extrasOf(pr: Element | null, modeled: ReadonlySet<string>, transform?: (el: Element) => Element | null): string | null {
  if (!pr || pr.namespaceURI !== NS.w) return null;
  let out = "";
  for (const c of kids(pr)) {
    if (c.namespaceURI === NS.w ? modeled.has(c.localName) : c.namespaceURI !== NS.w14) continue;
    const el = transform ? transform(c) : c;
    if (el) out += extraXml(el);
  }
  return out && !out.includes(PLACEHOLDER_NS) ? out : null;
}

/** Paragraf işaretinin koşu özellikleri: izlenen değişiklik öğeleri atılır, boş kalırsa yazılmaz. */
function markRPr(el: Element): Element | null {
  if (el.localName !== "rPr") return el;
  if (!el.firstElementChild) return null;
  if (!kids(el).some((k) => MARK_REVISIONS.has(k.localName))) return el;
  const clone = el.cloneNode(true) as Element;
  for (const k of kids(clone)) if (MARK_REVISIONS.has(k.localName)) clone.removeChild(k);
  return clone.firstElementChild ? clone : null;
}

/** Öğenin kapsamındaki ad alanı bildirimleri (en yakın geçerli). */
function scopeOf(el: Element): Map<string, string> {
  const map = new Map<string, string>();
  for (let e: Element | null = el; e; e = e.parentElement) {
    for (const a of Array.from(e.attributes)) if (a.prefix === "xmlns" && !map.has(a.localName)) map.set(a.localName, a.value);
  }
  return map;
}

/** Kapsamdaki mc:Ignorable ad alanları. */
function ignorableOf(el: Element): Set<string> {
  const out = new Set<string>();
  for (let e: Element | null = el; e; e = e.parentElement) {
    const ig = e.getAttributeNS(NS.mc, "Ignorable");
    if (!ig) continue;
    const scope = scopeOf(e);
    for (const p of ig.split(/\s+/)) if (scope.has(p)) out.add(scope.get(p)!);
  }
  return out;
}

const shapetypeCache = new WeakMap<Document, Map<string, Element>>();
function shapetypes(doc: Document): Map<string, Element> {
  let m = shapetypeCache.get(doc);
  if (!m) {
    m = new Map();
    for (const st of descendants(doc.documentElement, "shapetype")) {
      const id = st.getAttribute("id");
      if (id && st.namespaceURI === NS.v && !m.has(id)) m.set(id, st);
    }
    shapetypeCache.set(doc, m);
  }
  return m;
}

/**
 * Opak öğenin kendi kendine yeten XML'i: kullanılan (öğe/öznitelik ön ekleri ve MCE değerlerindeki) ad alanları
 * köke bildirilir, kapsamda yok sayılabilir olanlar köke mc:Ignorable olarak yazılır, başka yerde tanımlanmış
 * VML şekil türleri (v:shapetype) eklenir. Kök, çıktının bilmediği bir ad alanındaysa mc:AlternateContent ile
 * sarılır (anlayan okuyucu korur). Yer tutucu ad alanlı (bozuk) öğede null.
 */
function opaqueXml(el: Element): string | null {
  const scope = scopeOf(el);
  const clone = el.cloneNode(true) as Element;
  const all = [clone, ...Array.from(clone.getElementsByTagName("*"))];
  const used = new Set<string>();
  const ids = new Set<string>();
  const needTypes: Element[] = [];
  for (const e of all) {
    if (e.prefix) used.add(e.prefix);
    if (e.namespaceURI === NS.v && e.localName === "shapetype" && e.getAttribute("id")) ids.add(e.getAttribute("id")!);
    for (const a of Array.from(e.attributes)) {
      if (a.prefix === "xmlns" || a.name === "xmlns") continue;
      if (a.prefix) used.add(a.prefix);
      if (a.namespaceURI === NS.mc && MCE_ATTRS.has(a.localName)) for (const tok of a.value.split(/\s+/)) if (tok) used.add(tok.split(":")[0]!);
    }
    if (e.namespaceURI === NS.v && e.localName !== "shapetype" && /^#/.test(e.getAttribute("type") ?? "")) needTypes.push(e);
  }
  // Başka bir öğede tanımlı şekil türü (Word bir türü belgede bir kez yazar): başvuran şeklin önüne kopyalanır.
  const types = shapetypes(el.ownerDocument);
  for (const e of needTypes) {
    const id = e.getAttribute("type")!.slice(1);
    const st = !ids.has(id) ? types.get(id) : undefined;
    if (!st || !e.parentNode) continue;
    e.parentNode.insertBefore(st.cloneNode(true), e);
    ids.add(id);
    for (const s of [st, ...Array.from(st.getElementsByTagName("*"))]) {
      if (s.prefix) used.add(s.prefix);
      for (const a of Array.from(s.attributes)) if (a.prefix && a.prefix !== "xmlns") used.add(a.prefix);
    }
  }
  for (const p of used) {
    const uri = scope.get(p);
    if (uri && p !== "xml" && !clone.hasAttributeNS(XMLNS, p)) clone.setAttributeNS(XMLNS, `xmlns:${p}`, uri);
  }
  // Kapsamda yok sayılabilir ilan edilen ve öğede kullanılan ad alanları.
  const ign = ignorableOf(el);
  const prefixes = [...used].filter((p) => scope.has(p) && ign.has(scope.get(p)!));
  if (prefixes.length) {
    const mcPrefix = [...scope].find(([, u]) => u === NS.mc)?.[0] ?? "mc";
    if (!clone.hasAttributeNS(XMLNS, mcPrefix)) clone.setAttributeNS(XMLNS, `xmlns:${mcPrefix}`, NS.mc);
    const cur = (clone.getAttributeNS(NS.mc, "Ignorable") ?? "").split(/\s+/).filter(Boolean);
    clone.setAttributeNS(NS.mc, `${mcPrefix}:Ignorable`, [...new Set([...cur, ...prefixes])].join(" "));
  }
  let xml = serialize(clone);
  if (xml.includes(PLACEHOLDER_NS)) return null;
  if (el.namespaceURI && !OUTPUT_NAMESPACES.has(el.namespaceURI) && el.prefix) {
    xml = `<mc:AlternateContent xmlns:mc="${NS.mc}" xmlns:${el.prefix}="${escAttr(el.namespaceURI)}"><mc:Choice Requires="${el.prefix}">${xml}</mc:Choice></mc:AlternateContent>`;
  }
  return xml;
}

/** CSS uzunluğu (pt, px, in, cm, mm) → px. */
function cssPx(style: string, prop: string): number | null {
  const m = new RegExp(`(?:^|;)\\s*${prop}\\s*:\\s*(-?[\\d.]+)\\s*(pt|px|in|cm|mm)?`, "i").exec(style);
  if (!m) return null;
  const v = parseFloat(m[1]!);
  const unit = (m[2] ?? "px").toLowerCase(); // yerel-ayar: ascii
  const f = unit === "pt" ? 4 / 3 : unit === "in" ? 96 : unit === "cm" ? 96 / 2.54 : unit === "mm" ? 96 / 25.4 : 1;
  return Number.isFinite(v) ? Math.round(v * f) : null;
}

/** OLE ProgID → okunur tür (marka adı geçmez). */
function oleKind(progId: string): string {
  const p = progId.toLowerCase(); // yerel-ayar: ascii
  if (p.startsWith("excel.")) return "hesap tablosu";
  if (p.startsWith("word.")) return "belge";
  if (p.startsWith("powerpoint.")) return "sunu";
  if (p.includes("equation")) return "denklem";
  if (p === "package" || p.startsWith("package")) return "ekli dosya";
  if (p.startsWith("visio.")) return "çizim";
  if (p.startsWith("acroexch.") || p.includes("pdf")) return "PDF";
  if (p.startsWith("paint.") || p.startsWith("pbrush")) return "resim";
  if (p.includes("graph") || p.includes("chart")) return "grafik";
  return "";
}

const GRAPHIC_LABEL: [RegExp, string][] = [
  [/drawingml\/2006\/diagram$/, "Diyagram (SmartArt)"],
  [/wordprocessingGroup$/, "Çizim grubu"],
  [/wordprocessingCanvas$/, "Çizim tuvali"],
  [/wordprocessingShape$/, "Şekil"],
  [/drawingml\/2006\/chart$|chartex$/, "Grafik"],
  [/lockedCanvas$/, "Kilitli tuval"],
  [/ink$/, "Mürekkep çizimi"],
  [/drawingml\/2006\/picture$/, "Resim"],
];

/** Opak öğenin türü ve kullanıcıya gösterilen adı. */
function opaqueLabel(el: Element, rels: Map<string, Rel>): { kind: OpaqueKind; label: string } {
  const n = el.localName;
  if (n === "object" || n === "control") {
    if (n === "control" || descendants(el, "control").length) return { kind: "control", label: "Form denetimi" };
    const ole = descendants(el, "OLEObject")[0];
    const k = oleKind(ole?.getAttribute("ProgID") ?? "");
    const link = ole?.getAttribute("Type") === "Link";
    return { kind: "ole", label: `${link ? "Bağlantılı" : "Gömülü"} nesne${k ? ` (${k})` : ""}` };
  }
  if (n === "drawing" || n === "AlternateContent") {
    const uri = descendants(el, "graphicData")[0]?.getAttribute("uri") ?? "";
    const hit = GRAPHIC_LABEL.find(([re]) => re.test(uri));
    if (hit) return { kind: "drawing", label: hit[1] };
    if (descendants(el, "OLEObject").length) return { kind: "ole", label: "Gömülü nesne" };
    return { kind: "drawing", label: "Çizim" };
  }
  if (n === "pict") {
    const hr = Array.from(el.getElementsByTagName("*")).find((e) => e.getAttributeNS(NS.o, "hr") === "t");
    return hr ? { kind: "hr", label: "Yatay çizgi" } : { kind: "vml", label: "Çizim (VML)" };
  }
  if (n === "altChunk") {
    const rid = el.getAttributeNS(NS.r, "id") ?? attr(el, "id");
    const ext = /\.([a-z0-9]{1,5})$/i.exec(rid ? (rels.get(rid)?.target ?? "") : "")?.[1];
    return { kind: "altChunk", label: `İçe aktarılan içerik${ext ? ` (.${ext.toLowerCase()})` : ""}` }; // yerel-ayar: ascii
  }
  if (n === "subDoc") return { kind: "subDoc", label: "Alt belge" };
  if (n === "oMath" || n === "oMathPara") return { kind: "math", label: "Denklem" };
  if (n === "contentPart") return { kind: "drawing", label: "Mürekkep çizimi" };
  if (n === "ruby") return { kind: "other", label: "Fonetik kılavuz" };
  return { kind: "other", label: `${el.prefix ? `${el.prefix}:` : ""}${n}` };
}

/** Okuma boyunca korunamayan (kaynak paketi olmayan) gömülü nesne uyarısı verilen uyarı kümeleri. */
const objectLoss = new WeakSet<LoadWarnings>();

/** Sayfa/tarih alanı düğümü; tarih ve saatte \@ biçim resmi korunur (varsayılansa yazılmaz). */
function pageFieldNode(kw: string, instr: string): JSONContent {
  if (kw !== "DATE" && kw !== "TIME") return { type: "pageField", attrs: { field: kw } };
  const pic = pictureOfInstr(instr);
  return { type: "pageField", attrs: { field: kw, format: pic && pic !== fieldPicture(kw) ? pic : null } };
}

export class PartReader {
  private fields: FieldState[] = [];
  /** Paragraf içindeki kesmelerin türü (sayfa / sütun), parça sırasıyla. */
  private segKinds: ("pageBreak" | "columnBreak")[] = [];
  /** Gövdede bölüm sonları: sectPr taşıyan paragrafların ardından öğe sırası. */
  sectionEnds: { index: number; sectPr: Element }[] = [];
  private openComments = new Set<string>();
  /** Aralığı açılmış ve en az bir metne uygulanmış açıklamalar. */
  private rangedComments = new Set<string>();
  private markedComments = new Set<string>();
  /** Aralıksız (nokta) açıklamalar: ardından gelen ilk sözcüğe, yoksa paragrafın son metnine bağlanır. */
  pendingPoints: string[] = [];
  private listCounts = new Map<string, number>();
  watermark: DocSettings["watermark"] = null;

  constructor(
    readonly r: DocxReader,
    readonly rels: Map<string, Rel>,
  ) {}

  /** Not metni okunuyor: not stilinin (Dipnot Metni) biçimi doğrudan biçime çevrilmez. */
  noteMode = false;

  /** Opak öğe olarak alınan özgün öğeler. */
  private captured = new Set<Element>();
  /** collect() iç içeliği (0: parçanın kökü). */
  private depth = 0;

  /**
   * Parçanın sonunda hâlâ bağlanamayan nokta açıklamaları (ardından hiç metin gelmeyen, ör. boş paragraftaki
   * açıklama başvurusu) son paragrafın son satır içi öğesine, o da yoksa gizli bir yer işareti çapasına bağlanır;
   * böylece açıklama kaydederken kaybolmaz. Çapa yer işareti ikinci turda yeniden kullanılır (birikmez).
   */
  private anchorPoints(items: { node: JSONContent }[]) {
    const marks = this.pendingPoints.map((id) => ({ type: "comment", attrs: { id } }));
    for (const id of this.pendingPoints) this.markedComments.add(id);
    const first = this.pendingPoints[0]!;
    this.pendingPoints = [];
    let host = [...items].reverse().find((it) => it.node.type === "paragraph" || it.node.type === "heading")?.node;
    if (!host) {
      host = { type: "paragraph" };
      items.push({ node: host });
    }
    const content = host.content ?? [];
    const last = content[content.length - 1];
    if (last && last.type !== "text") content[content.length - 1] = { ...last, marks: [...(last.marks ?? []), ...marks] };
    else content.push({ type: "bookmark", attrs: { name: `_Aciklama_${first}` }, marks });
    host.content = content;
  }

  /**
   * Opak düğüm (blok ya da satır içi). Kurulamazsa null (çağıran bugünkü davranışa döner): bozuk ad alanı,
   * kaynakta karşılığı olmayan ilişki kimliği ya da özgün paket okunamadığı için iç ilişkilerin taşınamaması.
   */
  private opaque(el: Element, level: "block" | "run" | "para"): JSONContent | null {
    let xml: string | null;
    try {
      xml = opaqueXml(el);
    } catch {
      return null;
    }
    if (!xml) return null;
    const refs: OpaqueRef[] = [];
    for (const id of explicitRelIds(xml)) {
      const rel = this.rels.get(id);
      // Hedefsiz kimlik yazılırsa yazıcının yeni ilişki kimlikleriyle karışır.
      if (!rel) return null;
      refs.push({ id, type: rel.type, external: rel.external, target: rel.external ? rel.target : (this.r.source?.canonicalName(rel.target) ?? rel.target) });
    }
    // Kaydederken taşınamayacak (özgün paket okunamadı ya da hedef parça yok) öğe opak olmaz: bugünkü okuma uyarır.
    if (refs.some((r) => !r.external && !this.r.source?.has(r.target))) return null;
    const { kind, label } = opaqueLabel(el, this.rels);
    this.captured.add(el);
    const attrs: OpaqueAttrs = { xml, refs, src: this.r.kaynak, kind, label, ...this.opaqueView(el), ...(level === "block" ? {} : { level }) };
    return { type: level === "block" ? "opaqueBlock" : "opaqueInline", attrs };
  }

  /** Opak öğenin görünümü: tarayıcıda çizilebilen yedek görüntü (VML v:imagedata / a:blip) ve boyutu (px). */
  private opaqueView(el: Element): { preview: string | null; width: number | null; height: number | null } {
    const scopes = el.localName === "AlternateContent" ? [kid(el, "Fallback"), kid(el, "Choice")].filter((x): x is Element => !!x) : [el];
    let preview: string | null = null;
    for (const s of scopes) {
      for (const img of [...descendants(s, "imagedata"), ...descendants(s, "blip")]) {
        const rid = img.getAttributeNS(NS.r, "id") || img.getAttributeNS(NS.r, "embed") || img.getAttributeNS(NS.o, "relid");
        const target = rid ? this.rels.get(rid)?.target : undefined;
        const src = target ? this.r.media.get(target) : undefined;
        if (src && PREVIEWABLE.test(src)) {
          preview = src;
          break;
        }
      }
      if (preview) break;
    }
    let width: number | null = null;
    let height: number | null = null;
    const ext = descendants(el, "extent").find((e) => e.namespaceURI === NS.wp);
    if (ext) {
      width = Math.round((attrNum(ext, "cx") ?? 0) / 9525) || null;
      height = Math.round((attrNum(ext, "cy") ?? 0) / 9525) || null;
    } else {
      const shape = Array.from(el.getElementsByTagName("*")).find((e) => e.namespaceURI === NS.v && e.localName !== "shapetype" && /(?:^|;)\s*(width|height)\s*:/.test(e.getAttribute("style") ?? ""));
      const style = shape?.getAttribute("style") ?? "";
      width = cssPx(style, "width") || null;
      height = cssPx(style, "height") || null;
    }
    if (el.localName === "object") {
      width ??= Math.round((attrNum(el, "dxaOrig") ?? 0) / 15) || null;
      height ??= Math.round((attrNum(el, "dyaOrig") ?? 0) / 15) || null;
    }
    return { preview, width, height };
  }

  /** Opak düğüm; kurulamazsa `legacy` (bugünkü okuma) çalışır. */
  private opaqueOr(el: Element, level: "run" | "para", push: (n: JSONContent) => void, legacy: () => void) {
    const node = this.opaque(el, level);
    if (node) push(node);
    else legacy();
  }

  /** Kaydederken korunamayacak gömülü nesne / denetim uyarısı. */
  private loseObject(detail: string) {
    this.r.warnings.add("objects", detail);
    objectLoss.add(this.r.warnings);
  }

  /**
   * Gövdedeki bütün gömülü nesneler opak öğe olarak alındıysa okuyucunun "nesneler kaybolur" uyarısı geri alınır
   * (kaydederken özgün XML ve parçalarıyla korunurlar).
   */
  private settleObjectWarning(body: Element) {
    const w = this.r.warnings;
    if (!this.r.source || !w.has("objects") || objectLoss.has(w)) return;
    const inCaptured = (o: Element) => {
      for (let e: Element | null = o; e && e !== body; e = e.parentElement) if (this.captured.has(e)) return true;
      return false;
    };
    if (descendants(body, "OLEObject").every(inCaptured)) w.remove("objects");
  }

  /** Çizim modelce okunabiliyor mu (resim, grafik, metin kutusu); değilse opak alınır. */
  private drawingKnown(d: Element): boolean {
    const gd = descendants(d, "graphicData")[0];
    const uri = gd?.getAttribute("uri");
    const blip = descendants(d, "blip")[0];
    if (uri === PIC_NS) {
      const rid = blip ? (blip.getAttributeNS(NS.r, "embed") ?? attr(blip, "embed")) : null;
      const target = rid ? this.rels.get(rid)?.target : undefined;
      return !!target && this.r.media.has(target);
    }
    if (uri === CHART_NS) {
      const ref = gd ? descendants(gd, "chart")[0] : undefined;
      const rid = ref ? (ref.getAttributeNS(NS.r, "id") ?? attr(ref, "id")) : null;
      const target = rid ? this.rels.get(rid)?.target : undefined;
      return !!target && this.r.charts.has(target);
    }
    if (uri === NS.wps) return !blip && !!descendants(d, "txbxContent")[0];
    return false;
  }

  /** VML çizimi modelce okunabiliyor mu (filigran yazısı, resim, metin kutusu). */
  private pictKnown(pc: Element): boolean {
    if (descendants(pc, "textpath").some((x) => x.hasAttribute("string"))) return true;
    const img = descendants(pc, "imagedata")[0];
    if (img) {
      const rid = img.getAttributeNS(NS.r, "id") ?? attr(img, "id");
      const target = rid ? this.rels.get(rid)?.target : undefined;
      return !!target && this.r.media.has(target);
    }
    return !!descendants(pc, "txbxContent")[0];
  }

  /** VML çiziminde şekil var mı (yalnız şekil türü tanımı ya da boş w:pict içerik sayılmaz). */
  private static pictEmpty(pc: Element): boolean {
    return kids(pc).every((k) => k.localName === "shapetype");
  }

  blocks(container: Element): JSONContent[] {
    return this.foldLists(this.collect(container));
  }

  private collect(container: Element) {
    const items: { node: JSONContent; list?: { numId: string; ilvl: number; sig: ListSig; start: number }; ctx?: string }[] = [];
    // Büyük ilk harf: Word harfi çerçeveli ayrı bir paragrafa (w:framePr w:dropCap) yazar; sonraki paragrafa katılır.
    let drop: { items: typeof items; dc: Record<string, unknown> } | null = null;
    const flushDrop = () => {
      if (drop) items.push(...drop.items);
      drop = null;
    };
    const visit = (el: Element) => {
      for (const c of kids(el)) {
        if (c.localName !== "p" && c.localName !== "bookmarkStart" && c.localName !== "bookmarkEnd") flushDrop();
        switch (c.localName) {
          case "p": {
            const frame = kid(kid(c, "pPr"), "framePr");
            const dcMode = attr(frame, "dropCap");
            const its = this.paragraph(c);
            if ((dcMode === "drop" || dcMode === "margin") && its.length === 1 && its[0].node.content?.length) {
              flushDrop();
              drop = { items: its, dc: { mode: dcMode, lines: attrNum(frame, "lines") ?? 3, distance: Math.round(((attrNum(frame, "hSpace") ?? 0) / 20) * 100) / 100 } };
              break;
            }
            const pending = drop as { items: typeof items; dc: Record<string, unknown> } | null;
            const host = pending && its.length && !its[0].list && (its[0].node.type === "paragraph" || its[0].node.type === "heading") ? its[0].node : null;
            if (pending && host) {
              const letters = (pending.items[0].node.content ?? []).map(dropLetter);
              const font = letters.map((l) => l.font).find(Boolean) ?? null;
              host.content = mergeText([...letters.map((l) => l.node), ...(host.content ?? [])]);
              host.attrs = { ...host.attrs, dropCap: normalizeDropCap({ ...pending.dc, font }) };
              drop = null;
            } else flushDrop();
            for (const it of its) items.push(it);
            const sp = kid(kid(c, "pPr"), "sectPr");
            if (sp && container.localName === "body") this.sectionEnds.push({ index: items.length, sectPr: sp });
            break;
          }
          case "tbl":
            items.push({ node: this.table(c) });
            break;
          case "altChunk":
          case "subDoc": {
            // Başka biçimde gömülü içerik parçası / alt belge: opak blok (kaydederken korunur).
            const node = this.opaque(c, "block");
            if (node) items.push({ node });
            else this.loseObject(`w:${c.localName}`);
            break;
          }
          case "sdt": {
            const pr = kid(c, "sdtPr") ?? c;
            const gallery = attr(descendants(pr, "docPartGallery")[0], "val");
            const tag = attr(kid(pr, "tag"), "val");
            const content = kid(c, "sdtContent");
            const instr = content ? descendants(content, "instrText").map((t) => t.textContent ?? "").join("") : "";
            if (content && (tag === "OfficeHereIndex" || /^\s*INDEX\b/i.test(instr))) items.push({ node: indexNode(content, instr) });
            else if (content && (tag === "OfficeHereTableOfFigures" || /\\c\s+"/.test(instr))) items.push({ node: tofNode(this.toc(content), instr) });
            else if (gallery === "Table of Contents" && content) items.push({ node: this.toc(content) });
            else if (content && (gallery === "Bibliographies" || tag === "OfficeHereBibliography")) items.push({ node: bibliographyNode(content) });
            else if (content) visit(content);
            break;
          }
          case "customXml":
          case "ins":
          case "smartTag":
            visit(c);
            break;
          case "AlternateContent": {
            const ch = kid(c, "Choice") ?? kid(c, "Fallback");
            if (ch) visit(ch);
            break;
          }
          case "del":
          case "moveFrom":
          case "moveTo":
            break;
          default: {
            // Tanınmayan blok öğesi (gövdede denklem, bilinmeyen uzantı…): opak blok.
            if (MARKERS.has(c.localName)) break;
            const node = this.opaque(c, "block");
            if (node) items.push({ node });
          }
        }
      }
    };
    this.depth++;
    try {
      visit(container);
    } finally {
      this.depth--;
    }
    flushDrop();
    if (!this.depth && this.pendingPoints.length) this.anchorPoints(items);
    // Bağlamsal aralık (w:contextualSpacing): aynı stildeki ardışık paragraflar arasındaki aralık yok sayılır.
    // Modelde karşılığı olmadığından aradaki aralıklar doğrudan sıfırlanır (görünüm ve sayfa düzeni korunur).
    for (let i = 0; i + 1 < items.length; i++) {
      const a = items[i];
      const b = items[i + 1];
      if (!a.ctx || a.ctx !== b.ctx) continue;
      a.node.attrs = { ...a.node.attrs, spaceAfter: 0 };
      b.node.attrs = { ...b.node.attrs, spaceBefore: 0 };
    }
    return items;
  }

  /** Gövdeyi bölümlere ayırır (her bölümün listeleri ayrı katlanır). */
  sections(body: Element): { blocks: JSONContent[]; sectPr: Element | null }[] {
    this.sectionEnds = [];
    const items = this.collect(body);
    this.settleObjectWarning(body);
    const out: { blocks: JSONContent[]; sectPr: Element | null }[] = [];
    let start = 0;
    for (const e of this.sectionEnds) {
      out.push({ blocks: this.foldLists(items.slice(start, e.index)), sectPr: e.sectPr });
      start = e.index;
    }
    out.push({ blocks: this.foldLists(items.slice(start)), sectPr: kid(body, "sectPr") });
    return out;
  }

  private foldLists(items: { node: JSONContent; list?: { numId: string; ilvl: number; sig: ListSig; start: number } }[]): JSONContent[] {
    const out: JSONContent[] = [];
    type Frame = { numId: string; level: number; node: JSONContent; last: JSONContent | null };
    let stack: Frame[] = [];
    for (const it of items) {
      if (!it.list) {
        stack = [];
        out.push(it.node);
        continue;
      }
      const { numId, ilvl, sig, start } = it.list;
      while (stack.length && (stack[stack.length - 1].level > ilvl || (stack[stack.length - 1].level === ilvl && stack[stack.length - 1].numId !== numId))) stack.pop();
      const top = stack[stack.length - 1];
      if (!top || top.level < ilvl) {
        const key = `${numId}:${ilvl}`;
        const seen = this.listCounts.get(key) ?? 0;
        const attrs: Record<string, unknown> = { listStyle: sig.style, multilevel: sig.multilevel };
        if (sig.type === "ordered") attrs.start = seen ? start + seen : start;
        const listNode: JSONContent = { type: sig.type === "bullet" ? "bulletList" : "orderedList", attrs, content: [] };
        if (!top || !top.last) out.push(listNode);
        else top.last.content!.push(listNode);
        stack.push({ numId, level: ilvl, node: listNode, last: null });
      }
      const frame = stack[stack.length - 1];
      const li: JSONContent = { type: "listItem", content: [it.node] };
      frame.node.content!.push(li);
      frame.last = li;
      const key = `${numId}:${ilvl}`;
      this.listCounts.set(key, (this.listCounts.get(key) ?? 0) + 1);
    }
    return out;
  }

  private paragraph(p: Element): { node: JSONContent; list?: { numId: string; ilvl: number; sig: ListSig; start: number }; ctx?: string }[] {
    const pPr = kid(p, "pPr");
    const styleId = attr(kid(pPr, "pStyle"), "val");
    const direct = parsePPr(pPr);
    const noteStyle = this.noteMode && /(foot|end)note|dipnot|son ?not/i.test(`${styleId ?? ""} ${this.r.sheet.styles.get(styleId ?? "")?.name ?? ""}`);
    const mapped = noteStyle ? { type: "paragraph" as const, ourId: "Normal" } : this.r.mapParagraphStyle(styleId);
    const styleDiff = noteStyle ? { r: {}, p: {} } : this.r.styleDiff(styleId, mapped.ourId);
    // Liste: doğrudan numPr ya da stil üzerinden.
    const stylePP = this.r.sheet.effective(styleId).p;
    const numId = direct.numId ?? stylePP.numId;
    const ilvl = direct.ilvl ?? stylePP.ilvl ?? (numId ? listStyleLevel(this.r.numbering.styleLevel(numId, styleId), this.r.sheet.styles.get(styleId ?? "")?.name) : 0);
    const list = numId ? this.r.numbering.resolve(numId, ilvl) : null;
    const pp: ParaProps = { ...styleDiff.p, ...stripUndef(direct) };
    const attrs: Record<string, unknown> = {};
    if (mapped.type === "heading") attrs.level = mapped.level;
    else attrs.styleId = list && mapped.ourId === "ListParagraph" ? null : mapped.ourId === "Normal" ? null : mapped.ourId;
    if (pp.align) attrs.textAlign = pp.align;
    if (!list) {
      if (pp.indentLeft) attrs.indentLeft = pp.indentLeft;
      if (pp.indentFirst) attrs.indentFirst = pp.indentFirst;
    }
    if (pp.indentRight) attrs.indentRight = pp.indentRight;
    if (pp.spaceBefore !== undefined) attrs.spaceBefore = pp.spaceBefore;
    if (pp.spaceAfter !== undefined) attrs.spaceAfter = pp.spaceAfter;
    if (pp.lineHeight !== undefined) {
      attrs.lineHeight = pp.lineHeight;
      attrs.lineRule = pp.lineRule && pp.lineRule !== "auto" ? pp.lineRule : null;
    }
    if (pp.shading) attrs.shading = pp.shading;
    if (pp.borders) attrs.borders = pp.borders;
    if (direct.tabs?.length) attrs.tabs = direct.tabs;
    if (direct.outlineLevel && mapped.type !== "heading") attrs.outlineLevel = direct.outlineLevel;
    if (direct.keepNext) attrs.keepNext = true;
    if (direct.keepLines) attrs.keepLines = true;
    if (direct.pageBreakBefore) attrs.pageBreakBefore = true;
    if (direct.suppressLineNumbers) attrs.suppressLineNumbers = true;
    // Modelin okumadığı paragraf özellikleri (çerçeve, satır bölme, paragraf işaretinin biçimi…) taşınır.
    const pExtra = extrasOf(pPr, PPR_MODELED, markRPr);
    if (pExtra) attrs.ooxmlExtra = pExtra;

    const extra: JSONContent[] = [];
    const segments: JSONContent[][] = [[]];
    this.segKinds = [];
    this.boxes = [];
    const wmBefore = this.watermark;
    this.inline(p, styleDiff.r, [], segments, extra);
    if (this.pendingPoints.length) this.attachPointsToLast(segments);
    if (this.pointIds.size) {
      for (const seg of segments) extendCommentToWord(seg, this.pointIds);
      this.pointIds.clear();
    }
    // Yalnız filigran taşıyan üst bilgi paragrafı atlanır.
    if (this.watermark !== wmBefore && segments.every((x) => x.length === 0) && !this.boxes.length) return [];
    const type = mapped.type === "heading" ? "heading" : "paragraph";
    const out: { node: JSONContent; list?: { numId: string; ilvl: number; sig: ListSig; start: number }; ctx?: string }[] = [];
    const ctx = (direct.contextual ?? stylePP.contextual) ? (styleId ?? this.r.sheet.defaultParagraph) : undefined;
    // Yatay çizgi: metni olmayan, yalnız alt kenarlıklı paragraf.
    const onlyBottom = pp.borders && pp.borders.bottom && !pp.borders.top && !pp.borders.left && !pp.borders.right;
    if (onlyBottom && segments.length === 1 && segments[0].length === 0 && mapped.ourId === "Normal" && !list) {
      return [{ node: { type: "horizontalRule" } }];
    }
    segments.forEach((content, i) => {
      if (i > 0) out.push({ node: { type: this.segKinds[i - 1] ?? "pageBreak" } });
      const empty = content.length === 0;
      // Sayfa sonunun önündeki/arkasındaki boş parçalar paragraf oluşturmaz.
      if (empty && segments.length > 1 && (i === 0 || i === segments.length - 1)) return;
      const node: JSONContent = { type, attrs: { ...attrs }, ...(empty ? {} : { content: mergeText(content) }) };
      // Başlıklar (anahat numaralandırması olsa da) listeye girmez: liste öğesi paragrafla başlamalıdır.
      out.push({ node, list: i === 0 && list && numId && type !== "heading" ? { numId, ilvl, sig: list.sig, start: list.start } : undefined, ctx });
    });
    for (const e of extra) out.push({ node: e });
    // Paragrafa bağlı metin kutuları paragraftan önce yer alır.
    if (this.boxes.length) out.unshift(...this.boxes.map((node) => ({ node })));
    return out;
  }

  /** Geçerli paragrafa bağlı metin kutuları. */
  boxes: JSONContent[] = [];

  /** Satır içi öğeleri dolaşır; sayfa sonu yeni bir parça başlatır. */
  private inline(el: Element, base: RunProps, marks: Mark[], segs: JSONContent[][], extra: JSONContent[]) {
    const push = (n: JSONContent) => segs[segs.length - 1].push(n);
    for (const c of kids(el)) {
      switch (c.localName) {
        case "r":
          this.run(c, base, marks, segs, extra);
          break;
        case "oMath":
        case "oMathPara": {
          // Denklem: OMML → doğrusal giriş.
          const list = c.localName === "oMath" ? [c] : kids(c, "oMath");
          for (const m of list) push({ type: "equation", attrs: { latex: toLinear(fromOmml(m)), display: c.localName === "oMathPara" } });
          break;
        }
        case "hyperlink": {
          const rid = c.getAttributeNS(NS.r, "id") ?? attr(c, "id");
          const anchor = attr(c, "anchor");
          let href: string | null = null;
          if (rid && this.rels.get(rid)) href = this.rels.get(rid)!.target;
          if (anchor) href = (href ?? "") + `#${anchor}`;
          this.inline(c, base, href ? [...marks.filter((m) => m.type !== "link"), { type: "link", attrs: { href } }] : marks, segs, extra);
          break;
        }
        case "ins":
        case "moveTo":
          this.inline(c, base, [...marks, { type: "insertion", attrs: { author: attr(c, "author") ?? "", date: attr(c, "date") ?? "" } }], segs, extra);
          break;
        case "del":
        case "moveFrom":
          this.inline(c, base, [...marks, { type: "deletion", attrs: { author: attr(c, "author") ?? "", date: attr(c, "date") ?? "" } }], segs, extra);
          break;
        case "fldSimple": {
          const instr = (attr(c, "instr") ?? "").trim();
          const kw = instr.split(/\s+/)[0]?.toUpperCase();
          const fnode = fieldNode(instr, kids(c).map((r) => descendants(r, "t").map((t) => t.textContent ?? "").join("")).join(""));
          if (fnode) push(fnode);
          else if (PAGE_FIELDS.has(kw)) push(pageFieldNode(kw, instr));
          else if (kw === "HYPERLINK") {
            const href = parseHyperlinkInstr(instr);
            this.inline(c, base, href ? [...marks, { type: "link", attrs: { href } }] : marks, segs, extra);
          } else {
            // Tanınmayan alan: görünen sonucu metin olarak okunur (alan kodu kaydedilmez).
            if (kw && kw !== "SHAPE") this.r.warnings.add("fields", `alan: ${kw}`);
            this.inline(c, base, marks, segs, extra);
          }
          break;
        }
        case "bookmarkStart": {
          const name = attr(c, "name");
          if (name && name !== "_GoBack") push({ type: "bookmark", attrs: { name } });
          break;
        }
        case "commentRangeStart":
          this.openComments.add(`c${attr(c, "id")}`);
          this.rangedComments.add(`c${attr(c, "id")}`);
          break;
        case "commentRangeEnd": {
          const id = `c${attr(c, "id")}`;
          this.openComments.delete(id);
          // Boş aralık (hiçbir metni kapsamayan açıklama) nokta açıklaması gibi bağlanır.
          if (this.rangedComments.has(id) && !this.markedComments.has(id) && !this.pendingPoints.includes(id)) this.pendingPoints.push(id);
          break;
        }
        case "smartTag":
        case "customXml":
          this.inline(c, base, marks, segs, extra);
          break;
        case "sdt": {
          const content = kid(c, "sdtContent");
          if (content) this.inline(content, base, marks, segs, extra);
          break;
        }
        case "AlternateContent": {
          const ch = kid(c, "Choice") ?? kid(c, "Fallback");
          if (ch) this.inline(ch, base, marks, segs, extra);
          break;
        }
        case "dir":
        case "bdo":
          // Yazı yönü kapsayıcıları: içerikleri okunur.
          this.inline(c, base, marks, segs, extra);
          break;
        default:
          // Paragraf içinde tanınmayan öğe (alt belge, bilinmeyen uzantı…): opak.
          if (!MARKERS.has(c.localName)) this.opaqueOr(c, "para", (n) => push(marks.length ? { ...n, marks } : n), () => undefined);
      }
    }
  }

  private run(r: Element, base: RunProps, outer: Mark[], segs: JSONContent[][], extra: JSONContent[]) {
    const rPr = kid(r, "rPr");
    const rStyle = attr(kid(rPr, "rStyle"), "val");
    const charStyle = this.r.mapCharStyle(rStyle);
    const props: RunProps = { ...base, ...(charStyle ? {} : stripUndef(this.r.sheet.charProps(rStyle))), ...stripUndef(parseRPr(rPr, this.r.theme)) };
    if (rStyle && /hyperlink/i.test(rStyle)) {
      delete props.color;
      delete props.underline;
    }
    const field = this.fields[this.fields.length - 1];
    let marks = [...outer, ...runMarks(props)];
    // Modelin okumadığı koşu özellikleri (dil, ölçek, w14 efektleri…) textStyle işaretinde taşınır.
    const rExtra = extrasOf(rPr, RPR_MODELED);
    if (rExtra) {
      const ts = marks.find((m) => m.type === "textStyle");
      if (ts) marks = marks.map((m) => (m === ts ? { ...m, attrs: { ...m.attrs, ooxmlExtra: rExtra } } : m));
      else marks.push({ type: "textStyle", attrs: { ooxmlExtra: rExtra } });
    }
    if (charStyle) marks.push({ type: "charStyle", attrs: { styleId: charStyle } });
    const change = kid(rPr, "rPrChange");
    if (change) marks.push(this.formatChange(change, base));
    for (const id of this.openComments) marks.push({ type: "comment", attrs: { id } });
    if (field?.phase === "result" && field.kind === "link" && field.href) marks = [...marks.filter((m) => m.type !== "link"), { type: "link", attrs: { href: field.href } }];
    const push = (n: JSONContent) => {
      if (n.type === "text" && n.text) {
        for (const id of this.openComments) this.markedComments.add(id);
        if (this.pendingPoints.length) {
          // Bekleyen nokta açıklaması bu metnin ilk sözcüğüne bağlanır.
          const m = /^\s*\S+/.exec(n.text);
          const head = m ? m[0] : n.text;
          const points = this.takePoints();
          segs[segs.length - 1].push({ ...n, text: head, marks: [...marks, ...points] });
          if (head.length < n.text.length) segs[segs.length - 1].push(marks.length ? { ...n, text: n.text.slice(head.length), marks } : { ...n, text: n.text.slice(head.length) });
          return;
        }
      }
      segs[segs.length - 1].push(marks.length && n.type !== "pageBreak" ? { ...n, marks } : n);
    };
    for (const c of kids(r)) {
      const f = this.fields[this.fields.length - 1];
      const suppress = f && (f.phase === "instr" || f.kind === "page" || f.kind === "numpages");
      switch (c.localName) {
        case "fldChar": {
          const t = attr(c, "fldCharType");
          if (t === "begin") this.fields.push({ instr: "", phase: "instr", kind: "other" });
          else if (t === "separate" && f) {
            f.phase = "result";
            const kw = f.instr.trim().split(/\s+/)[0]?.toUpperCase();
            if (kw === "HYPERLINK") {
              f.kind = "link";
              f.href = parseHyperlinkInstr(f.instr) ?? undefined;
              if (f.href) marks = [...marks.filter((m) => m.type !== "link"), { type: "link", attrs: { href: f.href } }];
            } else if (PAGE_FIELDS.has(kw)) {
              f.kind = kw === "NUMPAGES" ? "numpages" : "page";
              push(pageFieldNode(kw, f.instr));
              f.emitted = true;
            } else {
              // SEQ / REF / PAGEREF / CITATION / MERGEFIELD: düğüm olarak alınır, sonuç metni atlanır.
              const fnode = fieldNode(f.instr, "");
              if (fnode) {
                f.kind = "page";
                push(fnode);
                f.emitted = true;
              }
            }
          } else if (t === "end" && f) {
            const kw = f.instr.trim().split(/\s+/)[0]?.toUpperCase();
            if (!f.emitted && f.phase === "instr" && PAGE_FIELDS.has(kw)) push(pageFieldNode(kw, f.instr));
            if (!f.emitted && f.phase === "instr" && kw === "XE") {
              const xe = xeNode(f.instr);
              if (xe) push(xe);
            } else if (!f.emitted && f.kind !== "link" && kw && kw !== "SHAPE" && !PAGE_FIELDS.has(kw)) this.r.warnings.add("fields", `alan: ${kw}`);
            this.fields.pop();
            if (f.kind === "link") marks = marks.filter((m) => m.type !== "link");
          }
          break;
        }
        case "instrText":
          if (f) f.instr += c.textContent ?? "";
          break;
        case "t":
        case "delText":
          if (!suppress && c.textContent) push({ type: "text", text: c.textContent });
          break;
        case "tab":
        case "ptab":
          if (!suppress) push({ type: "text", text: "\t" });
          break;
        case "br":
          if (suppress) break;
          if (attr(c, "type") === "page" || attr(c, "type") === "column") {
            this.segKinds.push(attr(c, "type") === "column" ? "columnBreak" : "pageBreak");
            segs.push([]);
          } else push({ type: "hardBreak" });
          break;
        case "cr":
          if (!suppress) push({ type: "hardBreak" });
          break;
        case "noBreakHyphen":
          if (!suppress) push({ type: "text", text: "‑" });
          break;
        case "softHyphen":
          if (!suppress) push({ type: "text", text: "­" });
          break;
        case "sym": {
          const code = parseInt(attr(c, "char") ?? "", 16);
          if (!suppress && code) push({ type: "text", text: String.fromCodePoint(code >= 0xf000 ? code - 0xf000 : code) });
          break;
        }
        case "footnoteRef":
        case "endnoteRef":
          // Not metninin başındaki kendi başvuru işareti atlanır.
          break;
        case "commentReference": {
          // Aralığı olmayan (yalnız başvurusu bulunan) açıklama: konumundaki sözcüğe bağlanır.
          const id = `c${attr(c, "id")}`;
          if (!this.rangedComments.has(id) && !this.pendingPoints.includes(id)) this.pendingPoints.push(id);
          break;
        }
        case "footnoteReference":
        case "endnoteReference": {
          if (suppress) break;
          const kind = c.localName === "footnoteReference" ? "footnote" : "endnote";
          const nid = attr(c, "id") ?? "";
          const note = (kind === "footnote" ? this.r.footnotes : this.r.endnotes).get(nid);
          segs[segs.length - 1].push({ type: "note", attrs: { kind, text: note?.text ?? "", content: note?.content ?? null, id: `n${kind[0]}${nid}` } });
          break;
        }
        case "drawing":
          // Resim, grafik ve metin kutusu modele okunur; şekil, grup, SmartArt, tuval… opak.
          if (suppress) break;
          if (this.drawingKnown(c)) readDrawing(this, c, push, extra);
          else this.opaqueOr(c, "run", push, () => readDrawing(this, c, push, extra));
          break;
        case "object":
          // Gömülü (OLE) nesne / denetim: opak (kaydederken nesnesi ve önizlemesiyle korunur).
          if (!suppress)
            this.opaqueOr(c, "run", push, () => {
              this.loseObject("w:object");
              readPict(this, c, push, extra);
            });
          break;
        case "control":
          if (!suppress) this.opaqueOr(c, "run", push, () => this.loseObject("w:control"));
          break;
        case "pict":
          if (suppress || PartReader.pictEmpty(c)) break;
          if (this.pictKnown(c)) readPict(this, c, push, extra);
          else this.opaqueOr(c, "run", push, () => readPict(this, c, push, extra));
          break;
        case "AlternateContent": {
          if (suppress) break;
          const ch = kid(c, "Choice") ?? kid(c, "Fallback");
          const dr = ch ? descendants(ch, "drawing")[0] : null;
          const pc = !dr && ch ? descendants(ch, "pict")[0] : undefined;
          const legacy = () => {
            if (dr) readDrawing(this, dr, push, extra);
            else if (pc) readPict(this, pc, push, extra);
          };
          // Anlaşılan seçenek (resim, grafik, metin kutusu) okunur; anlaşılmayan bütünüyle opak taşınır.
          if (dr ? this.drawingKnown(dr) : pc ? this.pictKnown(pc) || PartReader.pictEmpty(pc) : !ch) legacy();
          else this.opaqueOr(c, "run", push, legacy);
          break;
        }
        case "pgNum":
          if (!suppress) push(pageFieldNode("PAGE", "PAGE"));
          break;
        default:
          // Tanınmayan koşu çocuğu (tarih blokları, fonetik kılavuz, mürekkep…): opak.
          if (!suppress && !RUN_SKIP.has(c.localName)) this.opaqueOr(c, "run", push, () => undefined);
      }
    }
  }

  /** Bu paragrafta noktaya bağlanan açıklamalar (sözcüğün tamamına yayılır). */
  private pointIds = new Set<string>();

  private takePoints(): Mark[] {
    const out = this.pendingPoints.map((id) => ({ type: "comment", attrs: { id } }));
    for (const id of this.pendingPoints) {
      this.markedComments.add(id);
      this.pointIds.add(id);
    }
    this.pendingPoints = [];
    return out;
  }

  /** Paragraf sonunda bekleyen nokta açıklamaları paragrafın son metnine bağlanır (metin yoksa sonraki paragrafa kalır). */
  private attachPointsToLast(segs: JSONContent[][]) {
    for (let s = segs.length - 1; s >= 0; s--) {
      for (let i = segs[s].length - 1; i >= 0; i--) {
        const n = segs[s][i];
        if (n.type === "text" && n.text) {
          segs[s][i] = { ...n, marks: [...(n.marks ?? []), ...this.takePoints()] };
          return;
        }
      }
    }
  }

  /** w:rPrChange: değişiklik öncesi biçim, "formatChange" işaretinin "old" özniteliğine. */
  private formatChange(change: Element, base: RunProps): Mark {
    const oldPr = kid(change, "rPr");
    const oldStyle = attr(kid(oldPr, "rStyle"), "val");
    const oldChar = this.r.mapCharStyle(oldStyle);
    const props: RunProps = { ...base, ...(oldChar ? {} : stripUndef(this.r.sheet.charProps(oldStyle))), ...stripUndef(parseRPr(oldPr, this.r.theme)) };
    const old = runMarks(props);
    if (oldChar) old.push({ type: "charStyle", attrs: { styleId: oldChar } });
    return { type: "formatChange", attrs: { author: attr(change, "author") ?? "", date: attr(change, "date") ?? "", old: JSON.stringify(old) } };
  }

  private table(tbl: Element): JSONContent {
    const tblPr = kid(tbl, "tblPr");
    const styleId = attr(kid(tblPr, "tblStyle"), "val");
    const grid = kids(kid(tbl, "tblGrid"), "gridCol").map((g) => (attrNum(g, "w") ?? 0) / 15);
    const look = kid(tblPr, "tblLook");
    const lookVal = parseInt(attr(look, "val") ?? "0", 16) || 0;
    const flag = (name: string, bit: number) => {
      const v = attr(look, name);
      return v !== null ? v === "1" || v === "true" : (lookVal & bit) !== 0;
    };
    const ownBorders = kid(tblPr, "tblBorders");
    const bordersNone = ownBorders ? kids(ownBorders).every((b) => ["none", "nil"].includes(attr(b, "val") ?? "none")) : this.r.sheet.styles.get(styleId ?? "")?.tblBordersNone ?? !styleId;
    let tableStyle = tableStyleFromId(styleId);
    if (!tableStyle || (tableStyle === "none" && !bordersNone)) tableStyle = bordersNone ? "none" : "grid";
    const jc = attr(kid(tblPr, "jc"), "val");
    const tblW = kid(tblPr, "tblW");
    const fitWindow = attr(tblW, "type") === "pct" && (attrNum(tblW, "w") ?? 0) >= 4900;
    // Satırları ve hücre ızgara konumlarını topla.
    type Cell = { el: Element; col: number; span: number; vm: "restart" | "continue" | null };
    const trs = kids(tbl, "tr");
    const rows: Cell[][] = trs.map((tr) => {
      let col = 0;
      const cells: Cell[] = [];
      const gridBefore = attrNum(path(tr, "trPr", "gridBefore"), "val") ?? 0;
      col += gridBefore;
      for (const tc of kids(tr, "tc")) {
        const tcPr = kid(tc, "tcPr");
        // Geçersiz/dev birleştirme (bozuk dosya) sütun sınırına (63) indirgenir.
        const span = Math.max(1, Math.min(63, Math.round(attrNum(kid(tcPr, "gridSpan"), "val") ?? 1) || 1));
        const vmEl = kid(tcPr, "vMerge");
        const vm = vmEl ? (attr(vmEl, "val") === "restart" ? "restart" : "continue") : null;
        cells.push({ el: tc, col, span, vm });
        col += span;
      }
      return cells;
    });
    // Satır yüksekliği (trHeight, twip → px): kural yoksa ya da "atLeast" ise en az, "exact" ise tam; "auto" yok sayılır.
    const rowHeight = (tr: Element) => {
      const h = path(tr, "trPr", "trHeight");
      const tw = attrNum(h, "val") ?? 0;
      const rule = attr(h, "hRule");
      return h && tw > 0 && rule !== "auto" ? { height: Math.round((tw / 15) * 100) / 100, heightRule: rule === "exact" ? "exact" : null } : null;
    };
    const heights = trs.map(rowHeight);
    const content: JSONContent[] = rows.map((cells, ri) => ({
      type: "tableRow",
      ...(heights[ri] ? { attrs: heights[ri]! } : {}),
      content: cells
        .filter((c) => c.vm !== "continue")
        .map((c) => {
          let rowspan = 1;
          if (c.vm === "restart") {
            for (let k = ri + 1; k < rows.length; k++) {
              if (rows[k].some((x) => x.col === c.col && x.vm === "continue")) rowspan++;
              else break;
            }
          }
          const tcPr = kid(c.el, "tcPr");
          const bg = attr(kid(tcPr, "shd"), "fill");
          const va = attr(kid(tcPr, "vAlign"), "val");
          const widths = grid.length >= c.col + c.span ? grid.slice(c.col, c.col + c.span).map((w) => Math.round(w)) : null;
          let blocks = this.blocks(c.el);
          if (!blocks.length) blocks = [{ type: "paragraph" }];
          return {
            type: "tableCell",
            attrs: {
              colspan: c.span,
              rowspan,
              colwidth: !fitWindow && widths && widths.every((w) => w > 0) ? widths : null,
              background: bg && /^[0-9a-f]{6}$/i.test(bg) ? `#${bg.toUpperCase()}` : null,
              verticalAlign: va === "center" ? "middle" : va === "bottom" ? "bottom" : null,
            },
            content: blocks,
          };
        }),
    }));
    const firstTr = trs[0];
    const hdr = path(firstTr, "trPr", "tblHeader");
    const repeatHeader = !!hdr && !["0", "false", "off"].includes(attr(hdr, "val") ?? "");
    const mar = kid(tblPr, "tblCellMar");
    const pt = (name: string, alt: string, def: number) => {
      const el = kid(mar, name) ?? kid(mar, alt);
      const w = attrNum(el, "w");
      return el && w !== null && attr(el, "type") !== "pct" ? w / 20 : def;
    };
    const cellMargins = mar ? { top: pt("top", "top", 0), bottom: pt("bottom", "bottom", 0), left: pt("left", "start", 5.4), right: pt("right", "end", 5.4) } : null;
    const align = jc === "center" || jc === "right" ? jc : jc === "end" ? "right" : null;
    // Soldan girinti (tblInd, twip → px); yüzde ve "nil" girintisi yok sayılır.
    const ind = kid(tblPr, "tblInd");
    const indType = attr(ind, "type");
    const indent = ind && (!indType || indType === "dxa") ? Math.round((attrNum(ind, "w") ?? 0) / 15) || 0 : 0;
    return {
      type: "table",
      attrs: {
        tableStyle,
        headerRow: flag("firstRow", 0x20),
        bandedRows: !flag("noHBand", 0x200),
        firstColumn: flag("firstColumn", 0x80),
        totalRow: flag("lastRow", 0x40),
        align,
        ...(indent && !align ? { indent } : {}),
        repeatHeader,
        cellMargins,
      },
      content: content.filter((r) => (r.content?.length ?? 0) > 0),
    };
  }

  private toc(content: Element): JSONContent {
    const entries: TocEntry[] = [];
    let title = "İçindekiler";
    for (const p of descendants(content, "p")) {
      const sid = attr(path(p, "pPr", "pStyle"), "val") ?? "";
      const info = this.r.sheet.styles.get(sid);
      const name = (info?.name ?? sid).toLowerCase();
      const text = descendants(p, "t")
        .map((t) => t.textContent ?? "")
        .join("");
      const lm = /^toc\s*(\d)$/.exec(name) ?? /^TOC(\d)$/i.exec(sid);
      if (lm) {
        // Metin ve sayfa numarası son sekmeyle ayrılır.
        const parts: string[] = [""];
        for (const r of descendants(p, "r")) {
          for (const c of kids(r)) {
            if (c.localName === "tab") parts.push("");
            else if (c.localName === "t") parts[parts.length - 1] += c.textContent ?? "";
          }
        }
        const page = parseInt(parts.length > 1 ? parts[parts.length - 1] : "", 10);
        const label = (parts.length > 1 ? parts.slice(0, -1).join(" ") : parts[0]).trim();
        if (label) entries.push({ level: Number(lm[1]), text: label, page: Number.isFinite(page) ? page : 0, pos: 0 });
      } else if (text.trim() && !entries.length && /toc heading|tocheading/i.test(name + sid)) {
        title = text.trim();
      } else if (text.trim() && !entries.length) title = text.trim();
    }
    return { type: "toc", attrs: { entries, title } };
  }
}

function parseHyperlinkInstr(instr: string): string | null {
  const anchor = /\\l\s+"([^"]+)"/.exec(instr);
  const url = /HYPERLINK\s+"([^"]+)"/i.exec(instr);
  if (url) return url[1] + (anchor ? `#${anchor[1]}` : "");
  if (anchor) return `#${anchor[1]}`;
  return null;
}

/** Not metni paragraflarından not stili (Dipnot Metni) temizlenir. */
export function stripNoteStyle(b: JSONContent): JSONContent {
  if (b.type !== "paragraph" || !b.attrs) return b;
  const attrs = { ...b.attrs };
  if (attrs.styleId && /note|not/i.test(String(attrs.styleId))) attrs.styleId = null;
  return { ...b, attrs };
}

/** Koşu özelliklerinden işaretler. */
export function runMarks(p: RunProps): Mark[] {
  const m: Mark[] = [];
  if (p.bold) m.push({ type: "bold" });
  if (p.italic) m.push({ type: "italic" });
  if (p.underline) m.push({ type: "underline", attrs: { type: p.underline.type, color: p.underline.color } });
  if (p.strike) m.push({ type: "strike" });
  if (p.vertAlign === "superscript") m.push({ type: "superscript" });
  if (p.vertAlign === "subscript") m.push({ type: "subscript" });
  if (p.highlight) m.push({ type: "highlight", attrs: { color: p.highlight } });
  const ts: Record<string, unknown> = {};
  if (p.font) ts.fontFamily = p.font;
  if (p.size) ts.fontSize = `${p.size}pt`;
  if (p.color) ts.color = p.color;
  if (p.shd) ts.backgroundColor = p.shd;
  if (p.caps) ts.caps = true;
  if (p.smallCaps) ts.smallCaps = true;
  if (p.vanish) ts.hidden = true;
  if (p.dstrike) ts.doubleStrike = true;
  for (const k of ["shadow", "outline", "emboss", "imprint"] as const) if (p[k]) ts[k] = true;
  if (p.spacing) ts.letterSpacing = p.spacing;
  if (Object.keys(ts).length) m.push({ type: "textStyle", attrs: ts });
  return m;
}

/** Aynı işaretli bitişik metin düğümlerini birleştirir. */
function mergeText(nodes: JSONContent[]): JSONContent[] {
  const out: JSONContent[] = [];
  for (const n of nodes) {
    const last = out[out.length - 1];
    if (n.type === "text" && last?.type === "text" && JSON.stringify(last.marks ?? []) === JSON.stringify(n.marks ?? [])) last.text = (last.text ?? "") + (n.text ?? "");
    else if (n.type !== "text" || n.text) out.push({ ...n });
  }
  return out;
}

/**
 * Stilden gelen liste düzeyi. "List Bullet 2" / "Liste Madde İşareti 3" gibi stiller kendi numaralandırma
 * tanımlarında 0. düzeyde olsa da Word'de bir alt düzey gibi girintili görünür: adındaki sayıdan düzey çıkarılır.
 */
export function listStyleLevel(level: number | null, name: string | undefined): number {
  if (level && level > 0) return level;
  const m = /^(?:list|liste)\b.*?\s(\d)$/i.exec((name ?? "").trim());
  return m ? Math.max(0, Math.min(8, Number(m[1]) - 1)) : (level ?? 0);
}

/** Alan yönergesinden düğüm (SEQ, REF, PAGEREF, CITATION, MERGEFIELD); tanınmayan alan için null. */
export function fieldNode(instr: string, result: string): JSONContent | null {
  const t = instr.trim();
  const parts = t.match(/"[^"]*"|\S+/g) ?? [];
  const kw = parts[0]?.toUpperCase();
  const arg = (parts[1] ?? "").replace(/^"|"$/g, "");
  const sw = (s: string) => new RegExp(`\\\\${s}(\\s|$)`, "i").test(t);
  switch (kw) {
    case "SEQ":
      return { type: "seqField", attrs: { label: arg || "Şekil", num: parseInt(result, 10) || 1 } };
    case "REF":
      return arg ? { type: "crossRef", attrs: { ref: arg, kind: sw("p") ? "position" : "full", link: sw("h"), text: result } } : null;
    case "PAGEREF":
      return arg ? { type: "crossRef", attrs: { ref: arg, kind: "page", link: sw("h"), text: result } } : null;
    case "CITATION": {
      const pm = /\\p\s+("[^"]*"|\S+)/i.exec(t);
      return arg ? { type: "citation", attrs: { tag: arg, pages: pm ? pm[1].replace(/"/g, "") : "", text: result } } : null;
    }
    case "MERGEFIELD":
      return arg ? { type: "mergeField", attrs: { name: arg } } : null;
    default:
      return null;
  }
}

/** Şekiller Tablosu: İçindekiler ayrıştırıcısının girdileri + \\c etiketi. */
function tofNode(toc: JSONContent, instr: string): JSONContent {
  const label = /\\c\s+"([^"]+)"/.exec(instr)?.[1] ?? "Şekil";
  const entries = ((toc.attrs?.entries as TocEntry[]) ?? []).map((e) => ({ ...e, level: 1 }));
  return { type: "toc", attrs: { kind: "tof", label, title: "", entries } };
}

/** Kaynakça bloğu (girdiler kaynaklardan yeniden üretilir; başlık korunur). */
/** Kaynakça bloğu: başlık ve girdiler (italik bölüm ayrı; kaynaklar belgede varsa alan güncellemesi yeniler). */
function bibliographyNode(content: Element): JSONContent {
  const [first, ...rest] = descendants(content, "p");
  const title = first ? descendants(first, "t").map((t) => t.textContent ?? "").join("").trim() : "";
  const entries = rest
    .map((p) => {
      const e = { before: "", italic: "", after: "" };
      for (const r of descendants(p, "r")) {
        const text = kids(r, "t").map((t) => t.textContent ?? "").join("");
        if (!text) continue;
        const i = path(r, "rPr", "i");
        const italic = !!i && !["0", "false"].includes(attr(i, "val") ?? "");
        if (italic && !e.after) e.italic += text;
        else if (e.italic) e.after += text;
        else e.before += text;
      }
      return e;
    })
    .filter((e) => e.before || e.italic || e.after);
  return { type: "bibliography", attrs: { title: title || "Kaynakça", entries } };
}
