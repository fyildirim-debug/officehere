// DOCX parça yazıcısı: ProseMirror JSON blokları → WordprocessingML (paragraflar, koşular, izlenen
// değişiklikler, resimler, metin kutuları, tablolar, İçindekiler). İlişkiler (Rels) ve ortak sayaçlar (Ctx).
//
// Geçiş katmanı, öğe düzeyi: opak düğümlerin özgün XML'i yazılır. İlişkileri `reattach` ile yazma turunun
// taşıma planına (ctx.carry; parçalar carryParts sırasında kaynak paketten kopyalanır) bağlanır; belge
// genelinde benzersiz olması gereken kimlikler (wp:docPr/@id, w14:paraId, yer işareti ve düzeltme kimlikleri,
// yinelenen VML şekil kimlikleri) yeniden numaralanır. Bağlanamayan öğenin yerine yedek görünümü (önizleme
// resmi ya da yer tutucu metin) yazılır. Paragraf/koşu ek özellikleri (ooxmlExtra) şema sırasıyla katılır.
import type { JSONContent } from "@tiptap/core";
import { explicitRelIds, getSource, reattach, relativeTarget, type Carry } from "@officehere/core/opc";
import { opaquePlaceholder, type OpaqueRef } from "../../editor/opaque";
import type { DocSettings } from "../../model/document";
import { LINE_FACTOR, type BorderSpec } from "../../model/styles";
import type { ParaBorders } from "../../editor/paragraphFormat";
import { fieldNowText, fieldPicture, type TocEntry } from "../../editor/nodes";
import { firstGrapheme, normalizeDropCap, type DropCap } from "../../editor/dropCap";
import { parseMath } from "../../editor/math/ast";
import type { ChartData } from "../../editor/chart";
import { CHART_NS, REL_CHART, chartSpaceXml } from "./chartXml";
import { indexXml, xeXml } from "./indexXml";
import { MATH_NS, ommlOf } from "../../editor/math/render";
import { NS, REL, XML_HEAD, cleanText, esc, hex6, highlightName } from "./xml";
import { borderXml, fontsXml, tableStyleId } from "./styleXml";
import { sigKey, type ListSig } from "./numbering";

export type Mark = { type: string; attrs?: Record<string, unknown> };

export const DOC_NS = `xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="${NS.mc}" xmlns:o="${NS.o}" xmlns:r="${NS.r}" xmlns:v="${NS.v}" xmlns:w10="${NS.w10}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}" xmlns:w="${NS.w}" xmlns:w14="${NS.w14}" xmlns:wp14="${NS.wp14}" xmlns:wps="${NS.wps}" xmlns:m="${MATH_NS}"`;

/** Parça köklerinde (DOC_NS) bildirilen ön ekler → ad alanları. Opak XML ve ek özellikler bunlara göre yazılır. */
export const DOC_PREFIXES: ReadonlyMap<string, string> = new Map([...DOC_NS.matchAll(/xmlns:([\w.-]+)="([^"]+)"/g)].map((m) => [m[1]!, m[2]!]));

/**
 * Çıktı kökünün anladığı (şemadaki) ya da yok sayılabilir ilan ettiği (mc:Ignorable="w14 wp14") ad alanları.
 * Kökü bunların dışında kalan opak öğe mc:AlternateContent ile sarılarak yazılır.
 */
export const OUTPUT_NAMESPACES: ReadonlySet<string> = new Set([NS.w, NS.r, NS.wp, NS.a, NS.pic, NS.v, NS.o, NS.w10, NS.mc, MATH_NS, NS.w14, NS.wp14]);

const EMU_PER_PX = 9525;

export interface NoteData {
  text: string;
  content: JSONContent[] | null;
}

interface Media {
  name: string;
  data: Uint8Array;
  ext: string;
}

/** Bir parça (belge, üst/alt bilgi) için ilişki tablosu. */
export class Rels {
  private list: { id: string; type: string; target: string; external?: boolean }[] = [];
  add(type: string, target: string, external = false): string {
    const hit = this.list.find((r) => r.type === type && r.target === target);
    if (hit) return hit.id;
    const id = `rId${this.list.length + 1}`;
    this.list.push({ id, type, target, external });
    return id;
  }
  xml(): string {
    return `${XML_HEAD}<Relationships xmlns="${NS.rel}">${this.list
      .map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${esc(r.target)}"${r.external ? ' TargetMode="External"' : ""}/>`)
      .join("")}</Relationships>`;
  }
  get size() {
    return this.list.length;
  }
}

function dataUrlToBytes(src: string): { data: Uint8Array; ext: string } | null {
  const m = /^data:([^;,]+)(;base64)?,(.*)$/s.exec(src);
  if (!m) return null;
  const mime = m[1].toLowerCase();
  const ext = mime.includes("png") ? "png" : mime.includes("jpeg") || mime.includes("jpg") ? "jpeg" : mime.includes("gif") ? "gif" : mime.includes("bmp") ? "bmp" : mime.includes("svg") ? "svg" : mime.includes("webp") ? "webp" : "png";
  let bytes: Uint8Array;
  if (m[2]) {
    const bin = atob(m[3]);
    bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  } else bytes = new TextEncoder().encode(decodeURIComponent(m[3]));
  return { data: bytes, ext };
}

const ptTw = (pt: number) => Math.round(pt * 20);

/* ------------------------------------------------------------------ */
/* Opak XML ve ek özellikler                                             */
/* ------------------------------------------------------------------ */

/** CT_PPr alt öğe sırası (ECMA-376 Bölüm 1, §17.3.1.26). */
const PPR_ORDER = [
  "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
  "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
  "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
  "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
];
/** CT_RPr alt öğe sırası (§17.3.2.28). */
const RPR_ORDER = [
  "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof",
  "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd",
  "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange",
];
const rankOf = (order: string[]) => new Map(order.map((n, i) => [n, i]));
const PPR_RANK = rankOf(PPR_ORDER);
const RPR_RANK = rankOf(RPR_ORDER);
/** Ana ad alanı dışındaki (w14…) ek özelliklerin yeri: pPr'de cnfStyle'dan, rPr'de oMath'tan sonra. */
const PPR_OTHER = PPR_RANK.get("cnfStyle")! + 0.5;
const RPR_OTHER = RPR_RANK.get("oMath")! + 0.5;

const XML_TOKEN = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<!|<(\/?)([A-Za-z_][\w.-]*(?::[A-Za-z_][\w.-]*)?)((?:\s+[^\s=>/]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>/g;

/**
 * XML parçacığının üst düzey öğeleri; iyi biçimli değilse (dengesiz etiket, üst düzeyde metin, DOCTYPE ya da
 * çözülemeyen "<") null. Opak XML ve ek özellikler yazılmadan önce bununla denetlenir (belgenin yapısını bozamaz).
 */
export function topElements(xml: string): { prefix: string; local: string; xml: string }[] | null {
  const out: { prefix: string; local: string; xml: string }[] = [];
  const stack: string[] = [];
  let start = 0;
  let last = 0;
  XML_TOKEN.lastIndex = 0;
  for (let m = XML_TOKEN.exec(xml); m; m = XML_TOKEN.exec(xml)) {
    const between = xml.slice(last, m.index);
    if (between.includes("<") || (!stack.length && between.trim())) return null;
    last = m.index + m[0].length;
    const name = m[2];
    if (!name) {
      if (m[0] === "<!") return null;
      if (!stack.length && !m[0].startsWith("<!--") && !m[0].startsWith("<?")) return null;
      continue;
    }
    const top = (from: number) => {
      const i = name.indexOf(":");
      out.push({ prefix: i < 0 ? "" : name.slice(0, i), local: name.slice(i + 1), xml: xml.slice(from, last) });
    };
    if (m[1]) {
      if (stack.pop() !== name) return null;
      if (!stack.length) top(start);
    } else if (m[4]) {
      if (!stack.length) top(m.index);
    } else {
      if (!stack.length) start = m.index;
      stack.push(name);
    }
  }
  if (stack.length || xml.slice(last).includes("<") || xml.slice(last).trim()) return null;
  return out;
}

/** Ek özelliklerin üst düzey öğeleri ve kapsayıcı mc:Ignorable'ı (aynı dize çok koşuda yinelenir: önbellekli). */
const extraCache = new Map<string, { items: ReturnType<typeof topElements>; ignorable: string }>();
function extraInfo(extra: string) {
  let hit = extraCache.get(extra);
  if (!hit) {
    if (extraCache.size >= 2000) extraCache.clear();
    hit = { items: topElements(extra), ignorable: ignorableScan(extra) };
    extraCache.set(extra, hit);
  }
  return hit;
}
const extraElements = (extra: string) => extraInfo(extra).items;

/**
 * Yazıcının öğeleri (şema sırasında) ile ek özellik öğelerini şema sırasıyla birleştirir. Yazıcının da yazdığı
 * bir öğe ekte de varsa yazıcınınki geçerlidir. Ek iyi biçimli değilse yok sayılır.
 */
function mergeOrdered(parts: [string, string][], extra: unknown, rank: ReadonlyMap<string, number>, other: number): string {
  const own = parts.map(([, x]) => x).join("");
  if (typeof extra !== "string" || !extra) return own;
  const items = extraElements(extra);
  if (!items?.length) return own;
  const have = new Set(parts.map(([n]) => n));
  const add = items.filter((e) => !(e.prefix === "w" && have.has(e.local)));
  if (!add.length) return own;
  const all = [
    ...parts.map(([name, xml], i) => ({ k: rank.get(name) ?? other, i, xml })),
    ...add.map((e, j) => ({ k: e.prefix === "w" ? (rank.get(e.local) ?? other) : other, i: parts.length + j, xml: e.xml })),
  ];
  all.sort((p, q) => p.k - q.k || p.i - q.i);
  return all.map((x) => x.xml).join("");
}

/**
 * Ek özellikte ana ad alanı dışında (w14…) öğe varsa kapsayıcıya (w:pPr/w:rPr) yazılacak mc:Ignorable: üst/alt
 * bilgi ve not parçalarının kökünde bildirim yok; bu öğeleri anlamayan okuyucu da yok sayabilsin.
 */
function ignorableOf(extra: unknown): string {
  return typeof extra === "string" && extra ? extraInfo(extra).ignorable : "";
}

function ignorableScan(extra: string): string {
  const prefixes = new Set<string>();
  for (const m of extra.matchAll(/<([\w.-]+):|\s([\w.-]+):[\w.-]+\s*=/g)) {
    const p = (m[1] ?? m[2])!;
    if (p !== "xmlns" && IGNORABLE_PREFIXES.has(p)) prefixes.add(p);
  }
  return prefixes.size ? ` mc:Ignorable="${[...prefixes].join(" ")}"` : "";
}

/** Parça köklerinde bildirilen, şemada olmayan (yok sayılabilir) uzantı ad alanlarının önekleri. */
const IGNORABLE_PREFIXES = new Set([...DOC_PREFIXES].filter(([, uri]) => uri === NS.w14 || uri === NS.wp14).map(([p]) => p));

/** Düzeltme (izlenen değişiklik) öğeleri: w:id belge genelinde benzersiz. */
const REVISION_TAG = /<(?:[\w.-]+:)?(?:ins|del|moveFrom|moveTo|rPrChange|pPrChange|sectPrChange|tblPrChange|trPrChange|tcPrChange|tblGridChange|tblPrExChange|numberingChange|cellIns|cellDel|cellMerge)\b[^>]*>/g;
const ID_ATTR = /(\s(?:[\w.-]+:)?id\s*=\s*)("[^"]*"|'[^']*')/;
/** Opak öğede belgenin açıklama ve not parçalarına başvurular (bu parçalar yeniden üretilir; başvuru kırık kalırdı). */
const DANGLING_REF = /<(?:[\w.-]+:)?(?:commentRangeStart|commentRangeEnd|commentReference|footnoteReference|endnoteReference)\b[^>]*\/>/g;

/**
 * Opak XML'deki belge genelinde benzersiz kimlikler: wp:docPr/@id, w14:paraId, yer işareti (başlangıç ve bitiş
 * birlikte) ve düzeltme kimlikleri her yazımda yeniden verilir; VML şekil kimlikleri (v:shape/@id ↔
 * o:OLEObject/@ShapeID, o:spid) ve OLE nesne kimliği yalnız belgede yinelenirse yenilenir.
 */
function renumberIds(xml: string, ctx: Ctx): string {
  let out = xml.replace(DANGLING_REF, "");
  out = out.replace(/<(?:[\w.-]+:)?docPr\b[^>]*>/g, (tag) => tag.replace(/(\sid\s*=\s*)("[^"]*"|'[^']*')/, (_, h: string) => `${h}"${ctx.drawingId++}"`));
  out = out.replace(/(\s[\w.-]+:paraId\s*=\s*)("[^"]*"|'[^']*')/g, (_, h: string) => `${h}"${(0x10000000 + ++ctx.paraIds).toString(16).toUpperCase()}"`); // yerel-ayar: ascii
  const bookmarks = new Map<string, number>();
  out = out.replace(/<(?:[\w.-]+:)?bookmark(?:Start|End)\b[^>]*>/g, (tag) =>
    tag.replace(ID_ATTR, (_, h: string, v: string) => {
      let n = bookmarks.get(v);
      if (n === undefined) {
        n = ctx.bookmarkId++;
        bookmarks.set(v, n);
      }
      return `${h}"${n}"`;
    }),
  );
  out = out.replace(REVISION_TAG, (tag) => tag.replace(ID_ATTR, (_, h: string) => `${h}"${ctx.revId++}"`));
  return uniqueVmlIds(out, ctx);
}

function uniqueVmlIds(xml: string, ctx: Ctx): string {
  const found = new Set<string>();
  for (const m of xml.matchAll(/"(_x0000_[is]\d+)"/g)) found.add(m[1]!);
  for (const m of xml.matchAll(/\sObjectID\s*=\s*"(_\d+)"/g)) found.add(m[1]!);
  const map = new Map<string, string>();
  for (const id of found) {
    if (!ctx.vmlIds.has(id)) {
      ctx.vmlIds.add(id);
      continue;
    }
    let next: string;
    do next = id.startsWith("_x0000_") ? `${id.slice(0, 8)}${4096 + ++ctx.vmlNext}` : `_${1900000000 + ++ctx.vmlNext}`;
    while (ctx.vmlIds.has(next) || found.has(next));
    ctx.vmlIds.add(next);
    map.set(id, next);
  }
  return map.size ? xml.replace(/"(_x0000_[is]\d+|_\d+)"/g, (m, v: string) => (map.has(v) ? `"${map.get(v)}"` : m)) : xml;
}

/** Yazım bağlamı: belge genelinde sayaçlar. */
export class Ctx {
  media: Media[] = [];
  numbering = new Map<string, { abstractId: number; sig: ListSig }>();
  nums: { numId: number; abstractId: number; start: number }[] = [];
  bookmarkId = 0;
  drawingId = 1;
  commentIds = new Map<string, number>();
  openComments = new Set<string>();
  usedComments = new Set<string>();
  revId = 1;
  footnotes: NoteData[] = [];
  endnotes: NoteData[] = [];
  /** Grafik parçaları (word/charts/chartN.xml). */
  charts: { name: string; xml: string }[] = [];
  /** Öğe düzeyi taşıma planı (writeDocx kurar; yoksa ilişkili opak öğeler yedek görünümle yazılır). */
  carry?: Carry;
  /** Opak öğelerde yeniden verilen w14:paraId sayacı. */
  paraIds = 0;
  /** Yazılan VML şekil / OLE nesne kimlikleri (filigran varsa onun şekil kimliği baştan ayrılır). */
  vmlIds = new Set<string>();
  vmlNext = 0;
  constructor(public settings: DocSettings) {
    if (settings.watermark) this.vmlIds.add("_x0000_s1025");
  }

  commentNum(id: string): number {
    let n = this.commentIds.get(id);
    if (n === undefined) {
      n = this.commentIds.size;
      this.commentIds.set(id, n);
    }
    return n;
  }

  numFor(sig: ListSig, start: number): number {
    const key = sigKey(sig);
    let abs = this.numbering.get(key);
    if (!abs) {
      abs = { abstractId: this.numbering.size, sig };
      this.numbering.set(key, abs);
    }
    const numId = this.nums.length + 1;
    this.nums.push({ numId, abstractId: abs.abstractId, start });
    return numId;
  }
}

/** Bir parçanın (gövde/üst bilgi/alt bilgi) bloklarını yazar. */
export class PartWriter {
  /** Bölümün son paragrafı ve ona eklenecek bölüm özellikleri (OOXML: sectPr son paragrafın pPr'ındadır). */
  sectTarget: JSONContent | null = null;
  sectXml = "";
  constructor(
    private ctx: Ctx,
    private rels: Rels,
    public contentWidthTw: number,
  ) {}

  /** Bir sonraki paragrafın ilk koşusuna yazılacak bağlı nesneler (metin kutuları). */
  private pendingDrawings = "";

  blocks(nodes: JSONContent[] | undefined, list?: { numId: number; ilvl: number }): string {
    let out = "";
    const arr = nodes ?? [];
    arr.forEach((n, i) => {
      if (n.type === "textBox") {
        // Metin kutusu sonraki paragrafa bağlanır (OOXML: çapa bir koşu içindedir); yoksa kendi paragrafına.
        this.pendingDrawings += `<w:r>${this.textBox(n)}</w:r>`;
        const next = arr[i + 1];
        if (!next || (next.type !== "paragraph" && next.type !== "heading")) {
          out += `<w:p>${this.pendingDrawings}</w:p>`;
          this.pendingDrawings = "";
        }
        return;
      }
      out += this.block(n, list);
    });
    return out;
  }

  /** Metin kutusu: wps şekli (Word 2010+) ve VML yedeği. */
  textBox(n: JSONContent): string {
    const a = (n.attrs ?? {}) as Record<string, unknown>;
    const w = Number(a.w) || 240;
    const h = Number(a.h) || 120;
    const x = Number(a.x) || 0;
    const y = Number(a.y) || 0;
    const emu = (px: number) => Math.round(px * EMU_PER_PX);
    const pt = (px: number) => +(px * 0.75).toFixed(2);
    const border = hex6(a.border as string | null | undefined);
    const fill = hex6(a.fill as string | null | undefined);
    const bw = Number(a.borderWidth) || 0.75;
    const id = this.ctx.drawingId++;
    // VML şekil kimliği belgede benzersiz (filigran ve opak öğelerin şekilleriyle çakışmaz).
    let spid = 1024 + id;
    while (this.ctx.vmlIds.has(`_x0000_s${spid}`)) spid++;
    this.ctx.vmlIds.add(`_x0000_s${spid}`);
    const inner = new PartWriter(this.ctx, this.rels, Math.round((w - 19) * 15)).blocks(n.content) || "<w:p/>";
    const ln = border ? `<a:ln w="${Math.round(bw * 12700)}"><a:solidFill><a:srgbClr val="${border}"/></a:solidFill></a:ln>` : "<a:ln><a:noFill/></a:ln>";
    const fl = fill ? `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>` : "<a:noFill/>";
    const choice =
      `<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${251659264 + id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
      `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>${emu(x)}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>${emu(y)}</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="${emu(w)}" cy="${emu(h)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${id}" name="Metin Kutusu ${id}"/><wp:cNvGraphicFramePr/>` +
      `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${fl}${ln}</wps:spPr>` +
      `<wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>`;
    const style = `position:absolute;margin-left:${pt(x)}pt;margin-top:${pt(y)}pt;width:${pt(w)}pt;height:${pt(h)}pt;z-index:${id};mso-position-horizontal-relative:text;mso-position-vertical-relative:text`;
    const vml =
      `<w:pict><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>` +
      `<v:shape id="Metin Kutusu ${id}" o:spid="_x0000_s${spid}" type="#_x0000_t202" style="${style}"${fill ? ` fillcolor="#${fill}"` : ' filled="f"'}${border ? ` strokecolor="#${border}" strokeweight="${bw}pt"` : ' stroked="f"'}><v:textbox><w:txbxContent>${inner}</w:txbxContent></v:textbox></v:shape></w:pict>`;
    return `<mc:AlternateContent><mc:Choice Requires="wps">${choice}</mc:Choice><mc:Fallback>${vml}</mc:Fallback></mc:AlternateContent>`;
  }

  block(n: JSONContent, list?: { numId: number; ilvl: number }): string {
    switch (n.type) {
      case "paragraph":
      case "heading":
        return this.paragraph(n, list);
      case "bulletList":
      case "orderedList":
        return this.list(n, list ? list.ilvl + 1 : 0);
      case "table":
        return this.table(n);
      case "horizontalRule":
        return '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>';
      case "pageBreak":
        return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
      case "columnBreak":
        return '<w:p><w:r><w:br w:type="column"/></w:r></w:p>';
      case "toc":
        return this.toc(n);
      case "bibliography":
        return bibliographyXml(n);
      case "indexBlock":
        return indexXml(n);
      case "opaqueBlock":
        return this.opaqueXml(n.attrs ?? {}) ?? `<w:p>${this.opaqueFallback(n.attrs ?? {}, "")}</w:p>`;
      default:
        return n.content ? this.blocks(n.content, list) : "";
    }
  }

  /**
   * Opak öğenin yazılacak XML'i: iyi biçimliyse ilişkileri yeni pakete bağlanır (parçalar kaynak paketten bu öğe
   * için kopyalanır) ve kimlikleri yeniden numaralanır. Kaynak paket ya da başvurulan parça yoksa null.
   */
  private opaqueXml(a: Record<string, unknown>): string | null {
    const xml = typeof a.xml === "string" ? a.xml : "";
    if (topElements(xml)?.length !== 1) return null;
    const ids = explicitRelIds(xml);
    let out = xml;
    if (ids.size) {
      const refs = (Array.isArray(a.refs) ? (a.refs as OpaqueRef[]) : []).filter((r) => r && ids.has(r.id));
      if (new Set(refs.map((r) => r.id)).size !== ids.size || !this.ctx.carry) return null;
      const src = typeof a.src === "string" && a.src ? a.src : this.ctx.settings.kaynak;
      const pkg = src ? getSource(src) : undefined;
      if (refs.some((r) => !r.external && !pkg?.has(r.target))) return null;
      const r = reattach(xml, refs, this.ctx.carry.instance(pkg), (rel) => (rel.external ? this.rels.add(rel.type, rel.target, true) : this.rels.add(rel.type, relativeTarget("word/document.xml", rel.target))));
      if (r.missing.length) return null;
      out = r.xml;
    }
    return renumberIds(out, this.ctx);
  }

  /** Bağlanamayan opak öğenin yedek görünümü: önizleme resmi, yoksa yer tutucu metin. */
  private opaqueFallback(a: Record<string, unknown>, rpr: string): string {
    if (typeof a.preview === "string" && a.preview) {
      const img = this.image({ type: "image", attrs: { src: a.preview, width: a.width ?? 200, height: a.height ?? 150 } }, rpr);
      if (img) return img;
    }
    return `<w:r>${rpr}<w:t xml:space="preserve">${esc(cleanText(opaquePlaceholder(a)))}</w:t></w:r>`;
  }

  list(n: JSONContent, ilvl: number): string {
    const sig: ListSig = {
      type: n.type === "bulletList" ? "bullet" : "ordered",
      style: (n.attrs?.listStyle as string | null) ?? null,
      multilevel: (n.attrs?.multilevel as string | null) ?? null,
    };
    const numId = this.ctx.numFor(sig, (n.attrs?.start as number | undefined) ?? 1);
    let out = "";
    for (const li of n.content ?? []) {
      (li.content ?? []).forEach((child, i) => {
        if (i === 0 && (child.type === "paragraph" || child.type === "heading")) out += this.paragraph(child, { numId, ilvl });
        else if (child.type === "paragraph" || child.type === "heading") out += this.paragraph(child, undefined, 720 * (ilvl + 1));
        else out += this.block(child, { numId, ilvl });
      });
    }
    return out;
  }

  pPr(n: JSONContent, list?: { numId: number; ilvl: number }, contIndent?: number): string {
    const a = (n.attrs ?? {}) as Record<string, unknown>;
    // Öğeler şema sırasında; modelin okumadığı özellikler (ooxmlExtra) sıralarına katılır.
    const parts: [string, string][] = [];
    const put = (name: string, xml: string) => parts.push([name, xml]);
    let style: string | null = null;
    if (n.type === "heading") style = `Heading${a.level ?? 1}`;
    else if (a.styleId) style = a.styleId as string;
    else if (list || contIndent) style = "ListParagraph";
    if (style) put("pStyle", `<w:pStyle w:val="${esc(style)}"/>`);
    if (a.keepNext) put("keepNext", "<w:keepNext/>");
    if (a.keepLines) put("keepLines", "<w:keepLines/>");
    if (a.pageBreakBefore) put("pageBreakBefore", "<w:pageBreakBefore/>");
    if (list) put("numPr", `<w:numPr><w:ilvl w:val="${list.ilvl}"/><w:numId w:val="${list.numId}"/></w:numPr>`);
    // Şema sırası: numPr → suppressLineNumbers → pBdr.
    if (a.suppressLineNumbers) put("suppressLineNumbers", "<w:suppressLineNumbers/>");
    const b = a.borders as ParaBorders | null;
    if (b && (b.top || b.left || b.bottom || b.right)) {
      put("pBdr", `<w:pBdr>${b.top ? borderXml("top", b.top) : ""}${b.left ? borderXml("left", b.left as BorderSpec) : ""}${b.bottom ? borderXml("bottom", b.bottom) : ""}${b.right ? borderXml("right", b.right as BorderSpec) : ""}</w:pBdr>`);
    }
    if (a.shading) put("shd", `<w:shd w:val="clear" w:color="auto" w:fill="${hex6(a.shading as string) ?? "auto"}"/>`);
    const tabs = a.tabs as { pos: number; align: string; leader: string }[] | null;
    if (tabs?.length) put("tabs", `<w:tabs>${tabs.map((t) => `<w:tab w:val="${t.align}"${t.leader !== "none" ? ` w:leader="${t.leader}"` : ""} w:pos="${ptTw(t.pos)}"/>`).join("")}</w:tabs>`);
    const sp: string[] = [];
    if (a.spaceBefore != null) sp.push(`w:before="${ptTw(a.spaceBefore as number)}"`);
    if (a.spaceAfter != null) sp.push(`w:after="${ptTw(a.spaceAfter as number)}"`);
    if (a.lineHeight != null) {
      const rule = (a.lineRule as string | null) ?? "auto";
      if (rule === "auto") sp.push(`w:line="${Math.round((a.lineHeight as number) * 240)}" w:lineRule="auto"`);
      else sp.push(`w:line="${ptTw(a.lineHeight as number)}" w:lineRule="${rule === "exact" ? "exact" : "atLeast"}"`);
    }
    if (sp.length) put("spacing", `<w:spacing ${sp.join(" ")}/>`);
    const ind: string[] = [];
    if (contIndent) ind.push(`w:left="${contIndent}"`);
    else if (a.indentLeft != null && !list) ind.push(`w:left="${ptTw(a.indentLeft as number)}"`);
    if (a.indentRight != null) ind.push(`w:right="${ptTw(a.indentRight as number)}"`);
    if (a.indentFirst != null && !list) {
      const f = a.indentFirst as number;
      ind.push(f >= 0 ? `w:firstLine="${ptTw(f)}"` : `w:hanging="${ptTw(-f)}"`);
    }
    if (ind.length) put("ind", `<w:ind ${ind.join(" ")}/>`);
    const al = a.textAlign as string | null;
    if (al) put("jc", `<w:jc w:val="${al === "justify" ? "both" : al === "left" ? "left" : al}"/>`);
    if (a.outlineLevel && n.type !== "heading") put("outlineLvl", `<w:outlineLvl w:val="${(a.outlineLevel as number) - 1}"/>`);
    if (n === this.sectTarget && this.sectXml) put("sectPr", this.sectXml);
    const x = mergeOrdered(parts, a.ooxmlExtra, PPR_RANK, PPR_OTHER);
    return x ? `<w:pPr${ignorableOf(a.ooxmlExtra)}>${x}</w:pPr>` : "";
  }

  paragraph(n: JSONContent, list?: { numId: number; ilvl: number }, contIndent?: number): string {
    // Büyük ilk harf: harf, Word'deki gibi çerçeveli ayrı bir paragrafa (w:framePr w:dropCap) yazılır.
    const dc = list ? null : normalizeDropCap(n.attrs?.dropCap);
    const first = n.content?.[0];
    if (dc && first?.type === "text" && first.text) {
      const letter = firstGrapheme(first.text);
      if (letter.trim()) {
        const restText = first.text.slice(letter.length);
        const rest: JSONContent = { ...n, attrs: { ...n.attrs, dropCap: null }, content: [...(restText ? [{ ...first, text: restText }] : []), ...(n.content ?? []).slice(1)] };
        if (this.sectTarget === n) this.sectTarget = rest;
        return this.dropCapFrame(n, dc, { ...first, text: letter.trim() }) + this.paragraph(rest, list, contIndent);
      }
    }
    const drawings = this.pendingDrawings;
    this.pendingDrawings = "";
    return `<w:p>${this.pPr(n, list, contIndent)}${drawings}${this.inline(n.content ?? [])}</w:p>`;
  }

  /** Büyük ilk harf çerçevesi: harf boyutu, bırakılan satırların yüksekliğini kaplayacak biçimde hesaplanır. */
  private dropCapFrame(n: JSONContent, dc: DropCap, letter: JSONContent): string {
    const a = (n.attrs ?? {}) as Record<string, unknown>;
    const ts = letter.marks?.find((m) => m.type === "textStyle")?.attrs ?? {};
    const size = parseFloat(String(ts.fontSize ?? "")) || (n.type === "heading" ? 16 : 11);
    const factor = (!a.lineRule || a.lineRule === "auto") && typeof a.lineHeight === "number" ? a.lineHeight : 1.08;
    const line = a.lineRule === "exact" && typeof a.lineHeight === "number" ? a.lineHeight : size * LINE_FACTOR * factor;
    // Büyük harf yüksekliği ≈ 0,7 em: ilk satırın büyük harf üstünden son bırakılan satırın taban çizgisine.
    const capSize = Math.max(size, Math.round((((dc.lines - 1) * line + 0.7 * size) / 0.7) * 2) / 2);
    const marks = [...(letter.marks ?? []).filter((m) => m.type !== "textStyle"), { type: "textStyle", attrs: { ...ts, fontSize: `${capSize}pt`, fontFamily: dc.font ?? ts.fontFamily ?? null } }];
    let style = "";
    if (n.type === "heading") style = `<w:pStyle w:val="Heading${a.level ?? 1}"/>`;
    else if (a.styleId) style = `<w:pStyle w:val="${esc(String(a.styleId))}"/>`;
    const frame = `<w:framePr w:dropCap="${dc.mode}" w:lines="${dc.lines}" w:hSpace="${ptTw(dc.distance)}" w:wrap="around" w:vAnchor="text" w:hAnchor="${dc.mode === "margin" ? "page" : "text"}"/>`;
    const spacing = `<w:spacing w:before="0" w:after="0" w:line="${ptTw(dc.lines * line)}" w:lineRule="exact"/>`;
    return `<w:p><w:pPr>${style}<w:keepNext/>${frame}${spacing}<w:textAlignment w:val="baseline"/></w:pPr>${this.inline([{ ...letter, marks }])}</w:p>`;
  }

  /** Satır içi içerik: köprü, izlenen değişiklik ve açıklama sarmalayıcılarıyla koşular. */
  inline(nodes: JSONContent[]): string {
    let out = "";
    let link: string | null = null;
    let rev: "ins" | "del" | null = null;
    const closeRev = () => {
      if (rev) out += rev === "ins" ? "</w:ins>" : "</w:del>";
      rev = null;
    };
    const closeLink = () => {
      closeRev();
      if (link !== null) out += "</w:hyperlink>";
      link = null;
    };
    // Resim yazısı yer işareti ("_Ref…") etiketi ve numarayı kapsar (REF alanı "Şekil 1" verir).
    let refEnd: { id: number; at: number } | null = null;
    for (const [ni, node] of nodes.entries()) {
      if (refEnd && ni > refEnd.at) {
        out += `<w:bookmarkEnd w:id="${refEnd.id}"/>`;
        refEnd = null;
      }
      const marks = (node.marks ?? []) as Mark[];
      // Açıklama aralıkları
      const cids = new Set(marks.filter((m) => m.type === "comment").map((m) => String(m.attrs?.id ?? "")));
      for (const id of [...this.ctx.openComments]) {
        if (!cids.has(id)) {
          closeLink();
          const num = this.ctx.commentNum(id);
          out += `<w:commentRangeEnd w:id="${num}"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${num}"/></w:r>`;
          this.ctx.openComments.delete(id);
        }
      }
      for (const id of cids) {
        if (!this.ctx.openComments.has(id)) {
          closeLink();
          out += `<w:commentRangeStart w:id="${this.ctx.commentNum(id)}"/>`;
          this.ctx.openComments.add(id);
          this.ctx.usedComments.add(id);
        }
      }
      if (node.type === "bookmark") {
        const id = this.ctx.bookmarkId++;
        const name = String(node.attrs?.name ?? "");
        const seq: number = name.startsWith("_Ref") && !refEnd ? nodes.findIndex((x: JSONContent, k: number) => k > ni && x.type === "seqField") : -1;
        out += `<w:bookmarkStart w:id="${id}" w:name="${esc(name)}"/>`;
        if (seq > 0) refEnd = { id, at: seq };
        else out += `<w:bookmarkEnd w:id="${id}"/>`;
        continue;
      }
      const lm = marks.find((m) => m.type === "link");
      const href = lm ? String(lm.attrs?.href ?? "") : null;
      if (href !== link) {
        closeLink();
        if (href !== null) {
          if (href.startsWith("#")) out += `<w:hyperlink w:anchor="${esc(href.slice(1))}" w:history="1">`;
          else out += `<w:hyperlink r:id="${this.rels.add(REL.hyperlink, href, true)}" w:history="1">`;
          link = href;
        }
      }
      const r: "ins" | "del" | null = marks.some((m) => m.type === "insertion") ? "ins" : marks.some((m) => m.type === "deletion") ? "del" : null;
      if (r !== rev) {
        closeRev();
        if (r) {
          const m = marks.find((x) => x.type === (r === "ins" ? "insertion" : "deletion"))!;
          out += `<w:${r} w:id="${this.ctx.revId++}" w:author="${esc(String(m.attrs?.author || "Kullanıcı"))}" w:date="${esc(String(m.attrs?.date || new Date().toISOString().replace(/\.\d+Z$/, "Z")))}">`;
          rev = r;
        }
      }
      out += this.run(node, marks, link !== null, rev === "del");
    }
    if (refEnd) out += `<w:bookmarkEnd w:id="${refEnd.id}"/>`;
    closeLink();
    return out;
  }

  rPr(marks: Mark[], isLink: boolean): string {
    // Öğeler şema sırasında; modelin okumadığı özellikler (textStyle ooxmlExtra) sıralarına katılır.
    const parts: [string, string][] = [];
    const put = (name: string, xml: string) => parts.push([name, xml]);
    const has = (t: string) => marks.find((m) => m.type === t);
    const ts = (has("textStyle")?.attrs ?? {}) as Record<string, unknown>;
    const cs = has("charStyle");
    if (cs) put("rStyle", `<w:rStyle w:val="${esc(String(cs.attrs?.styleId))}"/>`);
    else if (isLink) put("rStyle", '<w:rStyle w:val="Hyperlink"/>');
    if (ts.fontFamily) put("rFonts", fontsXml(String(ts.fontFamily).replace(/["']/g, "").split(",")[0].trim()));
    if (has("bold")) put("b", "<w:b/><w:bCs/>");
    if (has("italic")) put("i", "<w:i/><w:iCs/>");
    if (ts.caps) put("caps", "<w:caps/>");
    if (ts.smallCaps) put("smallCaps", "<w:smallCaps/>");
    if (has("strike")) put("strike", "<w:strike/>");
    if (ts.doubleStrike) put("dstrike", "<w:dstrike/>");
    if (ts.outline) put("outline", "<w:outline/>");
    if (ts.shadow) put("shadow", "<w:shadow/>");
    if (ts.emboss) put("emboss", "<w:emboss/>");
    if (ts.imprint) put("imprint", "<w:imprint/>");
    if (ts.hidden) put("vanish", "<w:vanish/>");
    const color = hex6(ts.color as string | undefined);
    if (color) put("color", `<w:color w:val="${color}"/>`);
    if (ts.letterSpacing) put("spacing", `<w:spacing w:val="${ptTw(Number(ts.letterSpacing))}"/>`);
    if (ts.fontSize) {
      const pt = parseFloat(String(ts.fontSize));
      if (pt > 0) put("sz", `<w:sz w:val="${Math.round(pt * 2)}"/><w:szCs w:val="${Math.round(pt * 2)}"/>`);
    }
    const hl = has("highlight");
    let shdFill: string | null = hex6(ts.backgroundColor as string | undefined);
    if (hl) {
      const c = hex6((hl.attrs?.color as string | undefined) ?? "#FFFF00") ?? "FFFF00";
      const name = highlightName(`#${c}`);
      if (name) put("highlight", `<w:highlight w:val="${name}"/>`);
      else shdFill = c;
    }
    const u = has("underline");
    if (u) {
      const t = String(u.attrs?.type ?? "single");
      const val = t === "dashed" ? "dash" : t === "wave" ? "wave" : t;
      const uc = hex6(u.attrs?.color as string | undefined);
      put("u", `<w:u w:val="${val}"${uc ? ` w:color="${uc}"` : ""}/>`);
    }
    if (shdFill) put("shd", `<w:shd w:val="clear" w:color="auto" w:fill="${shdFill}"/>`);
    if (has("superscript")) put("vertAlign", '<w:vertAlign w:val="superscript"/>');
    else if (has("subscript")) put("vertAlign", '<w:vertAlign w:val="subscript"/>');
    // İzlenen biçim değişikliği: önceki biçim w:rPrChange içinde (rPr'nin son öğesi).
    const fc = has("formatChange");
    if (fc) {
      let old: Mark[] = [];
      try {
        old = JSON.parse(String(fc.attrs?.old ?? "[]")) as Mark[];
      } catch {
        old = [];
      }
      const inner = this.rPr(old, isLink) || "<w:rPr/>";
      put("rPrChange", `<w:rPrChange w:id="${this.ctx.revId++}" w:author="${esc(String(fc.attrs?.author || "Kullanıcı"))}" w:date="${esc(String(fc.attrs?.date || new Date().toISOString().replace(/\.\d+Z$/, "Z")))}">${inner}</w:rPrChange>`);
    }
    const x = mergeOrdered(parts, ts.ooxmlExtra, RPR_RANK, RPR_OTHER);
    return x ? `<w:rPr${ignorableOf(ts.ooxmlExtra)}>${x}</w:rPr>` : "";
  }

  run(node: JSONContent, marks: Mark[], isLink: boolean, deleted: boolean): string {
    const rpr = this.rPr(marks, isLink);
    switch (node.type) {
      case "text": {
        const text = cleanText(node.text ?? "");
        let body = "";
        const tTag = deleted ? "w:delText" : "w:t";
        for (const [i, seg] of text.split("\t").entries()) {
          if (i > 0) body += "<w:tab/>";
          if (seg) body += `<${tTag} xml:space="preserve">${esc(seg)}</${tTag}>`;
        }
        return `<w:r>${rpr}${body}</w:r>`;
      }
      case "hardBreak":
        return `<w:r>${rpr}<w:br/></w:r>`;
      case "pageField": {
        const kind = node.attrs?.field;
        const pic = fieldPicture(kind, node.attrs?.format);
        const f = kind === "NUMPAGES" ? "NUMPAGES" : (kind === "DATE" || kind === "TIME") && pic ? `${kind} \\@ "${pic}"` : "PAGE";
        return `<w:fldSimple w:instr=" ${esc(f)} \\* MERGEFORMAT "><w:r>${rpr}<w:t xml:space="preserve">${esc(fieldNowText(kind, new Date(), node.attrs?.format) ?? "1")}</w:t></w:r></w:fldSimple>`;
      }
      case "image":
        return this.image(node, rpr);
      case "chart":
        return this.chart(node, rpr);
      case "indexEntry":
        return xeXml(node);
      case "equation": {
        // OMML: satır içi m:oMath, kendi paragrafında görüntülü m:oMathPara.
        try {
          return ommlOf(parseMath(String(node.attrs?.latex ?? "")), !!node.attrs?.display);
        } catch {
          return `<w:r>${rpr}<w:t xml:space="preserve">${esc(String(node.attrs?.latex ?? ""))}</w:t></w:r>`;
        }
      }
      case "seqField":
      case "crossRef":
      case "citation":
      case "mergeField":
        return fieldXml(node, rpr);
      case "opaqueInline": {
        const a = node.attrs ?? {};
        const xml = this.opaqueXml(a);
        if (xml === null) return this.opaqueFallback(a, rpr);
        return a.level === "para" ? xml : `<w:r>${rpr}${xml}</w:r>`;
      }
      case "note": {
        const end = node.attrs?.kind === "endnote";
        const list = end ? this.ctx.endnotes : this.ctx.footnotes;
        list.push({ text: String(node.attrs?.text ?? ""), content: (node.attrs?.content as JSONContent[] | null) ?? null });
        const id = list.length;
        return end
          ? `<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="${id}"/></w:r>`
          : `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
      }
      default:
        return "";
    }
  }

  chart(node: JSONContent, rpr: string): string {
    return chartDrawing(this.ctx, this.rels, node, rpr);
  }

  image(node: JSONContent, rpr: string): string {
    const a = (node.attrs ?? {}) as Record<string, unknown>;
    const src = String(a.src ?? "");
    const bin = dataUrlToBytes(src);
    if (!bin) return "";
    const name = `image${this.ctx.media.length + 1}.${bin.ext}`;
    this.ctx.media.push({ name, data: bin.data, ext: bin.ext });
    const rid = this.rels.add(REL.image, `media/${name}`);
    const w = Math.max(1, Math.round(Number(a.width) || 200));
    const h = Math.max(1, Math.round(Number(a.height) || 150));
    const cx = w * EMU_PER_PX;
    const cy = h * EMU_PER_PX;
    const id = this.ctx.drawingId++;
    const descr = a.alt ? ` descr="${esc(String(a.alt))}"` : "";
    const rot = Number(a.rotate) || 0;
    const xfrmAttr = `${rot ? ` rot="${rot * 60000}"` : ""}${a.flipH ? ' flipH="1"' : ""}${a.flipV ? ' flipV="1"' : ""}`;
    const graphic =
      `<wp:docPr id="${id}" name="Resim ${id}"${descr}/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="${name}"${descr}/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${rid}"/>${srcRectXml(a.crop as { l: number; t: number; r: number; b: number } | null)}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm${xfrmAttr}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`;
    const wrap = (a.wrap as string | undefined) ?? "inline";
    if (wrap !== "inline") return `<w:r>${rpr}<w:drawing>${anchorXml(a, wrap, cx, cy, id, this.contentWidthTw / 15)}${graphic}</wp:anchor></w:drawing></w:r>`;
    return `<w:r>${rpr}<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${graphic}</wp:inline></w:drawing></w:r>`;
  }

  table(n: JSONContent): string {
    const a = (n.attrs ?? {}) as Record<string, unknown>;
    const rows = n.content ?? [];
    // Izgara: sütun sayısı ve genişlikler (ilk satırdan).
    const colWidths: (number | null)[] = [];
    for (const cell of rows[0]?.content ?? []) {
      const span = (cell.attrs?.colspan as number) ?? 1;
      const cw = (cell.attrs?.colwidth as number[] | null) ?? null;
      for (let j = 0; j < span; j++) colWidths.push(cw?.[j] ?? null);
    }
    const ncols = colWidths.length || 1;
    const known = colWidths.filter((w): w is number => !!w);
    const autoW = Math.round(this.contentWidthTw / ncols);
    const gridTw = colWidths.map((w) => (w ? Math.round(w * 15) : known.length === ncols ? autoW : autoW));
    const style = tableStyleId(a.tableStyle as string | null);
    const fr = a.headerRow === false ? 0 : 1;
    const lr = a.totalRow ? 1 : 0;
    const fc = a.firstColumn ? 1 : 0;
    const nh = a.bandedRows === false ? 1 : 0;
    const lookVal = (fr * 0x20) | (lr * 0x40) | (fc * 0x80) | (nh * 0x200) | 0x400;
    const look = `<w:tblLook w:val="${lookVal.toString(16).toUpperCase().padStart(4, "0")}" w:firstRow="${fr}" w:lastRow="${lr}" w:firstColumn="${fc}" w:lastColumn="0" w:noHBand="${nh}" w:noVBand="1"/>`;
    const allFixed = known.length === ncols;
    const tblW = allFixed ? `<w:tblW w:w="${gridTw.reduce((s, x) => s + x, 0)}" w:type="dxa"/>` : '<w:tblW w:w="5000" w:type="pct"/>';
    const jc = a.align ? `<w:jc w:val="${a.align}"/>` : "";
    const indent = !a.align || a.align === "left" ? Math.round((Number(a.indent) || 0) * 15) : 0;
    const tblInd = indent ? `<w:tblInd w:w="${indent}" w:type="dxa"/>` : "";
    const m = a.cellMargins as { top: number; bottom: number; left: number; right: number } | null;
    const tw = (pt: number) => Math.round(pt * 20);
    const cellMar = m ? `<w:tblCellMar><w:top w:w="${tw(m.top)}" w:type="dxa"/><w:left w:w="${tw(m.left)}" w:type="dxa"/><w:bottom w:w="${tw(m.bottom)}" w:type="dxa"/><w:right w:w="${tw(m.right)}" w:type="dxa"/></w:tblCellMar>` : "";
    let x = `<w:tbl><w:tblPr><w:tblStyle w:val="${style}"/>${tblW}${jc}${tblInd}<w:tblLayout w:type="fixed"/>${cellMar}${look}</w:tblPr>`;
    x += `<w:tblGrid>${gridTw.map((w) => `<w:gridCol w:w="${w}"/>`).join("")}</w:tblGrid>`;
    // rowspan → vMerge: her satırda dolu ızgara sütunlarını izle.
    const pending: { col: number; span: number; left: number; width: number }[] = [];
    rows.forEach((row, ri) => {
      x += ri === 0 && a.repeatHeader ? "<w:tr><w:trPr><w:tblHeader/></w:trPr>" : "<w:tr>";
      let col = 0;
      const cells = row.content ?? [];
      let ci = 0;
      const emitContinue = (p: { col: number; span: number; width: number }) => {
        x += `<w:tc><w:tcPr><w:tcW w:w="${p.width}" w:type="dxa"/>${p.span > 1 ? `<w:gridSpan w:val="${p.span}"/>` : ""}<w:vMerge/></w:tcPr><w:p/></w:tc>`;
      };
      while (ci < cells.length || pending.some((p) => p.col >= col && p.left > 0)) {
        const cont = pending.find((p) => p.col === col && p.left > 0);
        if (cont) {
          emitContinue(cont);
          cont.left--;
          col += cont.span;
          continue;
        }
        if (ci >= cells.length) break;
        const cell = cells[ci++];
        const ca = (cell.attrs ?? {}) as Record<string, unknown>;
        const span = (ca.colspan as number) ?? 1;
        const rowspan = (ca.rowspan as number) ?? 1;
        const width = gridTw.slice(col, col + span).reduce((s, w) => s + w, 0);
        let tcPr = `<w:tcW w:w="${width}" w:type="dxa"/>`;
        if (span > 1) tcPr += `<w:gridSpan w:val="${span}"/>`;
        if (rowspan > 1) {
          tcPr += '<w:vMerge w:val="restart"/>';
          pending.push({ col, span, left: rowspan - 1, width });
        }
        const bg = hex6(ca.background as string | undefined);
        if (bg) tcPr += `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>`;
        if (ca.verticalAlign) tcPr += `<w:vAlign w:val="${ca.verticalAlign === "middle" ? "center" : ca.verticalAlign}"/>`;
        let inner = this.blocks(cell.content);
        if (!inner.endsWith("</w:p>")) inner += "<w:p/>";
        x += `<w:tc><w:tcPr>${tcPr}</w:tcPr>${inner}</w:tc>`;
        col += span;
      }
      for (let i = pending.length - 1; i >= 0; i--) if (pending[i].left <= 0) pending.splice(i, 1);
      x += "</w:tr>";
    });
    return `${x}</w:tbl>`;
  }

  toc(n: JSONContent): string {
    const entries = (n.attrs?.entries as TocEntry[]) ?? [];
    const title = String(n.attrs?.title ?? "İçindekiler");
    // Şekiller Tablosu: resim yazısı etiketine göre TOC \c alanı (başlık paragrafı yok).
    const tof = n.attrs?.kind === "tof";
    const tab = `<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="${this.contentWidthTw}"/></w:tabs>`;
    let x = tof
      ? `<w:sdt><w:sdtPr><w:tag w:val="OfficeHereTableOfFigures"/></w:sdtPr><w:sdtContent>`
      : `<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>`;
    if (title && !tof) x += `<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr><w:r><w:t xml:space="preserve">${esc(title)}</w:t></w:r></w:p>`;
    const instr = tof ? ` TOC \\h \\z \\c "${esc(String(n.attrs?.label ?? "Şekil"))}" ` : ' TOC \\o "1-3" \\h \\z \\u ';
    const begin = `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">${instr}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>`;
    const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    if (!entries.length) {
      x += `<w:p>${begin}<w:r><w:t>İçindekiler girdisi bulunamadı.</w:t></w:r>${end}</w:p>`;
    } else {
      entries.forEach((e, i) => {
        const lvl = Math.min(3, Math.max(1, e.level));
        x += `<w:p><w:pPr><w:pStyle w:val="${tof ? "TableofFigures" : `TOC${lvl}`}"/>${tab}</w:pPr>${i === 0 ? begin : ""}<w:r><w:t xml:space="preserve">${esc(cleanText(e.text))}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>${e.page}</w:t></w:r>${i === entries.length - 1 ? end : ""}</w:p>`;
      });
    }
    return `${x}</w:sdtContent></w:sdt>`;
  }
}

/** Kırpma (a:srcRect; yüzde binde biri: 100000 = %100). */
function srcRectXml(c: { l: number; t: number; r: number; b: number } | null): string {
  if (!c || (!c.l && !c.t && !c.r && !c.b)) return "";
  const v = (x: number) => Math.round(Math.max(0, Math.min(0.99, x)) * 100000);
  return `<a:srcRect l="${v(c.l)}" t="${v(c.t)}" r="${v(c.r)}" b="${v(c.b)}"/>`;
}

/** Grafik: gömülü grafik parçası + satır içi çizim (c:chart başvurusu). */
function chartDrawing(ctx: Ctx, rels: Rels, node: JSONContent, rpr: string): string {
  const data = node.attrs?.data as ChartData | null;
  if (!data) return "";
  const n = ctx.charts.length + 1;
  const name = `chart${n}.xml`;
  ctx.charts.push({ name, xml: chartSpaceXml(data) });
  const rid = rels.add(REL_CHART, `charts/${name}`);
  const id = ctx.drawingId++;
  const cx = Math.round(data.width * EMU_PER_PX);
  const cy = Math.round(data.height * EMU_PER_PX);
  return (
    `<w:r>${rpr}<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="Grafik ${id}"${data.title ? ` descr="${esc(data.title)}"` : ""}/><wp:cNvGraphicFramePr/>` +
    `<a:graphic><a:graphicData uri="${CHART_NS}"><c:chart xmlns:c="${CHART_NS}" r:id="${rid}"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`
  );
}

/** Satır içi olmayan resmin wp:anchor başlangıcı (konum, metni kaydır, yığın sırası). */
function anchorXml(a: Record<string, unknown>, wrap: string, cx: number, cy: number, id: number, contentW: number): string {
  const free = wrap === "behind" || wrap === "front";
  const side = (a.floatSide as string | undefined) ?? "left";
  const z = 251658240 + Math.round(Number(a.z) || 0) * 1024 + id;
  const dist = 'distT="0" distB="0" distL="114300" distR="114300"';
  const px = Number(a.posX) || 0;
  // Kaydırmalı resim: kenara yaslı ve uzaklıksızsa hizalama; sürüklenmişse sütuna göre uzaklık.
  let offH: number | null = free ? px : null;
  if (!free && px > 0 && side !== "center") offH = side === "right" ? Math.max(0, contentW - cx / EMU_PER_PX - px) : px;
  const posH = offH !== null ? `<wp:positionH relativeFrom="column"><wp:posOffset>${Math.round(offH * EMU_PER_PX)}</wp:posOffset></wp:positionH>` : `<wp:positionH relativeFrom="column"><wp:align>${side === "right" ? "right" : side === "center" ? "center" : "left"}</wp:align></wp:positionH>`;
  const posV = `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${Math.round((Number(a.posY) || 0) * EMU_PER_PX)}</wp:posOffset></wp:positionV>`;
  let wrapXml = "<wp:wrapSquare wrapText=\"bothSides\"/>";
  if (free) wrapXml = "<wp:wrapNone/>";
  else if (wrap === "topBottom") wrapXml = "<wp:wrapTopAndBottom/>";
  else if (wrap === "tight") wrapXml = '<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>';
  return `<wp:anchor ${dist} simplePos="0" relativeHeight="${z}" behindDoc="${wrap === "behind" ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>${posH}${posV}<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${wrapXml}`;
}

/** Alan düğümleri: SEQ, REF/PAGEREF, CITATION, MERGEFIELD (sonuç metniyle birlikte). */
function fieldXml(node: JSONContent, rpr: string): string {
  const a = (node.attrs ?? {}) as Record<string, unknown>;
  let instr = "";
  let result = "";
  switch (node.type) {
    case "seqField":
      instr = ` SEQ ${String(a.label ?? "Şekil")} \\* ARABIC `;
      result = String(a.num ?? 1);
      break;
    case "crossRef":
      instr = ` ${a.kind === "page" ? "PAGEREF" : "REF"} ${String(a.ref ?? "")}${a.link !== false ? " \\h" : ""}${a.kind === "position" ? " \\p" : ""} `;
      result = String(a.text ?? "");
      break;
    case "citation":
      instr = ` CITATION ${String(a.tag ?? "")}${a.pages ? ` \\p ${String(a.pages)}` : ""} \\l 1055 `;
      result = String(a.text ?? "");
      break;
    case "mergeField":
      instr = ` MERGEFIELD ${String(a.name ?? "")} `;
      result = `«${String(a.name ?? "")}»`;
      break;
  }
  return `<w:fldSimple w:instr="${esc(instr)}"><w:r>${rpr}<w:t xml:space="preserve">${esc(cleanText(result))}</w:t></w:r></w:fldSimple>`;
}

/** Kaynakça: Word'deki gibi "Bibliographies" belge bölümü içinde BIBLIOGRAPHY alanı. */
function bibliographyXml(n: JSONContent): string {
  const entries = (n.attrs?.entries as { before: string; italic: string; after: string }[]) ?? [];
  const title = String(n.attrs?.title ?? "Kaynakça");
  let x = `<w:sdt><w:sdtPr><w:tag w:val="OfficeHereBibliography"/><w:docPartObj><w:docPartGallery w:val="Bibliographies"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>`;
  x += `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">${esc(title)}</w:t></w:r></w:p>`;
  const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> BIBLIOGRAPHY </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>';
  const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
  const run = (t: string, italic = false) => (t ? `<w:r>${italic ? "<w:rPr><w:i/></w:rPr>" : ""}<w:t xml:space="preserve">${esc(cleanText(t))}</w:t></w:r>` : "");
  if (!entries.length) x += `<w:p><w:pPr><w:pStyle w:val="Bibliography"/></w:pPr>${begin}${run("Belgede kaynak yok.")}${end}</w:p>`;
  entries.forEach((e, i) => {
    x += `<w:p><w:pPr><w:pStyle w:val="Bibliography"/><w:ind w:left="720" w:hanging="720"/></w:pPr>${i === 0 ? begin : ""}${run(e.before)}${run(e.italic, true)}${run(e.after)}${i === entries.length - 1 ? end : ""}</w:p>`;
  });
  return `${x}</w:sdtContent></w:sdt>`;
}
