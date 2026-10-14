/**
 * Geçiş katmanı, sayfa düzeyi: dosyadan okunan ama modelin tanımadığı XML parçacıkları (korumalı
 * aralıklar, OLE nesneleri, x14 veri doğrulama/koşullu biçim uzantıları, çizimdeki şekiller…).
 * Kaydederken özgün XML olarak geri konur (io/xlsxPost.ts). Başka parçalara ilişkileri `refs`'te
 * (kaynak paketteki mutlak adlarla) durur; bayt taşımaz.
 *
 * Yapısal değişiklikte (satır/sütun ekle-sil, sayfa yeniden adlandırma ya da silme) parçacıklardaki
 * hücre başvuruları (sqref/ref/r öznitelikleri, xm:sqref, formüller, çapa satır/sütunları) aynı anda
 * formül paketinin başvuru kaydırmasıyla dönüştürülür. Dönüştürülemeyen (aralığı tümüyle silinen,
 * formülünde #REF! oluşan ya da konumu kaydırılamayan) öğe düşürülür ve etiketi `dropped`a yazılır;
 * kaydederken kullanıcı uyarılır. `structVersion` parçacıkların hangi sayfa yapısına göre güncel
 * olduğunu gösterir: sayfanınkiyle aynı değilse konuma bağlı öğeler güvenle yazılamaz.
 */
import type { RawRef } from "@officehere/core/opc";
import { adjustFormula, colToName, parseRangeAddress, removeSheetFromFormula, renameSheetInFormula, type StructuralChange } from "@officehere/formula";
import { MAX_COLS, MAX_ROWS, shiftIndex, shiftRect } from "./rect";
import type { Rect } from "./types";

/** Taşınan XML parçacığı. */
export interface PtItem {
  xml: string;
  /** Parçacığın ilişkileri (r:id, r:embed, o:relid…). */
  refs?: RawRef[];
  /** Düşürülürse uyarıda gösterilecek ad. */
  label: string;
  /** Modelde karşılığı olan nesne (ör. gösterilen ama yazılamayan EMF resmi): silinirse parçacık da yazılmaz. */
  obj?: string;
}

/** Sayfanın tanınmayan üst düzey öğesi. */
export interface PtChild extends PtItem {
  /** Sıralama adı (mc:AlternateContent için içindeki öğenin adı). */
  name: string;
  /** Dosyada kendinden önce gelen bilinen öğe (şema sırasında yeri bilinmeyenler için). */
  after?: string;
}

/** Tanınmayan extLst girdisi (bütün ext öğesi). */
export interface PtExt extends PtItem {
  uri: string;
}

/** x14:conditionalFormatting girdisi; `rule` modelde karşılığı olan kural (kimlik), `sig` okunduğundaki imzası. */
export interface PtCf extends PtItem {
  rule?: string;
  sig?: string;
}

/** Ek parça (ör. tableSingleCells): sayfaya örtük ilişkiyle bağlı, hücre başvurusu taşıyan XML. */
export interface PtPart {
  type: string;
  contentType: string;
  xml: string;
  label: string;
}

/** Kök öğede bildirilen ve parçacıklarda kullanılan ad alanları. */
export interface PtNamespaces {
  ns: Record<string, string>;
  ignorable: string[];
}

export interface SheetPassthrough extends PtNamespaces {
  /** Kaynak paketteki sayfa parçası (paket düzeyi taşımada sayfa eşlemesi). */
  part?: string;
  /** Parçacıkların güncel olduğu sayfa yapısı sürümü (SheetModel.structVersion). */
  structVersion: number;
  /** sheetPr'nin model dışı öznitelikleri (codeName…). */
  sheetPr: [string, string][];
  children: PtChild[];
  exts: PtExt[];
  cf: PtCf[];
  /** x14:dataValidation girdileri. */
  dv: PtItem[];
  /** Çizimdeki tanınmayan çapalar (şekil, metin kutusu, desteklenmeyen resim…). */
  drawing: (PtNamespaces & { anchors: PtItem[] }) | null;
  /** VML'deki not dışı şekiller (OLE nesnesi ve form denetimi görünümleri). */
  vml: (PtNamespaces & { shapetypes: string[]; shapes: PtItem[] }) | null;
  parts: PtPart[];
  /** Kaynak tablo parçaları (tablo kimliği → parça) ve grafik parçaları (grafik kimliği → parça). */
  tables: Record<string, string>;
  charts: Record<string, string>;
  /** Yapısal değişiklikte düşürülen öğelerin etiketleri. */
  dropped: string[];
}

export function emptyPassthrough(structVersion = 0): SheetPassthrough {
  return { structVersion, sheetPr: [], ns: {}, ignorable: [], children: [], exts: [], cf: [], dv: [], drawing: null, vml: null, parts: [], tables: {}, charts: {}, dropped: [] };
}

/** Çalışma kitabı düzeyi (workbook.xml yaması). */
/**
 * Grafik sayfası (chartsheet): modelde sayfa olarak yok; <sheet> öğesi ve parçaları taşınır. Grafiklerin
 * veri başvuruları yapısal değişikliklerle güncel tutulur (yazarken kopyalanan grafik parçasına konur).
 */
export interface PtChartsheet extends PtItem {
  name: string;
  /** Kaynakta grafik sayfası → çizim zinciri (kopya adları bilinsin diye açıkça kopyalanır). */
  chain: string[];
  charts: { part: string; xml: string }[];
}

export interface WorkbookPassthrough extends PtNamespaces {
  fileVersion?: string;
  workbookPr: [string, string][];
  calcPr: [string, string][];
  children: PtChild[];
  exts: PtExt[];
  chartsheets: PtChartsheet[];
}

/* ------------------------------ Dönüşümler ------------------------------ */

const unesc = (s: string) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

/** sqref için açık A1 aralığı (tam sütun "A:A" değil "A1:A1048576"). */
export function sqrefText(r: Rect): string {
  const a = `${colToName(r.c1)}${r.r1 + 1}`;
  return r.r1 === r.r2 && r.c1 === r.c2 ? a : `${a}:${colToName(r.c2)}${r.r2 + 1}`;
}

/** Boşlukla ayrılmış aralık listesi. */
export function parseSqref(s: string): Rect[] {
  return s
    .split(/\s+/)
    .filter(Boolean)
    .map((p) => parseRangeAddress(p))
    .filter((r): r is Rect => !!r);
}

/** Hücre/aralık listesini kaydırır; tümü silinirse null. */
function shiftList(text: string, axis: "row" | "col", at: number, count: number): string | null {
  const out: string[] = [];
  let changed = false;
  for (const tok of text.trim().split(/\s+/)) {
    if (!tok) continue;
    const r = parseRangeAddress(tok);
    if (!r) {
      out.push(tok);
      continue;
    }
    const s = shiftRect(r, axis, at, count);
    if (!s) {
      changed = true;
      continue;
    }
    if (s.r1 === r.r1 && s.r2 === r.r2 && s.c1 === r.c1 && s.c2 === r.c2) out.push(tok);
    else {
      changed = true;
      out.push(tok.includes("$") ? sqrefText(s).replace(/([A-Z]+)(\d+)/g, "$$$1$$$2") : sqrefText(s));
    }
  }
  if (!out.length) return null;
  return changed ? out.join(" ") : text;
}

const REF_ATTRS = /(\s(?:sqref|ref|r|activeCell|topLeftCell)=")([^"]*)(")/g;
const SQREF_EL = /(<(?:[\w.-]+:)?sqref>)([^<]*)(<\/(?:[\w.-]+:)?sqref>)/g;
const FORMULA_EL = /(<((?:[\w.-]+:)?)(f|formula|formula1|formula2|FmlaLink|FmlaRange|FmlaMacro|FmlaGroup|FmlaTxbx)\b[^>/]*>)([^<]*)(<\/\2\3>)/g;
const FORMULA_ATTRS = /(\s(?:textlink|fmlaLink|fmlaRange|fmlaTxbx|fmlaGroup)=")([^"]*)(")/g;

/** Formül metinlerine (öğe içerikleri ve formül öznitelikleri) dönüşüm uygular; #REF! oluşursa null. */
function mapFormulas(xml: string, fn: (f: string) => string, allowRef = false): string | null {
  let bad = false;
  const apply = (raw: string) => {
    const f = unesc(raw);
    if (!f.trim()) return raw;
    const before = (f.match(/#REF!/g) ?? []).length;
    const g = fn(f);
    if (g === f) return raw;
    if (!allowRef && (g.match(/#REF!/g) ?? []).length > before) bad = true;
    return esc(g);
  };
  let out = xml.replace(FORMULA_EL, (_m, open: string, _p: string, _n: string, body: string, close: string) => open + apply(body) + close);
  out = out.replace(FORMULA_ATTRS, (_m, a: string, v: string, b: string) => a + apply(v).replace(/"/g, "&quot;") + b);
  return bad ? null : out;
}

/** DrawingML çapalarındaki (from/to) ve VML x:Anchor'daki satır/sütunları kaydırır. */
function shiftAnchors(xml: string, axis: "row" | "col", at: number, count: number): string {
  const key = axis === "row" ? "row" : "col";
  const max = axis === "row" ? MAX_ROWS - 1 : MAX_COLS - 1;
  const move = (i: number) => Math.min(max, Math.max(0, shiftIndex(i, at, count) ?? at));
  let out = xml.replace(/<((?:[\w.-]+:)?)(from|to)>([\s\S]*?)<\/\1\2>/g, (_m, p: string, tag: string, body: string) => {
    const nb = body.replace(new RegExp(`(<((?:[\\w.-]+:)?)${key}>)(\\d+)(</\\2${key}>)`), (_x, a: string, _q: string, n: string, b: string) => a + move(Number(n)) + b);
    return `<${p}${tag}>${nb}</${p}${tag}>`;
  });
  out = out.replace(/(<x:Anchor>)([^<]*)(<\/x:Anchor>)/g, (m, a: string, body: string, b: string) => {
    const v = body.split(",").map((s) => s.trim());
    if (v.length !== 8) return m;
    const idx = axis === "row" ? [2, 6] : [0, 4];
    for (const i of idx) v[i] = String(move(Number(v[i])));
    return `${a}\n    ${v.join(", ")}${b}`;
  });
  out = out.replace(new RegExp(`(<x:${axis === "row" ? "Row" : "Column"}>)(\\d+)(</x:${axis === "row" ? "Row" : "Column"}>)`, "g"), (_m, a: string, n: string, b: string) => a + move(Number(n)) + b);
  return out;
}

/** Bir yapısal değişiklik: aynı sayfada (`own`) ya da başka bir sayfada. */
interface Change extends StructuralChange {
  own: boolean;
  /** Parçacığın bulunduğu sayfanın adı (nitelenmemiş başvurular bu sayfayı gösterir). */
  formulaSheet: string;
}

/** Konum bilgisi kaydırılamayan öğeler: aynı sayfanın yapısı değişince düşürülür. */
const UNSHIFTABLE = new Set(["customSheetViews", "dataConsolidate", "controls"]);

function shiftXml(xml: string, ch: Change): string | null {
  let out: string | null = xml;
  if (ch.own) {
    let bad = false;
    out = out.replace(REF_ATTRS, (m, a: string, v: string, b: string) => {
      const s = shiftList(unesc(v), ch.axis, ch.at, ch.count);
      if (s === null) bad = true;
      return s === null || s === unesc(v) ? m : a + s + b;
    });
    out = out.replace(SQREF_EL, (m, a: string, v: string, b: string) => {
      const s = shiftList(v, ch.axis, ch.at, ch.count);
      if (s === null) bad = true;
      return s === null || s === v ? m : a + s + b;
    });
    if (bad) return null;
    out = shiftAnchors(out, ch.axis, ch.at, ch.count);
  }
  return mapFormulas(out, (f) => adjustFormula(f, ch.formulaSheet, ch));
}

/** Kapsayıcı öğenin (protectedRanges, oleObjects…) çocuklarını tek tek dönüştürür; boş kalırsa null. */
function mapContainer(xml: string, fn: (x: string) => string | null): string | null {
  const open = /^<[^>]*?(\/?)>/.exec(xml);
  if (!open) return fn(xml);
  if (open[1] === "/") return fn(xml);
  const name = /^<([\w.:-]+)/.exec(xml)![1];
  const close = `</${name}>`;
  const inner = xml.slice(open[0].length, xml.length - close.length);
  const kids = topElements(inner);
  if (!kids.length) return fn(xml);
  const out: string[] = [];
  let changed = false;
  for (const k of kids) {
    const r = fn(k);
    if (r !== k) changed = true;
    if (r !== null) out.push(r);
  }
  if (!out.length) return null;
  if (!changed) return xml;
  const head = open[0].replace(/(\scount=")\d+(")/, `$1${out.length}$2`);
  return head + out.join("") + close;
}

/** Bir XML parçasının üst düzey öğeleri (metin ve yorumlar atlanır). */
function topElements(xml: string): string[] {
  const out: string[] = [];
  let i = 0;
  let depth = 0;
  let start = -1;
  while (i < xml.length) {
    const lt = xml.indexOf("<", i);
    if (lt < 0) break;
    if (xml.startsWith("<!--", lt)) {
      const e = xml.indexOf("-->", lt);
      i = e < 0 ? xml.length : e + 3;
      continue;
    }
    if (xml.startsWith("<![CDATA[", lt)) {
      const e = xml.indexOf("]]>", lt);
      i = e < 0 ? xml.length : e + 3;
      continue;
    }
    let q = "";
    let j = lt + 1;
    for (; j < xml.length; j++) {
      const c = xml[j];
      if (q) {
        if (c === q) q = "";
      } else if (c === '"' || c === "'") q = c;
      else if (c === ">") break;
    }
    const end = j + 1;
    if (xml[lt + 1] === "?" || xml[lt + 1] === "!") {
      i = end;
      continue;
    }
    if (xml[lt + 1] === "/") {
      depth--;
      if (depth === 0 && start >= 0) {
        out.push(xml.slice(start, end));
        start = -1;
      }
    } else if (xml[end - 2] === "/") {
      if (depth === 0) out.push(xml.slice(lt, end));
    } else {
      if (depth === 0) start = lt;
      depth++;
    }
    i = end;
  }
  return out;
}

const CONTAINERS = new Set(["protectedRanges", "cellWatches", "ignoredErrors", "smartTags", "oleObjects", "webPublishItems", "AlternateContent"]);

type ItemFn = (xml: string, item: PtItem, kind: "child" | "ext" | "cf" | "dv" | "anchor" | "vml" | "part", name?: string) => string | null;

/** Bütün parçacıklara dönüşüm uygular; null dönenler düşürülür ve etiketleri kaydedilir. */
function mapAll(pt: SheetPassthrough, fn: ItemFn): void {
  const drop = (label: string) => {
    if (!pt.dropped.includes(label)) pt.dropped.push(label);
  };
  const run = <T extends PtItem>(list: T[], kind: Parameters<ItemFn>[2], nameOf?: (x: T) => string): T[] =>
    list.flatMap((it) => {
      const r = fn(it.xml, it, kind, nameOf?.(it));
      if (r === null) {
        drop(it.label);
        return [];
      }
      return r === it.xml ? [it] : [{ ...it, xml: r }];
    });
  pt.children = run(pt.children, "child", (c) => c.name);
  pt.exts = run(pt.exts, "ext");
  pt.cf = run(pt.cf, "cf");
  pt.dv = run(pt.dv, "dv");
  if (pt.drawing) pt.drawing.anchors = run(pt.drawing.anchors, "anchor");
  if (pt.vml) pt.vml.shapes = run(pt.vml.shapes, "vml");
  pt.parts = pt.parts.flatMap((p) => {
    const r = fn(p.xml, p, "part");
    if (r === null) {
      drop(p.label);
      return [];
    }
    return [{ ...p, xml: r }];
  });
}

/** Sayfanın kendi satır/sütun ekleme-silmesi. */
export function shiftOwnPassthrough(pt: SheetPassthrough, sheetName: string, axis: "row" | "col", at: number, count: number, version: number): void {
  const ch: Change = { own: true, sheet: sheetName, formulaSheet: sheetName, axis, at, count };
  mapAll(pt, (xml, _it, kind, name) => {
    if (kind === "child" && name && UNSHIFTABLE.has(name)) return null;
    if (kind === "part") return mapContainer(xml, (x) => shiftXml(x, ch));
    if (kind === "child" && name && (CONTAINERS.has(name) || /^<(?:[\w.-]+:)?AlternateContent\b/.test(xml))) return mapContainer(xml, (x) => shiftXml(x, ch));
    return shiftXml(xml, ch);
  });
  pt.structVersion = version;
}

/** Başka bir sayfanın satır/sütun ekleme-silmesi: yalnız o sayfaya nitelenmiş formül başvuruları. */
export function shiftForeignPassthrough(pt: SheetPassthrough, ownName: string, changed: string, axis: "row" | "col", at: number, count: number): void {
  const ch: Change = { own: false, sheet: changed, formulaSheet: ownName, axis, at, count };
  mapAll(pt, (xml) => mapFormulas(xml, (f) => adjustFormula(f, ownName, ch)));
}

/** Sayfa yeniden adlandırıldı. */
export function renamePassthrough(pt: SheetPassthrough, oldName: string, newName: string): void {
  mapAll(pt, (xml) => mapFormulas(xml, (f) => renameSheetInFormula(f, oldName, newName)) ?? xml);
}

/** Grafik sayfalarındaki grafiklerin veri başvuruları (hep nitelenmiş; #REF! olan seri grafikte kalır). */
function mapChartsheets(pt: WorkbookPassthrough | null, fn: (f: string) => string): void {
  for (const cs of pt?.chartsheets ?? []) for (const c of cs.charts) c.xml = mapFormulas(c.xml, fn, true) ?? c.xml;
}

/** Bir sayfanın satır/sütun ekleme-silmesi (grafik sayfalarının grafikleri). */
export function shiftWorkbookPassthrough(pt: WorkbookPassthrough | null, changed: string, axis: "row" | "col", at: number, count: number): void {
  mapChartsheets(pt, (f) => adjustFormula(f, "\u0000", { sheet: changed, axis, at, count }));
}

export function renameWorkbookPassthrough(pt: WorkbookPassthrough | null, oldName: string, newName: string): void {
  mapChartsheets(pt, (f) => renameSheetInFormula(f, oldName, newName));
}

export function removeSheetWorkbookPassthrough(pt: WorkbookPassthrough | null, name: string): void {
  mapChartsheets(pt, (f) => removeSheetFromFormula(f, name));
}

/** Sayfa silindi: ona başvuran formüllü öğeler düşürülür. */
export function removeSheetPassthrough(pt: SheetPassthrough, name: string): void {
  mapAll(pt, (xml) => mapFormulas(xml, (f) => removeSheetFromFormula(f, name)));
}
