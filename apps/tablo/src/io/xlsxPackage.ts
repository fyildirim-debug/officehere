/**
 * XLSX paketini okumadan önce sağlamlaştırır:
 * - Mutlak ilişki hedeflerini (Target="/xl/...", openpyxl vb. üretir) kaynak
 *   parçaya göre göreli yola çevirir (ECMA-376 Bölüm 2: ikisi de geçerlidir,
 *   ancak bazı okuyucular yalnız göreli yolu çözer).
 * - Okunamayan isteğe bağlı parçaları (açıklamalar, çizimler, tablolar…)
 *   ilişkilerden çıkarıp kalanının açılabilmesini sağlar.
 */
import JSZip from "jszip";
import { collectRefs, explicitRelIds, type OpcPackage, type RawRef } from "@officehere/core/opc";
import { parseQualifiedRange } from "@officehere/formula";
import type { CondRule } from "../model/types";
import type { SheetModel } from "../model/workbook";
import { emptyPassthrough, type PtChartsheet, type PtItem, type PtNamespaces, type SheetPassthrough, type WorkbookPassthrough } from "../model/passthrough";
import { unescapeXml } from "./ooxmlZip";
import { WORKBOOK_ORDER, WORKSHEET_ORDER, XmlPart, attrOf, childElements } from "./xmlPart";
import { loadZip } from "./zipCache";

/** İsteğe bağlı parça türleri: ilişki türü sonekine göre, kullanıcıya gösterilecek adla. */
export const OPTIONAL_PARTS: { kind: string; label: string; types: string[] }[] = [
  { kind: "comments", label: "açıklamalar (notlar)", types: ["/comments", "/vmlDrawing", "/threadedComment", "/person"] },
  { kind: "drawing", label: "resimler ve çizimler", types: ["/drawing"] },
  { kind: "table", label: "tablolar", types: ["/table"] },
  { kind: "pivot", label: "özet tablolar", types: ["/pivotTable", "/pivotCacheDefinition"] },
  { kind: "other", label: "diğer nesneler", types: ["/ctrlProp", "/oleObject", "/control", "/printerSettings", "/customProperty", "/slicer", "/timeline", "/webExtension"] },
];

function sourceOfRels(relsPath: string): string {
  // "xl/worksheets/_rels/sheet1.xml.rels" → "xl/worksheets/sheet1.xml"; "_rels/.rels" → ""
  const m = /^(.*?)_rels\/(.*)\.rels$/.exec(relsPath);
  if (!m) return "";
  return m[1] + m[2];
}

/** Hedefi kaynak parçanın klasörüne göre göreli yola çevirir. */
export function relativeTarget(source: string, absTarget: string): string {
  const target = absTarget.replace(/^\/+/, "");
  const from = source.split("/").slice(0, -1);
  const to = target.split("/");
  let i = 0;
  while (i < from.length && i < to.length - 1 && from[i] === to[i]) i++;
  const up = from.length - i;
  return [...Array<string>(up).fill(".."), ...to.slice(i)].join("/");
}

function relAttr(tag: string, name: string): string | undefined {
  return new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1];
}

/** .rels içindeki mutlak hedefleri göreli yapar. */
export function normalizeRelsXml(relsPath: string, xml: string): string {
  const source = sourceOfRels(relsPath);
  return xml.replace(/<Relationship\b[^>]*>/g, (tag) => {
    if (relAttr(tag, "TargetMode") === "External") return tag;
    const target = relAttr(tag, "Target");
    if (!target || !target.startsWith("/")) return tag;
    const rel = source ? relativeTarget(source, target) : target.replace(/^\/+/, "");
    return tag.replace(`Target="${target}"`, `Target="${rel}"`);
  });
}

/** Bilinen ad alanlarının alışılmış önekleri (önekli öğe adı bekleyen okuyucular için). */
const NS_PREFIX: Record<string, string> = {
  "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing": "xdr",
  "http://schemas.openxmlformats.org/drawingml/2006/main": "a",
  "http://schemas.openxmlformats.org/drawingml/2006/chart": "c",
  "http://schemas.openxmlformats.org/drawingml/2006/picture": "pic",
};

/**
 * Varsayılan ad alanıyla (öneksiz) yazılmış öğelere alışılmış öneki ekler:
 * <wsDr xmlns="…spreadsheetDrawing"><from> → <xdr:wsDr xmlns:xdr="…"><xdr:from>.
 */
export function prefixDefaultNamespaces(xml: string): string {
  const stack: (string | undefined)[] = [];
  let changed = false;
  const out = xml.replace(/<(\/?)([A-Za-z_][\w.-]*(?::[A-Za-z_][\w.-]*)?)((?:\s[^<>]*?)?)(\/?)>/g, (m, close: string, name: string, attrs: string, selfClose: string) => {
    if (close) {
      const prefix = stack.pop();
      if (name.includes(":") || !prefix) return m;
      changed = true;
      return `</${prefix}:${name}>`;
    }
    const declared = /\sxmlns="([^"]*)"/.exec(attrs)?.[1];
    const parentPrefix = stack.length ? stack[stack.length - 1] : undefined;
    let prefix: string | undefined;
    let newAttrs = attrs;
    if (declared !== undefined) {
      prefix = NS_PREFIX[declared];
      if (prefix) {
        // xmlns="…" → xmlns:önek="…" (önek zaten tanımlıysa yalnız kaldır)
        newAttrs = new RegExp(`\\sxmlns:${prefix}="`).test(attrs) ? attrs.replace(/\sxmlns="[^"]*"/, "") : attrs.replace(/\sxmlns="/, ` xmlns:${prefix}="`);
      }
    } else prefix = parentPrefix;
    const own = name.includes(":") ? undefined : prefix;
    if (!selfClose) stack.push(declared !== undefined ? prefix : parentPrefix);
    if (!own) return newAttrs === attrs ? m : `<${name}${newAttrs}${selfClose}>`;
    changed = true;
    return `<${own}:${name}${newAttrs}${selfClose}>`;
  });
  return changed ? out : xml;
}

const REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

/** Okuyucunun beklediği parça adları (tür → ad deseni ve yeni ad kalıbı). */
const PART_NAMES: { type: string; ok: RegExp; make: (n: number) => string }[] = [
  { type: "worksheet", ok: /^xl\/worksheets\/sheet\d+\.xml$/, make: (n) => `xl/worksheets/sheet${n}.xml` },
  { type: "comments", ok: /^xl\/comments\d+\.xml$/, make: (n) => `xl/comments${n}.xml` },
  { type: "drawing", ok: /^xl\/drawings\/[a-zA-Z0-9]+\.xml$/, make: (n) => `xl/drawings/drawing${n}.xml` },
  { type: "table", ok: /^xl\/tables\/table\d+\.xml$/, make: (n) => `xl/tables/table${n}.xml` },
];

async function renamePart(zip: JSZip, from: string, to: string): Promise<void> {
  const f = zip.file(from);
  if (!f) return;
  zip.file(to, await f.async("uint8array"));
  zip.remove(from);
  const rels = (p: string) => `${p.slice(0, p.lastIndexOf("/") + 1)}_rels/${p.slice(p.lastIndexOf("/") + 1)}.rels`;
  const oldRels = zip.file(rels(from));
  if (oldRels) {
    // Parçanın kendi ilişkileri: göreli hedefler yeni konuma göre yeniden yazılır
    const xml = await oldRels.async("string");
    const fixed = xml.replace(/<Relationship\b[^>]*>/g, (tag) => {
      const t = relAttr(tag, "Target");
      if (!t || relAttr(tag, "TargetMode") === "External") return tag;
      const abs = resolvePath(from, t);
      return tag.replace(`Target="${t}"`, `Target="${relativeTarget(to, abs)}"`);
    });
    zip.file(rels(to), fixed);
    zip.remove(rels(from));
  }
  // Bu parçayı gösteren tüm ilişkiler
  for (const name of Object.keys(zip.files)) {
    if (!name.endsWith(".rels")) continue;
    const src = sourceOfRels(name);
    const xml = await zip.file(name)!.async("string");
    const next = xml.replace(/<Relationship\b[^>]*>/g, (tag) => {
      const t = relAttr(tag, "Target");
      if (!t || relAttr(tag, "TargetMode") === "External" || resolvePath(src, t) !== from) return tag;
      return tag.replace(`Target="${t}"`, `Target="${src ? relativeTarget(src, to) : to}"`);
    });
    if (next !== xml) zip.file(name, next);
  }
  const ct = zip.file("[Content_Types].xml");
  if (ct) {
    const xml = await ct.async("string");
    zip.file("[Content_Types].xml", xml.split(`PartName="/${from}"`).join(`PartName="/${to}"`));
  }
}

/** Kaynak parçaya göre göreli ya da mutlak hedefi paket yoluna çevirir. */
export function resolvePath(source: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = source.split("/").slice(0, -1);
  for (const seg of target.split("/")) {
    if (seg === "..") parts.pop();
    else if (seg !== "." && seg !== "") parts.push(seg);
  }
  return parts.join("/");
}

/**
 * Paketi normalleştirir: mutlak hedefler göreli olur, parçalar alışılmış
 * adlarına taşınır, öneksiz DrawingML öğelerine önek eklenir ve VML not
 * çizimleri (yalnız görünüm bilgisi taşır) okuma için ayrılır.
 * Değişiklik yoksa özgün veriyi döndürür.
 */
export async function normalizePackage(data: ArrayBuffer): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(data);
  if (!(await normalizeZip(zip))) return data;
  return zip.generateAsync({ type: "arraybuffer" });
}

/** normalizePackage'ın açılmış paket üzerinde çalışan biçimi; değişiklik olduysa true. */
export async function normalizeZip(zip: JSZip): Promise<boolean> {
  let changed = false;
  for (const name of Object.keys(zip.files)) {
    if (!name.endsWith(".rels") || zip.files[name].dir) continue;
    const xml = await zip.file(name)!.async("string");
    const next = normalizeRelsXml(name, xml);
    if (next !== xml) {
      zip.file(name, next);
      changed = true;
    }
  }
  // Parça adları. Birden çok türde başvurulan (ör. döngüsel/bozuk ilişkiyle workbook.xml'i gösteren)
  // parçalar taşınmaz: yanlış bir ilişki ana parçayı sayfa klasörüne taşıyıp kitabı bozmasın.
  const typesOf = new Map<string, Set<string>>();
  for (const name of Object.keys(zip.files)) {
    if (!name.endsWith(".rels")) continue;
    const src = sourceOfRels(name);
    const xml = await zip.file(name)!.async("string");
    for (const m of xml.matchAll(/<Relationship\b[^>]*>/g)) {
      if (relAttr(m[0], "TargetMode") === "External") continue;
      const p = resolvePath(src, relAttr(m[0], "Target") ?? "");
      if (!typesOf.has(p)) typesOf.set(p, new Set());
      typesOf.get(p)!.add(relAttr(m[0], "Type") ?? "");
    }
  }
  const FIXED = /^(xl\/workbook\.xml|xl\/styles\.xml|xl\/sharedStrings\.xml|\[Content_Types\]\.xml)$|\.rels$/;
  for (const spec of PART_NAMES) {
    const targets: string[] = [];
    for (const name of Object.keys(zip.files)) {
      if (!name.endsWith(".rels")) continue;
      const src = sourceOfRels(name);
      const xml = await zip.file(name)!.async("string");
      for (const m of xml.matchAll(/<Relationship\b[^>]*>/g)) {
        if (relAttr(m[0], "Type") !== REL_TYPE + spec.type || relAttr(m[0], "TargetMode") === "External") continue;
        const p = resolvePath(src, relAttr(m[0], "Target") ?? "");
        if (!targets.includes(p)) targets.push(p);
      }
    }
    const bad = targets.filter((p) => !spec.ok.test(p) && zip.file(p) && !FIXED.test(p) && (typesOf.get(p)?.size ?? 0) <= 1);
    if (!bad.length) continue;
    let n = 1;
    for (const p of bad) {
      while (zip.file(spec.make(n))) n++;
      await renamePart(zip, p, spec.make(n));
      changed = true;
    }
  }
  // Öneksiz çizim/grafik XML'leri
  for (const name of Object.keys(zip.files)) {
    if (!/^xl\/(drawings|charts)\/[^/]+\.xml$/.test(name)) continue;
    const xml = await zip.file(name)!.async("string");
    const next = prefixDefaultNamespaces(xml);
    if (next !== xml) {
      zip.file(name, next);
      changed = true;
    }
  }
  // VML not çizimleri: yalnız Excel biçiminde (v:/x: önekli) olanlar kalır; diğerleri ilişkiden çıkarılır
  for (const name of Object.keys(zip.files)) {
    if (!/^xl\/worksheets\/_rels\/[^/]+\.rels$/.test(name)) continue;
    const src = sourceOfRels(name);
    const xml = await zip.file(name)!.async("string");
    const dropped: string[] = [];
    const next = await replaceAsync(xml, /<Relationship\b[^>]*>/g, async (tag) => {
      if (relAttr(tag, "Type") !== REL_TYPE + "vmlDrawing") return tag;
      const p = resolvePath(src, relAttr(tag, "Target") ?? "");
      const vml = zip.file(p) ? await zip.file(p)!.async("string") : "";
      if (/<v:shape\b/.test(vml) && /<x:ClientData\b/.test(vml) && /^xl\/drawings\/vmlDrawing\d+\.vml$/.test(p)) return tag;
      const id = relAttr(tag, "Id");
      if (id) dropped.push(id);
      return "";
    });
    if (!dropped.length) continue;
    zip.file(name, next);
    const sf = zip.file(src);
    if (sf) {
      let sx = await sf.async("string");
      for (const id of dropped) sx = sx.replace(new RegExp(`<(\\w+:)?legacyDrawing\\b[^>]*r:id="${id}"[^>]*/>`, "g"), "");
      zip.file(src, sx);
    }
    changed = true;
  }
  return changed;
}

async function replaceAsync(s: string, re: RegExp, fn: (m: string) => Promise<string>): Promise<string> {
  const parts: Promise<string>[] = [];
  let last = 0;
  for (const m of s.matchAll(re)) {
    parts.push(Promise.resolve(s.slice(last, m.index)));
    parts.push(fn(m[0]));
    last = m.index! + m[0].length;
  }
  parts.push(Promise.resolve(s.slice(last)));
  return (await Promise.all(parts)).join("");
}

/** Sayfa ilişkilerinden verilen türlerdeki parçaları çıkarır (ilgili XML öğeleriyle birlikte). */
export async function stripOptionalParts(data: ArrayBuffer, kinds: string[]): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(data);
  const types = OPTIONAL_PARTS.filter((p) => kinds.includes(p.kind)).flatMap((p) => p.types);
  const isDropped = (type: string | undefined) => !!type && types.some((t) => type.endsWith(t));
  for (const name of Object.keys(zip.files)) {
    if (!/^xl\/worksheets\/_rels\/[^/]+\.rels$/.test(name)) continue;
    const xml = await zip.file(name)!.async("string");
    const dropped = new Set<string>();
    const next = xml.replace(/<Relationship\b[^>]*>/g, (tag) => {
      if (!isDropped(relAttr(tag, "Type"))) return tag;
      const id = relAttr(tag, "Id");
      if (id) dropped.add(id);
      return "";
    });
    if (!dropped.size) continue;
    zip.file(name, next);
    const sheetPath = sourceOfRels(name);
    const sheetFile = zip.file(sheetPath);
    if (!sheetFile) continue;
    let sx = await sheetFile.async("string");
    // Kaldırılan ilişkilere başvuran öğeler (drawing, legacyDrawing, tableParts içindeki tablePart…)
    for (const id of dropped) {
      sx = sx.replace(new RegExp(`<(\\w+:)?(drawing|legacyDrawing|legacyDrawingHF|tablePart|control|oleObject)\\b[^>]*r:id="${id}"[^>]*/>`, "g"), "");
    }
    sx = sx.replace(/<tableParts\b[^>]*count="\d+"[^>]*>\s*<\/tableParts>/g, "").replace(/<tableParts\b[^>]*\/>/g, "");
    sx = sx.replace(/<tableParts\b([^>]*)count="\d+"([^>]*)>((?:\s*<tablePart\b[^>]*\/>)+)\s*<\/tableParts>/g, (_m, a: string, b: string, inner: string) => {
      const n = (inner.match(/<tablePart\b/g) ?? []).length;
      return `<tableParts${a}count="${n}"${b}>${inner}</tableParts>`;
    });
    zip.file(sheetPath, sx);
  }
  return zip.generateAsync({ type: "arraybuffer" });
}

/** Hücre notlarını (comments parçaları) doğrudan okur: düz <t> ya da biçimli <r><t> metinleri. */
export async function readCommentsFromData(data: ArrayBuffer): Promise<Map<string, { ref: string; text: string; author?: string }[]>> {
  const zip = await loadZip(data);
  const out = new Map<string, { ref: string; text: string; author?: string }[]>();
  const wbXml = await zip.file("xl/workbook.xml")?.async("string");
  const wbRels = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
  if (!wbXml || !wbRels) return out;
  const relTarget = new Map<string, string>();
  for (const m of wbRels.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = relAttr(m[0], "Id");
    const t = relAttr(m[0], "Target");
    if (id && t) relTarget.set(id, resolvePath("xl/workbook.xml", t));
  }
  const unesc = (s: string) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (_m, n: string) => String.fromCharCode(Number(n))).replace(/&amp;/g, "&");
  for (const m of wbXml.matchAll(/<(?:\w+:)?sheet\b[^>]*>/g)) {
    const name = relAttr(m[0], "name");
    const rid = /\br:id="([^"]+)"/.exec(m[0])?.[1];
    const sheetPath = rid ? relTarget.get(rid) : undefined;
    if (!name || !sheetPath) continue;
    const rp = `${sheetPath.slice(0, sheetPath.lastIndexOf("/") + 1)}_rels/${sheetPath.slice(sheetPath.lastIndexOf("/") + 1)}.rels`;
    const rels = await zip.file(rp)?.async("string");
    if (!rels) continue;
    for (const r of rels.matchAll(/<Relationship\b[^>]*>/g)) {
      if (!relAttr(r[0], "Type")?.endsWith("/comments")) continue;
      const cx = await zip.file(resolvePath(sheetPath, relAttr(r[0], "Target") ?? ""))?.async("string");
      if (!cx) continue;
      const authors = [...(/<(?:\w+:)?authors>([\s\S]*?)<\/(?:\w+:)?authors>/.exec(cx)?.[1] ?? "").matchAll(/<(?:\w+:)?author>([^<]*)<\/(?:\w+:)?author>/g)].map((a) => unesc(a[1]));
      const list: { ref: string; text: string; author?: string }[] = [];
      for (const c of cx.matchAll(/<(?:\w+:)?comment\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?comment>/g)) {
        const ref = relAttr(c[1], "ref");
        if (!ref) continue;
        const text = [...c[2].matchAll(/<(?:\w+:)?t(?:\s[^>]*)?>([^<]*)<\/(?:\w+:)?t>/g)].map((t) => unesc(t[1])).join("");
        const ai = Number(relAttr(c[1], "authorId") ?? -1);
        list.push({ ref, text, author: authors[ai] });
      }
      out.set(unesc(name), [...(out.get(unesc(name)) ?? []), ...list]);
    }
  }
  return out;
}

/** workbook.xml içindeki tanımlı adlar (sabit ve formül adları dahil). */
export async function readDefinedNamesFromData(data: ArrayBuffer): Promise<{ name: string; formula: string; localSheetId?: number }[]> {
  const zip = await loadZip(data);
  const xml = await zip.file("xl/workbook.xml")?.async("string");
  if (!xml) return [];
  const out: { name: string; formula: string; localSheetId?: number }[] = [];
  const unesc = (s: string) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
  for (const m of xml.matchAll(/<(?:\w+:)?definedName\b([^>]*)>([^<]*)<\/(?:\w+:)?definedName>/g)) {
    const name = relAttr(m[1], "name");
    if (!name || name.startsWith("_xlnm") || relAttr(m[1], "hidden") === "1") continue;
    const local = relAttr(m[1], "localSheetId");
    out.push({ name: unesc(name), formula: unesc(m[2]).trim(), localSheetId: local !== undefined ? Number(local) : undefined });
  }
  return out;
}

/* ============================== Geçiş katmanı: okuma ============================== */

/** Tablo adının dosyadaki biçimi: harf, rakam, alt çizgi, nokta ve ters bölü (Excel kuralları). */
export function tableFileName(name: string): string {
  const s = name.replace(/[^\p{L}\p{N}_.\\]/gu, "_");
  return /^[\p{L}_\\]/u.test(s) ? s : `_${s}`;
}

/**
 * Sayfa XML'inin hücre verisi çıkarılmış iskeleti (<sheetData/>): büyük sayfalarda üst düzey öğeler
 * hücreleri taramadan ayrıştırılabilsin.
 */
export function sheetSkeleton(xml: string): string {
  const m = /<((?:[\w.-]+:)?)sheetData\b[^>]*?(\/?)>/.exec(xml);
  if (!m || m[2] === "/") return xml;
  const close = `</${m[1]}sheetData>`;
  const end = xml.lastIndexOf(close);
  if (end < m.index) return xml;
  return `${xml.slice(0, m.index)}<${m[1]}sheetData/>${xml.slice(end + close.length)}`;
}

const NS_REL_URIS = ["http://schemas.openxmlformats.org/officeDocument/2006/relationships", "http://purl.oclc.org/ooxml/officeDocument/relationships"];
const NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/** Açılış etiketindeki ad alanı bildirimleri ve mc:Ignorable önekleri. */
export function rootNamespaces(open: string): PtNamespaces {
  const ns: Record<string, string> = {};
  for (const m of open.matchAll(/\sxmlns:([\w.-]+)="([^"]*)"/g)) ns[m[1]] = m[2];
  const mc = Object.entries(ns).find(([, u]) => u === NS_MC)?.[0] ?? "mc";
  const ign = new RegExp(`\\s${mc}:Ignorable="([^"]*)"`).exec(open)?.[1] ?? "";
  return { ns, ignorable: ign.split(/\s+/).filter(Boolean) };
}

const openTag = (xml: string) => /^<[^>]*>/.exec(xml)?.[0] ?? "";

/** Parçacıkta kullanılan önekler (öğe ve öznitelik adları). */
function usedPrefixes(xml: string): Set<string> {
  const out = new Set<string>();
  for (const m of xml.matchAll(/<\/?([A-Za-z_][\w.-]*):[A-Za-z_]/g)) out.add(m[1]);
  for (const m of xml.matchAll(/\s([A-Za-z_][\w.-]*):[A-Za-z_][\w.-]*\s*=\s*["']/g)) out.add(m[1]);
  out.delete("xmlns");
  out.delete("xml");
  return out;
}

/**
 * Parçacığı kendi başına geçerli yapar: kullandığı ama içinde bildirilmemiş öneklerin bildirimleri
 * (kapsamlardan, öndeki önce) ilk açılış etiketine eklenir.
 */
export function selfContained(xml: string, ...scopes: Record<string, string>[]): string {
  const declared = new Set([...xml.matchAll(/\sxmlns:([\w.-]+)=/g)].map((m) => m[1]));
  const add: string[] = [];
  for (const p of usedPrefixes(xml)) {
    if (declared.has(p)) continue;
    const uri = scopes.find((s) => p in s)?.[p];
    if (uri !== undefined) add.push(` xmlns:${p}="${uri}"`);
  }
  if (!add.length) return xml;
  return xml.replace(/^(<[A-Za-z_][\w.:-]*)/, `$1${add.join("")}`);
}

/** Parçacıkların kullandığı ve kökte mc:Ignorable olarak işaretli önekler (yazarken kökte yeniden bildirilir). */
function ignorableUsed(xmls: string[], root: PtNamespaces): PtNamespaces {
  const used = new Set<string>();
  for (const x of xmls) for (const p of usedPrefixes(x)) used.add(p);
  const ignorable = root.ignorable.filter((p) => used.has(p) && root.ns[p]);
  const ns: Record<string, string> = {};
  for (const p of ignorable) ns[p] = root.ns[p];
  return { ns, ignorable };
}

function relPrefixes(ns: Record<string, string>): string[] {
  return Object.entries(ns)
    .filter(([, u]) => NS_REL_URIS.includes(u))
    .map(([p]) => p);
}

/** Yazıcının ürettiği ya da modelin okuduğu sayfa öğeleri (geri kalanı taşınır). */
const OWNED_SHEET = new Set([
  "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetProtection", "scenarios", "autoFilter", "mergeCells",
  "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
  "drawing", "legacyDrawing", "tableParts", "extLst",
]);

const SHEET_LABELS: Record<string, string> = {
  sheetCalcPr: "sayfa hesaplama ayarları",
  protectedRanges: "korumalı aralıklar",
  sortState: "sıralama ayarları",
  dataConsolidate: "veri birleştirme ayarları",
  customSheetViews: "özel görünümler",
  phoneticPr: "fonetik ayarları",
  customProperties: "sayfa özel özellikleri",
  cellWatches: "izlenen hücreler",
  ignoredErrors: "yoksayılan hata işaretleri",
  smartTags: "akıllı etiketler",
  legacyDrawingHF: "üst/alt bilgi resimleri",
  drawingHF: "üst/alt bilgi resimleri",
  picture: "sayfa arka plan resmi",
  oleObjects: "gömülü nesneler (OLE)",
  controls: "form denetimleri",
  webPublishItems: "web yayımlama öğeleri",
};

/** Bilinen sayfa uzantıları (extLst ext uri, büyük harfle). */
export const EXT_URI = {
  cf: "{78C0D931-6437-407D-A8EE-F0AAD7539E65}",
  dv: "{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}",
  sparklines: "{05C60535-1F16-4FD2-B633-F4F36F0B64E0}",
};
const EXT_LABELS: Record<string, string> = {
  "{7E03D99A-DC04-49D9-9315-930204A7B6E9}": "zaman çizelgeleri",
  "{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}": "dilimleyiciler",
  "{3A4CF648-6AED-40F4-86FF-DC5316D8AED3}": "dilimleyiciler",
};
const LABEL_CF = "gelişmiş koşullu biçimlendirme kuralları";
const LABEL_DV = "genişletilmiş veri doğrulama kuralları";
const LABEL_ANCHOR = "çizim nesneleri (şekil, metin kutusu, resim)";
const LABEL_VML = "gömülü nesne ve form denetimi görünümleri";

/** Öğe düzeyinde kopyalanmaması gereken ilişki türleri: modelin parçaları (kopyası paketi bozar). */
const MODEL_PART_TYPES = /\/(worksheet|chartsheet|dialogsheet|pivotCacheDefinition|pivotCacheRecords|pivotTable|table|sharedStrings|styles|theme|calcChain)$/;

const upper = (s: string) => s.toUpperCase(); // yerel-ayar: ascii

/** mc:AlternateContent'in ilk seçeneğindeki ilk öğenin yerel adı (sıralama için). */
function alternateName(xml: string): string | undefined {
  const m = /<(?:[\w.-]+:)?Choice\b[^>]*>\s*<([\w.:-]+)/.exec(xml);
  return m ? m[1].slice(m[1].indexOf(":") + 1) : undefined;
}

/** Geçiş verisinin okunduğu kaynak: paket (ilişkiler) + metin (okuma önbelleğinden). */
export interface PtSource {
  pkg: OpcPackage;
  text(part: string): Promise<string | null>;
}

async function refsOf(src: PtSource, part: string, xml: string, ns: Record<string, string>): Promise<RawRef[] | undefined> {
  const refs = await collectRefs(src.pkg, part, xml, { prefixes: relPrefixes(ns) });
  return refs.length ? refs : undefined;
}

const touchesModel = (refs: RawRef[] | undefined) => !!refs?.some((r) => MODEL_PART_TYPES.test(r.type));

/** Öğenin (ör. <drawing r:id>) gösterdiği iç parça. */
async function relTarget(src: PtSource, part: string, xml: string, element: string, ns: Record<string, string>): Promise<string | undefined> {
  const tag = new RegExp(`<(?:[\\w.-]+:)?${element}\\b[^>]*>`).exec(xml)?.[0];
  if (!tag) return undefined;
  const id = [...explicitRelIds(tag, { prefixes: relPrefixes(ns) })][0];
  const rel = id ? (await src.pkg.rels(part)).find((r) => r.id === id && !r.external) : undefined;
  return rel?.part ? (src.pkg.canonicalName(rel.part) ?? rel.part) : undefined;
}

/** Grafiğin veri başvurusu okunabiliyor mu (xlsxCharts.readCharts ile aynı ölçüt). */
function chartReadable(cXml: string): boolean {
  for (const m of cXml.matchAll(/<c:f>([^<]+)<\/c:f>/g)) {
    const f = unescapeXml(m[1]).replace(/\$/g, "").replace(/^\(|\)$/g, "").split(",")[0];
    if (parseQualifiedRange(f)) return true;
  }
  return false;
}

/** Koşullu biçim kuralının (kimliği ve aralıkları dışında) imzası. */
export function condSig(rule: CondRule): string {
  const { id: _id, ranges: _r, ...rest } = rule;
  void _id;
  void _r;
  return JSON.stringify(rest);
}

/**
 * Sayfanın geçiş verisini kaynak paketten toplar: tanınmayan üst düzey öğeler, uzantılar, x14 koşullu
 * biçim ve veri doğrulama girdileri, çizimdeki tanınmayan çapalar, VML'deki not dışı şekiller ve
 * hücre başvurulu ek parçalar. `cfLinks`: exceljs'in x14 kural kimliği → modeldeki kural kimliği.
 */
export async function readSheetPassthrough(src: PtSource, sheet: SheetModel, path: string, cfLinks?: Map<string, string>): Promise<SheetPassthrough> {
  const pt = emptyPassthrough(sheet.structVersion);
  pt.part = path;
  const xml = await src.text(path);
  if (!xml) return pt;
  let part: XmlPart;
  try {
    part = new XmlPart(sheetSkeleton(xml), WORKSHEET_ORDER);
  } catch {
    return pt;
  }
  const root = rootNamespaces(part.open);
  const scope = root.ns;
  const isOrdered = (n: string) => (WORKSHEET_ORDER as readonly string[]).includes(n);
  let lastKnown: string | undefined;
  for (const c of part.children) {
    if (c.name === "sheetPr") {
      for (const m of openTag(c.xml).matchAll(/\s([\w.:-]+)="([^"]*)"/g)) if (!/^xmlns|^(filterMode|syncRef)$/.test(m[1])) pt.sheetPr.push([m[1], m[2]]);
    } else if (c.name === "extLst") {
      for (const e of childElements(c.xml)) {
        const uri = upper(attrOf(e.xml, "uri") ?? "");
        const extNs = { ...rootNamespaces(openTag(e.xml)).ns, ...scope };
        if (uri === EXT_URI.sparklines) continue;
        if (uri === EXT_URI.cf || uri === EXT_URI.dv) {
          for (const list of childElements(e.xml)) {
            const listNs = { ...rootNamespaces(openTag(list.xml)).ns, ...extNs };
            for (const item of childElements(list.xml)) {
              const x = selfContained(item.xml, listNs);
              const refs = await refsOf(src, path, x, scope);
              if (uri === EXT_URI.dv) {
                if (list.name === "dataValidations") pt.dv.push({ xml: x, refs, label: LABEL_DV });
                continue;
              }
              if (list.name !== "conditionalFormattings") continue;
              const rules = [...item.xml.matchAll(/<(?:[\w.-]+:)?cfRule\b[^>]*>/g)].map((m) => m[0]);
              // Veri çubukları ve hazır simge kümeleri modelde (yazıcı yeniden üretir)
              const custom = /<(?:[\w.-]+:)?iconSet\b[^>]*\scustom="(1|true)"/.test(item.xml);
              if (rules.length && rules.every((r) => attrOf(r, "type") === "dataBar" || (attrOf(r, "type") === "iconSet" && !custom))) continue;
              const id = rules.map((r) => attrOf(r, "id")).find(Boolean);
              const rule = id && cfLinks?.get(id) ? sheet.condFormats.find((r) => r.id === cfLinks.get(id)) : undefined;
              pt.cf.push({ xml: x, refs, label: LABEL_CF, ...(rule ? { rule: rule.id, sig: condSig(rule) } : {}) });
            }
          }
          continue;
        }
        const x = selfContained(e.xml, scope);
        const refs = await refsOf(src, path, x, scope);
        if (!touchesModel(refs)) pt.exts.push({ uri: attrOf(e.xml, "uri") ?? "", xml: x, refs, label: EXT_LABELS[uri] ?? "sayfa uzantıları" });
      }
    } else if (!OWNED_SHEET.has(c.name)) {
      const inner = c.name === "AlternateContent" ? alternateName(c.xml) : c.name;
      if (!inner || !OWNED_SHEET.has(inner)) {
        const x = selfContained(c.xml, scope);
        const refs = await refsOf(src, path, x, scope);
        const name = inner && isOrdered(inner) ? inner : c.name;
        if (!touchesModel(refs)) pt.children.push({ name, xml: x, refs, label: SHEET_LABELS[name] ?? "dosyadaki diğer öğeler", ...(isOrdered(name) ? {} : { after: lastKnown }) });
      }
    }
    if (isOrdered(c.name)) lastKnown = c.name;
  }
  Object.assign(pt, ignorableUsed([...pt.children, ...pt.exts, ...pt.cf, ...pt.dv].map((i) => i.xml), root));
  for (const r of await src.pkg.rels(path)) {
    if (r.external || !r.part) continue;
    // Sayfaya örtük ilişkiyle bağlı, hücre başvurusu taşıyan parça (XML eşlemeli tek hücreler)
    if (/\/tableSingleCells$/.test(r.type)) {
      const t = await src.text(r.part);
      if (t) pt.parts.push({ type: r.type, contentType: src.pkg.contentType(r.part) ?? CT_SINGLE_CELLS, xml: t.replace(/^\s*<\?xml[^>]*\?>\s*/, ""), label: "XML eşlemeli hücreler" });
    }
    // Tablolar: kaynak parçası (sorgu tablosu gibi örtük ilişkilerin taşınması ve tablo yaması için)
    if (/\/table$/.test(r.type)) {
      const t = await src.text(r.part);
      const display = t ? /<(?:\w+:)?table\b[^>]*?\sdisplayName="([^"]*)"/.exec(t)?.[1] : undefined;
      const def = display !== undefined ? sheet.tables.find((x) => x.name === unescapeXml(display)) : undefined;
      if (def) pt.tables[def.id] = src.pkg.canonicalName(r.part) ?? r.part;
    }
  }
  const skeleton = part.toString();
  await readDrawingPassthrough(src, sheet, path, skeleton, scope, pt);
  await readVmlPassthrough(src, path, skeleton, scope, pt);
  return pt;
}

const CT_SINGLE_CELLS = "application/vnd.openxmlformats-officedocument.spreadsheetml.tableSingleCells+xml";

/** Çizimdeki tanınmayan çapalar (modelin grafik ve resimleri hariç) ve grafik parçası eşlemesi. */
async function readDrawingPassthrough(src: PtSource, sheet: SheetModel, path: string, sheetXml: string, scope: Record<string, string>, pt: SheetPassthrough): Promise<void> {
  const dPath = await relTarget(src, path, sheetXml, "drawing", scope);
  const dXml = dPath ? await src.text(dPath) : null;
  if (!dPath || !dXml) return;
  let part: XmlPart;
  try {
    part = new XmlPart(dXml);
  } catch {
    return;
  }
  const dRoot = rootNamespaces(part.open);
  const rels = new Map((await src.pkg.rels(dPath)).map((r) => [r.id, r]));
  const target = (id: string | undefined) => {
    const r = id ? rels.get(id) : undefined;
    return r && !r.external && r.part ? (src.pkg.canonicalName(r.part) ?? r.part) : undefined;
  };
  const ANCHOR = /<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[\s\S]*?<\/xdr:\1>/g;
  // readCharts ile aynı tarama: veri başvurusu okunabilen grafikler modelde
  const modelChart = async (a: string) => {
    const p = target(/<c:chart\b[^>]*r:id="([^"]+)"/.exec(a)?.[1]);
    const c = p ? await src.text(p) : null;
    return c && chartReadable(c) ? p : undefined;
  };
  const chartPaths: string[] = [];
  for (const am of dXml.matchAll(ANCHOR)) {
    const p = await modelChart(am[0]);
    if (p) chartPaths.push(p);
  }
  if (chartPaths.length === sheet.charts.length) sheet.charts.forEach((c, i) => (pt.charts[c.id] = chartPaths[i]));
  const anchors: PtItem[] = [];
  let imageNo = 0;
  for (const c of part.children) {
    let chart = false;
    for (const am of c.xml.matchAll(ANCHOR)) if (await modelChart(am[0])) chart = true;
    if (chart) continue;
    // Resim: exceljs iki/tek hücreli çapadaki xdr:pic'i okur; png/jpeg/gif'i model yazar, gerisi taşınır
    let obj: string | undefined;
    if (/^<xdr:(twoCellAnchor|oneCellAnchor)\b/.test(c.xml) && /<xdr:pic\b/.test(c.xml)) {
      const media = [...c.xml.matchAll(/<a:blip\b[^>]*r:embed="([^"]+)"/g)]
        .map((m) => target(m[1]))
        .filter((p): p is string => !!p && /\/media\/.+\.[A-Za-z]{3,4}$/.test(p))
        .pop();
      if (media) {
        const img = sheet.images[imageNo++];
        if (/\.(png|jpe?g|gif)$/i.test(media)) continue;
        obj = img?.id;
      }
    }
    const x = selfContained(c.xml, dRoot.ns);
    const refs = await refsOf(src, dPath, x, dRoot.ns);
    if (!touchesModel(refs)) anchors.push({ xml: x, refs, label: LABEL_ANCHOR, ...(obj ? { obj } : {}) });
  }
  if (anchors.length) pt.drawing = { anchors, ...ignorableUsed(anchors.map((a) => a.xml), dRoot) };
}

/** VML'deki not dışı şekiller (OLE nesnelerinin ve form denetimlerinin görünümleri). */
async function readVmlPassthrough(src: PtSource, path: string, sheetXml: string, scope: Record<string, string>, pt: SheetPassthrough): Promise<void> {
  const vPath = await relTarget(src, path, sheetXml, "legacyDrawing", scope);
  const vXml = vPath ? await src.text(vPath) : null;
  if (!vPath || !vXml) return;
  const vRoot = rootNamespaces(/<xml\b[^>]*>/.exec(vXml)?.[0] ?? "");
  const shapes: PtItem[] = [];
  const types = new Set<string>();
  for (const m of vXml.matchAll(/<v:shape\b[\s\S]*?<\/v:shape>/g)) {
    if (/ObjectType="Note"/.test(m[0]) || !/<x:ClientData\b/.test(m[0])) continue;
    const x = selfContained(m[0], vRoot.ns);
    shapes.push({ xml: x, refs: await refsOf(src, vPath, x, vRoot.ns), label: LABEL_VML });
    const t = /\stype="#([^"]+)"/.exec(m[0])?.[1];
    if (t) types.add(t);
  }
  if (!shapes.length) return;
  const shapetypes = [...vXml.matchAll(/<v:shapetype\b[\s\S]*?<\/v:shapetype>/g)].map((m) => m[0]).filter((s) => types.has(/\sid="([^"]+)"/.exec(s)?.[1] ?? ""));
  pt.vml = { shapes, shapetypes: shapetypes.map((s) => selfContained(s, vRoot.ns)), ns: {}, ignorable: [] };
}

/** Grafik sayfaları: <sheet> öğesi, parça zinciri ve grafiklerinin XML'i. */
async function readChartsheets(src: PtSource, wbPath: string, sheetsXml: string, ns: Record<string, string>): Promise<PtChartsheet[]> {
  const out: PtChartsheet[] = [];
  const rels = new Map((await src.pkg.rels(wbPath)).map((r) => [r.id, r]));
  for (const s of childElements(sheetsXml)) {
    const id = [...explicitRelIds(s.xml, { prefixes: relPrefixes(ns) })][0];
    const rel = id ? rels.get(id) : undefined;
    if (!rel?.part || rel.external || !/\/chartsheet$/.test(rel.type)) continue;
    const part = src.pkg.canonicalName(rel.part) ?? rel.part;
    const chain = [part];
    const charts: { part: string; xml: string }[] = [];
    for (const d of await src.pkg.rels(part)) {
      if (d.external || !d.part || !/\/drawing$/.test(d.type)) continue;
      chain.push(src.pkg.canonicalName(d.part) ?? d.part);
      for (const c of await src.pkg.rels(d.part)) {
        if (c.external || !c.part || !/\/chart$/.test(c.type)) continue;
        const xml = await src.text(c.part);
        if (xml) charts.push({ part: src.pkg.canonicalName(c.part) ?? c.part, xml });
      }
    }
    const x = selfContained(s.xml, ns);
    out.push({ xml: x, refs: await refsOf(src, wbPath, x, ns), label: "grafik sayfaları", name: unescapeXml(attrOf(s.xml, "name") ?? ""), chain, charts });
  }
  return out;
}

// CT_WorkbookPr ve CT_CalcPr (ECMA-376 Transitional); date1904, calcId ve fullCalcOnLoad yazıcınındır
const WORKBOOK_PR_ATTRS = new Set([
  "showObjects", "showBorderUnselectedTables", "filterPrivacy", "promptedSolutions", "showInkAnnotation", "backupFile", "saveExternalLinkValues",
  "updateLinks", "codeName", "hidePivotFieldList", "showPivotChartFilter", "allowRefreshQuery", "publishItems", "checkCompatibility",
  "autoCompressPictures", "refreshAllConnections", "defaultThemeVersion",
]);
const CALC_PR_ATTRS = new Set([
  "calcMode", "refMode", "iterate", "iterateCount", "iterateDelta", "fullPrecision", "calcCompleted", "calcOnSave", "concurrentCalc",
  "concurrentManualCount", "forceFullCalc",
]);

/** workbook.xml'in modelin tanımadığı bölümleri (x15 workbookPr, revisionPtr, absPath, uzantılar…). */
export async function readWorkbookPassthrough(src: PtSource, wbPath: string): Promise<WorkbookPassthrough | null> {
  const xml = await src.text(wbPath);
  if (!xml) return null;
  let part: XmlPart;
  try {
    part = new XmlPart(xml, WORKBOOK_ORDER);
  } catch {
    return null;
  }
  const root = rootNamespaces(part.open);
  const pt: WorkbookPassthrough = { workbookPr: [], calcPr: [], children: [], exts: [], chartsheets: [], ns: {}, ignorable: [] };
  // Ön eksiz öznitelikler yalnız Transitional şemadakilerse taşınır (LibreOffice'in yazdığı Strict'e özgü
  // dateCompatibility gibileri şema hatası olur); ön ekli uzantı öznitelikleri olduğu gibi kalır.
  const attrs = (x: string, keep: Set<string>) =>
    [...openTag(x).matchAll(/\s([\w.:-]+)="([^"]*)"/g)].filter((m) => !/^xmlns/.test(m[1]) && (m[1].includes(":") || keep.has(m[1]))).map((m): [string, string] => [m[1], m[2]]);
  // Yazıcının ürettikleri: kitap koruması, görünümler, sayfalar, adlar, özet tablo önbellekleri
  const owned = new Set(["workbookProtection", "bookViews", "sheets", "definedNames", "pivotCaches"]);
  const isOrdered = (n: string) => (WORKBOOK_ORDER as readonly string[]).includes(n);
  let lastKnown: string | undefined;
  for (const c of part.children) {
    if (c.name === "fileVersion") pt.fileVersion = selfContained(c.xml, root.ns);
    // 1904 tarih sistemi yazılmaz: model değerleri 1900 sistemine çevrilmiş tutar
    else if (c.name === "workbookPr") pt.workbookPr = attrs(c.xml, WORKBOOK_PR_ATTRS);
    else if (c.name === "calcPr") pt.calcPr = attrs(c.xml, CALC_PR_ATTRS);
    else if (c.name === "sheets") pt.chartsheets = await readChartsheets(src, wbPath, c.xml, root.ns);
    else if (c.name === "extLst") {
      for (const e of childElements(c.xml)) {
        const x = selfContained(e.xml, root.ns);
        const refs = await refsOf(src, wbPath, x, root.ns);
        if (!touchesModel(refs)) pt.exts.push({ uri: attrOf(e.xml, "uri") ?? "", xml: x, refs, label: "çalışma kitabı uzantıları" });
      }
    } else if (!owned.has(c.name)) {
      const x = selfContained(c.xml, root.ns);
      const refs = await refsOf(src, wbPath, x, root.ns);
      if (!touchesModel(refs)) pt.children.push({ name: c.name, xml: x, refs, label: "çalışma kitabı ayarları", ...(isOrdered(c.name) ? {} : { after: lastKnown }) });
    }
    if (isOrdered(c.name)) lastKnown = c.name;
  }
  const attrXml = `<a ${pt.workbookPr.map(([k, v]) => `${k}="${v}"`).join(" ")}/>`;
  Object.assign(pt, ignorableUsed([...pt.children.map((c) => c.xml), ...pt.exts.map((e) => e.xml), pt.fileVersion ?? "", attrXml], root));
  return pt;
}
