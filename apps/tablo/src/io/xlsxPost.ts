/**
 * .xlsx paketinin son işlenmesi (kendi yazıcımızın çerçeveleri ve eski exceljs
 * yazıcısının çıktısı). Çalışma sayfası ve çalışma kitabı XML'ine eklenen tüm
 * öğeler (çizim, sayfa sonları, uzantılar, özet tablo önbellekleri, tanımlı
 * adlar, koruma…) XmlPart üzerinden yerleştirilir ve parçalar yazılırken
 * şemadaki sıraya dizilir (exceljs'in kendi yazdığı öğeler de: ör.
 * legacyDrawing'i extLst'ten sonra yazar).
 */
import JSZip from "jszip";
import { colToName } from "@officehere/formula";
import { disjointMerges } from "./bookUtil";
import type { SheetModel, WorkbookModel } from "../model/workbook";
import { sheetPaths } from "./ooxmlZip";
import { WORKBOOK_ORDER, WORKSHEET_ORDER, XmlPart, attrOf, childElements, getExt, rebuildElement, upsertExt } from "./xmlPart";
import { writeChartsInto } from "./xlsxCharts";
import { writePivotsInto } from "./xlsxPivot";
import { writeSparklinesInto, writeX14Sparklines } from "./xlsxSpark";
import { writePageBreaksInto } from "./xlsxBreaks";
import { writeDefinedNamesInto } from "./xlsxNames";
import { writeSheetExtras, writeWorkbookExtras } from "./xlsxExtras";
import { writeDataTablesInto } from "./xlsxWhatIf";
import { writeSlicersInto } from "./xlsxSlicers";
import { writeDynamicArrayMeta } from "./xlsxDynamic";
import { writeCheckboxesInto } from "./xlsxCheckbox";
import { patchThemeXml } from "../model/theme";
import { ContentTypes, RelsBuilder, parseRels, reattach, relsPathOf, type Carry, type CarryInstance, type OpcPackage } from "@officehere/core/opc";
import { decodeXString } from "./xlsxTables";
import { EXT_URI, condSig, tableFileName } from "./xlsxPackage";
import { parseSqref, sqrefText, type PtChild, type PtItem, type PtNamespaces, type SheetPassthrough } from "../model/passthrough";
import { subtractRects } from "../model/rect";

/** Paketteki XML parçalarının önbelleği: her parça bir kez ayrıştırılır, sonda sırasıyla yazılır. */
export class PackageParts {
  private cache = new Map<string, XmlPart>();
  constructor(readonly zip: JSZip) {}

  async part(path: string, order: readonly string[] = WORKSHEET_ORDER): Promise<XmlPart | null> {
    const hit = this.cache.get(path);
    if (hit) return hit;
    const f = this.zip.file(path);
    if (!f) return null;
    const p = new XmlPart(await f.async("string"), order);
    this.cache.set(path, p);
    return p;
  }

  /** Hazır parçayı önbelleğe koyar (yazıcının kurduğu sayfa iskeleti: yeniden ayrıştırılmaz). */
  seed(path: string, part: XmlPart): void {
    this.cache.set(path, part);
  }

  workbook(): Promise<XmlPart | null> {
    return this.part("xl/workbook.xml", WORKBOOK_ORDER);
  }

  /** Önbellekteki parçaları pakete yazar; `skip` verilen parçalar (yazıcının kendisi akıtacakları) atlanır. */
  flush(skip?: (path: string) => boolean): void {
    for (const [path, p] of this.cache) if (!skip?.(path)) this.zip.file(path, p.toString());
  }
}

/**
 * Birleştirmelerin tamamı mergeCells'e doğrudan yazılır (exceljs'e verilmez: mergeCells'i her
 * birleştirmede öncekilerin hepsini gezer). Geçersiz, tek hücrelik ve çakışanlar atılır.
 */
function writeMerges(part: XmlPart, sheet: SheetModel): void {
  const list = sheet.merges.length ? disjointMerges(sheet.merges).merges : [];
  if (!list.length) {
    part.set("mergeCells", null);
    return;
  }
  const items = list.map((m) => `<mergeCell ref="${colToName(m.c1)}${m.r1 + 1}:${colToName(m.c2)}${m.r2 + 1}"/>`);
  part.set("mergeCells", `<mergeCells count="${items.length}">${items.join("")}</mergeCells>`);
}

/** İçeriği olmayan ama gizli/yüksekliği özel/anahat düzeyli satırlar sheetData'ya eklenir. */
function fixRows(part: XmlPart, sheet: SheetModel): void {
  const rows = new Set<number>([...sheet.rows.customItems().map(([r]) => r), ...sheet.rows.hiddenItems(), ...sheet.rows.levelItems().map(([r]) => r)]);
  if (!rows.size) return;
  let xml = part.get("sheetData") ?? "<sheetData></sheetData>";
  // İçeriğe göre (otomatik) yükseklikteki satırlar customHeight taşımaz
  xml = xml.replace(/<row r="(\d+)"([^>]*?) customHeight="1"/g, (m0, r: string, rest: string) => (sheet.rows.isManual(Number(r) - 1) ? m0 : `<row r="${r}"${rest}`));
  const existing = new Set<number>([...xml.matchAll(/<row r="(\d+)"/g)].map((m) => Number(m[1]) - 1));
  const missing = [...rows].filter((r) => !existing.has(r)).sort((a, b) => a - b);
  if (missing.length) {
    if (/^<sheetData\s*\/>$/.test(xml)) xml = "<sheetData></sheetData>";
    // Mevcut satır başlangıçları bir kez taranır; eksikler sırayla araya yerleştirilir
    const starts = [...xml.matchAll(/<row r="(\d+)"/g)].map((m) => ({ r: Number(m[1]) - 1, at: m.index! }));
    const endAt = xml.lastIndexOf("</sheetData>");
    const pieces: string[] = [];
    let last = 0;
    let mi = 0;
    const tagFor = (r: number) => `<row r="${r + 1}"${sheet.rows.isCustom(r) ? ` ht="${Math.round(((sheet.rows.rawSize(r) * 72) / 96) * 100) / 100}"${sheet.rows.isManual(r) ? ' customHeight="1"' : ""}` : ""}${sheet.rows.isHidden(r) ? ' hidden="1"' : ""}${sheet.rows.level(r) ? ` outlineLevel="${sheet.rows.level(r)}"` : ""}/>`;
    for (const s of starts) {
      let add = "";
      while (mi < missing.length && missing[mi] < s.r) add += tagFor(missing[mi++]);
      if (add) {
        pieces.push(xml.slice(last, s.at), add);
        last = s.at;
      }
    }
    let tail = "";
    while (mi < missing.length) tail += tagFor(missing[mi++]);
    pieces.push(xml.slice(last, endAt), tail, xml.slice(endAt));
    xml = pieces.join("");
  }
  part.set("sheetData", xml);
}

const X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const XM_NS = "http://schemas.microsoft.com/office/excel/2006/main";
export const CF_EXT_URI = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}";
const RULE_EXT_URI = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}";

export function guid(): string {
  const h = Array.from({ length: 32 }, () => Math.floor(Math.random() * 16).toString(16).toUpperCase()).join("");
  return `{${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20)}}`;
}

/**
 * Veri çubukları: kuraldaki x14:id ile sayfa düzeyindeki x14:conditionalFormatting
 * eşleştirilir (2010 sonrası çubuk görünümü: kenarlık, negatif değer rengi, eksen).
 * exceljs boş <x14:id/> yazar ve sayfa uzantısını hiç yazmaz.
 */
export function fixDataBars(part: XmlPart): void {
  const entries: string[] = [];
  part.map("conditionalFormatting", (cf) => {
    if (!/type="dataBar"/.test(cf)) return cf;
    const sqref = attrOf(cf, "sqref") ?? "";
    const rules = childElements(cf).map((c) => {
      if (c.name !== "cfRule" || attrOf(c.xml, "type") !== "dataBar") return c.xml;
      const id = guid();
      const color = /<(?:\w+:)?color\b[^>]*\brgb="([0-9A-Fa-f]{8})"/.exec(c.xml)?.[1] ?? "FF638EC6";
      const inner = childElements(c.xml).filter((k) => k.name !== "extLst").map((k) => k.xml);
      inner.push(`<extLst><ext uri="${RULE_EXT_URI}" xmlns:x14="${X14_NS}"><x14:id>${id}</x14:id></ext></extLst>`);
      entries.push(
        `<x14:conditionalFormatting xmlns:xm="${XM_NS}"><x14:cfRule type="dataBar" id="${id}"><x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:borderColor rgb="${color}"/><x14:negativeFillColor rgb="FFFF0000"/><x14:negativeBorderColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule><xm:sqref>${sqref}</xm:sqref></x14:conditionalFormatting>`,
      );
      return rebuildElement(c.xml, inner);
    });
    return rebuildElement(cf, rules);
  });
  if (!entries.length) return;
  // exceljs'in yazdığı (ör. 3 Üçgen simge kümesi) koşullu biçim uzantısıyla birleştirilir
  const ext = getExt(part, CF_EXT_URI);
  let list = ext ? childElements(ext).find((c) => c.name === "conditionalFormattings")?.xml : undefined;
  const items = list ? [...childElements(list).map((c) => c.xml), ...entries] : entries;
  list = list ? rebuildElement(list, items) : `<x14:conditionalFormattings>${items.join("")}</x14:conditionalFormattings>`;
  const others = ext ? childElements(ext).filter((c) => c.name !== "conditionalFormattings").map((c) => c.xml) : [];
  const extXml = ext ? rebuildElement(ext, [list, ...others]) : `<ext uri="${CF_EXT_URI}" xmlns:x14="${X14_NS}">${list}</ext>`;
  upsertExt(part, CF_EXT_URI, extXml);
}

/* ============================== Geçiş katmanı: geri koyma ============================== */

/** Yazma turunun geçiş bağlamı: kaynak paket, öğe düzeyi taşıma planı, sayfa/parça eşlemesi ve uyarılar. */
export interface PassthroughContext {
  /** Kaynak paket (kayıtta yoksa undefined: ilişkili parçacıklar yazılamaz). */
  src?: OpcPackage;
  carry: Carry;
  /** Kaynak parça → çıktı parçası (sayfalar, tablolar, grafikler): carryParts'a verilir. */
  partMap: [string, string][];
  /** Kaydedilemeyen öğelerin etiketleri (Türkçe). */
  warnings: string[];
}

const NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const NS_XDR_ = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const NS_A_ = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_R_ = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const CT_DRAWING_ = "application/vnd.openxmlformats-officedocument.drawing+xml";
const CT_VML = "application/vnd.openxmlformats-officedocument.vmlDrawing";
const VML_ROOT = '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">';
const X14_NS_ = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const XM_NS_ = "http://schemas.microsoft.com/office/excel/2006/main";

/** Açılış etiketine mc:Ignorable önekleri ve bildirimleri ekler (önek başka ad alanına bağlıysa atlanır). */
function mergeIgnorable(open: string, ns: PtNamespaces): string {
  let out = open;
  for (const p of ns.ignorable) {
    const uri = ns.ns[p];
    if (!uri) continue;
    const have = new RegExp(`\\sxmlns:${p}="([^"]*)"`).exec(out);
    if (have && have[1] !== uri) continue;
    if (!have) out = out.replace(/^<([^\s>/]+)/, `<$1 xmlns:${p}="${uri}"`);
    if (!/\sxmlns:mc="/.test(out)) out = out.replace(/^<([^\s>/]+)/, `<$1 xmlns:mc="${NS_MC}"`);
    const ig = /\smc:Ignorable="([^"]*)"/.exec(out);
    if (!ig) out = out.replace(/^<([^\s>/]+)/, `<$1 mc:Ignorable="${p}"`);
    else if (!ig[1].split(/\s+/).includes(p)) out = out.replace(ig[0], ` mc:Ignorable="${`${ig[1]} ${p}`.trim()}"`);
  }
  return out;
}

/** Açılış etiketinde olmayan öznitelikleri ekler (önekli olanlar yalnız önek bildirilmişse). */
function addAttrs(xml: string, attrs: [string, string][], root: string): string {
  const open = /^<[^>]*?(\/?)>/.exec(xml);
  if (!open) return xml;
  let tag = open[0];
  for (const [k, v] of attrs) {
    if (new RegExp(`\\s${k.replace(/[.]/g, "\\.")}="`).test(tag)) continue;
    const p = k.includes(":") ? k.slice(0, k.indexOf(":")) : "";
    if (p && !new RegExp(`\\sxmlns:${p}="`).test(root) && !new RegExp(`\\sxmlns:${p}="`).test(tag)) continue;
    tag = tag.replace(/(\/?>)$/, ` ${k}="${v}"$1`);
  }
  return tag + xml.slice(open[0].length);
}

/**
 * Şema sırasında yeri olmayan öğeyi (mc:AlternateContent, xr:revisionPtr…) dosyadaki konumuna koyar:
 * kaynakta kendinden önce gelen bilinen öğenin (ya da sırası ondan küçük en büyük öğenin) arkasına.
 */
function placeChild(part: XmlPart, order: readonly string[], c: PtChild, xml: string): void {
  if (order.includes(c.name)) {
    part.add(c.name, xml);
    return;
  }
  const limit = c.after ? order.indexOf(c.after) : -1;
  let idx = -1;
  let best = -1;
  part.children.forEach((x, i) => {
    const r = order.indexOf(x.name);
    if (r >= 0 && r <= limit && r >= best) {
      best = r;
      idx = i;
    }
  });
  while (idx + 1 < part.children.length && order.indexOf(part.children[idx + 1].name) < 0) idx++;
  part.children.splice(idx + 1, 0, { name: c.name, xml });
}

/** Parçacığın ilk etiketindeki, kökte aynısı bildirilmiş ad alanı bildirimlerini çıkarır. */
function trimNs(xml: string, root: string): string {
  return xml.replace(/^<[^>]*>/, (t) => t.replace(/\sxmlns:([\w.-]+)="([^"]*)"/g, (m, p: string, u: string) => (root.includes(` xmlns:${p}="${u}"`) ? "" : m)));
}

/** Parçanın .rels kurucusu (JSZip'teki mevcut ilişkilerle). */
async function relsOf(zip: JSZip, part: string): Promise<RelsBuilder> {
  const xml = await zip.file(relsPathOf(part))?.async("string");
  return new RelsBuilder(part, parseRels(xml ?? null, part));
}

/** Aynı paket içi konuma (Target="#Sayfa!A1") giden ilişkiler: TargetMode="External" yazılmaz (özgündeki gibi). */
const localTargets = new WeakMap<RelsBuilder, Set<string>>();

function saveRels(zip: JSZip, b: RelsBuilder): void {
  const local = localTargets.get(b);
  let xml = b.xml();
  if (local?.size) xml = xml.replace(/<Relationship\b[^>]*>/g, (tag) => (local.has(/\sId="([^"]*)"/.exec(tag)?.[1] ?? "") ? tag.replace(/\sTargetMode="External"/, "") : tag));
  zip.file(relsPathOf(b.source), xml);
}

function freePath(zip: JSZip, make: (n: number) => string): string {
  let n = 1;
  while (zip.file(make(n))) n++;
  return make(n);
}

/** Parçacığın ilişkilerini yeni pakete bağlar; kaynak yoksa ya da bir hedef kopyalanamazsa null. */
function attach(it: PtItem, inst: CarryInstance | null, rels: () => RelsBuilder): string | null {
  if (!it.refs?.length) return it.xml;
  if (!inst?.source) return null;
  // Paket içi konum hedefi (ör. şekil köprüsü "#Giriş!A1") parça değildir: hedefiyle aynen yazılır
  const refs = it.refs.map((r) => (!r.external && r.target.startsWith("#") ? { ...r, external: true, local: true } : r));
  const r = reattach(it.xml, refs, inst, (x) => {
    if (!x.external) return rels().addPart(x.type, x.target);
    const id = rels().addExternal(x.type, x.target);
    if (refs.some((f) => "local" in f && f.target === x.target)) {
      const set = localTargets.get(rels()) ?? new Set<string>();
      set.add(id);
      localTargets.set(rels(), set);
    }
    return id;
  });
  return r.missing.length ? null : r.xml;
}

/** Sayfa XML'indeki öğenin r:id'sinin gösterdiği parça (çıktı paketinde). */
function relTargetOf(b: RelsBuilder, partXml: string | undefined): string | undefined {
  const id = partXml ? /\sr:id="([^"]+)"/.exec(partXml)?.[1] : undefined;
  return id ? (b.relationships.find((r) => r.id === id && !r.external)?.part ?? undefined) : undefined;
}

/** Bağlı x14 kuralı değişmemişse modeldeki kuralın yerine özgün x14 girdisi yazılır (yazıcı kuralı atlar). */
export function passthroughCfRules(sheet: SheetModel): Set<string> {
  const out = new Set<string>();
  const pt = sheet.passthrough;
  if (!pt || pt.structVersion !== sheet.structVersion) return out;
  for (const e of pt.cf) {
    if (!e.rule) continue;
    const rule = sheet.condFormats.find((r) => r.id === e.rule);
    const sq = /<(?:[\w.-]+:)?sqref>([^<]*)</.exec(e.xml)?.[1] ?? "";
    const norm = (s: string) => s.trim().split(/\s+/).sort().join(" ");
    if (rule && condSig(rule) === e.sig && norm(rule.ranges.map(sqrefText).join(" ")) === norm(sq)) out.add(rule.id);
  }
  return out;
}

/** Dosyadan gelen, modelin tanımadığı öğeleri yazıcının çıktısına geri koyar (sayfalar ve workbook.xml). */
export async function restorePassthrough(zip: JSZip, wb: WorkbookModel, paths: Map<string, string>, parts: PackageParts, ctx: PassthroughContext): Promise<void> {
  const ct = ContentTypes.parse((await zip.file("[Content_Types].xml")?.async("string")) ?? "");
  let ctChanged = false;
  const mapped = new Set<string>();
  for (const sheet of wb.sheets) {
    const pt = sheet.passthrough;
    const path = paths.get(sheet.name);
    const part = path ? await parts.part(path) : null;
    if (!pt || !path || !part) continue;
    if (pt.part && !mapped.has(pt.part.toLowerCase())) { // yerel-ayar: ascii
      mapped.add(pt.part.toLowerCase()); // yerel-ayar: ascii
      ctx.partMap.push([pt.part, path]);
    }
    for (const label of pt.dropped) ctx.warnings.push(label);
    const inst = ctx.src ? ctx.carry.instance(ctx.src) : null;
    // Yapı sürümü tutmuyorsa (parçacıklar sayfanın güncel yapısına göre değil) konuma bağlı öğeler yazılamaz
    if (pt.structVersion !== sheet.structVersion) {
      for (const it of [...pt.children, ...pt.exts, ...pt.cf, ...pt.dv, ...(pt.drawing?.anchors ?? []), ...(pt.vml?.shapes ?? []), ...pt.parts]) ctx.warnings.push(it.label);
      continue;
    }
    const sheetRels = await relsOf(zip, path);
    const sheetRelsUsed = { v: false };
    const rels = () => ((sheetRelsUsed.v = true), sheetRels);
    const drop = (it: { label: string }) => ctx.warnings.push(it.label);
    await mapTablesAndCharts(zip, sheet, pt, sheetRels, ctx);
    if (pt.sheetPr.length) {
      const cur = part.get("sheetPr");
      part.set("sheetPr", cur ? addAttrs(cur, pt.sheetPr, part.open) : addAttrs("<sheetPr/>", pt.sheetPr, part.open));
    }
    // VML şekilleri (OLE/form denetimi görünümleri): şekil kimlikleri çakışmasın diye yeniden numaralanır
    const shapeIds = new Map<string, string>();
    if (pt.vml?.shapes.length) {
      const vmlPath = relTargetOf(sheetRels, part.get("legacyDrawing"));
      const vmlFile = vmlPath ? zip.file(vmlPath) : null;
      let vml = vmlFile ? await vmlFile.async("string") : null;
      const vPath = vml && vmlPath ? vmlPath : freePath(zip, (n) => `xl/drawings/vmlDrawing${n}.vml`);
      const vRels = await relsOf(zip, vPath);
      let next = Math.max(1024, ...[...(vml ?? "").matchAll(/_x0000_s(\d+)/g)].map((m) => Number(m[1])));
      const vmlRoot = (vml && /<xml\b[^>]*>/.exec(vml)?.[0]) || VML_ROOT;
      const shapes: string[] = [];
      for (const s of pt.vml.shapes) {
        let x = attach(s, inst, () => vRels);
        if (x === null) {
          drop(s);
          continue;
        }
        if (vml) {
          const old = /\sid="_x0000_s(\d+)"/.exec(x)?.[1];
          if (old) {
            const id = String(++next);
            shapeIds.set(old, id);
            x = x.replace(/(\s(?:id|o:spid)="_x0000_s)\d+(")/g, `$1${id}$2`);
          }
        }
        shapes.push(trimNs(x, vmlRoot));
      }
      if (shapes.length) {
        const types = pt.vml.shapetypes.filter((t) => !vml?.includes(`id="${/\sid="([^"]+)"/.exec(t)?.[1]}"`)).map((t) => trimNs(t, vmlRoot));
        const blocks = [...new Set(shapes.map((x) => Math.floor(Number(/\sid="_x0000_s(\d+)"/.exec(x)?.[1] ?? 1025) / 1024)))];
        if (vml) {
          const idmap = /<o:idmap\b[^>]*\sdata="([^"]*)"/.exec(vml);
          if (idmap) {
            const all = [...new Set([...idmap[1].split(",").map((v) => Number(v.trim())).filter((v) => v > 0), ...blocks])];
            vml = vml.replace(idmap[0], idmap[0].replace(`data="${idmap[1]}"`, `data="${all.join(",")}"`));
          }
          vml = vml.replace(/<\/xml>\s*$/, `${types.join("")}${shapes.join("")}</xml>`);
        } else {
          vml = `${VML_ROOT}<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${blocks.join(",")}"/></o:shapelayout>${types.join("")}${shapes.join("")}</xml>`;
          const rid = rels().addPart(REL + "vmlDrawing", vPath);
          part.set("legacyDrawing", `<legacyDrawing r:id="${rid}"/>`);
          part.ensureNamespace("r", NS_R_);
          if (!ct.getDefault("vml")) {
            ct.setDefault("vml", CT_VML);
            ctChanged = true;
          }
        }
        zip.file(vPath, vml);
        if (vRels.size) saveRels(zip, vRels);
      }
    }
    const remapShapes = (x: string) => (shapeIds.size ? x.replace(/(\sshapeId=")(\d+)(")/g, (m, a: string, n: string, b: string) => (shapeIds.has(n) ? a + shapeIds.get(n) + b : m)).replace(/(\sspid="_x0000_s)(\d+)(")/g, (m, a: string, n: string, b: string) => (shapeIds.has(n) ? a + shapeIds.get(n) + b : m)) : x);
    // Üst düzey öğeler (yazıcının ürettiği aynı adlı öğe varsa onunki geçerli)
    const written = new Set(part.children.map((c) => c.name));
    for (const c of pt.children) {
      if (written.has(c.name)) continue;
      const x = attach(c, inst, rels);
      if (x === null) drop(c);
      else placeChild(part, WORKSHEET_ORDER, c, trimNs(remapShapes(x), part.open));
    }
    // x14 koşullu biçim girdileri (yazıcının uzantısıyla birleştirilir; öncelikler model kurallarından sonra)
    const keep = passthroughCfRules(sheet);
    const cfItems: string[] = [];
    let priority = sheet.condFormats.filter((r) => !keep.has(r.id)).length;
    for (const e of pt.cf) {
      if (e.rule && !keep.has(e.rule)) continue;
      const x = attach(e, inst, rels);
      if (x === null) drop(e);
      else cfItems.push(x.replace(/(<(?:[\w.-]+:)?cfRule\b[^>]*\spriority=")\d+(")/g, (_m, a: string, b: string) => `${a}${++priority}${b}`));
    }
    if (cfItems.length) {
      const ext = getExt(part, EXT_URI.cf);
      const list = ext ? childElements(ext).find((c) => c.name === "conditionalFormattings")?.xml : undefined;
      const items = list ? [...childElements(list).map((c) => c.xml), ...cfItems] : cfItems;
      const listXml = list ? rebuildElement(list, items) : `<x14:conditionalFormattings>${items.join("")}</x14:conditionalFormattings>`;
      const others = ext ? childElements(ext).filter((c) => c.name !== "conditionalFormattings").map((c) => c.xml) : [];
      upsertExt(part, EXT_URI.cf, ext ? rebuildElement(ext, [listXml, ...others]) : `<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="${X14_NS_}">${listXml}</ext>`);
    }
    // x14 veri doğrulamaları: modeldeki doğrulamaların hücreleri çıkarılır (sonra eklenen kural geçerli)
    const dvItems: string[] = [];
    const modelDv = sheet.validations.flatMap((v) => v.ranges);
    for (const e of pt.dv) {
      let x = attach(e, inst, rels);
      if (x === null) {
        drop(e);
        continue;
      }
      if (modelDv.length) {
        const sq = /<((?:[\w.-]+:)?)sqref>([^<]*)<\/\1sqref>/.exec(x);
        if (sq) {
          const rest = subtractRects(parseSqref(sq[2]), modelDv);
          if (!rest.length) continue;
          x = x.replace(sq[0], `<${sq[1]}sqref>${rest.map(sqrefText).join(" ")}</${sq[1]}sqref>`);
        }
      }
      dvItems.push(x);
    }
    if (dvItems.length) upsertExt(part, EXT_URI.dv, `<ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="${X14_NS_}"><x14:dataValidations count="${dvItems.length}" xmlns:xm="${XM_NS_}">${dvItems.join("")}</x14:dataValidations></ext>`);
    // Diğer uzantılar (yazıcınınki geçerli)
    for (const e of pt.exts) {
      if (getExt(part, e.uri)) continue;
      const x = attach(e, inst, rels);
      if (x === null) drop(e);
      else upsertExt(part, e.uri, trimNs(remapShapes(x), part.open));
    }
    if (pt.cf.length || pt.dv.length || pt.exts.length || pt.children.length) part.open = mergeIgnorable(part.open, pt);
    // Çizimdeki tanınmayan çapalar
    const anchors = (pt.drawing?.anchors ?? []).filter((a) => !a.obj || sheet.images.some((i) => i.id === a.obj));
    if (pt.drawing && anchors.length) {
      let dPath = relTargetOf(sheetRels, part.get("drawing"));
      let dXml = dPath ? await zip.file(dPath)?.async("string") : undefined;
      if (!dPath || dXml === undefined) {
        dPath = freePath(zip, (n) => `xl/drawings/drawing${n}.xml`);
        dXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="${NS_XDR_}" xmlns:a="${NS_A_}"></xdr:wsDr>`;
        const rid = rels().addPart(REL + "drawing", dPath);
        part.set("drawing", `<drawing r:id="${rid}"/>`);
        part.ensureNamespace("r", NS_R_);
        ct.setOverride(dPath, CT_DRAWING_);
        ctChanged = true;
      }
      const dRels = await relsOf(zip, dPath);
      let nv = Math.max(0, ...[...dXml.matchAll(/<(?:[\w.-]+:)?cNvPr\b[^>]*\sid="(\d+)"/g)].map((m) => Number(m[1])));
      const out: string[] = [];
      for (const a of anchors) {
        let x = attach(a, inst, () => dRels);
        if (x === null) {
          drop(a);
          continue;
        }
        // Çizim içi şekil kimlikleri benzersiz olmalı; bağlayıcı uçları aynı eşlemeyle
        const ids = new Map<string, number>();
        x = x.replace(/(<(?:[\w.-]+:)?cNvPr\b[^>]*\sid=")(\d+)(")/g, (_m, p: string, n: string, q: string) => {
          if (!ids.has(n)) ids.set(n, ++nv);
          return p + ids.get(n) + q;
        });
        x = x.replace(/(<a:(?:stCxn|endCxn)\b[^>]*\sid=")(\d+)(")/g, (m, p: string, n: string, q: string) => (ids.has(n) ? p + ids.get(n) + q : m));
        out.push(trimNs(remapShapes(x), /<(?:[\w.-]+:)?wsDr\b[^>]*>/.exec(dXml)?.[0] ?? ""));
      }
      if (out.length) {
        const close = /<\/((?:[\w.-]+:)?)wsDr>\s*$/.exec(dXml);
        if (close) {
          dXml = dXml.slice(0, close.index) + out.join("") + dXml.slice(close.index);
          dXml = dXml.replace(/<((?:[\w.-]+:)?)wsDr\b[^>]*>/, (open) => mergeIgnorable(open, pt.drawing!));
          if (!/<(?:[\w.-]+:)?wsDr\b[^>]*\sxmlns:r=/.test(dXml)) dXml = dXml.replace(/<((?:[\w.-]+:)?)wsDr\b/, `<$1wsDr xmlns:r="${NS_R_}"`);
          zip.file(dPath, dXml);
          if (dRels.size) saveRels(zip, dRels);
        }
      }
    }
    // Hücre başvurulu ek parçalar (XML eşlemeli tek hücreler)
    for (const p of pt.parts) {
      const to = freePath(zip, (n) => `xl/tables/tableSingleCells${n}.xml`);
      zip.file(to, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${p.xml}`);
      rels().addPart(p.type, to);
      ct.setOverride(to, p.contentType);
      ctChanged = true;
    }
    if (sheetRelsUsed.v) saveRels(zip, sheetRels);
  }
  await restoreWorkbook(zip, wb, parts, ctx);
  if (ctChanged) zip.file("[Content_Types].xml", ct.toXml());
}

/**
 * Tablo ve grafik eşlemesi (carryParts örtük ilişkileri bunlardan da taşısın: sorgu tablosu, grafik
 * stili ve renkleri). Sorgu/XML tablolarında yazıcının sütun kimlikleri ve özniteliklerinin özgünle
 * aynı olması gerekir (queryTableField tableColumnId): uyuşmazsa eşlenmez, uyarılır.
 */
async function mapTablesAndCharts(zip: JSZip, sheet: SheetModel, pt: SheetPassthrough, sheetRels: RelsBuilder, ctx: PassthroughContext): Promise<void> {
  if (!ctx.src) return;
  for (const r of sheetRels.relationships) {
    if (r.external || !r.part || !/\/table$/.test(r.type)) continue;
    const xml = await zip.file(r.part)?.async("string");
    const display = xml ? /<(?:\w+:)?table\b[^>]*?\sdisplayName="([^"]*)"/.exec(xml)?.[1] : undefined;
    const t = display !== undefined ? sheet.tables.find((x) => tableFileName(x.name) === display) : undefined;
    const from = t ? pt.tables[t.id] : undefined;
    if (!xml || !from) continue;
    const srcXml = await ctx.src.text(from).catch(() => null);
    if (!srcXml) continue;
    const special = /<(?:\w+:)?table\b[^>]*\stableType="(queryTable|xml)"/.test(srcXml) || /queryTableFieldId=|<(?:\w+:)?xmlColumnPr\b/.test(srcXml);
    if (special) {
      const patched = patchTableXml(xml, srcXml);
      if (!patched) {
        ctx.warnings.push("dış veri sorgusu bağlantıları");
        continue;
      }
      zip.file(r.part, patched);
    }
    ctx.partMap.push([from, r.part]);
  }
  // Grafikler: çıktı çizimindeki grafik çapaları sheet.charts sırasıyla yazılır (writeChartsInto)
  const charts = Object.keys(pt.charts).length ? sheet.charts : [];
  if (!charts.length) return;
  const d = sheetRels.relationships.find((r) => !r.external && r.part && /\/drawing$/.test(r.type));
  const dXml = d?.part ? await zip.file(d.part)?.async("string") : undefined;
  if (!d?.part || !dXml) return;
  const dRels = parseRels((await zip.file(relsPathOf(d.part))?.async("string")) ?? null, d.part);
  const out = [...dXml.matchAll(/<c:chart\b[^>]*\sr:id="([^"]+)"/g)].map((m) => dRels.find((x) => x.id === m[1])?.part).filter((p): p is string => !!p);
  if (out.length !== charts.length) return;
  charts.forEach((c, i) => {
    const from = pt.charts[c.id];
    if (from) ctx.partMap.push([from, out[i]]);
  });
}

/** Sorgu/XML tablosunun özgün sütun kimliklerini ve bağlarını yazıcının tablo XML'ine taşır; uyuşmazsa null. */
function patchTableXml(out: string, src: string): string | null {
  const colsOf = (x: string) => [...x.matchAll(/<(?:\w+:)?tableColumn\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?tableColumn>)/g)].map((m) => ({ all: m[0], attrs: m[1], body: m[2] ?? "", name: decodeXString((/\sname="([^"]*)"/.exec(m[1])?.[1] ?? "").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"')) }));
  const a = colsOf(out);
  const b = colsOf(src);
  if (a.length !== b.length || a.some((c, i) => c.name !== b[i].name)) return null;
  let i = 0;
  let xml = out.replace(/<(?:\w+:)?tableColumn\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?tableColumn>)/g, () => {
    const s = b[i];
    const o = a[i++];
    let attrs = o.attrs.replace(/\sid="\d+"/, ` id="${/\sid="(\d+)"/.exec(s.attrs)?.[1] ?? i}"`);
    for (const k of ["uniqueName", "queryTableFieldId"]) {
      const v = new RegExp(`\\s${k}="([^"]*)"`).exec(s.attrs)?.[1];
      if (v !== undefined && !new RegExp(`\\s${k}=`).test(attrs)) attrs += ` ${k}="${v}"`;
    }
    const xmlCol = /<(?:\w+:)?xmlColumnPr\b[^>]*?(?:\/>|>[\s\S]*?<\/(?:\w+:)?xmlColumnPr>)/.exec(s.body)?.[0] ?? "";
    const body = o.body + xmlCol;
    return body ? `<tableColumn${attrs}>${body}</tableColumn>` : `<tableColumn${attrs}/>`;
  });
  const head = /<(?:\w+:)?table\b[^>]*>/.exec(src)?.[0] ?? "";
  const extra: [string, string][] = [];
  for (const k of ["tableType", "connectionId"]) {
    const v = new RegExp(`\\s${k}="([^"]*)"`).exec(head)?.[1];
    if (v !== undefined) extra.push([k, v]);
  }
  xml = xml.replace(/<((?:\w+:)?)table\b[^>]*>/, (open) => addAttrs(open, extra, open));
  return xml;
}

/** workbook.xml yaması: fileVersion, workbookPr/calcPr öznitelikleri, tanınmayan öğeler ve uzantılar. */
async function restoreWorkbook(zip: JSZip, wb: WorkbookModel, parts: PackageParts, ctx: PassthroughContext): Promise<void> {
  const pt = wb.passthrough;
  const part = pt ? await parts.workbook() : null;
  if (!pt || !part) return;
  const path = "xl/workbook.xml";
  const b = await relsOf(zip, path);
  const used = { v: false };
  const rels = () => ((used.v = true), b);
  const inst = ctx.src ? ctx.carry.instance(ctx.src) : null;
  if (pt.fileVersion) part.set("fileVersion", pt.fileVersion);
  if (pt.workbookPr.length) {
    const cur = part.get("workbookPr");
    part.set("workbookPr", addAttrs(cur ?? "<workbookPr/>", pt.workbookPr, part.open));
  }
  if (pt.calcPr.length) {
    const cur = part.get("calcPr");
    part.set("calcPr", addAttrs(cur ?? "<calcPr/>", pt.calcPr, part.open));
  }
  const written = new Set(part.children.map((c) => c.name));
  for (const c of pt.children) {
    if (written.has(c.name) && c.name !== "AlternateContent") continue;
    const x = attach(c, inst, rels);
    if (x === null) ctx.warnings.push(c.label);
    else placeChild(part, WORKBOOK_ORDER, c, trimNs(x, part.open));
  }
  for (const e of pt.exts) {
    if (getExt(part, e.uri)) continue;
    const x = attach(e, inst, rels);
    if (x === null) ctx.warnings.push(e.label);
    else upsertExt(part, e.uri, trimNs(x, part.open), []);
  }
  // Grafik sayfaları: sayfa listesinin sonuna (dizine bağlı localSheetId/activeTab kaymasın)
  const sheets = part.get("sheets");
  if (pt.chartsheets.length && sheets) {
    const items = childElements(sheets).map((c) => c.xml);
    const names = new Set(items.map((x) => unescapeXmlAttr(attrOf(x, "name") ?? "").toLocaleUpperCase("tr-TR")));
    let sheetId = Math.max(0, ...items.map((x) => Number(attrOf(x, "sheetId") ?? 0)));
    for (const cs of pt.chartsheets) {
      if (names.has(cs.name.toLocaleUpperCase("tr-TR"))) {
        ctx.warnings.push(cs.label);
        continue;
      }
      let x = attach(cs, inst, rels);
      if (x === null || !inst) {
        ctx.warnings.push(cs.label);
        continue;
      }
      for (const p of cs.chain.slice(1)) inst.copy(p);
      for (const c of cs.charts) {
        const to = inst.copy(c.part);
        if (to) ctx.carry.patch(to, () => c.xml);
      }
      x = x.replace(/(\ssheetId=")\d+(")/, `$1${++sheetId}$2`);
      items.push(trimNs(x, part.open));
      names.add(cs.name.toLocaleUpperCase("tr-TR"));
    }
    part.set("sheets", rebuildElement(sheets, items));
  }
  part.open = mergeIgnorable(part.open, pt);
  if (used.v) saveRels(zip, b);
}

const unescapeXmlAttr = (s: string) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");

/** Sonradan işlem seçenekleri. */
export interface PostOptions {
  /**
   * Hücre verisi yazıcının kendisinde (yeni yazıcı): satırlar, veri tabloları, dinamik dizi üst verisi
   * ve onay kutulu biçimler sayfa akıtılırken yazılır; sheetData'ya dokunulmaz.
   */
  ownSheetData?: boolean;
}

/**
 * Paketi (JSZip; sayfalar PackageParts önbelleğinde) son işler: birleştirmeler, veri çubukları,
 * grafikler, özet tablolar, uzantılar, sayfa sonları, adlar, geçiş katmanı ve tema. Parçalar
 * önbellekte kalır; yazmak çağıranın işidir (PackageParts.flush).
 */
export async function postProcessParts(zip: JSZip, wb: WorkbookModel, paths: Map<string, string>, parts: PackageParts, ctx?: PassthroughContext, opts: PostOptions = {}): Promise<void> {
  const own = !!opts.ownSheetData;
  for (const sheet of wb.sheets) {
    const path = paths.get(sheet.name);
    const part = path ? await parts.part(path) : null;
    if (!part) continue;
    if (!own) fixRows(part, sheet);
    fixDataBars(part);
    writeMerges(part, sheet);
  }
  if (wb.sheets.some((s) => s.charts.length)) await writeChartsInto(zip, wb, paths, parts);
  if (wb.sheets.some((s) => s.pivots.length)) await writePivotsInto(zip, wb, paths, parts);
  await writeSparklinesInto(zip, wb);
  await writeSlicersInto(zip, wb);
  await writeX14Sparklines(wb, paths, parts);
  await writePageBreaksInto(wb, paths, parts);
  await writeSheetExtras(wb, paths, parts);
  if (!own) {
    await writeDataTablesInto(wb, paths, parts);
    await writeDynamicArrayMeta(zip, wb, paths, parts);
    await writeCheckboxesInto(zip, wb, paths, parts);
  }
  await writeDefinedNamesInto(wb, parts);
  await writeWorkbookExtras(wb, parts);
  if (ctx) await restorePassthrough(zip, wb, paths, parts, ctx);
  // Çalışma kitabı teması (renk ve yazı tipi şemaları)
  const themeFile = zip.file("xl/theme/theme1.xml");
  if (themeFile) zip.file("xl/theme/theme1.xml", patchThemeXml(await themeFile.async("string"), wb.theme));
}

/** exceljs çıktısını son işler: satırlar, çizimler, uzantılar, sayfa sonları, adlar ve şema sırası. */
export async function postProcess(bytes: Uint8Array, wb: WorkbookModel, ctx?: PassthroughContext): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes);
  const paths = await sheetPaths(zip);
  const parts = new PackageParts(zip);
  await postProcessParts(zip, wb, paths, parts, ctx);
  parts.flush();
  // Büyük kitaplarda daha hızlı sıkıştırma düzeyi (200 bin satırda: boyut %10 artar, süre yarıya iner)
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: bytes.length > 24 * 1024 * 1024 ? 3 : 6 } });
}
