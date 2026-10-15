/**
 * Kendi XLSX yazıcımız (exceljs'in yerine): çalışma kitabı → paket. Aşamalar:
 *  1) Sayfalar planlanır (sayfa iskeletleri, ilişkiler, notlar, resimler, tablolar), küçük parçalar
 *     (kitap, ilişkiler, içerik türleri, tema, belge özellikleri) bellekteki pakete yazılır.
 *  2) Sonradan işlem (xlsxPost: grafikler, özet tablolar, uzantılar, adlar, geçiş katmanı) iskeletler
 *     üzerinde, sayfa verisine dokunmadan çalışır.
 *  3) Sayfa verisi satır sırasıyla üretilip parça parça sıkıştırılır (tek sıkıştırma); ardından
 *     paylaşılan metinler, stiller ve gerekirse dinamik dizi üst verisi / onay kutusu özellik torbası.
 *  4) Paket kurulur; kaynak paket ya da makro varsa geçiş katmanı (carryParts) ham kopyayla ekler.
 */
import JSZip from "jszip";
import { Carry, ZipWriter, carryParts } from "@officehere/core/opc";
import type { WorkbookModel } from "../../model/workbook";
import { PackageParts, postProcessParts, xlsxSource, type PassthroughContext } from "../xlsxPost";
import { CT, addOverride, contentTypesXml } from "./contentTypes";
import { printNames } from "./pageSetup";
import { REL, Rels, appendRel, relsXml } from "./rels";
import { SharedStrings } from "./sharedStrings";
import { SheetWriter, type WriteContext } from "./sheet";
import { StyleSheet } from "./styles";
import { themeXml } from "./theme";
import { DeflateSink, addPacked, packText, type PackedEntry } from "./zip";
import { esc, XML_DECL } from "./xml";

export interface XlsxWriterOptions {
  macro?: boolean;
  report?: { warnings: string[] };
  /** Evre süreleri (ms), ölçüm betikleri için: plan, sonradan işlem, sayfalar, paket, taşıma. */
  sureler?: Record<string, number>;
}

const METADATA_XML =
  XML_DECL +
  '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">' +
  '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>' +
  '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>' +
  '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>' +
  "</metadata>";

const FPB_PATH = "xl/featurePropertyBag/featurePropertyBag.xml";
const FPB_XML =
  XML_DECL +
  '<FeaturePropertyBags xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag">' +
  '<bag type="Checkbox"/>' +
  '<bag type="XFControls"><bagId k="CellControl">0</bagId></bag>' +
  '<bag type="XFComplement"><bagId k="XFControls">1</bagId></bag>' +
  '<bag type="XFComplements" extRef="XFComplementsMapperExtRef"><a k="MappedFeaturePropertyBags"><bagId>2</bagId></a></bag>' +
  "</FeaturePropertyBags>";

function isoDate(d: Date | undefined): string {
  const t = d instanceof Date && !Number.isNaN(d.getTime()) ? d : new Date();
  return t.toISOString().replace(/[.]\d{3}/, "");
}

function coreXml(wb: WorkbookModel): string {
  const p = wb.props;
  const who = esc(p.author ?? "OfficeHere Tablo");
  return (
    XML_DECL +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    `<dc:creator>${who}</dc:creator><dc:title>${esc(p.title || "")}</dc:title><dc:subject>${esc(p.subject || "")}</dc:subject><dc:description></dc:description><cp:keywords></cp:keywords><cp:category></cp:category>` +
    `<cp:lastModifiedBy>${who}</cp:lastModifiedBy>` +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${isoDate(p.created)}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${isoDate(new Date())}</dcterms:modified>` +
    "</cp:coreProperties>"
  );
}

function appXml(wb: WorkbookModel): string {
  const names = wb.sheets.map((s) => `<vt:lpstr>${esc(s.name)}</vt:lpstr>`).join("");
  return (
    XML_DECL +
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
    "<Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>" +
    `<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>${wb.sheets.length}</vt:i4></vt:variant></vt:vector></HeadingPairs>` +
    `<TitlesOfParts><vt:vector size="${wb.sheets.length}" baseType="lpstr">${names}</vt:vector></TitlesOfParts>` +
    `<Company>${esc(wb.props.company || "")}</Company><Manager></Manager>` +
    "<LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion>" +
    "</Properties>"
  );
}

function workbookXml(wb: WorkbookModel, sheetRids: string[]): string {
  const sheets = wb.sheets.map((s, i) => `<sheet sheetId="${i + 1}" name="${esc(s.name)}" state="${s.hidden ? "hidden" : "visible"}" r:id="${sheetRids[i]}"/>`).join("");
  const names = wb.sheets.flatMap((s, i) => printNames(s.page, s.name, i));
  return (
    XML_DECL +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">' +
    '<fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9303"/><workbookPr defaultThemeVersion="164011" filterPrivacy="1"/>' +
    `<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="15000" firstSheet="0" activeTab="${wb.activeIndex}"/></bookViews>` +
    `<sheets>${sheets}</sheets>` +
    (names.length ? `<definedNames>${names.join("")}</definedNames>` : "") +
    '<calcPr calcId="171027" fullCalcOnLoad="1"/>' +
    "</workbook>"
  );
}

/** Kitabı .xlsx (ya da makro taşınacaksa .xlsm) baytlarına yazar. */
export async function writeXlsxPackage(wb: WorkbookModel, opts: XlsxWriterOptions = {}): Promise<Uint8Array<ArrayBuffer>> {
  let t = performance.now();
  const evre = (ad: string) => {
    const n = performance.now();
    if (opts.sureler) opts.sureler[ad] = Math.round(n - t);
    t = n;
  };
  wb.engine.recalc();
  const checkbox = new Set<number>();
  wb.styles.all().forEach((s, i) => {
    if (s.checkbox) checkbox.add(i);
  });
  const ctx: WriteContext = { wb, styles: new StyleSheet(wb), sst: new SharedStrings(), media: [], tables: 0, drawings: 0, dynamic: false, checkbox };
  const writers = wb.sheets.map((s, i) => new SheetWriter(ctx, s, i));
  for (const w of writers) w.plan();
  const hasStrings = writers.some((w) => w.hasStrings);

  // 1) Bellekteki paket (küçük parçalar) ve sayfa iskeletleri
  const zip = new JSZip();
  const wbRels = new Rels();
  wbRels.add(REL.styles, "styles.xml");
  wbRels.add(REL.theme, "theme/theme1.xml");
  if (hasStrings) wbRels.add(REL.sharedStrings, "sharedStrings.xml");
  const sheetRids = writers.map((w) => wbRels.add(REL.worksheet, w.path.slice("xl/".length)));
  const mediaExts: string[] = [];
  for (const m of ctx.media) if (!mediaExts.includes(m.ext)) mediaExts.push(m.ext);
  zip.file(
    "[Content_Types].xml",
    contentTypesXml({
      media: mediaExts,
      sheets: writers.map((w) => w.path),
      sharedStrings: hasStrings,
      tables: writers.flatMap((w) => w.tables.map((t) => t.path)),
      drawings: writers.flatMap((w) => (w.drawing ? [w.drawing.path] : [])),
      comments: writers.flatMap((w) => (w.comments ? [w.comments.path] : [])),
    }),
  );
  zip.file(
    "_rels/.rels",
    relsXml([
      { id: "rId1", type: REL.officeDocument, target: "xl/workbook.xml" },
      { id: "rId2", type: REL.coreProperties, target: "docProps/core.xml" },
      { id: "rId3", type: REL.extendedProperties, target: "docProps/app.xml" },
    ]),
  );
  zip.file("xl/_rels/workbook.xml.rels", wbRels.xml());
  zip.file("xl/workbook.xml", workbookXml(wb, sheetRids));
  zip.file("xl/theme/theme1.xml", themeXml(wb.theme));
  zip.file("docProps/app.xml", appXml(wb));
  zip.file("docProps/core.xml", coreXml(wb));
  const parts = new PackageParts(zip);
  const paths = new Map<string, string>();
  const own = new Set<string>();
  for (const w of writers) {
    paths.set(w.sheet.name, w.path);
    own.add(w.path);
    parts.seed(w.path, w.frame);
    zip.file(w.path, w.frame.toString());
    if (w.rels.size) zip.file(`xl/worksheets/_rels/${w.path.slice("xl/worksheets/".length)}.rels`, w.rels.xml());
    if (w.comments) {
      zip.file(w.comments.path, w.comments.xml);
      zip.file(w.comments.vmlPath, w.comments.vml);
    }
    if (w.drawing) {
      zip.file(w.drawing.path, w.drawing.xml);
      zip.file(`xl/drawings/_rels/${w.drawing.path.slice("xl/drawings/".length)}.rels`, w.drawing.rels.xml());
    }
    for (const t of w.tables) zip.file(t.path, t.xml);
  }
  for (const m of ctx.media) zip.file(m.path, m.data);
  evre("plan");

  // 2) Sonradan işlem (sayfa verisine dokunmadan)
  const src = xlsxSource(wb);
  const pctx: PassthroughContext = { src, carry: new Carry(), partMap: [], warnings: [] };
  await postProcessParts(zip, wb, paths, parts, pctx, { ownSheetData: true });
  parts.flush((p) => own.has(p));
  evre("sonradan");

  // 3) Sayfa verisi (akıtılarak sıkıştırılır), paylaşılan metinler, stiller
  // Sıkıştırıcı sayfanın sonunu işlerken (Node'da ayrı iş parçacığında) paylaşılan metinler ve
  // stiller hazırlanır: sayfaların bitişi en sonda beklenir
  const packed = new Map<string, PackedEntry>();
  const pending: Promise<unknown>[] = [];
  for (const w of writers) {
    const sink = new DeflateSink();
    await w.stream(sink);
    pending.push(sink.finish().then((e) => packed.set(w.path, e)));
  }
  evre("sayfalar");
  const rel = async (path: string) => (await zip.file(path)?.async("string")) ?? "";
  let ct = await rel("[Content_Types].xml");
  let wbRelsXml = await rel("xl/_rels/workbook.xml.rels");
  if (ctx.sst.count || hasStrings) {
    pending.push(ctx.sst.pack().then((e) => packed.set("xl/sharedStrings.xml", e)));
    if (!hasStrings) {
      ct = addOverride(ct, "/xl/sharedStrings.xml", CT.sharedStrings);
      wbRelsXml = appendRel(wbRelsXml, REL.sharedStrings, "sharedStrings.xml");
    }
  }
  if (ctx.dynamic && !zip.file("xl/metadata.xml")) {
    zip.file("xl/metadata.xml", METADATA_XML);
    ct = addOverride(ct, "/xl/metadata.xml", CT.sheetMetadata);
    if (!wbRelsXml.includes(REL.sheetMetadata)) wbRelsXml = appendRel(wbRelsXml, REL.sheetMetadata, "metadata.xml");
  }
  if (ctx.styles.hasBoxed) {
    zip.file(FPB_PATH, FPB_XML);
    ct = addOverride(ct, `/${FPB_PATH}`, CT.featurePropertyBag);
    if (!wbRelsXml.includes(REL.featurePropertyBag)) wbRelsXml = appendRel(wbRelsXml, REL.featurePropertyBag, "featurePropertyBag/featurePropertyBag.xml");
  }
  pending.push(packText(ctx.styles.xml()).then((e) => packed.set("xl/styles.xml", e)));
  await Promise.all(pending);
  evre("sikistirma");
  zip.file("[Content_Types].xml", ct);
  zip.file("xl/_rels/workbook.xml.rels", wbRelsXml);

  // 4) Paket: içerik türleri ilk parça; sayfalar ham (zaten sıkıştırılmış) eklenir
  const zw = new ZipWriter();
  const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir && !own.has(n));
  const first = ["[Content_Types].xml", "_rels/.rels", "xl/workbook.xml", "xl/_rels/workbook.xml.rels"];
  const ordered = [...first.filter((n) => names.includes(n)), ...names.filter((n) => !first.includes(n))];
  for (const n of ordered) {
    zw.add(n, await zip.file(n)!.async("uint8array"));
    if (n === "xl/_rels/workbook.xml.rels") {
      addPacked(zw, "xl/styles.xml", packed.get("xl/styles.xml")!);
      for (const w of writers) addPacked(zw, w.path, packed.get(w.path)!);
      const sst = packed.get("xl/sharedStrings.xml");
      if (sst) addPacked(zw, "xl/sharedStrings.xml", sst);
    }
  }
  let bytes: Uint8Array = await zw.finish();
  evre("paket");
  if (src || opts.macro) {
    const r = await carryParts({ app: "tablo", source: src ?? null, output: bytes, partMap: pctx.partMap, carry: pctx.carry, macro: !!opts.macro, overrides: { own: ["tableSingleCells"] } });
    bytes = r.bytes;
    for (const d of r.report.dropped) {
      if (d.reason === "macro") pctx.warnings.push("makrolar (.xlsx biçiminde saklanamaz; korumak için .xlsm olarak kaydedin)");
      else if (d.reason === "comments") pctx.warnings.push("yorum yanıt zincirleri (notlar korunur)");
    }
  }
  if (opts.report) opts.report.warnings = [...new Set(pctx.warnings)];
  evre("tasima");
  return new Uint8Array(bytes);
}
