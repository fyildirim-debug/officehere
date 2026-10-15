/**
 * Geçiş katmanı (P3.3): modelin tanımadığı parça ve öğelerin kaydederken korunması, yapısal
 * değişiklikte başvuruların kaydırılması, makrolu kitap, dijital imza, ikinci tur kararlılığı.
 * POI dosyaları: test-derlem/poi (Apache-2.0, NOTICE orada).
 */
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import JSZip from "jszip";
import { OpcPackage } from "@officehere/core/opc";
import { Controller } from "../src/controller";
import { openBook, writeBook } from "../src/io/formats";
import { readXlsx, writeXlsx, type XlsxWriteReport } from "../src/io/xlsx";
import { writeOds } from "../src/io/ods/writer";
import { serializeDraftSync, workbookFromDraft } from "../src/io/draft";
import { WorkbookModel } from "../src/model/workbook";
import { emptyPassthrough } from "../src/model/passthrough";

const poi = (name: string): ArrayBuffer => {
  const b = readFileSync(new URL(`../../../test-derlem/poi/${name}`, import.meta.url));
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer;
};
const ab = (u: Uint8Array): ArrayBuffer => u.buffer.slice(u.byteOffset, u.byteOffset + u.byteLength) as ArrayBuffer;
const SLOW = { timeout: 60_000 };

async function text(bytes: Uint8Array | ArrayBuffer, part: string): Promise<string> {
  return (await (await JSZip.loadAsync(bytes)).file(part)?.async("string")) ?? "";
}

/** Paket tutarlılığı: her iç ilişkinin hedefi var, her parçanın içerik türü var. */
async function opcErrors(bytes: Uint8Array): Promise<string[]> {
  const pkg = await OpcPackage.open(bytes);
  const out: string[] = [];
  for (const name of pkg.names) {
    if (name.endsWith(".rels") || name === "[Content_Types].xml") continue;
    if (!pkg.contentType(name)) out.push(`tür yok: ${name}`);
  }
  for (const src of ["", ...pkg.names.filter((n) => !n.endsWith(".rels"))]) {
    for (const r of await pkg.rels(src)) if (!r.external && r.part && !pkg.has(r.part) && !r.target.startsWith("#")) out.push(`hedef yok: ${src} → ${r.target}`);
  }
  return out;
}

const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const XM = "http://schemas.microsoft.com/office/excel/2006/main";

/** Başka sayfaya başvuran x14 veri doğrulaması ve özel simge kümesi (x14) içeren kitap. */
async function x14Book(): Promise<ArrayBuffer> {
  const wb = new WorkbookModel();
  const s = wb.sheets[0];
  wb.addSheet("Liste");
  const l = wb.sheets[1];
  for (let i = 0; i < 5; i++) wb.setInputRaw(l, i, 0, `Seçenek ${i + 1}`);
  for (let i = 0; i < 12; i++) wb.setInputRaw(s, i, 1, i * 10);
  const bytes = await writeXlsx(wb);
  const zip = await JSZip.loadAsync(bytes);
  const path = "xl/worksheets/sheet1.xml";
  const xml = await zip.file(path)!.async("string");
  const ext =
    `<extLst>` +
    `<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="${X14}"><x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="${XM}"><x14:cfRule type="iconSet" priority="1" id="{F6A45A70-DEA9-40B0-83E9-E53A5F349BE7}"><x14:iconSet custom="1"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>30</xm:f></x14:cfvo><x14:cfIcon iconSet="3Signs" iconId="0"/><x14:cfIcon iconSet="3Flags" iconId="1"/><x14:cfIcon iconSet="3Symbols2" iconId="2"/></x14:iconSet></x14:cfRule><xm:sqref>B1:B12</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext>` +
    `<ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="${X14}"><x14:dataValidations count="1" xmlns:xm="${XM}"><x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1"><x14:formula1><xm:f>Liste!$A$1:$A$5</xm:f></x14:formula1><xm:sqref>D3:D10</xm:sqref></x14:dataValidation></x14:dataValidations></ext>` +
    `</extLst>`;
  zip.file(path, xml.replace(/<extLst>[\s\S]*?<\/extLst>|(?=<\/worksheet>)/, ext));
  return ab(await zip.generateAsync({ type: "uint8array" }));
}

describe("Geçiş katmanı: yapısal değişiklikte başvurular (birim)", () => {
  function book() {
    const wb = new WorkbookModel();
    const s = wb.sheets[0];
    const l = wb.addSheet("Liste");
    const pt = emptyPassthrough();
    pt.dv.push({ label: "dv", xml: `<x14:dataValidation xmlns:x14="${X14}" xmlns:xm="${XM}" type="list"><x14:formula1><xm:f>Liste!$A$1:$A$5</xm:f></x14:formula1><xm:sqref>A3:A10 C1</xm:sqref></x14:dataValidation>` });
    pt.cf.push({ label: "cf", xml: `<x14:conditionalFormatting xmlns:x14="${X14}" xmlns:xm="${XM}"><x14:cfRule type="expression" priority="1" id="{1}"><xm:f>$B$5&gt;0</xm:f></x14:cfRule><xm:sqref>B5:B9</xm:sqref></x14:conditionalFormatting>` });
    pt.children.push({ name: "protectedRanges", label: "korumalı aralıklar", xml: `<protectedRanges><protectedRange sqref="A2:B4" name="a"/><protectedRange sqref="E6" name="b"/></protectedRanges>` });
    pt.drawing = { ns: {}, ignorable: [], anchors: [{ label: "şekil", xml: `<xdr:twoCellAnchor xmlns:xdr="x"><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp textlink="$C$7"/></xdr:twoCellAnchor>` }] };
    pt.vml = { ns: {}, ignorable: [], shapetypes: [], shapes: [{ label: "ole", xml: `<v:shape id="_x0000_s1025"><x:ClientData ObjectType="Pict"><x:Anchor>\n    0, 0, 4, 0, 5, 1, 10, 1</x:Anchor></x:ClientData></v:shape>` }] };
    s.passthrough = pt;
    return { wb, s, l, pt };
  }

  it("satır eklenince sqref, formül, çapa ve VML konumu kayar", () => {
    const { wb, s, pt } = book();
    // 3. satırın önüne 3 satır (indeks 2)
    s.shiftStructure("row", 2, 3);
    wb.shiftDataRefs(s, "row", 2, 3);
    expect(pt.structVersion).toBe(s.structVersion);
    expect(pt.dv[0].xml).toContain("<xm:sqref>A6:A13 C1</xm:sqref>");
    expect(pt.dv[0].xml).toContain("Liste!$A$1:$A$5");
    expect(pt.cf[0].xml).toContain("<xm:sqref>B8:B12</xm:sqref>");
    expect(pt.cf[0].xml).toContain("<xm:f>$B$8&gt;0</xm:f>");
    expect(pt.children[0].xml).toContain('sqref="A2:B7"');
    expect(pt.children[0].xml).toContain('sqref="E9"');
    expect(pt.drawing!.anchors[0].xml).toMatch(/<xdr:from>.*<xdr:row>7<\/xdr:row>.*<\/xdr:from>/);
    expect(pt.drawing!.anchors[0].xml).toMatch(/<xdr:to>.*<xdr:row>11<\/xdr:row>.*<\/xdr:to>/);
    expect(pt.drawing!.anchors[0].xml).toContain('textlink="$C$10"');
    expect(pt.vml!.shapes[0].xml).toContain("0, 0, 7, 0, 5, 1, 13, 1");
    expect(pt.dropped).toEqual([]);
  });

  it("sütun silmede aralığı tümüyle silinen öğe düşürülür ve etiketi kalır", () => {
    const { wb, s, pt } = book();
    s.shiftStructure("col", 4, -1);
    wb.shiftDataRefs(s, "col", 4, -1);
    // E6 silindi: kapsayıcıda tek korumalı aralık kalır
    expect(pt.children[0].xml).not.toContain('name="b"');
    expect(pt.children[0].xml).toContain('name="a"');
    s.shiftStructure("col", 0, -3);
    wb.shiftDataRefs(s, "col", 0, -3);
    expect(pt.cf).toHaveLength(0);
    expect(pt.dropped).toContain("cf");
  });

  it("düşürülen öğeler kaydederken Türkçe uyarı olarak bildirilir", SLOW, async () => {
    const { wb, s } = book();
    s.shiftStructure("col", 0, -3);
    wb.shiftDataRefs(s, "col", 0, -3);
    const report: XlsxWriteReport = { warnings: [] };
    await writeXlsx(wb, { report });
    expect(report.warnings).toContain("cf");
  });

  it("başka sayfada ekleme, sayfa adlandırma ve silme formüllere yansır", () => {
    const { wb, l, pt } = book();
    l.shiftStructure("row", 0, 1);
    wb.shiftDataRefs(l, "row", 0, 1);
    expect(pt.dv[0].xml).toContain("Liste!$A$2:$A$6");
    expect(pt.dv[0].xml).toContain("<xm:sqref>A3:A10 C1</xm:sqref>");
    wb.renameSheet(1, "Kaynak");
    expect(pt.dv[0].xml).toContain("Kaynak!$A$2:$A$6");
    wb.removeSheet(1);
    expect(pt.dv).toHaveLength(0);
    expect(pt.dropped).toContain("dv");
  });

  it("geri alma geçiş verisini de geri yükler", () => {
    const ctl = new Controller();
    const { wb, s } = book();
    ctl.load(wb, "a.xlsx", null);
    ctl.ops.insertRows(s, 0, 2);
    expect(s.passthrough!.dv[0].xml).toContain("A5:A12 C3");
    ctl.ops.undo();
    const back = ctl.wb.sheets[0].passthrough!;
    expect(back.dv[0].xml).toContain("<xm:sqref>A3:A10 C1</xm:sqref>");
    expect(back.structVersion).toBe(ctl.wb.sheets[0].structVersion);
  });
});

describe("Geçiş katmanı: x14 uzantıları", () => {
  it("x14 veri doğrulama ve özel simge kümesi iki turda korunur", SLOW, async () => {
    const src = await x14Book();
    const a = await readXlsx(src);
    const out1 = await writeXlsx(a);
    const b = await readXlsx(ab(out1));
    const out2 = await writeXlsx(b);
    for (const out of [out1, out2]) {
      const xml = await text(out, "xl/worksheets/sheet1.xml");
      expect(xml.match(/<x14:dataValidation\b/g)).toHaveLength(1);
      expect(xml).toContain("Liste!$A$1:$A$5");
      expect(xml.match(/<x14:cfIcon\b/g)).toHaveLength(3);
      expect(xml.match(/<x14:conditionalFormatting\b/g)).toHaveLength(1);
      expect(await opcErrors(out)).toEqual([]);
    }
    // Özel simge kümesinin modeldeki yaklaşık kuralı ayrıca yazılmaz
    expect((await text(out2, "xl/worksheets/sheet1.xml")).match(/<cfRule\b[^>]*type="iconSet"/g) ?? []).toHaveLength(0);
  });

  it("satır eklendikten sonra kaydedince x14 sqref'leri kayar", SLOW, async () => {
    const ctl = new Controller();
    ctl.load(await readXlsx(await x14Book()), "a.xlsx", null);
    ctl.ops.insertRows(ctl.wb.sheets[0], 0, 2);
    const xml = await text(await writeXlsx(ctl.wb), "xl/worksheets/sheet1.xml");
    expect(xml).toContain("<xm:sqref>D5:D12</xm:sqref>");
    expect(xml).toContain("<xm:sqref>B3:B14</xm:sqref>");
  });

  it("modeldeki doğrulama aynı hücrelerde x14 doğrulamanın yerine geçer", SLOW, async () => {
    const wb = await readXlsx(await x14Book());
    wb.sheets[0].validations.push({ id: "v9", ranges: [{ r1: 2, c1: 3, r2: 5, c2: 3 }], rule: { type: "whole", operator: "greaterThan", formula1: "0" } });
    const xml = await text(await writeXlsx(wb), "xl/worksheets/sheet1.xml");
    expect(xml).toContain("<xm:sqref>D7:D10</xm:sqref>");
  });
});

describe("Geçiş katmanı: POI dosyaları", () => {
  it(".xlsm → .xlsm: makro projesi ve kod adları korunur; .xlsx'te düşürülür ve uyarılır", SLOW, async () => {
    const o = await openBook("SimpleMacro.xlsm", poi("SimpleMacro.xlsm"));
    expect(o.format).toBe("xlsm");
    expect(o.notice).toMatch(/makrolar kaydederken korunur, ancak OfficeHere.da çalıştırılmaz/);
    const m = (await writeBook(o.wb, o.wb.sheets[0], "xlsm")) as Uint8Array;
    const zip = await JSZip.loadAsync(m);
    expect(zip.file("xl/vbaProject.bin")).toBeTruthy();
    expect(await text(m, "[Content_Types].xml")).toContain("application/vnd.ms-excel.sheet.macroEnabled.main+xml");
    expect(await text(m, "xl/workbook.xml")).toMatch(/<workbookPr\b[^>]*codeName="ThisWorkbook"/);
    expect(await text(m, "xl/worksheets/sheet1.xml")).toMatch(/<sheetPr\b[^>]*codeName="Sheet1"/);
    expect(await opcErrors(m)).toEqual([]);
    const report: XlsxWriteReport = { warnings: [] };
    const x = (await writeBook(o.wb, o.wb.sheets[0], "xlsx", undefined, report)) as Uint8Array;
    expect((await JSZip.loadAsync(x)).file("xl/vbaProject.bin")).toBeNull();
    expect(await text(x, "[Content_Types].xml")).not.toContain("macroEnabled");
    expect(report.warnings.join(" ")).toMatch(/makrolar/);
  });

  it("workbookPr/calcPr'de yalnız Transitional şema öznitelikleri taşınır", SLOW, async () => {
    const b = readFileSync(new URL("./fixtures/openpyxl/karmasik_lo.xlsx", import.meta.url));
    const o = await openBook("karmasik_lo.xlsx", ab(b));
    const out = (await writeBook(o.wb, o.wb.sheets[0], "xlsx")) as Uint8Array;
    const wb = await text(out, "xl/workbook.xml");
    expect(wb).toMatch(/<workbookPr\b[^>]*backupFile="false"/);
    expect(wb).toMatch(/<calcPr\b[^>]*iterateCount="100"/);
    expect(wb).not.toContain("dateCompatibility");
  });

  it("dijital imza: açılışta uyarılır (otomatik kayıt bekler), kaydedince kaldırılır", SLOW, async () => {
    const o = await openBook("hello-world-signed.xlsx", poi("hello-world-signed.xlsx"));
    expect(o.notice).toMatch(/imza/);
    expect(o.wb.loadWarnings.join(" ")).toMatch(/dijital imza/);
    const out = (await writeBook(o.wb, o.wb.sheets[0], "xlsx")) as Uint8Array;
    const names = Object.keys((await JSZip.loadAsync(out)).files);
    expect(names.some((n) => /signature|origin\.sigs/i.test(n))).toBe(false);
    expect(await opcErrors(out)).toEqual([]);
  });

  it("OLE nesneleri ve VML görünümleri taşınır; satır eklenince çapalar kayar", SLOW, async () => {
    const ctl = new Controller();
    const o = await openBook("WithEmbeded.xlsx", poi("WithEmbeded.xlsx"));
    ctl.load(o.wb, "WithEmbeded.xlsx", null);
    ctl.ops.insertRows(ctl.wb.sheets[0], 0, 1);
    const out = await writeXlsx(ctl.wb);
    const sheet = await text(out, "xl/worksheets/sheet1.xml");
    expect(sheet.match(/<oleObject\b/g)).toHaveLength(2);
    const vmlPath = Object.keys((await JSZip.loadAsync(out)).files).find((n) => n.endsWith(".vml"))!;
    const vml = await text(out, vmlPath);
    expect(vml).toContain("0, 0, 5, 0, 5, 1, 11, 1");
    // oleObject shapeId'leri VML şekil kimlikleriyle eşleşir
    for (const m of sheet.matchAll(/shapeId="(\d+)"/g)) expect(vml).toContain(`id="_x0000_s${m[1]}"`);
    const zip = await JSZip.loadAsync(out);
    expect(Object.keys(zip.files).filter((n) => n.startsWith("xl/embeddings/"))).toHaveLength(2);
    expect(await opcErrors(out)).toEqual([]);
  });

  it("çizim metin kutusu, XML eşlemeleri, bağlantılar ve grafik sayfası taşınır", SLOW, async () => {
    const cases: [string, RegExp][] = [
      ["WithTextBox.xlsx", /^xl\/drawings\/drawing\d+\.xml$/],
      ["CustomXMLMappings.xlsx", /^xl\/xmlMaps\.xml$/],
      ["ExcelPivotTableSample.xlsx", /^xl\/connections\.xml$/],
      ["SimpleScatterChart.xlsx", /^xl\/chartsheets\/.*\.xml$/],
    ];
    for (const [name, re] of cases) {
      const o = await openBook(name, poi(name));
      const out = (await writeBook(o.wb, o.wb.sheets[0], "xlsx")) as Uint8Array;
      expect(Object.keys((await JSZip.loadAsync(out)).files).some((n) => re.test(n)), name).toBe(true);
      expect(await opcErrors(out), name).toEqual([]);
      // Yeniden açılır
      expect((await openBook(name, ab(out))).wb.sheets.length).toBeGreaterThan(0);
    }
    const sc = await openBook("SimpleScatterChart.xlsx", poi("SimpleScatterChart.xlsx"));
    const wbXml = await text((await writeBook(sc.wb, sc.wb.sheets[0], "xlsx")) as Uint8Array, "xl/workbook.xml");
    expect(wbXml).toMatch(/<sheet\b[^>]*name="Chart1"/);
  });

  it("veri doğrulamaları ikinci turda değişmez; hücre başvurulu formüller korunur", SLOW, async () => {
    const o = await openBook("DataValidationEvaluations.xlsx", poi("DataValidationEvaluations.xlsx"));
    const count = (wb: WorkbookModel) => wb.sheets[0].validations.reduce((n, v) => n + v.ranges.length, 0);
    expect(count(o.wb)).toBe(17);
    expect(o.wb.sheets[0].validations.some((v) => v.rule.formula1 === "=E1")).toBe(true);
    const out1 = await writeXlsx(o.wb);
    const b = await readXlsx(ab(out1));
    expect(count(b)).toBe(17);
    const out2 = await writeXlsx(b);
    const dv = async (u: Uint8Array) => /<dataValidations[\s\S]*?<\/dataValidations>/.exec(await text(u, "xl/worksheets/sheet1.xml"))?.[0];
    expect(await dv(out2)).toBe(await dv(out1));
    expect(await dv(out1)).not.toContain("NaN");
  });

  it("yapılandırılmış başvurulu tablo adı (\\_Prime.1) korunur, ikinci tur kararlı", SLOW, async () => {
    const o = await openBook("StructuredReferences.xlsx", poi("StructuredReferences.xlsx"));
    const out1 = await writeXlsx(o.wb);
    const tables = Object.keys((await JSZip.loadAsync(out1)).files).filter((n) => /^xl\/tables\/table\d+\.xml$/.test(n));
    expect((await Promise.all(tables.map((t) => text(out1, t)))).join("")).toContain('displayName="\\_Prime.1"');
    const b = await readXlsx(ab(out1));
    const f = b.sheets.find((s) => s.name === "Formulas")!;
    expect(b.getValue(f, 0, 0)).toBe(o.wb.getValue(o.wb.sheets.find((s) => s.name === "Formulas")!, 0, 0));
    const out2 = await writeXlsx(b);
    const meta = async (u: Uint8Array) => !!(await JSZip.loadAsync(u)).file("xl/metadata.xml");
    expect(await meta(out2)).toBe(await meta(out1));
  });

  it("kopyalanan sayfa, kaynaksız taslak ve biçim değişimi paketi bozmaz", SLOW, async () => {
    const ctl = new Controller();
    const o = await openBook("WithEmbeded.xlsx", poi("WithEmbeded.xlsx"));
    ctl.load(o.wb, "WithEmbeded.xlsx", null);
    ctl.ops.copySheet(0, 1);
    const out = await writeXlsx(ctl.wb);
    expect(await opcErrors(out)).toEqual([]);
    expect((await text(out, "xl/worksheets/sheet2.xml")).match(/<oleObject\b/g)).toHaveLength(2);
    // Kurtarılan taslakta kaynak paket yok: ilişkili öğeler yazılmaz ve uyarılır
    const draft = workbookFromDraft(serializeDraftSync(o.wb));
    const report: XlsxWriteReport = { warnings: [] };
    const d = await writeXlsx(draft, { report });
    expect(await opcErrors(d)).toEqual([]);
    expect(report.warnings.join(" ")).toMatch(/gömülü nesne/);
    // .ods'ye kaydetmede taşıma yapılmaz
    expect((await writeOds(o.wb)).byteLength).toBeGreaterThan(0);
  });
});
