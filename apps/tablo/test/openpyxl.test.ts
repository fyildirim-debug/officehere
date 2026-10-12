/** Başka araçlarla (openpyxl, LibreOffice) üretilmiş XLSX dosyalarının içe aktarılması. */
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { cellKey } from "@officehere/formula";
import JSZip from "jszip";
import { readXlsx, writeXlsx } from "../src/io/xlsx";
import { normalizeRelsXml, prefixDefaultNamespaces, relativeTarget } from "../src/io/xlsxPackage";
import type { WorkbookModel } from "../src/model/workbook";

const dir = new URL("./fixtures/openpyxl/", import.meta.url);
function load(name: string): Promise<WorkbookModel> {
  const b = readFileSync(new URL(name, dir));
  return readXlsx(b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer);
}
const sheet = (wb: WorkbookModel, name: string) => wb.sheetByName(name)!;

describe("paket normalleştirme", () => {
  it("mutlak ilişki hedefleri göreli olur", () => {
    expect(relativeTarget("xl/worksheets/sheet1.xml", "/xl/drawings/drawing1.xml")).toBe("../drawings/drawing1.xml");
    expect(relativeTarget("xl/workbook.xml", "/xl/worksheets/sheet1.xml")).toBe("worksheets/sheet1.xml");
    const xml = '<Relationships><Relationship Id="rId1" Type="t" Target="/xl/tables/table1.xml"/><Relationship Id="rId2" Type="h" Target="https://a.b" TargetMode="External"/></Relationships>';
    const out = normalizeRelsXml("xl/worksheets/_rels/sheet1.xml.rels", xml);
    expect(out).toContain('Target="../tables/table1.xml"');
    expect(out).toContain('Target="https://a.b"');
    expect(normalizeRelsXml("_rels/.rels", '<Relationship Id="r" Type="t" Target="/xl/workbook.xml"/>')).toContain('Target="xl/workbook.xml"');
  });
  it("öneksiz DrawingML öğelerine önek eklenir", () => {
    const x = '<wsDr xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><oneCellAnchor><from><col>3</col></from><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:t>x</a:t></a:graphic><clientData/></oneCellAnchor></wsDr>';
    const y = prefixDefaultNamespaces(x);
    expect(y).toContain('<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">');
    expect(y).toContain("<xdr:from><xdr:col>3</xdr:col></xdr:from>");
    expect(y).toContain("<a:t>x</a:t>");
    expect(y).toContain("<xdr:clientData/>");
    expect(y).toContain("</xdr:oneCellAnchor></xdr:wsDr>");
  });
});

describe("kabul testi dosyaları (openpyxl)", () => {
  it("chart_only.xlsx: grafik okunur", async () => {
    const wb = await load("chart_only.xlsx");
    const s = wb.sheets[0];
    expect(s.charts).toHaveLength(1);
    expect(s.charts[0].type).toBe("column");
    expect(s.charts[0].title).toBe("Satış");
    expect(wb.loadWarnings).toEqual([]);
  });
  it("table_only.xlsx: tablo okunur", async () => {
    const wb = await load("table_only.xlsx");
    expect(wb.sheets[0].tables.map((t) => t.name)).toEqual(["Tablo1"]);
  });
  it("karmasik_nochart.xlsx: not, tablo, birleştirme, koşullu biçim, ad", async () => {
    const wb = await load("karmasik_nochart.xlsx");
    const b = sheet(wb, "Bütçe");
    expect(b.noteAt(2, 0)?.text).toContain("Kira sözleşmesi");
    expect(b.merges).toHaveLength(1);
    expect(b.condFormats.map((c) => c.type).sort()).toEqual(["cellIs", "colorScale", "dataBar"]);
    expect(sheet(wb, "Satışlar").tables[0].name).toBe("SatisTablosu");
    expect(wb.engine.getNames().map((n) => n.name)).toContain("Toplam_Tutar");
    expect(wb.getValue(b, 9, 2)).toBeGreaterThan(30000);
  });
  it("karmasik.xlsx: grafikler ve notlar birlikte", async () => {
    const wb = await load("karmasik.xlsx");
    const b = sheet(wb, "Bütçe");
    expect(b.charts.map((c) => c.type)).toEqual(["column", "pie"]);
    expect(sheet(wb, "Satışlar").charts.map((c) => c.type)).toEqual(["line"]);
    expect(b.noteAt(2, 0)?.text).toContain("Kira");
    expect(wb.loadWarnings).toEqual([]);
  });
});

describe("çeşitli openpyxl dosyaları", () => {
  it("açıklamalar düz <t> metniyle okunur", async () => {
    const wb = await load("aciklama.xlsx");
    const s = wb.sheets[0];
    expect(s.notes.get(cellKey(1, 0))?.text).toBe("İlk ay notu");
    expect(s.notes.get(cellKey(2, 1))?.text).toBe("Çok satırlı\nnot metni");
  });
  it("tablo ve tablo formülü", async () => {
    const wb = await load("tablo.xlsx");
    const s = wb.sheets[0];
    expect(s.tables[0].name).toBe("Satislar");
    expect(s.tables[0].ref).toEqual({ r1: 0, c1: 0, r2: 6, c2: 2 });
    expect(wb.getValue(s, 1, 4)).toBe(100 + 110 + 120 + 130 + 140 + 150);
  });
  it("beş grafik türü iki sayfada", async () => {
    const wb = await load("grafikler.xlsx");
    expect(sheet(wb, "Grafik").charts.map((c) => c.type)).toEqual(["column", "line", "pie", "area"]);
    expect(sheet(wb, "Dağılım").charts.map((c) => c.type)).toEqual(["scatter"]);
  });
  it("birleştirmeler ve formül", async () => {
    const wb = await load("birlesik.xlsx");
    const s = wb.sheets[0];
    expect(s.merges).toEqual(expect.arrayContaining([{ r1: 0, c1: 0, r2: 0, c2: 3 }, { r1: 2, c1: 0, r2: 5, c2: 0 }]));
    expect(wb.getValue(s, 4, 1)).toBe(3);
  });
  it("tanımlı adlar (aralık ve sabit)", async () => {
    const wb = await load("adlar.xlsx");
    const s = wb.sheets[0];
    const names = wb.engine.getNames().map((n) => n.name);
    expect(names).toEqual(expect.arrayContaining(["SatisAraligi", "Oran"]));
    expect(wb.getValue(s, 0, 4)).toBe(100 + 110 + 120 + 130 + 140 + 150);
    expect(wb.getValue(s, 1, 4)).toBeCloseTo(750 * 0.18);
  });
  it("koşullu biçimler (simge kümesi dahil)", async () => {
    const wb = await load("kosullu.xlsx");
    const types = wb.sheets[0].condFormats.map((c) => c.type);
    expect(types).toEqual(expect.arrayContaining(["cellIs", "colorScale", "dataBar", "expression", "iconSet"]));
  });
  it("hepsi bir arada", async () => {
    const wb = await load("hepsi.xlsx");
    const s = wb.sheets[0];
    expect(s.charts).toHaveLength(1);
    expect(s.tables).toHaveLength(1);
    expect(s.notes.size).toBe(1);
    expect(s.merges).toHaveLength(1);
    expect(wb.loadWarnings).toEqual([]);
  });
  it("bozuk açıklama parçası tüm dosyayı reddettirmez; atlanan öğe bildirilir", async () => {
    const wb = await load("bozuk_aciklama.xlsx");
    const s = wb.sheets[0];
    expect(s.charts).toHaveLength(1);
    expect(s.tables).toHaveLength(1);
    expect(wb.getValue(s, 1, 1)).toBe(100);
    expect(wb.loadWarnings.join(" ")).toContain("açıklamalar");
  });
  it("açılan dosya yeniden kaydedilip açılabilir (adlar korunur)", async () => {
    const wb = await load("adlar.xlsx");
    const bytes = await writeXlsx(wb);
    const zip = await JSZip.loadAsync(bytes);
    const wbXml = await zip.file("xl/workbook.xml")!.async("string");
    expect(wbXml).toContain('<definedName name="Oran">0.18</definedName>');
    const again = await readXlsx(bytes.buffer);
    expect(again.engine.getNames().map((n) => n.name)).toEqual(expect.arrayContaining(["SatisAraligi", "Oran"]));
  });
});
