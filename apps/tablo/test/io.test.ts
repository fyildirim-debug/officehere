import { describe, expect, it } from "vitest";
import { writeFileSync } from "node:fs";
import { cellKey, isError } from "@officehere/formula";
import { WorkbookModel } from "../src/model/workbook";
import { readXlsx, writeXlsx, addXlfnPrefixes, compressCells, hfToFile, stripHF } from "../src/io/xlsx";
import { csvToInputs, detectDelimiter, decodeText, parseCsv, readCsv, writeCsv } from "../src/io/csv";
import { chartXml } from "../src/io/xlsxCharts";

const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFklEQVR42mP8z8Dwn4GBgYGJAQoAAAD//wMAAP//AwDsAD8AAAAASUVORK5CYII=";

function sample(): WorkbookModel {
  const wb = new WorkbookModel(true);
  const s = wb.addSheet("Satışlar");
  const o = wb.addSheet("Özet");
  const set = (sh = s, r: number, c: number, v: Parameters<WorkbookModel["setInputRaw"]>[3]) => wb.setInputRaw(sh, r, c, v);
  set(s, 0, 0, "Ürün");
  set(s, 0, 1, "Adet");
  set(s, 0, 2, "Fiyat");
  set(s, 0, 3, "Tutar");
  const items: [string, number, number][] = [
    ["Çay", 10, 12.5],
    ["Şeker", 4, 30],
    ["Ğofret", 7, 5.75],
  ];
  items.forEach(([n, q, p], i) => {
    set(s, i + 1, 0, n);
    set(s, i + 1, 1, q);
    set(s, i + 1, 2, p);
    set(s, i + 1, 3, { formula: `=B${i + 2}*C${i + 2}` });
  });
  set(s, 4, 0, "Toplam");
  set(s, 4, 3, { formula: "=SUM(D2:D4)" });
  set(s, 5, 0, true);
  set(s, 5, 1, 45366);
  set(s, 6, 0, { formula: "=XLOOKUP(\"Şeker\",A2:A4,D2:D4)" });
  set(s, 7, 0, { formula: "=SEQUENCE(3)" });
  set(s, 8, 3, { formula: "=1/0" });
  set(o, 0, 0, "Genel Toplam");
  set(o, 0, 1, { formula: "='Satışlar'!D5*2" });
  // Stiller
  const bold = wb.styles.intern({ bold: true, fill: "#C6EFCE", color: "#006100", bBottom: { style: "medium", color: "#000000" }, hAlign: "center" });
  for (let c = 0; c < 4; c++) wb.setStyleRaw(s, 0, c, bold);
  wb.setStyleRaw(s, 5, 1, wb.styles.intern({ numFmt: "dd.mm.yyyy" }));
  wb.setStyleRaw(s, 1, 2, wb.styles.intern({ numFmt: "#,##0.00 \"₺\"", italic: true, font: "Arial", size: 14, wrap: true, vAlign: "top" }));
  s.merges.push({ r1: 10, c1: 0, r2: 11, c2: 2 });
  set(s, 10, 0, "Birleşik");
  s.cols.set(0, 120);
  s.cols.setHidden(5, true);
  s.rows.set(0, 30);
  s.rows.setHidden(20, true);
  s.frozen = { rows: 1, cols: 1 };
  s.tabColor = "#E97132";
  s.autoFilter = { range: { r1: 0, c1: 0, r2: 3, c2: 3 }, columns: {} };
  s.notes.set(cellKey(0, 0), { text: "Ürün adı" });
  s.links.set(cellKey(12, 0), "https://example.com/");
  set(s, 12, 0, "Bağlantı");
  s.condFormats.push({ id: "cf1", type: "cellIs", operator: "greaterThan", values: ["=100"], ranges: [{ r1: 1, c1: 3, r2: 3, c2: 3 }], style: { fill: "#FFC7CE", color: "#9C0006" } });
  s.condFormats.push({ id: "cf2", type: "dataBar", color: "#638EC6", ranges: [{ r1: 1, c1: 1, r2: 3, c2: 1 }] });
  s.condFormats.push({ id: "cf3", type: "colorScale", colors: ["#F8696B", "#FFEB84", "#63BE7B"], ranges: [{ r1: 1, c1: 2, r2: 3, c2: 2 }] });
  s.validations.push({ id: "v1", ranges: [{ r1: 1, c1: 4, r2: 3, c2: 4 }], rule: { type: "list", formula1: "Evet;Hayır" } });
  s.validations.push({ id: "v2", ranges: [{ r1: 1, c1: 5, r2: 3, c2: 5 }], rule: { type: "whole", operator: "between", formula1: "1", formula2: "10" } });
  wb.engine.defineName("Fiyatlar", "='Satışlar'!$C$2:$C$4");
  o.tables.push({ id: "t1", name: "Tablo1", ref: { r1: 3, c1: 0, r2: 5, c2: 1 }, style: "TableStyleMedium2", headerRow: true, totalRow: false, bandedRows: true, bandedCols: false, firstColumn: false, lastColumn: false, filterButton: true, columns: ["Ad", "Değer"] });
  set(o, 3, 0, "Ad");
  set(o, 3, 1, "Değer");
  set(o, 4, 0, "x");
  set(o, 4, 1, 1);
  set(o, 5, 0, "y");
  set(o, 5, 1, 2);
  o.images.push({ id: "i1", name: "Resim 1", src: PNG, anchor: { row: 8, col: 3, dx: 0, dy: 0 }, width: 40, height: 30 });
  s.charts.push({ id: "c1", type: "column", range: "'Satışlar'!A1:B4", seriesInRows: false, title: "Satışlar", showTitle: true, showLegend: true, legendPos: "bottom", showDataLabels: false, showGridlines: true, palette: 0, style: 1, anchor: { row: 2, col: 6, dx: 0, dy: 0 }, width: 400, height: 240 });
  s.page = { ...s.page, orientation: "landscape", paper: "A4", printArea: { r1: 0, c1: 0, r2: 5, c2: 3 } };
  wb.props.title = "Deneme";
  wb.engine.recalc();
  return wb;
}

describe("XLSX gidiş-dönüş", () => {
  it("yaz → oku → karşılaştır", async () => {
    const wb = sample();
    const bytes = await writeXlsx(wb);
    expect(bytes.length).toBeGreaterThan(1000);
    if (process.env.TABLO_OUT) writeFileSync(process.env.TABLO_OUT, bytes);
    const back = await readXlsx(bytes.buffer);
    expect(back.sheets.map((s) => s.name)).toEqual(["Satışlar", "Özet"]);
    const s = back.sheets[0];
    const o = back.sheets[1];
    expect(back.getValue(s, 0, 0)).toBe("Ürün");
    expect(back.getValue(s, 2, 0)).toBe("Şeker");
    expect(back.getFormula(s, 1, 3)).toBe("=B2*C2");
    expect(back.getValue(s, 1, 3)).toBe(125);
    expect(back.getValue(s, 4, 3)).toBe(125 + 120 + 40.25);
    expect(back.getValue(s, 5, 0)).toBe(true);
    expect(back.getValue(s, 5, 1)).toBe(45366);
    expect(back.styleAt(s, 5, 1).numFmt).toBe("dd.mm.yyyy");
    expect(back.getFormula(s, 6, 0)).toBe('=XLOOKUP("Şeker",A2:A4,D2:D4)');
    expect(back.getValue(s, 6, 0)).toBe(120);
    expect(back.getFormula(s, 7, 0)).toBe("=SEQUENCE(3)");
    expect(back.getValue(s, 9, 0)).toBe(3);
    expect(isError(back.getValue(s, 8, 3))).toBe(true);
    expect(back.getFormula(o, 0, 1)).toBe("='Satışlar'!D5*2");
    expect(back.getValue(o, 0, 1)).toBe(570.5);
    // Stiller
    const h = back.styleAt(s, 0, 1);
    expect(h).toMatchObject({ bold: true, fill: "#C6EFCE", color: "#006100", hAlign: "center" });
    expect(h.bBottom).toEqual({ style: "medium", color: "#000000" });
    expect(back.styleAt(s, 1, 2)).toMatchObject({ italic: true, font: "Arial", size: 14, wrap: true, vAlign: "top", numFmt: '#,##0.00 "₺"' });
    // Yapı
    expect(s.merges).toContainEqual({ r1: 10, c1: 0, r2: 11, c2: 2 });
    expect(Math.abs(s.cols.size(0) - 120)).toBeLessThanOrEqual(1);
    expect(s.cols.isHidden(5)).toBe(true);
    expect(s.rows.size(0)).toBe(30);
    expect(s.rows.isHidden(20)).toBe(true);
    expect(s.frozen).toEqual({ rows: 1, cols: 1 });
    expect(s.tabColor).toBe("#E97132");
    expect(s.autoFilter?.range).toEqual({ r1: 0, c1: 0, r2: 3, c2: 3 });
    expect(s.notes.get(cellKey(0, 0))?.text).toBe("Ürün adı");
    expect(s.links.get(cellKey(12, 0))).toBe("https://example.com/");
    expect(back.getValue(s, 12, 0)).toBe("Bağlantı");
    // Koşullu biçim ve doğrulama
    expect(s.condFormats.map((c) => c.type).sort()).toEqual(["cellIs", "colorScale", "dataBar"]);
    const cf = s.condFormats.find((c) => c.type === "cellIs")!;
    expect(cf.operator).toBe("greaterThan");
    expect(cf.values).toEqual(["=100"]);
    expect(cf.style?.fill).toBe("#FFC7CE");
    const list = s.validations.find((v) => v.rule.type === "list")!;
    expect(list.rule.formula1).toBe("Evet;Hayır");
    expect(list.ranges).toEqual([{ r1: 1, c1: 4, r2: 3, c2: 4 }]);
    const whole = s.validations.find((v) => v.rule.type === "whole")!;
    expect(whole.rule.operator).toBe("between");
    // Ad, tablo, resim, grafik, sayfa yapısı
    expect(back.engine.getNames().map((n) => n.name)).toContain("Fiyatlar");
    expect(o.tables[0]).toMatchObject({ name: "Tablo1", ref: { r1: 3, c1: 0, r2: 5, c2: 1 }, style: "TableStyleMedium2" });
    expect(back.getValue(o, 4, 1)).toBe(1);
    expect(o.images.length).toBe(1);
    expect(o.images[0].src.startsWith("data:image/png;base64,")).toBe(true);
    expect(s.charts.length).toBe(1);
    expect(s.charts[0]).toMatchObject({ type: "column", title: "Satışlar", showLegend: true, legendPos: "bottom" });
    expect(s.charts[0].range).toBe("Satışlar!A1:B4");
    expect(s.page.orientation).toBe("landscape");
    expect(s.page.printArea).toEqual({ r1: 0, c1: 0, r2: 5, c2: 3 });
    expect(back.props.title).toBe("Deneme");
  });
  it("_xlfn önekleri", () => {
    expect(addXlfnPrefixes("XLOOKUP(1,A:A,B:B)+SUM(1)")).toBe("_xlfn.XLOOKUP(1,A:A,B:B)+SUM(1)");
    expect(addXlfnPrefixes("IFS(A1>1,1)&CONCAT(A1)")).toBe("_xlfn.IFS(A1>1,1)&_xlfn.CONCAT(A1)");
    expect(addXlfnPrefixes('"XLOOKUP("&1')).toBe('"XLOOKUP("&1');
    // LibreOffice ile doğrulandı: AGGREGATE/T.INV.2T/GAMMA önekli, NETWORKDAYS.INTL öneksiz
    expect(addXlfnPrefixes("AGGREGATE(9,6,A1:A2)+T.INV.2T(0.05,10)+GAMMA(2)+NETWORKDAYS.INTL(A1,B1)")).toBe("_xlfn.AGGREGATE(9,6,A1:A2)+_xlfn.T.INV.2T(0.05,10)+_xlfn.GAMMA(2)+NETWORKDAYS.INTL(A1,B1)");
  });
  it("hücre sıkıştırma", () => {
    expect(compressCells([[0, 0], [1, 0], [2, 0], [0, 1], [1, 1], [2, 1], [5, 0]])).toEqual([
      { r1: 0, c1: 0, r2: 2, c2: 1 },
      { r1: 5, c1: 0, r2: 5, c2: 0 },
    ]);
  });
  it("grafik XML'i", () => {
    const wb = sample();
    const xml = chartXml(wb, wb.sheets[0].charts[0]);
    expect(xml).toContain("<c:barChart>");
    expect(xml).toContain("<c:f>Satışlar!$B$1</c:f>");
    expect(xml).toContain("<c:f>Satışlar!$B$2:$B$4</c:f>");
    expect(xml).toContain("<c:f>Satışlar!$A$2:$A$4</c:f>");
  });
});

describe("CSV", () => {
  it("Türkçe CSV yazar (BOM, ; ayırıcı, görünen değerler)", () => {
    const wb = new WorkbookModel();
    const s = wb.active;
    wb.setInputRaw(s, 0, 0, "Ad;Soyad");
    wb.setInputRaw(s, 0, 1, 1234.5);
    wb.setInputRaw(s, 1, 0, 'Tırnak "x"');
    wb.setInputRaw(s, 1, 1, { formula: "=B1*2" });
    wb.setStyleRaw(s, 0, 1, wb.styles.intern({ numFmt: "#,##0.00" }));
    wb.engine.recalc();
    const out = writeCsv(wb, s);
    expect(out.startsWith("﻿")).toBe(true);
    expect(out).toBe('﻿"Ad;Soyad";1.234,50\r\n"Tırnak ""x""";2469\r\n');
  });
  it("ayırıcı ve kodlama tahmini", () => {
    expect(detectDelimiter("a;b;c\n1;2;3")).toBe(";");
    expect(detectDelimiter("a,b,c\n1,2,3")).toBe(",");
    expect(detectDelimiter("a\tb\n1\t2")).toBe("\t");
    expect(decodeText(new Uint8Array([0xef, 0xbb, 0xbf, 0xc5, 0x9f]))).toBe("ş");
    expect(decodeText(new Uint8Array([0xfe, 0xdd]))).toBe("şİ");
  });
  it("tırnaklı alanlar ve satır sonları", () => {
    expect(parseCsv('a;"b;c";"d""e"\r\n"çok\nsatır";2', ";")).toEqual([
      ["a", "b;c", 'd"e'],
      ["çok\nsatır", "2"],
    ]);
  });
  it("içe aktarma türleri", () => {
    const { rows } = csvToInputs("Ad;Tutar;Tarih;Oran\nÇay;1.234,5;15.03.2024;%12\n");
    expect(rows[1][1]).toMatchObject({ input: 1234.5 });
    expect(rows[1][2]).toMatchObject({ input: 45366, numFmt: "dd.mm.yyyy" });
    expect(rows[1][3]).toMatchObject({ input: 0.12 });
    const en = csvToInputs("a,b\n1.5,\"1,234\"\n");
    expect(en.rows[1][0].input).toBe(1.5);
    const wb = readCsv(new TextEncoder().encode("x;y\n1;=A2*2\n"));
    expect(wb.getValue(wb.active, 1, 1)).toBe(2);
  });
});

describe("üst/alt bilgi kodları", () => {
  it("arayüz kodları dosya kodlarına çevrilir ve geri okunur", () => {
    expect(hfToFile("Sayfa &[Sayfa] / &[Sayfalar] - &[Sekme]")).toBe("&CSayfa &P / &N - &A");
    expect(stripHF("&CSayfa &P / &N - &A")).toBe("Sayfa &[Sayfa] / &[Sayfalar] - &[Sekme]");
    expect(stripHF('&L&"Arial,Bold"&12Rapor&R&D')).toBe("Rapor&[Tarih]");
    expect(stripHF(hfToFile("A & B"))).toBe("A & B");
  });
});
