/**
 * Kendi .xlsx yazıcımız (io/xlsxWrite, P5.1): eski (exceljs) yazıcıyla eşdeğer model, paket
 * tutarlılığı, akıtılarak yazılan büyük sayfa, yazıcı seçimi (OH_XLSX_YAZICI) ve eski yazıcının
 * bozuk ürettiği durumların düzeltilmesi. Ayrıntılı fark testi: scripts/derlem/xlsx_fark.mjs.
 */
import { afterEach, describe, expect, it } from "vitest";
import JSZip from "jszip";
import { cellKey, isError } from "@officehere/formula";
import { OpcPackage } from "@officehere/core/opc";
import { WorkbookModel } from "../src/model/workbook";
import { readXlsx, writeXlsx, xlsxWriterChoice } from "../src/io/xlsx";
import { crc32Update } from "../src/io/xlsxWrite/zip";

const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFklEQVR42mP8z8Dwn4GBgYGJAQoAAAD//wMAAP//AwDsAD8AAAAASUVORK5CYII=";
const ab = (u: Uint8Array): ArrayBuffer => u.buffer.slice(u.byteOffset, u.byteOffset + u.byteLength) as ArrayBuffer;

async function yaz(wb: WorkbookModel, yazici: "eski" | "yeni"): Promise<Uint8Array> {
  const once = process.env.OH_XLSX_YAZICI;
  process.env.OH_XLSX_YAZICI = yazici;
  try {
    return await writeXlsx(wb);
  } finally {
    if (once === undefined) delete process.env.OH_XLSX_YAZICI;
    else process.env.OH_XLSX_YAZICI = once;
  }
}

afterEach(() => {
  delete process.env.OH_XLSX_YAZICI;
});

/** Karşılaştırılabilir model özeti (hücreler, stiller, sayfa öğeleri). */
function ozet(wb: WorkbookModel) {
  wb.engine.recalc();
  return wb.sheets.map((s) => {
    const cells: Record<string, unknown> = {};
    wb.engine.forEachCell(s.id, (r, c) => {
      const v = wb.getValue(s, r, c);
      cells[`${r},${c}`] = [wb.getFormula(s, r, c) ?? null, isError(v) ? v.code : v];
    });
    const st: Record<string, unknown> = {};
    for (const [k, id] of s.styles) st[k] = wb.styles.get(id);
    return {
      name: s.name,
      cells,
      st,
      colStyles: [...s.colStyles].map(([c, id]) => [c, wb.styles.get(id)]),
      cols: s.cols.customItems(),
      rows: s.rows.customItems(),
      hiddenRows: s.rows.hiddenItems(),
      merges: s.merges,
      notes: [...s.notes].map(([k, n]) => [k, n.text]),
      links: [...s.links],
      validations: s.validations.map((v) => ({ ...v, id: "" })),
      condFormats: s.condFormats.map((c) => ({ ...c, id: "" })),
      tables: s.tables.map((t) => ({ ...t, id: "" })),
      images: s.images.map((i) => ({ ...i, id: "" })),
      autoFilter: s.autoFilter,
      frozen: s.frozen,
      page: s.page,
      tabColor: s.tabColor,
      hidden: s.hidden,
    };
  });
}

function zengin(ikinci = "Liste"): WorkbookModel {
  const wb = new WorkbookModel(true);
  const s = wb.addSheet("Satışlar");
  const l = wb.addSheet(ikinci);
  const items: [string, number][] = [
    ["Çay", 10],
    ["Şeker <&>", 4],
    [" boşluklu ", 7],
  ];
  ["Ürün", "Adet", "Tutar"].forEach((h, c) => wb.setInputRaw(s, 0, c, h));
  items.forEach(([n, q], i) => {
    wb.setInputRaw(s, i + 1, 0, n);
    wb.setInputRaw(s, i + 1, 1, q);
    wb.setInputRaw(s, i + 1, 2, { formula: `=B${i + 2}*2` });
  });
  s.tables.push({ id: "t1", name: "Satış Tablosu", ref: { r1: 0, c1: 0, r2: 4, c2: 2 }, style: "TableStyleMedium2", headerRow: true, totalRow: true, bandedRows: true, bandedCols: false, firstColumn: false, lastColumn: false, filterButton: true, columns: ["Ürün", "Adet", "Tutar"] });
  wb.setInputRaw(s, 4, 0, "Toplam");
  wb.setInputRaw(s, 4, 2, { formula: "=SUBTOTAL(109,C2:C4)" });
  wb.setInputRaw(s, 6, 0, { formula: "=SEQUENCE(3)" });
  wb.setInputRaw(s, 6, 1, true);
  wb.setInputRaw(s, 6, 2, { formula: "=1/0" });
  wb.setInputRaw(s, 7, 1, -0.000001234);
  wb.setInputRaw(s, 8, 1, 123456789012345);
  const kalin = wb.styles.intern({ bold: true, fill: "#C6EFCE", bBottom: { style: "medium", color: "#000000" }, hAlign: "center", vAlign: "middle", rotation: -45 });
  const yuzde = wb.styles.intern({ numFmt: "0.00%", italic: true, font: "Arial", size: 10.5, locked: false });
  for (let c = 0; c < 3; c++) s.styles.set(cellKey(0, c), kalin);
  s.colStyles.set(5, yuzde);
  s.rowStyles.set(12, kalin);
  s.notes.set(cellKey(12, 5), { text: "satır ve sütun stilinin kesişimi" });
  s.merges.push({ r1: 14, c1: 0, r2: 15, c2: 2 });
  wb.setInputRaw(s, 14, 0, "Birleşik");
  s.styles.set(cellKey(14, 0), kalin);
  s.links.set(cellKey(1, 0), "https://example.com/?a=1&b=2");
  s.links.set(cellKey(20, 0), `#'${ikinci.replace(/'/g, "''")}'!A1`);
  s.cols.set(0, 140);
  s.cols.set(1, 90);
  s.cols.set(2, 90);
  s.cols.setHidden(7, true);
  s.rows.set(0, 32);
  s.rows.setHidden(21, true);
  s.rows.setLevel(22, 1);
  s.frozen = { rows: 1, cols: 0 };
  s.tabColor = "#E97132";
  s.condFormats.push({ id: "c1", type: "cellIs", operator: "greaterThan", values: ["=5"], ranges: [{ r1: 1, c1: 1, r2: 3, c2: 1 }], style: { fill: "#FFC7CE", color: "#9C0006" } });
  s.condFormats.push({ id: "c2", type: "dataBar", color: "#638EC6", ranges: [{ r1: 1, c1: 2, r2: 3, c2: 2 }] });
  s.condFormats.push({ id: "c3", type: "iconSet", iconSet: "3Triangles", ranges: [{ r1: 1, c1: 1, r2: 3, c2: 1 }] });
  s.condFormats.push({ id: "c4", type: "containsText", text: 'a"b', ranges: [{ r1: 1, c1: 0, r2: 3, c2: 0 }], style: { bold: true } });
  s.validations.push({ id: "v1", ranges: [{ r1: 1, c1: 4, r2: 3, c2: 4 }], rule: { type: "list", formula1: "Evet;Hayır" } });
  s.images.push({ id: "i1", name: "Resim 1", src: PNG, anchor: { row: 2, col: 4, dx: 3, dy: 2 }, width: 40, height: 20 });
  s.page.printArea = { r1: 0, c1: 0, r2: 4, c2: 2 };
  s.page.header = "Rapor & Özet";
  s.page.fitMode = "cols";
  l.page.printArea = { r1: 1, c1: 1, r2: 1, c2: 1 };
  l.page.titleRows = { r1: 0, r2: 0 };
  wb.setInputRaw(l, 0, 0, "Liste");
  wb.engine.defineName("Oran", "=0.18");
  wb.syncTables();
  wb.engine.recalc(true);
  return wb;
}

describe("Kendi .xlsx yazıcısı", () => {
  it("eski yazıcıyla aynı modeli üretir", async () => {
    const wb = zengin();
    const eski = await readXlsx(ab(await yaz(wb, "eski")));
    const yeni = await readXlsx(ab(await yaz(wb, "yeni")));
    const a = ozet(eski);
    const b = ozet(yeni);
    // Eski yazıcı tek hücrelik yazdırma alanını bozuk yazıyordu ($B2:$undefined → anlamsız aralık)
    expect(a[1].page.printArea).not.toEqual({ r1: 1, c1: 1, r2: 1, c2: 1 });
    expect(b[1].page.printArea).toEqual({ r1: 1, c1: 1, r2: 1, c2: 1 });
    a[1].page.printArea = b[1].page.printArea = undefined;
    expect(b).toEqual(a);
    expect(yeni.engine.getNames().map((n) => n.name)).toEqual(eski.engine.getNames().map((n) => n.name));
  });

  it("kesme işaretli sayfa adında yazdırma alanı ve başlıklar okunabilir (eski yazıcının çıktısı açılmıyordu)", async () => {
    const wb = zengin("O'Brien");
    await expect(readXlsx(ab(await yaz(wb, "eski")))).rejects.toThrow();
    const yeni = await readXlsx(ab(await yaz(wb, "yeni")));
    expect(yeni.sheets[1].name).toBe("O'Brien");
    expect(yeni.sheets[1].page.printArea).toEqual({ r1: 1, c1: 1, r2: 1, c2: 1 });
    expect(yeni.sheets[1].page.titleRows).toEqual({ r1: 0, r2: 0 });
    expect(yeni.sheets[0].links.get(cellKey(20, 0))).toBe("#'O''Brien'!A1");
  });

  it("paket tutarlı: içerik türleri ilk parça, her parçanın türü ve her ilişkinin hedefi var", async () => {
    const out = await yaz(zengin("O'Brien"), "yeni");
    const pkg = await OpcPackage.open(out);
    expect(pkg.names[0]).toBe("[Content_Types].xml");
    for (const name of pkg.names) if (!name.endsWith(".rels") && name !== "[Content_Types].xml") expect(pkg.contentType(name), name).toBeTruthy();
    for (const src of ["", ...pkg.names.filter((n) => !n.endsWith(".rels"))])
      for (const r of await pkg.rels(src)) if (!r.external && r.part) expect(pkg.has(r.part), `${src} → ${r.target}`).toBe(true);
    const zip = await JSZip.loadAsync(out);
    const sheet = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
    // Dinamik dizi: cm + metadata.xml; köprüler; tablo; resim çizimi (oneCellAnchor'da editAs yok)
    expect(sheet).toMatch(/<c r="A7" cm="1"><f t="array" ref="A7:A9">_xlfn.SEQUENCE\(3\)<\/f>/);
    expect(zip.file("xl/metadata.xml")).toBeTruthy();
    expect(sheet).toContain('<hyperlink ref="A2" r:id="rId1"/>');
    expect(await zip.file("xl/drawings/drawing1.xml")!.async("string")).not.toContain("editAs");
    // Metin kuralı: geçerli tür, işleç, metin ve kaçışlı formül
    expect(sheet).toContain('type="containsText"');
    expect(sheet).toContain('text="a&quot;b"');
    expect(sheet).toContain("NOT(ISERROR(SEARCH(&quot;a&quot;&quot;b&quot;,A2)))");
    const wbXml = await zip.file("xl/workbook.xml")!.async("string");
    expect(wbXml).toContain("&apos;O&apos;&apos;Brien&apos;!$B$2:$B$2");
  });

  it("büyük sayfayı parça parça yazar (satır sırası, paylaşılan metinler, CRC)", async () => {
    const wb = new WorkbookModel(true);
    const s = wb.addSheet("Veri");
    const N = 30_000;
    wb.engine.autoCalc = false;
    // Satırlar ters sırada girilir: yazıcı anahtarları sıralamalı
    for (let r = N - 1; r >= 0; r--) {
      wb.setInputRaw(s, r, 0, r);
      wb.setInputRaw(s, r, 1, `Müşteri ${r % 1000}`);
      wb.setInputRaw(s, r, 2, r % 3 === 0 ? { formula: `=A${r + 1}*2` } : r / 7);
    }
    wb.engine.autoCalc = true;
    wb.engine.recalc(true);
    const out = await yaz(wb, "yeni");
    // ZIP açıcı CRC'yi doğrular (JSZip checkCRC32)
    const zip = await JSZip.loadAsync(out, { checkCRC32: true });
    const xml = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
    const rows = [...xml.matchAll(/<row r="(\d+)"/g)].map((m) => Number(m[1]));
    expect(rows.length).toBe(N);
    expect(rows.every((r, i) => r === i + 1)).toBe(true);
    const sst = await zip.file("xl/sharedStrings.xml")!.async("string");
    expect(sst).toContain(`count="${N}" uniqueCount="1000"`);
    const back = await readXlsx(ab(out));
    const b = back.sheets[0];
    expect(back.getValue(b, N - 1, 1)).toBe(`Müşteri ${(N - 1) % 1000}`);
    expect(back.getFormula(b, 3, 2)).toBe("=A4*2");
    expect(back.getValue(b, 3, 2)).toBe(6);
    expect(back.getValue(b, 1, 2)).toBeCloseTo(1 / 7, 12);
  });

  it("CRC-32 sürdürmesi tek seferlik hesapla aynı", () => {
    const data = new TextEncoder().encode("OfficeHere ığüşöç ".repeat(1000));
    const whole = crc32Update(0, data);
    const parts = crc32Update(crc32Update(crc32Update(0, data.subarray(0, 7)), data.subarray(7, 4099)), data.subarray(4099));
    expect(parts).toBe(whole);
    expect(crc32Update(0, new TextEncoder().encode("123456789"))).toBe(0xcbf43926);
  });

  it("yazıcı seçimi: OH_XLSX_YAZICI=eski exceljs yolunu seçer", async () => {
    expect(xlsxWriterChoice()).toBe("");
    process.env.OH_XLSX_YAZICI = "eski";
    expect(xlsxWriterChoice()).toBe("eski");
    const wb = new WorkbookModel(true);
    wb.setInputRaw(wb.addSheet("A"), 0, 0, 1);
    // exceljs (JSZip) çıktısında dizin girdileri vardır, yeni yazıcınınkinde yok
    const eski = await JSZip.loadAsync(await writeXlsx(wb));
    expect(Object.keys(eski.files)).toContain("xl/");
    delete process.env.OH_XLSX_YAZICI;
    const yeni = await JSZip.loadAsync(await writeXlsx(wb));
    expect(Object.keys(yeni.files)).not.toContain("xl/");
  });
});
