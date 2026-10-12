/** 4. tur: tablo düzeltmeleri (Y1, Y2, O1–O7), Durum Çözümlemesi, Gelişmiş Filtre, kitap koruması, dilimleyici, yazdırma başlıkları. */
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { Controller } from "../src/controller";
import * as D from "../src/actions/data";
import { tableNameError, looksLikeCellRef } from "../src/model/tableName";
import { structuredPointRef } from "../src/model/structPoint";
import { readXlsx, writeXlsx } from "../src/io/xlsx";
import { calculatedColumnFormula, dataValidationsXml } from "../src/io/xlsxExtras";
import { deleteTableRows, insertTableCols, insertTableRows } from "../src/actions/tableEdit";
import { createDataTable, saveScenario, scenarioSummary, showScenario } from "../src/actions/whatIfOps";
import { advancedMatches } from "../src/model/advancedFilter";
import { applyAdvancedFilter } from "../src/actions/advFilter";
import { checkProtectionPassword, hashPassword, legacyPasswordHash } from "../src/model/protection";
import { clickSlicerItem, insertSlicers, slicerItems } from "../src/actions/slicers";
import { sheetPages, titledLists } from "../src/io/printPages";
import { learnProgram, runProgram } from "../src/model/flashFill";
import { flashPreviewFor } from "../src/actions/flashPreview";
import { pickOrphans } from "../src/app/draftStore";
import { parseX14Sparklines, x14SparklinesExt } from "../src/io/xlsxSpark";

function type(ctl: Controller, row: number, col: number, text: string): void {
  ctl.selectCell(row, col);
  ctl.startEdit(text);
  ctl.commitEdit(null);
}

function table(): Controller {
  const ctl = new Controller();
  const rows: (string | number)[][] = [
    ["Bölge", "Adet", "Fiyat"],
    ["Ege", 2, 10],
    ["Akdeniz", 3, 12],
    ["Marmara", 4, 15],
  ];
  ctl.ops.setInputs(ctl.sheet, rows.flatMap((r, i) => r.map((v, j) => ({ row: i, col: j, input: v }))), "örnek");
  D.createTable(ctl, { r1: 0, c1: 0, r2: 3, c2: 2 }, true);
  return ctl;
}

const v = (ctl: Controller, r: number, c: number) => ctl.wb.getValue(ctl.sheet, r, c);
const buf = (b: Uint8Array) => b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer;

describe("Y1: tablo adı doğrulaması", () => {
  it("boşluk, geçersiz karakter, hücre başvurusu ve çakışma reddedilir", () => {
    const ctl = table();
    const id = ctl.sheet.tables[0].id;
    expect(tableNameError(ctl.wb, "Satış Tablosu", id)).toMatch(/boşluk/);
    expect(tableNameError(ctl.wb, "1Satis", id)).toMatch(/harf/);
    expect(tableNameError(ctl.wb, "A1", id)).toMatch(/hücre başvurusu/);
    expect(tableNameError(ctl.wb, "R2C3", id)).toMatch(/hücre başvurusu/);
    expect(tableNameError(ctl.wb, "", id)).toMatch(/boş/);
    expect(tableNameError(ctl.wb, "Satis_2026", id)).toBeNull();
    ctl.wb.engine.defineName("Fiyatlar", "=Sayfa1!$C$2:$C$4");
    expect(tableNameError(ctl.wb, "fiyatlar", id)).toMatch(/zaten/);
    expect(looksLikeCellRef("XFD1048576")).toBe(true);
    expect(looksLikeCellRef("XFE1")).toBe(false);
  });
});

describe("Y2: Aralığa Dönüştür yapısal başvuruları A1'e çevirir", () => {
  it("TOPLA(Tablo1[Fiyat]) ve [@Adet]*[@Fiyat] doğru sonuçla kalır; geri alınabilir", () => {
    const ctl = table();
    const n = ctl.sheet.tables[0].name;
    type(ctl, 0, 3, "Tutar");
    type(ctl, 1, 3, "=[@Adet]*[@Fiyat]");
    type(ctl, 6, 0, `=TOPLA(${n}[Fiyat])`);
    expect(v(ctl, 6, 0)).toBe(37);
    expect(v(ctl, 3, 3)).toBe(60);
    D.convertTableToRange(ctl, ctl.sheet.tables[0]);
    expect(ctl.sheet.tables).toHaveLength(0);
    expect(ctl.wb.getFormula(ctl.sheet, 6, 0)).toBe("=SUM($C$2:$C$4)");
    expect(ctl.wb.getFormula(ctl.sheet, 2, 3)).toBe("=$B3*$C3");
    expect(v(ctl, 6, 0)).toBe(37);
    expect(v(ctl, 2, 3)).toBe(36);
    ctl.ops.undo();
    expect(ctl.sheet.tables).toHaveLength(1);
    expect(ctl.wb.getFormula(ctl.sheet, 6, 0)).toBe(`=SUM(${n}[Fiyat])`);
  });
});

describe("O1: tablonun son hücresinde Tab yeni satır ekler", () => {
  it("satır sonunda sonraki satırın ilk hücresine, son hücrede yeni satıra geçilir", () => {
    const ctl = table();
    ctl.selectCell(1, 2);
    ctl.advance("right");
    expect(ctl.sel.active).toEqual({ row: 2, col: 0 });
    ctl.selectCell(3, 2);
    ctl.advance("right");
    expect(ctl.sheet.tables[0].ref.r2).toBe(4);
    expect(ctl.sel.active).toEqual({ row: 4, col: 0 });
    ctl.selectCell(2, 0);
    ctl.advance("left");
    expect(ctl.sel.active).toEqual({ row: 1, col: 2 });
  });
  it("toplam satırı aşağı kayar", () => {
    const ctl = table();
    D.updateTable(ctl, ctl.sheet.tables[0], { totalRow: true });
    expect(ctl.sheet.tables[0].ref.r2).toBe(4);
    ctl.selectCell(3, 2);
    ctl.advance("right");
    const t = ctl.sheet.tables[0];
    expect(t.ref.r2).toBe(5);
    expect(v(ctl, 5, 0)).toBe("Toplam");
    expect(ctl.wb.isEmpty(ctl.sheet, 4, 0)).toBe(true);
  });
});

describe("O2: tablo genişleyince doğrulama ve koşullu biçim yeni satırı kapsar; çakışan doğrulama yazılmaz", () => {
  it("alta yazılan satıra liste doğrulaması uzar", () => {
    const ctl = table();
    ctl.sheet.validations.push({ id: "dv1", ranges: [{ r1: 1, c1: 0, r2: 3, c2: 0 }], rule: { type: "list", formula1: "Ege;Akdeniz;Marmara;Trakya" } });
    ctl.sheet.condFormats.push({ id: "cf1", type: "dataBar", color: "#638EC6", ranges: [{ r1: 1, c1: 1, r2: 3, c2: 1 }] });
    type(ctl, 4, 0, "Trakya");
    expect(ctl.sheet.tables[0].ref.r2).toBe(4);
    expect(ctl.sheet.validations[0].ranges).toEqual([{ r1: 1, c1: 0, r2: 4, c2: 0 }]);
    expect(ctl.sheet.condFormats[0].ranges).toEqual([{ r1: 1, c1: 1, r2: 4, c2: 1 }]);
  });
  it("sonra eklenen kural çakışan hücrelerde öncekinin yerine geçer; aralık sınırı yok", () => {
    const ctl = new Controller();
    ctl.sheet.validations.push({ id: "a", ranges: [{ r1: 1, c1: 2, r2: 24, c2: 2 }], rule: { type: "list", formula1: "A;B" } });
    ctl.sheet.validations.push({ id: "b", ranges: [{ r1: 9, c1: 2, r2: 24, c2: 2 }], rule: { type: "whole", operator: "between", formula1: "1", formula2: "9" } });
    ctl.sheet.validations.push({ id: "c", ranges: [{ r1: 0, c1: 5, r2: 1048575, c2: 5 }], rule: { type: "decimal", operator: "greaterThan", formula1: "0" } });
    const xml = dataValidationsXml(ctl.sheet)!;
    expect(xml).toContain('sqref="C2:C9"');
    expect(xml).toContain('sqref="C10:C25"');
    expect(xml).not.toContain("C2:C25");
    expect(xml).toContain('sqref="F1:F1048576"');
    expect(xml).toContain("<formula1>&quot;A,B&quot;</formula1>");
  });
});

describe("O3: tabloya özgü ekle/sil", () => {
  it("tablo satırı yukarı eklenir, sütun sola eklenir, satır silinir (sayfanın geri kalanı kaymaz)", () => {
    const ctl = table();
    type(ctl, 2, 6, "dışarıda");
    ctl.selectCell(2, 1);
    insertTableRows(ctl, ctl.sheet.tables[0]);
    expect(ctl.sheet.tables[0].ref.r2).toBe(4);
    expect(v(ctl, 3, 0)).toBe("Akdeniz");
    expect(ctl.wb.isEmpty(ctl.sheet, 2, 0)).toBe(true);
    expect(v(ctl, 2, 6)).toBe("dışarıda");
    ctl.selectCell(1, 1);
    insertTableCols(ctl, ctl.sheet.tables[0]);
    expect(ctl.sheet.tables[0].columns).toEqual(["Bölge", "Sütun4", "Adet", "Fiyat"]);
    expect(v(ctl, 0, 1)).toBe("Sütun4");
    ctl.selectCell(2, 0);
    deleteTableRows(ctl, ctl.sheet.tables[0]);
    expect(ctl.sheet.tables[0].ref.r2).toBe(3);
    expect(v(ctl, 2, 0)).toBe("Akdeniz");
    // Sütun eklemede tablonun satırlarındaki hücreler sağa kayar (Excel'deki gibi), satır silmede yerinde kalır
    expect(v(ctl, 2, 7)).toBe("dışarıda");
  });
});

describe("O5: formülde tablo hücresine işaret etmek yapısal başvuru üretir", () => {
  it("aynı satır [@Sütun], sütunun verisi Tablo1[Sütun], başlık [#Başlıklar]", () => {
    const ctl = table();
    const sh = ctl.sheet;
    const n = sh.tables[0].name;
    expect(structuredPointRef(sh, { r1: 2, c1: 1, r2: 2, c2: 1 }, { sheetId: sh.id, row: 2, col: 5 })).toBe(`${n}[@Adet]`);
    expect(structuredPointRef(sh, { r1: 2, c1: 1, r2: 2, c2: 2 }, { sheetId: sh.id, row: 2, col: 3 })).toBe(`${n}[@[Adet]:[Fiyat]]`);
    expect(structuredPointRef(sh, { r1: 1, c1: 2, r2: 3, c2: 2 }, { sheetId: sh.id, row: 6, col: 0 })).toBe(`${n}[Fiyat]`);
    expect(structuredPointRef(sh, { r1: 0, c1: 2, r2: 0, c2: 2 }, { sheetId: sh.id, row: 6, col: 0 })).toBe(`${n}[[#Başlıklar];[Fiyat]]`);
    expect(structuredPointRef(sh, { r1: 2, c1: 1, r2: 2, c2: 1 }, { sheetId: sh.id, row: 6, col: 0 })).toBeNull();
    // Tablonun içindeki formül: tablo adı yazılmaz
    expect(structuredPointRef(sh, { r1: 1, c1: 1, r2: 1, c2: 1 }, { sheetId: sh.id, row: 1, col: 2 })).toBe("[@Adet]");
  });
});

describe("O7: hesaplanmış sütun ve toplam satırı tablo parçasına yazılır", () => {
  it("calculatedColumnFormula ve totalsRowFunction", async () => {
    const ctl = table();
    type(ctl, 0, 3, "Tutar");
    type(ctl, 1, 3, "=[@Adet]*[@Fiyat]");
    D.updateTable(ctl, ctl.sheet.tables[0], { totalRow: true });
    const t = ctl.sheet.tables[0];
    expect(calculatedColumnFormula(ctl.wb, ctl.sheet, t, 3)).toContain("[#This Row]");
    const zip = await JSZip.loadAsync(await writeXlsx(ctl.wb));
    const tx = await zip.file("xl/tables/table1.xml")!.async("string");
    expect(tx).toMatch(/<tableColumn id="4" name="Tutar" totalsRowFunction="sum"><calculatedColumnFormula>Tablo\d+\[\[#This Row\],\[Adet\]\]\*Tablo\d+\[\[#This Row\],\[Fiyat\]\]<\/calculatedColumnFormula><\/tableColumn>/);
    expect(tx).toContain('totalsRowLabel="Toplam"');
  });
});

describe("Durum Çözümlemesi: Senaryo Yöneticisi", () => {
  it("ekle, göster, özet ve XLSX gidiş-dönüşü", async () => {
    const ctl = new Controller();
    type(ctl, 0, 1, "100");
    type(ctl, 1, 1, "5");
    type(ctl, 2, 1, "=B1*B2");
    expect(saveScenario(ctl, { name: "İyi", cells: [{ row: 0, col: 1, value: 120 }, { row: 1, col: 1, value: 8 }] })).toBeNull();
    expect(saveScenario(ctl, { name: "Kötü", cells: [{ row: 0, col: 1, value: 80 }, { row: 1, col: 1, value: 3 }], comment: "düşük" })).toBeNull();
    expect(saveScenario(ctl, { name: "iyi", cells: [{ row: 0, col: 1, value: 1 }] })).toMatch(/zaten/);
    showScenario(ctl, "İyi");
    expect(v(ctl, 2, 1)).toBe(960);
    ctl.ops.undo();
    expect(v(ctl, 2, 1)).toBe(500);
    scenarioSummary(ctl, [{ row: 2, col: 1 }]);
    const sum = ctl.sheet;
    expect(sum.name).toBe("Senaryo Özeti");
    // Sonuç hücresi satırı: geçerli, İyi, Kötü
    const vals = [2, 3, 4].map((c) => ctl.wb.getValue(sum, 8, c));
    expect(vals).toEqual([500, 960, 240]);
    const src = ctl.wb.sheets.find((s) => s.name === "Sayfa1")!;
    expect(ctl.wb.getValue(src, 0, 1)).toBe(100);
    const bytes = await writeXlsx(ctl.wb);
    const zip = await JSZip.loadAsync(bytes);
    const paths = Object.keys(zip.files).filter((p) => /worksheets\/sheet\d+\.xml$/.test(p));
    const xmls = await Promise.all(paths.map((p) => zip.file(p)!.async("string")));
    expect(xmls.some((x) => x.includes('<scenario name="İyi"'))).toBe(true);
    const back = await readXlsx(buf(bytes));
    const s1 = back.sheetByName("Sayfa1")!;
    expect(s1.scenarios.map((s) => s.name)).toEqual(["İyi", "Kötü"]);
    expect(s1.scenarios[1]).toMatchObject({ comment: "düşük", cells: [{ row: 0, col: 1, value: 80 }, { row: 1, col: 1, value: 3 }] });
  });
});

describe("Durum Çözümlemesi: Veri Tablosu", () => {
  it("tek değişkenli sütun tablosu hesaplanır, giriş değişince yenilenir ve dosyaya TABLE olarak yazılır", async () => {
    const ctl = new Controller();
    type(ctl, 0, 1, "100");
    type(ctl, 1, 1, "5");
    type(ctl, 2, 1, "=B1*B2");
    type(ctl, 0, 3, "=B3");
    type(ctl, 1, 2, "2");
    type(ctl, 2, 2, "4");
    expect(createDataTable(ctl, { r1: 0, c1: 2, r2: 2, c2: 3 }, undefined, { row: 1, col: 1 })).toBeNull();
    expect([v(ctl, 1, 3), v(ctl, 2, 3)]).toEqual([200, 400]);
    expect(v(ctl, 2, 1)).toBe(500);
    expect(ctl.wb.editText(ctl.sheet, 1, 3)).toBe("{=TABLO(;B2)}");
    // Hücresi tek başına değiştirilemez
    ctl.selectCell(1, 3);
    ctl.startEdit("9");
    expect(ctl.edit).toBeNull();
    const zip = await JSZip.loadAsync(await writeXlsx(ctl.wb));
    const xml = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
    expect(xml).toMatch(/<c r="D2"[^>]*><f t="dataTable" ref="D2:D3" dt2D="0" dtr="0" r1="B2"\/><v>200<\/v><\/c>/);
    const back = await readXlsx(buf(await writeXlsx(ctl.wb)));
    expect(back.sheets[0].dataTables).toHaveLength(1);
    expect(back.sheets[0].dataTables[0]).toMatchObject({ ref: { r1: 1, c1: 3, r2: 2, c2: 3 }, colInput: { row: 1, col: 1 } });
  });
  it("iki değişkenli tablo", () => {
    const ctl = new Controller();
    type(ctl, 0, 0, "2");
    type(ctl, 0, 1, "3");
    type(ctl, 2, 2, "=A1*B1");
    type(ctl, 2, 3, "10");
    type(ctl, 2, 4, "20");
    type(ctl, 3, 2, "1");
    type(ctl, 4, 2, "5");
    expect(createDataTable(ctl, { r1: 2, c1: 2, r2: 4, c2: 4 }, { row: 0, col: 0 }, { row: 0, col: 1 })).toBeNull();
    expect([v(ctl, 3, 3), v(ctl, 3, 4), v(ctl, 4, 3), v(ctl, 4, 4)]).toEqual([10, 20, 50, 100]);
  });
});

describe("Gelişmiş Filtre", () => {
  function data(): Controller {
    const ctl = new Controller();
    const rows: (string | number)[][] = [["Bölge", "Ürün", "Tutar"], ["Ege", "Kalem", 100], ["Akdeniz", "Defter", 250], ["Ege", "Defter", 300], ["Marmara", "Kalem", 50], ["Ege", "Kalem", 100]];
    ctl.ops.setInputs(ctl.sheet, rows.flatMap((r, i) => r.map((x, j) => ({ row: i, col: j, input: x }))), "örnek");
    return ctl;
  }
  it("VE/VEYA ölçütleri, ile başlar, benzersiz kayıtlar", () => {
    const ctl = data();
    const s = ctl.sheet;
    const list = { r1: 0, c1: 0, r2: 5, c2: 2 };
    ctl.ops.setInputs(s, [{ row: 10, col: 0, input: "Bölge" }, { row: 10, col: 1, input: "Tutar" }, { row: 11, col: 0, input: "Ege" }, { row: 11, col: 1, input: ">150" }, { row: 12, col: 0, input: "Mar" }], "ölçüt");
    expect(advancedMatches(ctl.wb, s, list, { r1: 10, c1: 0, r2: 12, c2: 1 }, false)).toEqual([3, 4]);
    expect(advancedMatches(ctl.wb, s, list, null, true)).toEqual([1, 2, 3, 4]);
  });
  it("yerinde filtre satırları gizler, Temizle gösterir; kopyalama hedefe yazar", () => {
    const ctl = data();
    const s = ctl.sheet;
    ctl.ops.setInputs(s, [{ row: 10, col: 0, input: "Ürün" }, { row: 11, col: 0, input: "=Kalem" }], "ölçüt");
    expect(applyAdvancedFilter(ctl, { list: { r1: 0, c1: 0, r2: 5, c2: 2 }, criteria: { r1: 10, c1: 0, r2: 11, c2: 0 }, unique: false })).toBeNull();
    expect([1, 2, 3, 4, 5].map((r) => s.rows.isHidden(r))).toEqual([false, true, true, false, false]);
    D.clearAllFilters(ctl);
    expect([1, 2, 3, 4, 5].some((r) => s.rows.isHidden(r))).toBe(false);
    expect(applyAdvancedFilter(ctl, { list: { r1: 0, c1: 0, r2: 5, c2: 2 }, criteria: null, copyTo: { row: 0, col: 5 }, unique: true })).toBeNull();
    expect([0, 1, 2, 3, 4, 5].map((r) => ctl.wb.getValue(s, r, 5))).toEqual(["Bölge", "Ege", "Akdeniz", "Ege", "Marmara", null]);
  });
});

describe("Çalışma Kitabını Koru", () => {
  it("parola karması (SHA-512, tuz, döngü) ve eski 16 bit karma", async () => {
    const h = await hashPassword("abc", undefined, 50);
    expect(await checkProtectionPassword({ structure: true, hash: h }, "abc")).toBe(true);
    expect(await checkProtectionPassword({ structure: true, hash: h }, "abd")).toBe(false);
    // ECMA-376 Bölüm 4'teki eski algoritma (openpyxl ile aynı)
    expect(legacyPasswordHash("test")).toBe("CBEB");
    expect(await checkProtectionPassword({ structure: true, legacy: "CBEB" }, "test")).toBe(true);
  });
  it("yapı korunurken sayfa eklenemez/silinemez/adlandırılamaz; dosyaya yazılır ve okunur", async () => {
    const ctl = new Controller();
    const msgs: string[] = [];
    ctl.alert = (m) => msgs.push(m);
    ctl.wb.protection = { structure: true, hash: await hashPassword("x", undefined, 10) };
    ctl.ops.addSheet();
    ctl.ops.renameSheet(0, "Yeni");
    expect(ctl.wb.sheets.map((s) => s.name)).toEqual(["Sayfa1"]);
    expect(msgs.length).toBe(2);
    const bytes = await writeXlsx(ctl.wb);
    const zip = await JSZip.loadAsync(bytes);
    const wbXml = await zip.file("xl/workbook.xml")!.async("string");
    expect(wbXml).toMatch(/<workbookProtection workbookAlgorithmName="SHA-512" workbookHashValue="[^"]+" workbookSaltValue="[^"]+" workbookSpinCount="10" lockStructure="1"\/>/);
    // Şema sırası: workbookPr'den sonra, bookViews'tan önce
    expect(wbXml.indexOf("<workbookProtection")).toBeLessThan(wbXml.indexOf("<bookViews"));
    const back = await readXlsx(buf(bytes));
    expect(back.protection?.structure).toBe(true);
    expect(await checkProtectionPassword(back.protection!, "x")).toBe(true);
  });
});

describe("Dilimleyici", () => {
  it("tablo dilimleyicisi filtreler, Ctrl ile çoklu seçim, XLSX'te filtre ve dilimleyici korunur", async () => {
    const ctl = table();
    type(ctl, 4, 0, "Ege");
    const t = ctl.sheet.tables[0];
    insertSlicers(ctl, { kind: "table", tableId: t.id }, [{ name: "Bölge" }]);
    const s = () => ctl.sheet.slicers[0];
    expect(slicerItems(ctl, s()).map((i) => i.label)).toEqual(["Akdeniz", "Ege", "Marmara"]);
    clickSlicerItem(ctl, s(), "Ege", false);
    expect([1, 2, 3, 4].map((r) => ctl.sheet.rows.isHidden(r))).toEqual([false, true, true, false]);
    clickSlicerItem(ctl, s(), "Marmara", true);
    expect([1, 2, 3, 4].map((r) => ctl.sheet.rows.isHidden(r))).toEqual([false, true, false, false]);
    const bytes = await writeXlsx(ctl.wb);
    const zip = await JSZip.loadAsync(bytes);
    expect(await zip.file("xl/tables/table1.xml")!.async("string")).toContain('<filterColumn colId="0"><filters><filter val="Ege"/><filter val="Marmara"/></filters></filterColumn>');
    const back = await readXlsx(buf(bytes));
    const c2 = new Controller(back);
    c2.load(back, "x.xlsx", null);
    expect(back.sheets[0].slicers).toHaveLength(1);
    expect(slicerItems(c2, back.sheets[0].slicers[0]).filter((i) => i.selected).map((i) => i.label)).toEqual(["Ege", "Marmara"]);
    // Filtre ölçütü geri geldi: Temizle satırları gösterir
    expect(back.sheets[0].rows.isHidden(2)).toBe(true);
    D.clearAllFilters(c2);
    expect(back.sheets[0].rows.isHidden(2)).toBe(false);
  });
});

describe("Yazdırma başlıkları ve sayfa sırası (D10)", () => {
  it("başlık satırları sonraki sayfaların başına eklenir; önce yana sırası", async () => {
    const ctl = new Controller();
    const s = ctl.sheet;
    ctl.ops.setInputs(s, Array.from({ length: 120 }, (_, i) => ({ row: i, col: 0, input: i ? i : "Başlık" })), "örnek");
    s.page.titleRows = { r1: 0, r2: 0 };
    s.page.rowBreaks = [50];
    const pages = sheetPages(ctl.wb, s);
    expect(pages.length).toBeGreaterThan(1);
    expect(titledLists(s, pages[0].rect).rowList).toBeUndefined();
    expect(titledLists(s, pages[1].rect).rowList?.slice(0, 2)).toEqual([0, pages[1].rect.r1]);
    s.page.pageOrder = "overThenDown";
    s.page.blackAndWhite = true;
    const back = await readXlsx(buf(await writeXlsx(ctl.wb)));
    expect(back.sheets[0].page).toMatchObject({ titleRows: { r1: 0, r2: 0 }, pageOrder: "overThenDown", blackAndWhite: true });
  });
});

describe("O13: Akıllı Doldurma", () => {
  it("soyad son sözcükten; büyük harf dönüşümü; canlı önizleme", () => {
    const rows = [["Ali Veli"], ["Ayşe Kaya"], ["Mehmet Can Öz"]];
    const prog = learnProgram([{ row: rows[0], output: "Veli" }])!;
    expect(runProgram(prog, rows[2])).toBe("Öz");
    const up = learnProgram([{ row: rows[0], output: "ALİ VELİ" }])!;
    expect(runProgram(up, rows[1])).toBe("AYŞE KAYA");
    const ctl = new Controller();
    ctl.ops.setInputs(ctl.sheet, [["Ad Soyad", "Ad"], ["Ali Veli", "Ali"], ["Ayşe Kaya", null], ["Mehmet Can Öz", null]].flatMap((r, i) => r.map((x, j) => ({ row: i, col: j, input: x }))), "örnek");
    ctl.selectCell(2, 1);
    ctl.startEdit("Ay");
    const p = flashPreviewFor(ctl);
    expect(p?.text).toBe("Ayşe");
    expect(p?.fills).toEqual([{ row: 3, text: "Mehmet" }]);
  });
});

describe("Y3: kurtarma taslakları", () => {
  it("canlı ya da nabzı taze pencerelerin taslakları önerilmez", () => {
    const all = [
      { session: "a", savedAt: 1 },
      { session: "b", savedAt: 3 },
      { session: "c", savedAt: 2 },
      { session: "me", savedAt: 9 },
    ];
    const age = (id: string) => (id === "c" ? 1000 : Infinity);
    expect(pickOrphans(all, new Set(["a"]), "me", age).map((d) => d.session)).toEqual(["b"]);
  });
});

describe("K1: mini grafikler x14 uzantısı", () => {
  it("yazılır ve başka uygulamanın dosyasından okunur", () => {
    const ctl = new Controller();
    const ext = x14SparklinesExt("Sayfa1", [{ id: "g", type: "column", items: [{ row: 1, col: 5, range: "B2:E2" }], color: "#FF0000", highPoint: true, negative: true }])!;
    expect(ext).toContain('type="column"');
    expect(ext).toContain("<xm:f>Sayfa1!B2:E2</xm:f><xm:sqref>F2</xm:sqref>");
    const groups = parseX14Sparklines(`<worksheet><extLst>${ext}</extLst></worksheet>`, ctl.wb);
    expect(groups).toHaveLength(1);
    expect(groups[0]).toMatchObject({ type: "column", color: "#FF0000", highPoint: true, negative: true, items: [{ row: 1, col: 5, range: "Sayfa1!B2:E2" }] });
  });
});
