/** 3. tur kabul bulguları: tablolar, yapısal başvurular, biçim devralma, F4, nesneler, grafik serileri. */
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { Controller } from "../src/controller";
import * as D from "../src/actions/data";
import * as E from "../src/actions/edit";
import * as F from "../src/actions/format";
import { insertChart } from "../src/actions/objects";
import { copyObject, handleObjectKey, pasteObject } from "../src/actions/objectClip";
import { readXlsx, writeXlsx } from "../src/io/xlsx";
import { extractChartData } from "../src/charts/chartData";
import { requestDeleteSheet } from "../src/dialogs/MiscDialogs";
import { displayText } from "../src/model/display";

function type(ctl: Controller, row: number, col: number, text: string): void {
  ctl.selectCell(row, col);
  ctl.startEdit(text);
  ctl.commitEdit(null);
}

/** A1:E8 bölge tablosu (Ctrl+T). */
function regionTable(): Controller {
  const ctl = new Controller();
  const rows: (string | number)[][] = [
    ["Bölge", "Ocak", "Şubat", "Mart", "Tutar"],
    ["Marmara", 1, 2, 3, 100],
    ["Ege", 1, 2, 3, 200],
    ["Akdeniz", 1, 2, 3, 300],
    ["Karadeniz", 1, 2, 3, 400],
    ["Doğu", 1, 2, 3, 500],
    ["Güneydoğu", 1, 2, 3, 600],
    ["Trakya", 1, 2, 3, 700],
  ];
  ctl.ops.setInputs(ctl.sheet, rows.flatMap((r, i) => r.map((v, j) => ({ row: i, col: j, input: v }))), "örnek");
  D.createTable(ctl, { r1: 0, c1: 0, r2: 7, c2: 4 }, true);
  return ctl;
}

const v = (ctl: Controller, r: number, c: number) => ctl.wb.getValue(ctl.sheet, r, c);

describe("Tablo: otomatik genişleme ve toplam satırı (#3, #14)", () => {
  it("alta yazılan satır tabloya katılır; toplam satırı veriyi ezmez", () => {
    const ctl = regionTable();
    const t = ctl.sheet.tables[0];
    type(ctl, 8, 0, "İç Anadolu");
    type(ctl, 8, 4, "500");
    expect(ctl.sheet.tables[0].ref.r2).toBe(8);
    D.updateTable(ctl, ctl.sheet.tables[0], { totalRow: true });
    const t2 = ctl.sheet.tables[0];
    expect(t2.ref.r2).toBe(9);
    expect(v(ctl, 8, 0)).toBe("İç Anadolu");
    expect(v(ctl, 9, 0)).toBe("Toplam");
    expect(v(ctl, 9, 4)).toBe(3300);
    expect(ctl.wb.getFormula(ctl.sheet, 9, 4)).toBe(`=SUBTOTAL(109,${t.name}[Tutar])`);
    // Tek adımda geri alınır
    ctl.ops.undo();
    expect(ctl.sheet.tables[0].totalRow).toBe(false);
    expect(v(ctl, 9, 0)).toBe(null);
  });

  it("toplam satırı açılırken altındaki veri aşağı kaydırılır", () => {
    const ctl = regionTable();
    ctl.ops.setInputs(ctl.sheet, [{ row: 8, col: 9, input: "uzak" }], "x");
    // Tablonun hemen altında (tablo sütunlarında değil, ama bitişik satırda) veri
    ctl.ops.setInputs(ctl.sheet, [{ row: 8, col: 2, input: "Not" }], "x");
    ctl.ops.history.clear();
    // "Not" yazısı tabloya katılmasın diye doğrudan setInputs ile yazıldı; toplam satırı onu ezmemeli
    D.updateTable(ctl, ctl.sheet.tables[0], { totalRow: true });
    expect(v(ctl, 8, 0)).toBe("Toplam");
    expect(v(ctl, 9, 2)).toBe("Not");
    expect(v(ctl, 8, 9)).toBe("uzak");
  });

  it("sağa yazılan başlık sütunu tabloya katar; formül hesaplanmış sütun olur", () => {
    const ctl = regionTable();
    type(ctl, 0, 5, "KDV");
    const t = ctl.sheet.tables[0];
    expect(t.ref.c2).toBe(5);
    expect(t.columns[5]).toBe("KDV");
    type(ctl, 1, 5, "=E2*0,2");
    for (let r = 1; r <= 7; r++) expect(v(ctl, r, 5)).toBeCloseTo((r * 100) * 0.2);
    expect(ctl.wb.getFormula(ctl.sheet, 7, 5)).toBe("=E8*0.2");
    // Yeni satırda hesaplanmış sütun sürer
    type(ctl, 8, 0, "Yeni");
    type(ctl, 8, 4, "1000");
    expect(ctl.wb.getFormula(ctl.sheet, 8, 5)).toBe("=E9*0.2");
    expect(v(ctl, 8, 5)).toBe(200);
  });

  it("başlık değişince sütun adı ve yapısal başvurular güncellenir", () => {
    const ctl = regionTable();
    const name = ctl.sheet.tables[0].name;
    type(ctl, 10, 0, `=TOPLA(${name}[Tutar])`);
    expect(v(ctl, 10, 0)).toBe(2800);
    type(ctl, 0, 4, "Toplam Tutar");
    expect(ctl.sheet.tables[0].columns[4]).toBe("Toplam Tutar");
    expect(ctl.wb.getFormula(ctl.sheet, 10, 0)).toBe(`=SUM(${name}[Toplam Tutar])`);
    expect(v(ctl, 10, 0)).toBe(2800);
  });
});

describe("Yapısal başvurular uygulamada (#4)", () => {
  it("TOPLA, @satır, #Tümü, SATIRSAY ve tablo genişleyince güncellenme", () => {
    const ctl = regionTable();
    const n = ctl.sheet.tables[0].name;
    type(ctl, 12, 0, `=TOPLA(${n}[Tutar])`);
    type(ctl, 12, 1, `=TOPLA(${n}[[#Tümü];[Tutar]])`);
    type(ctl, 12, 2, `=SATIRSAY(${n})`);
    expect(v(ctl, 12, 0)).toBe(2800);
    expect(v(ctl, 12, 1)).toBe(2800);
    expect(v(ctl, 12, 2)).toBe(7);
    type(ctl, 1, 5, "=[@Tutar]*2");
    expect(ctl.wb.getFormula(ctl.sheet, 1, 5)).toBe(`=${n}[[#This Row],[Tutar]]*2`);
    expect(ctl.wb.editText(ctl.sheet, 1, 5)).toBe(`=${n}[@Tutar]*2`);
    expect(v(ctl, 1, 5)).toBe(200);
    type(ctl, 8, 0, "Ek");
    type(ctl, 8, 4, "1000");
    expect(v(ctl, 12, 0)).toBe(3800);
    expect(v(ctl, 12, 2)).toBe(8);
  });

  it("XLSX'e İngilizce yazılır ve geri okunur", async () => {
    const ctl = regionTable();
    const n = ctl.sheet.tables[0].name;
    type(ctl, 12, 0, `=TOPLA(${n}[[#Tümü];[Tutar]])`);
    const bytes = await writeXlsx(ctl.wb);
    const wb2 = await readXlsx(bytes.buffer);
    const s2 = wb2.sheets[0];
    expect(wb2.getFormula(s2, 12, 0)).toBe(`=SUM(${n}[[#All],[Tutar]])`);
    expect(wb2.getValue(s2, 12, 0)).toBe(2800);
  });
});

describe("Formül sonucu biçim devralma (#5)", () => {
  it("TOPLA para biçimini, BUGÜN tarihi, ŞİMDİ tarih-saati alır", () => {
    const ctl = new Controller();
    ctl.ops.setInputs(ctl.sheet, [1, 2, 3].map((x, i) => ({ row: i, col: 2, input: x * 100 })), "x");
    ctl.ops.applyStyle(ctl.sheet, [{ r1: 0, c1: 2, r2: 2, c2: 2 }], { numFmt: '#,##0.00 "₺"' });
    type(ctl, 4, 2, "=TOPLA(C1:C3)");
    expect(ctl.wb.styleAt(ctl.sheet, 4, 2).numFmt).toBe('#,##0.00 "₺"');
    expect(displayText(ctl.wb, ctl.sheet, 4, 2)).toBe("600,00 ₺");
    type(ctl, 5, 0, "=BUGÜN()");
    expect(ctl.wb.styleAt(ctl.sheet, 5, 0).numFmt).toBe("dd.mm.yyyy");
    type(ctl, 6, 0, "=ŞİMDİ()");
    expect(ctl.wb.styleAt(ctl.sheet, 6, 0).numFmt).toBe("dd.mm.yyyy hh:mm");
    type(ctl, 7, 0, "=BAĞ_DEĞ_SAY(C1:C3)");
    expect(ctl.wb.styleAt(ctl.sheet, 7, 0).numFmt).toBeUndefined();
  });
});

describe("F4 / Ctrl+Y son işlemi yinele (#6)", () => {
  it("biçim ve satır ekleme yinelenir; geri alınan varsa Ctrl+Y yineler", () => {
    const ctl = new Controller();
    ctl.selectCell(0, 0);
    F.toggle(ctl, "bold");
    ctl.selectCell(3, 3);
    expect(ctl.repeatLast()).toBe(true);
    expect(ctl.wb.styleAt(ctl.sheet, 3, 3).bold).toBe(true);
    ctl.selectRange({ r1: 5, c1: 0, r2: 6, c2: 2 });
    F.setFill(ctl, "#FF0000");
    ctl.selectCell(9, 9);
    E.redo(ctl);
    expect(ctl.wb.styleAt(ctl.sheet, 9, 9).fill).toBe("#FF0000");
    ctl.ops.undo();
    expect(ctl.wb.styleAt(ctl.sheet, 9, 9).fill).toBeUndefined();
    E.redo(ctl);
    expect(ctl.wb.styleAt(ctl.sheet, 9, 9).fill).toBe("#FF0000");
    ctl.ops.setInputs(ctl.sheet, [{ row: 20, col: 0, input: "x" }], "x");
    ctl.selectCell(20, 0);
    E.insertRows(ctl);
    expect(v(ctl, 21, 0)).toBe("x");
    ctl.repeatLast();
    expect(v(ctl, 22, 0)).toBe("x");
  });
});

describe("Boş sayfa silme (#7)", () => {
  it("boş sayfa sorulmadan silinir, dolu sayfa onay ister", () => {
    const ctl = new Controller();
    ctl.ops.addSheet();
    let opened = "";
    ctl.openDialog = (id) => {
      opened = id;
    };
    requestDeleteSheet(ctl, 1);
    expect(ctl.wb.sheets).toHaveLength(1);
    expect(opened).toBe("");
    ctl.ops.addSheet();
    ctl.ops.setInputs(ctl.wb.sheets[1], [{ row: 0, col: 0, input: 1 }], "x");
    requestDeleteSheet(ctl, 1);
    expect(opened).toBe("deleteSheet");
    expect(ctl.wb.sheets).toHaveLength(2);
  });
});

describe("Grafik nesnesi ve dış grafik serileri (#8, #9)", () => {
  it("karmasik_lo Bütçe grafiği tek seri ve metin kategorilerle okunur", async () => {
    const b = readFileSync(new URL("./fixtures/openpyxl/karmasik_lo.xlsx", import.meta.url));
    const wb = await readXlsx(b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer);
    const s = wb.sheetByName("Bütçe")!;
    const col = s.charts.find((c) => c.type === "column")!;
    expect(col.series).toHaveLength(1);
    wb.activeIndex = wb.sheets.indexOf(s);
    const data = extractChartData(wb, col);
    expect(data.series).toHaveLength(1);
    expect(data.categories[0]).toBe("Kira");
    expect(data.series[0].values[0]).toBe(15000);
    // Kaydedip açınca seriler korunur
    const again = await readXlsx((await writeXlsx(wb)).buffer);
    expect(again.sheetByName("Bütçe")!.charts.find((c) => c.type === "column")!.series).toHaveLength(1);
  });

  it("Delete siler, Ctrl+C/V kopyalar, geri alınabilir", () => {
    const ctl = regionTable();
    ctl.selectRange({ r1: 0, c1: 0, r2: 7, c2: 4 });
    insertChart(ctl, "column");
    const id = ctl.sheet.charts[0].id;
    expect(ctl.selObj).toEqual({ kind: "chart", id });
    expect(copyObject(ctl, false)).toBe(true);
    expect(pasteObject(ctl)).toBe(true);
    expect(ctl.sheet.charts).toHaveLength(2);
    expect(ctl.sheet.charts[1].id).not.toBe(id);
    const ev = { key: "Delete", ctrlKey: false, metaKey: false, altKey: false, shiftKey: false } as KeyboardEvent;
    expect(handleObjectKey(ctl, ev)).toBe(true);
    expect(ctl.sheet.charts).toHaveLength(1);
    // Hücre içerikleri silinmedi
    expect(v(ctl, 0, 0)).toBe("Bölge");
    ctl.ops.undo();
    expect(ctl.sheet.charts).toHaveLength(2);
    // Bir hücre seçilip yapıştırılınca grafik o hücreye yerleşir
    ctl.selObj = null;
    ctl.selectRange({ r1: 19, c1: 7, r2: 19, c2: 7 });
    expect(pasteObject(ctl)).toBe(true);
    expect(ctl.sheet.charts[2].anchor).toMatchObject({ row: 19, col: 7, dx: 0, dy: 0 });
  });
});

describe("Otomatik Kurtarma taslağı", () => {
  it("taslak gidiş-dönüşü: değerler, formüller, stiller, tablolar, grafikler, adlar", async () => {
    const { serializeDraftSync, serializeDraft, workbookFromDraft } = await import("../src/io/draft");
    const ctl = regionTable();
    const s = ctl.sheet;
    ctl.selectRange({ r1: 0, c1: 0, r2: 7, c2: 4 });
    insertChart(ctl, "line");
    ctl.ops.applyStyle(s, [{ r1: 1, c1: 4, r2: 7, c2: 4 }], { numFmt: '#,##0.00 "₺"', bold: true });
    type(ctl, 10, 0, `=TOPLA(${s.tables[0].name}[Tutar])`);
    ctl.wb.engine.defineName("Oran", "=0.18");
    ctl.ops.setHidden(s, "row", [3], true);
    s.cols.set(1, 120);
    ctl.ops.addSheet();
    ctl.ops.setInputs(ctl.wb.sheets[1], [{ row: 0, col: 0, input: true }], "x");
    const d = serializeDraftSync(ctl.wb);
    // Yapılandırılmış kopyaya uygun (IndexedDB)
    const copy = structuredClone(d);
    const wb = workbookFromDraft(copy);
    const s2 = wb.sheets[0];
    expect(wb.sheets.map((x) => x.name)).toEqual(ctl.wb.sheets.map((x) => x.name));
    expect(wb.getValue(s2, 10, 0)).toBe(2800);
    expect(wb.styleAt(s2, 2, 4)).toMatchObject({ numFmt: '#,##0.00 "₺"', bold: true });
    expect(s2.tables[0].name).toBe(s.tables[0].name);
    expect(s2.charts).toHaveLength(1);
    expect(s2.rows.isHidden(3)).toBe(true);
    expect(s2.cols.size(1)).toBe(120);
    expect(wb.getValue(wb.sheets[1], 0, 0)).toBe(true);
    expect(wb.engine.getNames().map((n) => n.name)).toContain("Oran");
    // Yeni nesne kimlikleri çakışmaz
    expect(wb.newId("chart")).not.toBe(s2.charts[0].id);
    // Üreteç parça parça ilerler
    const g = serializeDraft(ctl.wb, 5);
    let steps = 0;
    for (let r = g.next(); !r.done; r = g.next()) steps++;
    expect(steps).toBeGreaterThan(3);
  });
});
