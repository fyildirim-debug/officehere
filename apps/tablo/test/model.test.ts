import { describe, expect, it } from "vitest";
import { dateToSerial, formatValue, isError } from "@officehere/formula";
import { Controller } from "../src/controller";
import { WorkbookModel } from "../src/model/workbook";
import { addDate, generateFill } from "../src/model/fill";
import { findAll } from "../src/model/find";
import { buildClipboard, parseTsv, pastePayload } from "../src/model/clipboard";
import { currentRegion } from "../src/model/selection";
import { displayText } from "../src/model/display";
import * as D from "../src/actions/data";
import * as E from "../src/actions/edit";
import * as F from "../src/actions/format";
import { buildCode, detectCategory } from "../src/dialogs/numberFormats";
import { computePageBreaks, expandHeaderFooter, hfCss } from "../src/io/print";
import { freeze, gotoNote } from "../src/actions/view";
import { insertChart } from "../src/actions/objects";
import { extractChartData } from "../src/charts/chartData";

/** Örnek bölge satışları tablosu. */
function sample(): Controller {
  const ctl = new Controller();
  const rows: (string | number)[][] = [
    ["Bölge", "Ocak", "Şubat"],
    ["Marmara", 120, 135],
    ["Ege", 90, 95],
    ["Akdeniz", 70, 88],
    ["Ege", 90, 95],
  ];
  const cells = rows.flatMap((r, i) => r.map((v, j) => ({ row: i, col: j, input: v })));
  ctl.ops.setInputs(ctl.sheet, cells, "örnek");
  return ctl;
}

const val = (ctl: Controller, r: number, c: number) => ctl.wb.getValue(ctl.sheet, r, c);

describe("doldurma serileri", () => {
  const fill = (inputs: (string | number)[], count: number, numFmt?: string) => generateFill(inputs.map((input) => ({ input, numFmt })), count, false, false, () => [0, 0]);

  it("sayı serileri doğrusal artar", () => {
    expect(fill([1, 2], 3)).toEqual([3, 4, 5]);
    expect(fill([10, 7], 2)).toEqual([4, 1]);
  });
  it("Türkçe ay ve gün adları", () => {
    expect(fill(["Ocak"], 3)).toEqual(["Şubat", "Mart", "Nisan"]);
    expect(fill(["Cuma"], 3)).toEqual(["Cumartesi", "Pazar", "Pazartesi"]);
    expect(fill(["ARALIK"], 1)).toEqual(["OCAK"]);
  });
  it("metin + sayı ('Ürün 1')", () => {
    expect(fill(["Ürün 1"], 3)).toEqual(["Ürün 2", "Ürün 3", "Ürün 4"]);
  });
  it("tarih biçimli tek hücre gün gün artar", () => {
    const d = dateToSerial(2024, 2, 28);
    expect(fill([d], 2, "dd.mm.yyyy")).toEqual([d + 1, d + 2]);
  });
  it("Seriler: ay/yıl/hafta içi adımları", () => {
    const jan31 = dateToSerial(2024, 1, 31);
    expect(addDate(jan31, "month", 1)).toBe(dateToSerial(2024, 2, 29));
    expect(addDate(dateToSerial(2024, 2, 29), "year", 1)).toBe(dateToSerial(2025, 2, 28));
    // Cuma + 1 iş günü = Pazartesi
    expect(addDate(dateToSerial(2024, 3, 1), "weekday", 1)).toBe(dateToSerial(2024, 3, 4));
  });
});

describe("düzenleme ve geri alma", () => {
  it("Türkçe formül girişi kanonik saklanır, geri al/yinele çalışır", () => {
    const ctl = sample();
    ctl.selectCell(5, 1);
    ctl.startEdit("=TOPLA(B2:B5)");
    expect(ctl.commitEdit("down")).toBe(true);
    expect(ctl.wb.getFormula(ctl.sheet, 5, 1)).toBe("=SUM(B2:B5)");
    expect(val(ctl, 5, 1)).toBe(370);
    expect(ctl.sel.active).toEqual({ row: 6, col: 1 });
    E.undo(ctl);
    expect(val(ctl, 5, 1)).toBeNull();
    E.redo(ctl);
    expect(val(ctl, 5, 1)).toBe(370);
  });
  it("ondalık virgül ve tarih girişi", () => {
    const ctl = new Controller();
    ctl.startEdit("12,5");
    ctl.commitEdit(null);
    expect(val(ctl, 0, 0)).toBe(12.5);
    ctl.selectCell(1, 0);
    ctl.startEdit("15.03.2024");
    ctl.commitEdit(null);
    expect(val(ctl, 1, 0)).toBe(dateToSerial(2024, 3, 15));
    expect(displayText(ctl.wb, ctl.sheet, 1, 0)).toBe("15.03.2024");
  });
  it("satır eklemek formülleri kaydırır, silmek #BAŞV! üretir", () => {
    const ctl = sample();
    ctl.ops.enterText(ctl.sheet, [{ row: 6, col: 1 }], "=B2+B3");
    ctl.selectRows(1, 1);
    E.insertRows(ctl);
    expect(ctl.wb.getFormula(ctl.sheet, 7, 1)).toBe("=B3+B4");
    expect(val(ctl, 7, 1)).toBe(210);
    ctl.selectRows(2, 2);
    E.deleteRows(ctl);
    const v = val(ctl, 6, 1);
    expect(isError(v) && v.code).toBe("#REF!");
  });
  it("Ctrl+Enter seçimi göreli formülle doldurur", () => {
    const ctl = sample();
    ctl.selectRange({ r1: 1, c1: 3, r2: 4, c2: 3 }, { row: 1, col: 3 });
    ctl.startEdit("=B2+C2");
    ctl.commitEdit(null, true);
    expect([1, 2, 3, 4].map((r) => val(ctl, r, 3))).toEqual([255, 185, 158, 185]);
  });
  it("korumalı sayfada kilitli hücre düzenlenemez, kilidi açık hücre düzenlenir", () => {
    const ctl = sample();
    const alerts: string[] = [];
    ctl.alert = (m) => alerts.push(m);
    ctl.selectCell(1, 1);
    F.applyStyle(ctl, { locked: false });
    ctl.sheet.protected = true;
    ctl.startEdit("5");
    expect(ctl.edit).not.toBeNull();
    ctl.commitEdit(null);
    expect(val(ctl, 1, 1)).toBe(5);
    ctl.selectCell(1, 2);
    ctl.startEdit("7");
    expect(ctl.edit).toBeNull();
    expect(alerts.length).toBe(1);
  });
});

describe("geri alma birleştirme", () => {
  it("birden çok adım tek geri almada geri alınır", () => {
    const ctl = sample();
    const mark = ctl.ops.history.mark();
    ctl.selectCell(1, 1);
    F.applyStyle(ctl, { bold: true });
    F.borders(ctl, "outside");
    ctl.ops.history.squash(mark, "Hücreleri Biçimlendir");
    expect(ctl.ops.history.undoLabel).toBe("Hücreleri Biçimlendir");
    E.undo(ctl);
    const st = ctl.wb.styleAt(ctl.sheet, 1, 1);
    expect(st.bold).toBeUndefined();
    expect(st.bTop).toBeUndefined();
    E.redo(ctl);
    expect(ctl.wb.styleAt(ctl.sheet, 1, 1).bold).toBe(true);
    expect(ctl.wb.styleAt(ctl.sheet, 1, 1).bTop).toBeTruthy();
  });
});

describe("veri araçları", () => {
  it("çok düzeyli sıralama başlığı korur", () => {
    const ctl = sample();
    D.sortRange(ctl, { r1: 0, c1: 0, r2: 4, c2: 2 }, [{ col: 1, desc: true }, { col: 0, desc: false }], true);
    expect([1, 2, 3, 4].map((r) => val(ctl, r, 0))).toEqual(["Marmara", "Ege", "Ege", "Akdeniz"]);
    expect(val(ctl, 0, 0)).toBe("Bölge");
  });
  it("Türkçe alfabetik sıralama (Ç, Ğ, İ, Ö, Ş, Ü)", () => {
    const ctl = new Controller();
    const words = ["Zeytin", "Çilek", "Ayva", "İncir", "Üzüm", "Şeftali", "Ihlamur", "Ödül"];
    ctl.ops.setInputs(ctl.sheet, words.map((w, i) => ({ row: i, col: 0, input: w })), "x");
    D.sortRange(ctl, { r1: 0, c1: 0, r2: words.length - 1, c2: 0 }, [{ col: 0, desc: false }], false);
    expect(words.map((_, i) => val(ctl, i, 0))).toEqual(["Ayva", "Çilek", "Ihlamur", "İncir", "Ödül", "Şeftali", "Üzüm", "Zeytin"]);
  });
  it("yinelenenleri kaldırır", () => {
    const ctl = sample();
    const res = D.removeDuplicates(ctl, { r1: 0, c1: 0, r2: 4, c2: 2 }, [0, 1, 2], true);
    expect(res).toEqual({ removed: 1, kept: 3 });
    expect(val(ctl, 4, 0)).toBeNull();
  });
  it("metni sütunlara dönüştürür (tırnak içindeki ayırıcıyı korur)", () => {
    expect(D.splitDelimited('a;"b;c";d', [";"], false)).toEqual(["a", "b;c", "d"]);
    expect(D.splitDelimited("a  b", [" "], true)).toEqual(["a", "b"]);
    const ctl = new Controller();
    ctl.ops.setInputs(ctl.sheet, [{ row: 0, col: 0, input: "Ali;34;12,5" }], "x");
    D.textToColumns(ctl, { r1: 0, c1: 0, r2: 0, c2: 0 }, [";"], false);
    expect([0, 1, 2].map((c) => val(ctl, 0, c))).toEqual(["Ali", 34, 12.5]);
  });
  it("otomatik filtre ve özel ölçüt satırları gizler", () => {
    const ctl = sample();
    ctl.selectCell(0, 0);
    D.toggleAutoFilter(ctl);
    const f = ctl.sheet.autoFilter!;
    expect(f.range).toEqual({ r1: 0, c1: 0, r2: 4, c2: 2 });
    D.setColumnFilter(ctl, f, 1, { kind: "custom", op1: "ge", v1: "90" });
    expect([...ctl.sheet.filterHidden]).toEqual([3]);
    D.setColumnFilter(ctl, f, 0, { kind: "values", values: ["Ege"], blanks: false });
    expect([...ctl.sheet.filterHidden].sort()).toEqual([1, 3]);
    // ALTTOPLAM gizli satırları atlar
    ctl.ops.enterText(ctl.sheet, [{ row: 6, col: 1 }], "=ALTTOPLAM(109;B2:B5)");
    expect(val(ctl, 6, 1)).toBe(180);
    D.clearAllFilters(ctl);
    expect(ctl.sheet.filterHidden.size).toBe(0);
    expect(val(ctl, 6, 1)).toBe(370);
  });
  it("tablo oluşturma ve toplam satırı", () => {
    const ctl = sample();
    D.createTable(ctl, { r1: 0, c1: 0, r2: 4, c2: 2 }, true);
    const t = ctl.sheet.tables[0];
    expect(t.name).toBe("Tablo1");
    D.updateTable(ctl, t, { totalRow: true });
    expect(ctl.sheet.tables[0].ref.r2).toBe(5);
    expect(val(ctl, 5, 2)).toBe(413);
  });
});

describe("pano", () => {
  it("TSV ayrıştırma (tırnak ve satır sonu)", () => {
    expect(parseTsv('a\t"b\nc"\r\n1\t2\r\n')).toEqual([
      ["a", "b\nc"],
      ["1", "2"],
    ]);
  });
  it("kopyalanan formül göreli kayar, değer yapıştırma formülü atar", () => {
    const ctl = sample();
    ctl.ops.enterText(ctl.sheet, [{ row: 1, col: 3 }], "=B2*2");
    const { payload, tsv } = buildClipboard(ctl.wb, ctl.sheet, { r1: 1, c1: 3, r2: 1, c2: 3 });
    expect(tsv.trim()).toBe("240");
    pastePayload(ctl.ops, ctl.sheet, { r1: 2, c1: 3, r2: 3, c2: 3 }, payload);
    expect(ctl.wb.getFormula(ctl.sheet, 3, 3)).toBe("=B4*2");
    expect(val(ctl, 3, 3)).toBe(140);
    pastePayload(ctl.ops, ctl.sheet, { r1: 2, c1: 4, r2: 2, c2: 4 }, payload, { what: "values" });
    expect(ctl.wb.getFormula(ctl.sheet, 2, 4)).toBeUndefined();
    expect(val(ctl, 2, 4)).toBe(240);
  });
});

describe("bul ve değiştir", () => {
  it("joker karakterli arama ve büyük/küçük harf", () => {
    const ctl = sample();
    const hits = findAll(ctl.wb, [ctl.sheet], { query: "e*", matchCase: false, wholeCell: true, lookIn: "values", byColumns: false });
    expect(hits.map((h) => h.row)).toEqual([2, 4]);
    expect(findAll(ctl.wb, [ctl.sheet], { query: "ege", matchCase: true, wholeCell: false, lookIn: "values", byColumns: false })).toHaveLength(0);
    expect(findAll(ctl.wb, [ctl.sheet], { query: "?ge", matchCase: false, wholeCell: true, lookIn: "values", byColumns: false })).toHaveLength(2);
  });
  it("formüllerde arama yerel adla yapılır", () => {
    const ctl = sample();
    ctl.ops.enterText(ctl.sheet, [{ row: 6, col: 1 }], "=TOPLA(B2:B3)");
    const hits = findAll(ctl.wb, [ctl.sheet], { query: "TOPLA", matchCase: false, wholeCell: false, lookIn: "formulas", byColumns: false });
    expect(hits).toHaveLength(1);
    const n = ctl.ops.replaceAll(ctl.sheet, [{ row: 6, col: 1 }], /TOPLA/giu, "ORTALAMA");
    expect(n).toBe(1);
    expect(val(ctl, 6, 1)).toBe(105);
  });
});

describe("sayı biçimi kategorileri", () => {
  it("kod üretimi ve geri tanıma", () => {
    const o = { decimals: 2, thousands: true, negative: 1, symbol: "₺" };
    expect(buildCode("number", o)).toBe("#,##0.00;[Red]#,##0.00");
    expect(detectCategory("#,##0.00;[Red]#,##0.00")).toEqual({ cat: "number", opts: o });
    expect(buildCode("currency", { ...o, negative: 0 })).toBe('#,##0.00 "₺"');
    expect(detectCategory('#,##0.00 "₺"').cat).toBe("currency");
    expect(detectCategory("0.0%")).toEqual({ cat: "percent", opts: { decimals: 1, thousands: true, negative: 0, symbol: "₺" } });
    expect(detectCategory("dd.mm.yyyy").cat).toBe("date");
    expect(detectCategory("hh:mm").cat).toBe("time");
    expect(detectCategory("@").cat).toBe("text");
    expect(detectCategory('0.00 "kg"').cat).toBe("custom");
    expect(formatValue(-1234.1, buildCode("currency", { decimals: 2, thousands: true, negative: 2, symbol: "₺" })).text).toBe("-1.234,10 ₺");
    const acc = buildCode("accounting", { decimals: 2, thousands: true, negative: 0, symbol: "₺" });
    expect(detectCategory(acc).cat).toBe("accounting");
  });
});

describe("yazdırma ve görünüm", () => {
  it("üst/alt bilgi kodları", () => {
    expect(expandHeaderFooter("Sayfa &[Sayfa] / &[Sayfalar] - &[Sekme]", { page: 2, pages: 5, sheet: "Gelir", file: "x" })).toBe("Sayfa 2 / 5 - Gelir");
    expect(hfCss("Sayfa &[Sayfa] / &[Sayfalar]", "S", "f")).toBe('"Sayfa " counter(page) " / " counter(pages)');
  });
  it("sayfa sonları ve tek sayfaya sığdırma", () => {
    const wb = new WorkbookModel();
    const s = wb.active;
    for (let r = 0; r < 120; r++) for (let c = 0; c < 20; c++) wb.setInputRaw(s, r, c, r * c);
    const b = computePageBreaks(wb, s);
    expect(b.rows.length).toBeGreaterThan(1);
    expect(b.cols.length).toBeGreaterThan(1);
    s.page.fitToWidth = true;
    expect(computePageBreaks(wb, s).cols).toEqual([20]);
  });
  it("bölmeleri dondur ve notlar arasında gezinme", () => {
    const ctl = sample();
    ctl.selectCell(1, 1);
    freeze(ctl, "panes");
    expect(ctl.sheet.frozen).toEqual({ rows: 1, cols: 1 });
    freeze(ctl, "none");
    expect(ctl.sheet.frozen).toEqual({ rows: 0, cols: 0 });
    ctl.ops.setNote(ctl.sheet, 3, 2, "not 1");
    ctl.ops.setNote(ctl.sheet, 1, 0, "not 2");
    ctl.selectCell(0, 0);
    gotoNote(ctl, 1);
    expect(ctl.sel.active).toEqual({ row: 1, col: 0 });
    gotoNote(ctl, 1);
    expect(ctl.sel.active).toEqual({ row: 3, col: 2 });
  });
  it("grafik ekleme geçerli bölgeyi kullanır", () => {
    const ctl = sample();
    ctl.selectCell(2, 1);
    expect(currentRegion(ctl.wb, ctl.sheet, { row: 2, col: 1 })).toEqual({ r1: 0, c1: 0, r2: 4, c2: 2 });
    insertChart(ctl, "column");
    const ch = ctl.sheet.charts[0];
    expect(ch.range).toBe("Sayfa1!$A$1:$C$5");
    const data = extractChartData(ctl.wb, ch);
    expect(data.series.map((s) => s.name)).toEqual(["Ocak", "Şubat"]);
    expect(data.categories).toEqual(["Marmara", "Ege", "Akdeniz", "Ege"]);
    E.undo(ctl);
    expect(ctl.sheet.charts).toHaveLength(0);
  });
});
