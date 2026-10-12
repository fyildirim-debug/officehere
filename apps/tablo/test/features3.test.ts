/** 3. tur özellikleri: anahat, alt toplam, hedef ara, akıllı doldurma, özel git, simge kümeleri, yazdırma sayfaları. */
import { describe, expect, it } from "vitest";
import { Controller } from "../src/controller";
import { addSubtotals, clearOutline, groupSelection, removeSubtotals, showOutlineLevel, toggleGroup, ungroupSelection } from "../src/actions/outline";
import { outlineGroups } from "../src/model/outline";
import { flashFill, runGoalSeek, applyGoalSeek } from "../src/actions/whatIf";
import { learnProgram, runProgram } from "../src/model/flashFill";
import { goSpecial } from "../src/model/goSpecial";
import { readXlsx, writeXlsx } from "../src/io/xlsx";
import { buildPrintHtml, printPages } from "../src/io/printPages";
import { fitScale } from "../src/io/print";
import { CondFormatter } from "../src/model/condfmt";

function setRows(ctl: Controller, rows: (string | number | null)[][]) {
  ctl.ops.setInputs(ctl.sheet, rows.flatMap((r, i) => r.map((v, j) => ({ row: i, col: j, input: v }))), "x");
}
const v = (ctl: Controller, r: number, c: number) => ctl.wb.getValue(ctl.sheet, r, c);

describe("Anahat: gruplandır, daralt, düzeyler", () => {
  it("satırları gruplandırır, daraltır/genişletir, düzey düğmeleri ve geri alma", () => {
    const ctl = new Controller();
    const s = ctl.sheet;
    ctl.selectRows(2, 5);
    groupSelection(ctl);
    expect(s.rows.level(2)).toBe(1);
    expect(s.rows.level(5)).toBe(1);
    ctl.selectRows(3, 4);
    groupSelection(ctl);
    expect(s.rows.level(3)).toBe(2);
    const gs = outlineGroups(s.rows);
    expect(gs.map((g) => [g.level, g.start, g.end, g.summary])).toEqual([
      [1, 2, 5, 6],
      [2, 3, 4, 5],
    ]);
    showOutlineLevel(ctl, "row", 1);
    expect([2, 3, 4, 5].every((r) => s.rows.isHidden(r))).toBe(true);
    showOutlineLevel(ctl, "row", 2);
    expect(s.rows.isHidden(2)).toBe(false);
    expect(s.rows.isHidden(3)).toBe(true);
    // İçteki grup daraltılmışken dış grubu daraltıp açmak içtekini kapalı bırakır
    toggleGroup(ctl, "row", outlineGroups(s.rows)[0], true);
    toggleGroup(ctl, "row", outlineGroups(s.rows)[0], false);
    expect(s.rows.isHidden(2)).toBe(false);
    expect(s.rows.isHidden(3)).toBe(true);
    ctl.ops.undo();
    expect(s.rows.isHidden(2)).toBe(true);
    ctl.selectRows(2, 5);
    ungroupSelection(ctl);
    expect(s.rows.level(2)).toBe(0);
    expect(s.rows.level(3)).toBe(1);
    clearOutline(ctl);
    expect(s.rows.maxLevel()).toBe(0);
    expect(s.rows.isHidden(3)).toBe(false);
  });

  it("sütun gruplandırma ve XLSX outlineLevel gidiş-dönüşü", async () => {
    const ctl = new Controller();
    setRows(ctl, [[1, 2, 3, 4]]);
    ctl.selectColumns(1, 2);
    groupSelection(ctl);
    ctl.selectRows(3, 4);
    groupSelection(ctl);
    showOutlineLevel(ctl, "row", 1);
    const wb2 = await readXlsx((await writeXlsx(ctl.wb)).buffer);
    const s2 = wb2.sheets[0];
    expect(s2.cols.level(1)).toBe(1);
    expect(s2.cols.level(2)).toBe(1);
    expect(s2.rows.level(3)).toBe(1);
    expect(s2.rows.isHidden(4)).toBe(true);
  });
});

describe("Alt Toplam", () => {
  it("her değişiklikte ALTTOPLAM satırı ve genel toplam ekler; 1-2-3 düzeyleri; tümünü kaldır", () => {
    const ctl = new Controller();
    setRows(ctl, [
      ["Bölge", "Ürün", "Tutar"],
      ["Ege", "A", 10],
      ["Ege", "B", 20],
      ["Marmara", "A", 5],
      ["Marmara", "C", 7],
      ["Marmara", "D", 3],
    ]);
    ctl.selectCell(0, 0);
    addSubtotals(ctl, { range: { r1: 0, c1: 0, r2: 5, c2: 2 }, groupCol: 0, fn: 9, cols: [2] }, true);
    expect(v(ctl, 3, 0)).toBe("Ege Toplam");
    expect(v(ctl, 3, 2)).toBe(30);
    expect(v(ctl, 7, 0)).toBe("Marmara Toplam");
    expect(v(ctl, 7, 2)).toBe(15);
    expect(v(ctl, 8, 0)).toBe("Genel Toplam");
    expect(v(ctl, 8, 2)).toBe(45);
    expect(ctl.sheet.rows.maxLevel()).toBe(2);
    showOutlineLevel(ctl, "row", 2);
    expect(ctl.sheet.rows.isHidden(1)).toBe(true);
    expect(ctl.sheet.rows.isHidden(3)).toBe(false);
    expect(v(ctl, 8, 2)).toBe(45);
    removeSubtotals(ctl, { r1: 0, c1: 0, r2: 8, c2: 2 });
    expect(v(ctl, 3, 0)).toBe("Marmara");
    expect(v(ctl, 6, 0)).toBe(null);
    expect(ctl.sheet.rows.maxLevel()).toBe(0);
    ctl.ops.undo();
    expect(v(ctl, 8, 0)).toBe("Genel Toplam");
  });
});

describe("Hedef Ara", () => {
  it("formülün istenen değeri vermesi için girdiyi bulur", () => {
    const ctl = new Controller();
    setRows(ctl, [[100, null]]);
    ctl.ops.setInputs(ctl.sheet, [{ row: 0, col: 1, input: { formula: "=A1*1.18" } }], "f");
    const res = runGoalSeek(ctl, { row: 0, col: 1 }, 236, { row: 0, col: 0 });
    expect(res.found).toBe(true);
    expect(res.value).toBeCloseTo(200, 3);
    // Arama özgün değeri bozmaz; uygulama geri alınabilir
    expect(v(ctl, 0, 0)).toBe(100);
    applyGoalSeek(ctl, { row: 0, col: 0 }, res.value);
    expect(v(ctl, 0, 1) as number).toBeCloseTo(236, 3);
    ctl.ops.undo();
    expect(v(ctl, 0, 0)).toBe(100);
  });
  it("doğrusal olmayan formül (kare kök)", () => {
    const ctl = new Controller();
    setRows(ctl, [[1]]);
    ctl.ops.setInputs(ctl.sheet, [{ row: 0, col: 1, input: { formula: "=A1^2+3*A1" } }], "f");
    const res = runGoalSeek(ctl, { row: 0, col: 1 }, 40, { row: 0, col: 0 });
    expect(res.found).toBe(true);
    expect(Math.abs((res.value * res.value + 3 * res.value) - 40)).toBeLessThan(0.001);
  });
});

describe("Akıllı Doldurma", () => {
  it("desen öğrenme: ayır, birleştir, baş harf, büyük harf", () => {
    const p1 = learnProgram([{ row: ["Ahmet Yılmaz"], output: "Ahmet" }]);
    expect(runProgram(p1!, ["Ayşe Demir"])).toBe("Ayşe");
    const p2 = learnProgram([{ row: ["ahmet", "yılmaz"], output: "Yılmaz, Ahmet" }]);
    expect(runProgram(p2!, ["ayşe", "demir"])).toBe("Demir, Ayşe");
    const p3 = learnProgram([
      { row: ["Ahmet Yılmaz"], output: "A. YILMAZ" },
      { row: ["Can Öz"], output: "C. ÖZ" },
    ]);
    expect(runProgram(p3!, ["İrem Işık"])).toBe("İ. IŞIK");
    const p4 = learnProgram([{ row: ["ali.kaya@ornek.com"], output: "ali.kaya" }]);
    expect(runProgram(p4!, ["veli@ornek.com"])).toBe("veli");
  });

  it("Ctrl+E sütunu örnekten doldurur (geri alınabilir)", () => {
    const ctl = new Controller();
    setRows(ctl, [
      ["Ad Soyad", "Ad"],
      ["Ahmet Yılmaz", "Ahmet"],
      ["Ayşe Demir", null],
      ["Mehmet Can Öz", null],
    ]);
    ctl.selectCell(2, 1);
    expect(flashFill(ctl)).toBe(2);
    expect(v(ctl, 2, 1)).toBe("Ayşe");
    expect(v(ctl, 3, 1)).toBe("Mehmet");
    ctl.ops.undo();
    expect(v(ctl, 2, 1)).toBe(null);
  });
});

describe("Özel Git", () => {
  it("boşluklar, formüller, sabitler (alt türler), son hücre", () => {
    const ctl = new Controller();
    setRows(ctl, [
      [1, "a", null],
      [null, true, 3],
    ]);
    ctl.ops.setInputs(ctl.sheet, [{ row: 2, col: 2, input: { formula: "=A1+C2" } }], "f");
    const s = ctl.sheet;
    const all = [{ r1: 0, c1: 0, r2: 0, c2: 0 }];
    const blanks = goSpecial(ctl.wb, s, all, { row: 0, col: 0 }, "blanks")!;
    expect(blanks.ranges.length).toBeGreaterThan(0);
    const f = goSpecial(ctl.wb, s, all, { row: 0, col: 0 }, "formulas")!;
    expect(f.ranges).toEqual([{ r1: 2, c1: 2, r2: 2, c2: 2 }]);
    const nums = goSpecial(ctl.wb, s, all, { row: 0, col: 0 }, "constants", { text: false, logicals: false, errors: false })!;
    expect(nums.ranges).toEqual(expect.arrayContaining([{ r1: 0, c1: 0, r2: 0, c2: 0 }, { r1: 1, c1: 2, r2: 1, c2: 2 }]));
    expect(nums.ranges).toHaveLength(2);
    expect(goSpecial(ctl.wb, s, all, { row: 0, col: 0 }, "lastCell")!.ranges).toEqual([{ r1: 2, c1: 2, r2: 2, c2: 2 }]);
    const prec = goSpecial(ctl.wb, s, [{ r1: 2, c1: 2, r2: 2, c2: 2 }], { row: 2, col: 2 }, "precedents")!;
    expect(prec.ranges).toHaveLength(2);
  });
});

describe("Simge kümesi", () => {
  it("yüzde eşiklerine göre simge dizini", () => {
    const ctl = new Controller();
    setRows(ctl, [[0], [50], [100]]);
    ctl.sheet.condFormats.push({ id: "cf1", type: "iconSet", iconSet: "3Arrows", ranges: [{ r1: 0, c1: 0, r2: 2, c2: 0 }] });
    const cf = new CondFormatter(ctl.wb);
    expect(cf.evaluate(ctl.sheet, 0, 0)?.icon?.index).toBe(0);
    expect(cf.evaluate(ctl.sheet, 1, 0)?.icon?.index).toBe(1);
    expect(cf.evaluate(ctl.sheet, 2, 0)?.icon?.index).toBe(2);
  });
});

describe("Yazdırma sayfaları", () => {
  it("sayfa sayısı, sığdırma ölçeği, sayfa aralığı ve kopyalar", () => {
    const ctl = new Controller();
    const s = ctl.sheet;
    ctl.ops.setInputs(s, Array.from({ length: 120 }, (_, i) => ({ row: i, col: 0, input: i })), "x");
    ctl.ops.setInputs(s, [{ row: 0, col: 25, input: "son" }], "x");
    const pages = printPages(ctl.wb, s);
    expect(pages.length).toBeGreaterThan(2);
    s.page.fitMode = "sheet";
    expect(printPages(ctl.wb, s)).toHaveLength(1);
    expect(fitScale(s, { r1: 0, c1: 0, r2: 119, c2: 25 })).toBeLessThan(1);
    s.page.fitMode = undefined;
    const html = buildPrintHtml(ctl.wb, s, { from: 1, to: 2, copies: 2 });
    expect(html.match(/class="tb-print-page"/g)).toHaveLength(4);
  });
});

describe("Sayfa sonları (el ile)", () => {
  it("sayfalamaya katılır, satır eklemeyle kayar, sürüklenir, ekle/kaldır ve XLSX gidiş-dönüşü", async () => {
    const { computePageBreaks } = await import("../src/io/print");
    const { pageBreakDrop, insertPageBreak, resetPageBreaks } = await import("../src/grid/pageBreakDrag");
    const { injectBreaks, parseBreaks } = await import("../src/io/xlsxBreaks");
    const ctl = new Controller();
    const s = ctl.sheet;
    ctl.ops.setInputs(s, Array.from({ length: 30 }, (_, i) => ({ row: i, col: 0, input: i })), "x");
    ctl.ops.setInputs(s, [{ row: 0, col: 3, input: "d" }], "x");
    expect(computePageBreaks(ctl.wb, s).rows).toEqual([30]);
    // Ekle: etkin hücrenin üstüne ve soluna
    ctl.selectCell(10, 2);
    insertPageBreak(ctl, false);
    expect(s.page.rowBreaks).toEqual([10]);
    expect(s.page.colBreaks).toEqual([2]);
    const pb = computePageBreaks(ctl.wb, s);
    expect(pb.rows).toEqual([10, 30]);
    expect(pb.manualRows).toEqual([10]);
    expect(pb.cols).toEqual([2, 4]);
    expect(printPages(ctl.wb, s)).toHaveLength(4);
    // Satır eklenince sayfa sonu kayar; geri alınınca döner
    ctl.ops.insertRows(s, 3, 2);
    expect(s.page.rowBreaks).toEqual([12]);
    ctl.ops.undo();
    expect(s.page.rowBreaks).toEqual([10]);
    // Sürükle: 10 → 15; yazdırma alanı dışına sürüklenince kaldırılır
    pageBreakDrop(ctl, { axis: "row", from: 10, kind: "break", at: 15 });
    expect(s.page.rowBreaks).toEqual([15]);
    pageBreakDrop(ctl, { axis: "row", from: 15, kind: "break", at: 40 });
    expect(s.page.rowBreaks).toEqual([]);
    ctl.ops.undo();
    expect(s.page.rowBreaks).toEqual([15]);
    // Yazdırma alanının alt kenarını sürükle
    pageBreakDrop(ctl, { axis: "row", from: 30, kind: "end", at: 20 });
    expect(s.page.printArea).toEqual({ r1: 0, c1: 0, r2: 19, c2: 3 });
    // XLSX
    const bytes = await writeXlsx(ctl.wb);
    const again = await readXlsx(bytes.buffer);
    expect(again.sheets[0].page.rowBreaks).toEqual([15]);
    expect(again.sheets[0].page.colBreaks).toEqual([2]);
    resetPageBreaks(ctl);
    expect(s.page.rowBreaks).toBeUndefined();
    // XML yerleşimi: sayfa sonları drawing/tableParts'tan önce
    const xml = injectBreaks('<worksheet><sheetData/><pageMargins/><drawing r:id="rId1"/><tableParts count="0"/></worksheet>', [5], []);
    expect(xml).toBe('<worksheet><sheetData/><pageMargins/><rowBreaks count="1" manualBreakCount="1"><brk id="5" max="16383" man="1"/></rowBreaks><drawing r:id="rId1"/><tableParts count="0"/></worksheet>');
    expect(parseBreaks(xml, "rowBreaks")).toEqual([5]);
  });
});
