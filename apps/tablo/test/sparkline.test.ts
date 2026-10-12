import { describe, expect, it } from "vitest";
import type { Scalar } from "@officehere/formula";
import { Controller } from "../src/controller";
import { createSparklines, clearSparklines, groupSparklines, planItems, ungroupSparklines } from "../src/sparkline/sparkOps";
import { sparkAt, sparkPrims, sparkValues } from "../src/sparkline/geometry";
import { shiftRangeText, renameRangeText } from "../src/model/rangeRefs";
import type { SparklineGroup } from "../src/model/types";

function setup(): Controller {
  const ctl = new Controller();
  const sh = ctl.sheet;
  const rows: Scalar[][] = [
    ["Ürün", "Oca", "Şub", "Mar", "Nis"],
    ["Kalem", 10, 14, 9, 20],
    ["Defter", -5, 3, -2, 8],
    ["Silgi", 1, null, 3, 4],
  ];
  rows.forEach((r, i) => r.forEach((v, j) => v !== null && ctl.wb.setInputRaw(sh, i, j, v)));
  return ctl;
}

const G = (p: Partial<SparklineGroup>): SparklineGroup => ({ id: "g", type: "line", items: [], color: "#123456", ...p });

describe("Mini grafik yerleşimi", () => {
  it("dikey konumda veri satırlara, yatay konumda sütunlara paylaştırılır", () => {
    const ctl = setup();
    const data = { sheet: ctl.sheet, rect: { r1: 1, c1: 1, r2: 3, c2: 4 } };
    const v = planItems(data, { r1: 1, c1: 5, r2: 3, c2: 5 })!;
    expect(v.map((i) => [i.row, i.col, i.range])).toEqual([
      [1, 5, "Sayfa1!B2:E2"],
      [2, 5, "Sayfa1!B3:E3"],
      [3, 5, "Sayfa1!B4:E4"],
    ]);
    const h = planItems(data, { r1: 5, c1: 1, r2: 5, c2: 4 })!;
    expect(h.map((i) => i.range)).toEqual(["Sayfa1!B2:B4", "Sayfa1!C2:C4", "Sayfa1!D2:D4", "Sayfa1!E2:E4"]);
    // Boyut eşleşmezse ya da konum iki boyutluysa geçersiz
    expect(planItems(data, { r1: 1, c1: 5, r2: 2, c2: 5 })).toBeNull();
    expect(planItems(data, { r1: 1, c1: 5, r2: 3, c2: 6 })).toBeNull();
    expect(planItems(data, { r1: 1, c1: 5, r2: 1, c2: 5 })).toBeNull();
  });

  it("oluşturma, geri alma, temizleme, gruplama", () => {
    const ctl = setup();
    expect(createSparklines(ctl, "line", "B2:E4", "$F$2:$F$4")).toBeNull();
    const sh = ctl.sheet;
    expect(sh.sparklines).toHaveLength(1);
    expect(sh.sparklines[0].items).toHaveLength(3);
    expect(sparkAt(sh, 2, 5)?.item.range).toBe("Sayfa1!B3:E3");
    expect(createSparklines(ctl, "column", "B2:E4", "F2:G3")).toMatch(/geçerli değil/);
    // Aynı hücreye yeni grup eskisinin yerini alır
    expect(createSparklines(ctl, "column", "B2:E2", "F2")).toBeNull();
    expect(sh.sparklines.map((g) => [g.type, g.items.length])).toEqual([
      ["line", 2],
      ["column", 1],
    ]);
    ctl.ops.undo();
    expect(sh.sparklines.map((g) => g.items.length)).toEqual([3]);
    ctl.selectRange({ r1: 1, c1: 5, r2: 3, c2: 5 });
    ungroupSparklines(ctl);
    expect(sh.sparklines).toHaveLength(3);
    groupSparklines(ctl);
    expect(sh.sparklines).toHaveLength(1);
    ctl.selectCell(2, 5);
    clearSparklines(ctl, false);
    expect(sh.sparklines[0].items.map((i) => i.row)).toEqual([1, 3]);
    clearSparklines(ctl, true);
    expect(sh.sparklines).toHaveLength(1);
    ctl.selectCell(1, 5);
    clearSparklines(ctl, true);
    expect(sh.sparklines).toHaveLength(0);
  });

  it("satır ekleme ve sayfa adı değişikliği veri aralığını günceller", () => {
    const ctl = setup();
    createSparklines(ctl, "line", "B2:E3", "F2:F3");
    const sh = ctl.sheet;
    ctl.ops.insertRows(sh, 0, 2);
    expect(sh.sparklines[0].items.map((i) => [i.row, i.range])).toEqual([
      [3, "Sayfa1!B4:E4"],
      [4, "Sayfa1!B5:E5"],
    ]);
    ctl.ops.renameSheet(0, "Satış 2026");
    expect(sh.sparklines[0].items[0].range).toBe("'Satış 2026'!B4:E4");
    expect(sparkValues(ctl.wb, sh, sh.sparklines[0], sh.sparklines[0].items[0])).toEqual([10, 14, 9, 20]);
  });

  it("doldurma tutamacı mini grafiği göreli aralıkla kopyalar; geri alınabilir", () => {
    const ctl = setup();
    createSparklines(ctl, "line", "B2:E2", "F2");
    const sh = ctl.sheet;
    ctl.ops.fill(sh, { r1: 1, c1: 5, r2: 1, c2: 5 }, { r1: 1, c1: 5, r2: 3, c2: 5 });
    expect(sh.sparklines).toHaveLength(1);
    expect(sh.sparklines[0].items.map((i) => [i.row, i.range])).toEqual([
      [1, "Sayfa1!B2:E2"],
      [2, "Sayfa1!B3:E3"],
      [3, "Sayfa1!B4:E4"],
    ]);
    ctl.ops.undo();
    expect(sh.sparklines[0].items).toHaveLength(1);
    ctl.ops.redo();
    expect(sh.sparklines[0].items).toHaveLength(3);
  });

  it("aralık metni yardımcıları", () => {
    expect(shiftRangeText("Sayfa1!$A$1:$C$5", "Sayfa1", "Sayfa1", "row", 0, 1)).toBe("Sayfa1!$A$2:$C$6");
    expect(shiftRangeText("A1:C5", "Sayfa1", "Sayfa2", "row", 0, 1)).toBe("A1:C5");
    expect(shiftRangeText("A1:C5", "Sayfa1", "Sayfa1", "col", 1, -1)).toBe("A1:B5");
    expect(renameRangeText("'Eski Ad'!B2:E2", "Eski Ad", "Yeni")).toBe("Yeni!B2:E2");
  });
});

describe("Mini grafik geometrisi", () => {
  const box = { x: 0, y: 0, w: 100, h: 20 };
  it("boş hücre boşluk, sıfır ya da bağlantı olarak çizilir", () => {
    const ctl = setup();
    const sh = ctl.sheet;
    const item = { row: 3, col: 5, range: "B4:E4" };
    expect(sparkValues(ctl.wb, sh, G({}), item)).toEqual([1, null, 3, 4]);
    expect(sparkValues(ctl.wb, sh, G({ emptyAs: "zero" }), item)).toEqual([1, 0, 3, 4]);
    const gap = sparkPrims(G({}), [1, null, 3, 4], box, 1, null).filter((p) => p.k === "path");
    expect(gap).toHaveLength(1); // tek noktalı ilk parça çizilmez, 3→4 parçası çizilir
    const span = sparkPrims(G({ emptyAs: "span" }), [1, null, 3, 4], box, 1, null).filter((p) => p.k === "path");
    expect(span).toHaveLength(1);
    expect(span[0].k === "path" && span[0].pts.length).toBe(3);
  });

  it("işaretçiler ve vurgu noktaları", () => {
    const prims = sparkPrims(G({ highPoint: true, lowPoint: true, highColor: "#00FF00", lowColor: "#0000FF" }), [5, 9, 2, 7], box, 1, null);
    const dots = prims.filter((p) => p.k === "dot");
    expect(dots.map((d) => d.k === "dot" && d.color)).toEqual(["#00FF00", "#0000FF"]);
    const all = sparkPrims(G({ markers: true }), [5, 9, 2, 7], box, 1, null).filter((p) => p.k === "dot");
    expect(all).toHaveLength(4);
  });

  it("sütun ve kazanç/kayıp", () => {
    const cols = sparkPrims(G({ type: "column", negative: true, negativeColor: "#FF0000" }), [-5, 3, -2, 8], box, 1, null).filter((p) => p.k === "rect");
    expect(cols).toHaveLength(4);
    expect(cols.map((p) => p.k === "rect" && p.color)).toEqual(["#FF0000", "#123456", "#FF0000", "#123456"]);
    const wl = sparkPrims(G({ type: "winLoss", negativeColor: "#FF0000" }), [-5, 3, 0, 8], box, 1, null).filter((p) => p.k === "rect");
    expect(wl).toHaveLength(3);
    // Kazançlar üst yarıda, kayıplar alt yarıda
    const [a, b] = wl;
    expect(a.k === "rect" && b.k === "rect" && a.y > b.y).toBe(true);
  });
});

describe("Mini grafik XLSX gidiş-dönüşü", () => {
  it("özel XML bölümünde saklanır ve geri okunur", async () => {
    const { writeXlsx, readXlsx } = await import("../src/io/xlsx");
    const JSZip = (await import("jszip")).default;
    const ctl = setup();
    createSparklines(ctl, "column", "B2:E4", "F2:F4");
    const g = ctl.sheet.sparklines[0];
    g.highPoint = true;
    g.highColor = "#00B050";
    g.weight = 1.5;
    g.minAxis = "group";
    const bytes = await writeXlsx(ctl.wb);
    const zip = await JSZip.loadAsync(bytes);
    const item = zip.file("customXml/item1.xml");
    expect(item).toBeTruthy();
    expect(await item!.async("string")).toContain('<sparkline ref="F3" data="Sayfa1!B3:E3"/>');
    expect(zip.file("customXml/itemProps1.xml")).toBeTruthy();
    expect(await zip.file("xl/_rels/workbook.xml.rels")!.async("string")).toContain('Target="../customXml/item1.xml"');
    expect(await zip.file("[Content_Types].xml")!.async("string")).toContain("customXmlProperties+xml");
    const back = await readXlsx(bytes.buffer as ArrayBuffer);
    const s = back.sheets[0].sparklines;
    expect(s).toHaveLength(1);
    expect(s[0]).toMatchObject({ type: "column", highPoint: true, highColor: "#00B050", weight: 1.5, minAxis: "group" });
    expect(s[0].items.map((i) => [i.row, i.col])).toEqual([
      [1, 5],
      [2, 5],
      [3, 5],
    ]);
  });
});
