import { describe, expect, it } from "vitest";
import type { Scalar } from "@officehere/formula";
import { computePivot, newPivotDef, type PivotSource } from "../src/pivot/compute";
import type { PivotDef } from "../src/model/types";

const src: PivotSource = {
  headers: ["Bölge", "Ürün", "Yıl", "Satış", "Adet"],
  records: [
    ["Ege", "Kalem", 2024, 100, 1],
    ["Ege", "Defter", 2024, 50, 2],
    ["Ege", "Kalem", 2025, 30, 3],
    ["Marmara", "Kalem", 2024, 200, 4],
    ["marmara", "Silgi", 2025, 20, 5],
    ["Akdeniz", "Defter", 2025, 70, 6],
    ["Akdeniz", null, 2025, 5, 7],
  ],
  numFmts: [undefined, undefined, undefined, "#,##0", undefined],
};

function grid(def: PivotDef): string[][] {
  const lay = computePivot(def, src);
  const g: string[][] = Array.from({ length: lay.rows }, () => Array.from({ length: lay.cols }, () => ""));
  for (const c of lay.cells) g[c.row][c.col] = c.value === null ? "" : `${"  ".repeat(c.indent ?? 0)}${fmt(c.value)}`;
  return g;
}

function fmt(v: Scalar): string {
  if (typeof v === "number") return String(Math.round(v * 100) / 100);
  if (v && typeof v === "object") return (v as { code: string }).code;
  return String(v);
}

function def(p: Partial<PivotDef>): PivotDef {
  return { ...newPivotDef("p1", "PivotTable1", 1, { r1: 0, c1: 0, r2: 7, c2: 4 }, { row: 2, col: 0 }), ...p };
}

describe("PivotTable hesaplama", () => {
  it("satırlar + toplam (sıkıştırılmış)", () => {
    const g = grid(def({ rows: [0], values: [{ field: 3, fn: "sum" }] }));
    expect(g).toEqual([
      ["Satır Etiketleri", "Toplam Satış"],
      ["Akdeniz", "75"],
      ["Ege", "180"],
      ["Marmara", "220"],
      ["Genel Toplam", "475"],
    ]);
  });

  it("iç içe satırlar üstte alt toplam ve boş öğe", () => {
    const g = grid(def({ rows: [0, 1], values: [{ field: 3, fn: "sum" }] }));
    expect(g[1]).toEqual(["Akdeniz", "75"]);
    expect(g[2]).toEqual(["  Defter", "70"]);
    expect(g[3]).toEqual(["  (boş)", "5"]);
    expect(g[g.length - 1]).toEqual(["Genel Toplam", "475"]);
  });

  it("sütun alanı ve genel toplam sütunu", () => {
    const g = grid(def({ rows: [0], cols: [2], values: [{ field: 3, fn: "sum" }] }));
    expect(g[0]).toEqual(["Toplam Satış", "Sütun Etiketleri", "", ""]);
    expect(g[1]).toEqual(["Satır Etiketleri", "2024", "2025", "Genel Toplam"]);
    expect(g[2]).toEqual(["Akdeniz", "", "75", "75"]);
    expect(g[5]).toEqual(["Genel Toplam", "350", "125", "475"]);
  });

  it("say / ortalama / en büyük / en küçük ve birden çok değer", () => {
    const g = grid(def({ rows: [0], values: [{ field: 3, fn: "count" }, { field: 3, fn: "average" }, { field: 4, fn: "max" }, { field: 4, fn: "min" }] }));
    expect(g[0]).toEqual(["Satır Etiketleri", "Say Satış", "Ortalama Satış", "En Büyük Adet", "En Küçük Adet"]);
    expect(g[2]).toEqual(["Ege", "3", "60", "3", "1"]);
    expect(g[4]).toEqual(["Genel Toplam", "7", "67.86", "7", "1"]);
  });

  it("rapor filtresi ve gizli öğe", () => {
    const g = grid(def({ rows: [1], values: [{ field: 3, fn: "sum" }], filters: [{ field: 0, selected: ["sEGE"] }], hidden: { 1: ["sDEFTER"] } }));
    expect(g[0]).toEqual(["Bölge", "Ege"]);
    expect(g[1]).toEqual(["", ""]);
    expect(g[2]).toEqual(["Satır Etiketleri", "Toplam Satış"]);
    expect(g[3]).toEqual(["Kalem", "130"]);
    expect(g[4]).toEqual(["Genel Toplam", "130"]);
  });

  it("tablo biçimi: alanlar ayrı sütunlarda, altta alt toplam", () => {
    const g = grid(def({ rows: [0, 1], values: [{ field: 3, fn: "sum" }], layout: "tabular", subtotals: "bottom" }));
    expect(g[0]).toEqual(["Bölge", "Ürün", "Toplam Satış"]);
    expect(g[1]).toEqual(["Akdeniz", "Defter", "70"]);
    expect(g[2]).toEqual(["", "(boş)", "5"]);
    expect(g[3]).toEqual(["Akdeniz Toplam", "", "75"]);
  });

  it("genel toplamlar kapalı, alt toplam yok", () => {
    const g = grid(def({ rows: [0, 1], values: [{ field: 3, fn: "sum" }], subtotals: "none", grandRows: false }));
    expect(g[1]).toEqual(["Akdeniz", ""]);
    expect(g.some((r) => r[0] === "Genel Toplam")).toBe(false);
  });

  it("değer alanı olmadan yalnız satır ve sütun öğeleri", () => {
    const g = grid(def({ rows: [0] }));
    expect(g.map((r) => r[0])).toEqual(["Satır Etiketleri", "Akdeniz", "Ege", "Marmara", "Genel Toplam"]);
    const g2 = grid(def({ rows: [0], cols: [2] }));
    expect(g2[1].slice(0, 3)).toEqual(["Satır Etiketleri", "2024", "2025"]);
  });

  it("birden çok değer satırlarda (Σ Değerler)", () => {
    const g = grid(def({ rows: [0], values: [{ field: 3, fn: "sum" }, { field: 4, fn: "sum" }], valuesOnRows: true }));
    expect(g[1]).toEqual(["Akdeniz", ""]);
    expect(g[2]).toEqual(["  Toplam Satış", "75"]);
    expect(g[3]).toEqual(["  Toplam Adet", "13"]);
  });

  it("boş tanım boş düzen verir", () => {
    const lay = computePivot(def({}), src);
    expect(lay.empty).toBe(true);
  });
});

describe("PivotTable XLSX gidiş-dönüşü", () => {
  it("tanım, çıktı ve filtreler korunur; pivot parçaları yazılır", async () => {
    const { Controller } = await import("../src/controller");
    const { createPivot, applyPivot } = await import("../src/pivot/pivotOps");
    const { writeXlsx, readXlsx } = await import("../src/io/xlsx");
    const JSZip = (await import("jszip")).default;
    const ctl = new Controller();
    const sh = ctl.sheet;
    const rows: Scalar[][] = [["Bölge", "Ürün", "Satış"], ["Ege", "Kalem", 100], ["Ege", "Defter", 50], ["Marmara", "Kalem", 200], ["Akdeniz", "Silgi", 70]];
    rows.forEach((r, i) => r.forEach((v, j) => ctl.wb.setInputRaw(sh, i, j, v)));
    createPivot(ctl, sh, { r1: 0, c1: 0, r2: 4, c2: 2 }, { sheet: sh, row: 2, col: 5 });
    const p = sh.pivots[0];
    applyPivot(ctl, sh, p.id, (d) => {
      d.rows = [0];
      d.values = [{ field: 2, fn: "sum" }];
      d.filters = [{ field: 1, selected: ["sKALEM"] }];
    });
    // Rapor filtresi bağlantı hücresinin (F3) üstünde
    expect(ctl.wb.getValue(sh, 0, 5)).toBe("Ürün");
    expect(ctl.wb.getValue(sh, 0, 6)).toBe("Kalem");
    expect(ctl.wb.getValue(sh, 2, 5)).toBe("Satır Etiketleri");
    expect(ctl.wb.getValue(sh, 5, 6)).toBe(300);
    const bytes = await writeXlsx(ctl.wb);
    const zip = await JSZip.loadAsync(bytes);
    expect(zip.file("xl/pivotTables/pivotTable1.xml")).toBeTruthy();
    expect(zip.file("xl/pivotCache/pivotCacheDefinition1.xml")).toBeTruthy();
    const wbXml = await zip.file("xl/workbook.xml")!.async("string");
    expect(wbXml).toMatch(/<pivotCaches><pivotCache cacheId="1" r:id="rId\d+"\/><\/pivotCaches>/);
    const ptXml = await zip.file("xl/pivotTables/pivotTable1.xml")!.async("string");
    expect(ptXml).toContain('<location ref="F3:G6"');
    expect(ptXml).toContain('<dataField name="Toplam Satış" fld="2"');
    const back = await readXlsx(bytes.buffer as ArrayBuffer);
    const s2 = back.sheets[0];
    expect(s2.pivots).toHaveLength(1);
    expect(s2.pivots[0].rows).toEqual([0]);
    expect(s2.pivots[0].filters[0].selected).toEqual(["sKALEM"]);
    expect(s2.pivots[0].output).toEqual({ r1: 0, c1: 5, r2: 5, c2: 6 });
    expect(back.getValue(s2, 5, 6)).toBe(300);
  });
});

describe("Başka uygulamanın yazdığı PivotTable", () => {
  it("LibreOffice'in kaydettiği PivotTable canlı olarak okunur", async () => {
    const { readFileSync } = await import("node:fs");
    const { readXlsx } = await import("../src/io/xlsx");
    const buf = readFileSync(new URL("./fixtures/lo-pivot.xlsx", import.meta.url));
    const wb = await readXlsx(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer);
    const s = wb.sheetByName("Sayfa2")!;
    expect(s.pivots).toHaveLength(1);
    const p = s.pivots[0];
    expect(p.rows).toEqual([0]);
    expect(p.cols).toEqual([2]);
    expect(p.values).toEqual([{ field: 3, fn: "sum", name: "Toplam Satış" }]);
    expect(p.filters).toEqual([{ field: 1, selected: null }]);
    expect(wb.sheetById(p.sourceSheet)?.name).toBe("Sayfa1");
    expect(p.output).toEqual({ r1: 0, c1: 0, r2: 7, c2: 3 });
  });
});
