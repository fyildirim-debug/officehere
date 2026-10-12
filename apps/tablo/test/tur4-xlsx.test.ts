/** 4. tur K1: aynı sayfada tüm özellikler → sayfa XML'indeki üst düzey öğe sırası şemaya uygun. */
import { describe, expect, it } from "vitest";
import { mkdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import JSZip from "jszip";
import { cellKey } from "@officehere/formula";
import { Controller } from "../src/controller";
import * as D from "../src/actions/data";
import { insertChart } from "../src/actions/objects";
import { groupSelection } from "../src/actions/outline";
import { createSparklines } from "../src/sparkline/sparkOps";
import { applyPivot, createPivot } from "../src/pivot/pivotOps";
import { readXlsx, writeXlsx } from "../src/io/xlsx";
import { WORKBOOK_ORDER, WORKSHEET_ORDER, XmlPart, orderViolation, topLevelNames } from "../src/io/xmlPart";
import { injectBreaks } from "../src/io/xlsxBreaks";
import { fixDataBars } from "../src/io/xlsxPost";

/** Veri çubuğu + simge kümeleri + grafik + mini grafik + özet tablo + not + köprü + doğrulama + sayfa sonu + anahat + tablo. */
export function kitchenSink(): Controller {
  const ctl = new Controller();
  const sh = ctl.sheet;
  const rows: (string | number)[][] = [
    ["Bölge", "Ocak", "Şubat", "Mart", "Tutar"],
    ["Marmara", 10, -2, 3, 100],
    ["Ege", 4, 2, 6, 200],
    ["Akdeniz", 7, 5, 3, 300],
    ["Karadeniz", 1, 9, 3, 400],
    ["Doğu", 3, 2, 8, 500],
  ];
  ctl.ops.setInputs(sh, rows.flatMap((r, i) => r.map((v, j) => ({ row: i, col: j, input: v }))), "örnek");
  D.createTable(ctl, { r1: 0, c1: 0, r2: 5, c2: 4 }, true);
  sh.condFormats.push({ id: "cf1", type: "dataBar", color: "#638EC6", ranges: [{ r1: 1, c1: 1, r2: 5, c2: 1 }] });
  sh.condFormats.push({ id: "cf2", type: "iconSet", iconSet: "3Triangles", ranges: [{ r1: 1, c1: 2, r2: 5, c2: 2 }] });
  sh.condFormats.push({ id: "cf3", type: "iconSet", iconSet: "3Arrows", ranges: [{ r1: 1, c1: 3, r2: 5, c2: 3 }] });
  sh.condFormats.push({ id: "cf4", type: "cellIs", operator: "greaterThan", values: ["=250"], ranges: [{ r1: 1, c1: 4, r2: 5, c2: 4 }], style: { fill: "#FFC7CE" } });
  ctl.selectRange({ r1: 0, c1: 0, r2: 5, c2: 1 });
  insertChart(ctl, "column");
  expect(createSparklines(ctl, "line", "B2:D6", "F2:F6")).toBeNull();
  createPivot(ctl, sh, { r1: 0, c1: 0, r2: 5, c2: 4 }, { sheet: sh, row: 1, col: 8 });
  applyPivot(ctl, sh, sh.pivots[0].id, (d) => {
    d.rows = [0];
    d.values = [{ field: 4, fn: "sum" }];
  });
  sh.notes.set(cellKey(0, 0), { text: "Bölge adı" });
  ctl.wb.setInputRaw(sh, 12, 0, "Site");
  sh.links.set(cellKey(12, 0), "https://example.com/");
  sh.validations.push({ id: "v1", ranges: [{ r1: 14, c1: 0, r2: 16, c2: 0 }], rule: { type: "list", formula1: "Evet;Hayır" } });
  sh.page.rowBreaks = [8];
  sh.page.colBreaks = [3];
  sh.page.header = "Rapor";
  sh.page.printGridlines = true;
  sh.merges.push({ r1: 20, c1: 0, r2: 20, c2: 2 });
  sh.frozen = { rows: 1, cols: 0 };
  ctl.selectRows(18, 19);
  groupSelection(ctl, "row");
  ctl.selectCell(0, 0);
  return ctl;
}

async function sheetXml(bytes: Uint8Array, n = 1): Promise<string> {
  const zip = await JSZip.loadAsync(bytes);
  return zip.file(`xl/worksheets/sheet${n}.xml`)!.async("string");
}

describe("K1: sayfa XML öğe sırası (ECMA-376 CT_Worksheet)", () => {
  it("tüm özellikler aynı sayfada: üst düzey öğeler şema sırasında", async () => {
    const ctl = kitchenSink();
    const bytes = await writeXlsx(ctl.wb);
    if (process.env.TABLO_DUMP_DIR) {
      mkdirSync(process.env.TABLO_DUMP_DIR, { recursive: true });
      writeFileSync(join(process.env.TABLO_DUMP_DIR, "k1-hepsi.xlsx"), bytes);
    }
    const xml = await sheetXml(bytes);
    const names = topLevelNames(xml);
    expect(orderViolation(names, WORKSHEET_ORDER)).toBeNull();
    // Her öğe şema tablosunda bilinen bir öğe olmalı ve en çok bir kez geçmeli (conditionalFormatting hariç)
    for (const n of names) expect(WORKSHEET_ORDER as readonly string[]).toContain(n);
    for (const n of new Set(names)) if (n !== "conditionalFormatting") expect(names.filter((x) => x === n)).toHaveLength(1);
    for (const n of ["sheetViews", "sheetFormatPr", "sheetData", "mergeCells", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "drawing", "legacyDrawing", "tableParts", "extLst"]) expect(names).toContain(n);
    // Sıra listesiyle birebir: şema tablosundan bu sayfadaki öğelerin alt dizisi
    const expected = (WORKSHEET_ORDER as readonly string[]).filter((n) => names.includes(n));
    expect([...new Set(names)]).toEqual(expected);
    // drawing hiçbir zaman cfRule içinde değil; veri çubuğu x14:id dolu ve sayfa uzantısıyla eşleşir
    expect(xml).not.toMatch(/<cfRule[^>]*>(?:(?!<\/cfRule>)[\s\S])*<drawing/);
    expect(xml).not.toContain("<x14:id/>");
    const id = /<x14:id>(\{[0-9A-F-]+\})<\/x14:id>/.exec(xml)?.[1];
    expect(id).toBeTruthy();
    expect(xml).toContain(`<x14:cfRule type="dataBar" id="${id}">`);
    // 3 Üçgen simge kümesi ve veri çubuğu aynı x14:conditionalFormattings içinde
    expect(xml.match(/<x14:conditionalFormattings>/g)).toHaveLength(1);
    expect(xml).toContain('iconSet="3Triangles"');
    // Mini grafikler sayfa uzantısında
    expect(xml).toContain("<x14:sparklineGroups");
    expect(xml).toMatch(/<xm:f>(?:'?Sayfa1'?)!B2:D2<\/xm:f><xm:sqref>F2<\/xm:sqref>/);
    const zip = await JSZip.loadAsync(bytes);
    const wbXml = await zip.file("xl/workbook.xml")!.async("string");
    expect(orderViolation(topLevelNames(wbXml), WORKBOOK_ORDER)).toBeNull();
  });

  it("gidiş-dönüşte her şey korunur", async () => {
    const ctl = kitchenSink();
    const back = await readXlsx((await writeXlsx(ctl.wb)).buffer as ArrayBuffer);
    const s = back.sheets[0];
    expect(s.condFormats.map((r) => r.type).sort()).toEqual(["cellIs", "dataBar", "iconSet", "iconSet"]);
    expect(s.condFormats.find((r) => r.type === "dataBar")!.color).toBe("#638EC6");
    expect(s.charts).toHaveLength(1);
    expect(s.sparklines).toHaveLength(1);
    expect(s.sparklines[0].items).toHaveLength(5);
    expect(s.pivots).toHaveLength(1);
    expect(s.tables.map((t) => t.name)).toEqual([ctl.sheet.tables[0].name]);
    expect(s.noteAt(0, 0)?.text).toBe("Bölge adı");
    expect(s.linkAt(12, 0)).toBe("https://example.com/");
    expect(s.validations).toHaveLength(1);
    expect(s.page.rowBreaks).toEqual([8]);
    expect(s.page.colBreaks).toEqual([3]);
    expect(s.rows.level(18)).toBe(1);
    // İkinci kez yazmak da geçerli sıra üretir
    const xml2 = await sheetXml(await writeXlsx(back));
    expect(orderViolation(topLevelNames(xml2), WORKSHEET_ORDER)).toBeNull();
  });

  it("XmlPart: bilinmeyen öğe öncekinin arkasında kalır, set/kaldır yerinde çalışır", () => {
    const p = new XmlPart('<?xml version="1.0"?>\n<worksheet xmlns="m"><sheetData/><extLst><ext uri="a"/></extLst><mc:AlternateContent>x</mc:AlternateContent><pageMargins/></worksheet>', WORKSHEET_ORDER);
    p.set("drawing", '<drawing r:id="rId1"/>');
    p.set("rowBreaks", "<rowBreaks/>");
    expect(p.toString()).toBe('<?xml version="1.0"?>\n<worksheet xmlns="m"><sheetData/><pageMargins/><rowBreaks/><drawing r:id="rId1"/><extLst><ext uri="a"/></extLst><mc:AlternateContent>x</mc:AlternateContent></worksheet>');
    p.set("rowBreaks", null);
    expect(p.has("rowBreaks")).toBe(false);
    // Önekli öğeler ve yorumlar
    const q = new XmlPart('<x:worksheet xmlns:x="m"><!-- not --><x:sheetData><x:row r="1"/></x:sheetData><x:cols/></x:worksheet>', WORKSHEET_ORDER);
    expect(q.children.map((c) => c.name)).toEqual(["sheetData", "cols"]);
    expect(q.toString()).toBe('<x:worksheet xmlns:x="m"><x:cols/><x:sheetData><x:row r="1"/></x:sheetData></x:worksheet>');
  });

  it("sayfa sonları cfRule içindeki extLst'e değil kök düzeye yerleşir", () => {
    const xml = '<worksheet><sheetData/><conditionalFormatting sqref="A1"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"><x14:id/></ext></extLst></cfRule></conditionalFormatting><pageMargins/></worksheet>';
    const out = injectBreaks(xml, [4], []);
    expect(topLevelNames(out)).toEqual(["sheetData", "conditionalFormatting", "pageMargins", "rowBreaks"]);
    const p = new XmlPart(out, WORKSHEET_ORDER);
    fixDataBars(p);
    const s = p.toString();
    expect(topLevelNames(s)).toEqual(["sheetData", "conditionalFormatting", "pageMargins", "rowBreaks", "extLst"]);
    expect(s).toMatch(/<cfRule type="dataBar" priority="1"><dataBar>[\s\S]*<\/dataBar><extLst><ext uri="\{B025F937[^"]*"[^>]*><x14:id>\{/);
  });
});
