/** LibreOffice Calc'ın ürettiği XLSX dosyasını içe aktarma (dış üretici fikstürü). */
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { dateToSerial } from "@officehere/formula";
import { readXlsx, writeXlsx } from "../src/io/xlsx";
import { displayText } from "../src/model/display";

const bytes = readFileSync(new URL("./fixtures/lo-generated.xlsx", import.meta.url));
const ab = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;

describe("LibreOffice XLSX fikstürü", () => {
  it("sayfalar, değerler, formüller ve yeniden hesap", async () => {
    const wb = await readXlsx(ab);
    expect(wb.sheets.map((s) => s.name)).toEqual(["Gelirler", "Özet"]);
    const [g, o] = wb.sheets;
    expect(wb.getValue(g, 0, 3)).toBe("Kâr");
    expect(wb.getFormula(g, 1, 3)).toBe("=B2-C2");
    expect(wb.getValue(g, 1, 3)).toBe(4200.5);
    expect(wb.getValue(g, 4, 1)).toBe(41400.5);
    expect(wb.getValue(g, 4, 3)).toBe(14100.25);
    expect(wb.getFormula(o, 0, 1)).toBe("=Gelirler!D5");
    expect(wb.getValue(o, 0, 1)).toBe(14100.25);
    expect(wb.getValue(o, 1, 1)).toBe(dateToSerial(2024, 3, 15));
    // LibreOffice (en-US) tarihi "mm/dd/yy" özel biçimiyle yazdı; biçim olduğu gibi uygulanır
    expect(displayText(wb, o, 1, 1)).toBe("03/15/24");
    expect(wb.getValue(o, 3, 1)).toBe("iyi");
    // Değer değişince bağımlılar yeniden hesaplanır
    wb.setInputRaw(g, 1, 1, 20000);
    wb.engine.recalc();
    expect(wb.getValue(o, 0, 1)).toBe(21599.75);
  });
  it("biçimler, birleştirme ve sütun genişliği", async () => {
    const wb = await readXlsx(ab);
    const g = wb.sheets[0];
    const h = wb.styleAt(g, 0, 0);
    expect(h.bold).toBe(true);
    expect(h.fill?.toUpperCase()).toBe("#1F4E79");
    expect(h.color?.toUpperCase()).toBe("#FFFFFF");
    expect(wb.styleAt(g, 1, 1).numFmt).toBe("#,##0.00");
    expect(displayText(wb, g, 1, 1)).toBe("12.500,50");
    expect(wb.styleAt(g, 4, 0).bTop?.style).toBeTruthy();
    expect(g.merges).toEqual([{ r1: 0, c1: 5, r2: 0, c2: 7 }]);
    expect(g.cols.size(0)).toBeGreaterThan(g.cols.size(1));
  });
  it("yeniden yazılan dosya tekrar aynı içeriği verir", async () => {
    const wb = await readXlsx(ab);
    const out = await writeXlsx(wb);
    const wb2 = await readXlsx(out.buffer as ArrayBuffer);
    expect(wb2.getFormula(wb2.sheets[1], 2, 1)).toBe("=AVERAGE(Gelirler!B2:B4)");
    expect(wb2.getValue(wb2.sheets[1], 2, 1)).toBeCloseTo(13800.1666667, 6);
    expect(wb2.sheets[0].merges).toHaveLength(1);
  });
});
