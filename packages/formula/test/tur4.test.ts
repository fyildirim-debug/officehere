/** 4. tur: aralığa dönüştürmede yapısal başvuru çevirisi, tablo içi gösterim, yerel ayar adları, Türkçe geçen süre kodları. */
import { describe, expect, it } from "vitest";
import { FormulaEngine } from "../src/engine";
import { structRefsToA1, unqualifyStructRefs } from "../src/transform";
import { formatValue } from "../src/numfmt";
import { formatFromLocal, looksLocalFormat } from "../src/numfmt-locale";
import type { TableInfo } from "../src/structref";

const T: TableInfo = { name: "Tablo1", sheet: 1, r1: 0, c1: 0, r2: 4, c2: 2, headerRow: true, totalRow: false, columns: ["Ürün", "Adet", "Tutar"] };
const names = (id: number) => (id === 1 ? "Satış Verisi" : "Özet");

describe("structRefsToA1 (Aralığa Dönüştür)", () => {
  it("sütun başvuruları mutlak, bu satır başvuruları sütunu sabit A1 olur", () => {
    expect(structRefsToA1("=SUM(Tablo1[Tutar])", T, { sheet: 1, row: 7, col: 5 }, names)).toBe("=SUM($C$2:$C$5)");
    expect(structRefsToA1("=Tablo1[[#This Row],[Adet]]*Tablo1[[#This Row],[Tutar]]", T, { sheet: 1, row: 2, col: 3 }, names)).toBe("=$B3*$C3");
    expect(structRefsToA1("=SUM(Tablo1[[#All],[Adet]:[Tutar]])", T, { sheet: 1, row: 7, col: 5 }, names)).toBe("=SUM($B$1:$C$5)");
    expect(structRefsToA1("=Tablo1[[#Headers],[Ürün]]", T, { sheet: 1, row: 7, col: 5 }, names)).toBe("=$A$1");
  });
  it("başka sayfadan başvuruda sayfa adı eklenir; tablosuz başvuru yalnız tablonun içinde çevrilir", () => {
    expect(structRefsToA1("=COUNT(Tablo1[Adet])", T, { sheet: 2, row: 0, col: 0 }, names)).toBe("=COUNT('Satış Verisi'!$B$2:$B$5)");
    expect(structRefsToA1("=[@Adet]*2", T, { sheet: 1, row: 3, col: 4 }, names)).toBe("=[@Adet]*2");
    expect(structRefsToA1("=[[#This Row],[Adet]]*2", T, { sheet: 1, row: 3, col: 2 }, names)).toBe("=$B4*2");
    expect(structRefsToA1("=Tablo2[Adet]", T, { sheet: 1, row: 3, col: 2 }, names)).toBe("=Tablo2[Adet]");
  });
  it("olmayan sütun #REF! olur", () => {
    expect(structRefsToA1("=SUM(Tablo1[Yok])", T, { sheet: 1, row: 7, col: 5 }, names)).toBe("=SUM(#REF!)");
  });
  it("motor: tableRefsToA1 hücreleri ve tanımlı adları çevirir, sonuç değişmez", () => {
    const e = new FormulaEngine();
    const s = e.addSheet("Sayfa1");
    const rows: (string | number)[][] = [["Ürün", "Adet", "Tutar"], ["Çay", 2, 100], ["Şeker", 3, 250]];
    rows.forEach((r, i) => r.forEach((v, j) => e.setValue(s, i, j, v)));
    e.setTables([{ name: "Tablo1", sheet: s, r1: 0, c1: 0, r2: 2, c2: 2, headerRow: true, totalRow: false, columns: ["Ürün", "Adet", "Tutar"] }]);
    e.setFormula(s, 5, 0, "=SUM(Tablo1[Tutar])");
    e.defineName("Toplam", "=SUM(Tablo1[Adet])");
    e.setFormula(s, 6, 0, "=Toplam");
    e.recalc();
    expect(e.getValue(s, 5, 0)).toBe(350);
    e.tableRefsToA1("Tablo1");
    e.setTables([]);
    e.recalc();
    expect(e.getFormula(s, 5, 0)).toBe("=SUM($C$2:$C$3)");
    expect(e.getValue(s, 5, 0)).toBe(350);
    expect(e.getValue(s, 6, 0)).toBe(5);
  });
});

describe("unqualifyStructRefs (tablo içinde [@Adet] gösterimi)", () => {
  it("yalnız verilen tabloya ait başvurulardan ad kalkar", () => {
    expect(unqualifyStructRefs("=Tablo1[[#This Row],[Adet]]*2", "Tablo1")).toBe("=[[#This Row],[Adet]]*2");
    expect(unqualifyStructRefs("=Tablo2[Adet]+Tablo1[Tutar]", "tablo1")).toBe("=Tablo2[Adet]+[Tutar]");
  });
});

describe("yerel ayar etiketli biçimler (D3) ve Türkçe geçen süre (D4)", () => {
  it("tr/en dışındaki yerel ayarlar Intl ile adlandırılır", () => {
    const d = 46307; // 12.10.2026 Pazartesi
    expect(formatValue(d, "[$-de-DE]dddd").text).toBe("Montag");
    expect(formatValue(d, "[$-407]mmmm").text).toBe("Oktober");
    expect(formatValue(d, "[$-fr-FR]dddd d mmmm").text).toBe("lundi 12 octobre");
    expect(formatValue(d, "[$-tr-TR]dddd").text).toBe("Pazartesi");
    expect(formatValue(d, "[$-409]dddd").text).toBe("Monday");
  });
  it("[s]:dd Türkçe kodlardır (saat:dakika)", () => {
    expect(looksLocalFormat("[s]:dd")).toBe(true);
    expect(formatFromLocal("[s]:dd")).toBe("[h]:mm");
    expect(looksLocalFormat("[h]:mm")).toBe(false);
  });
});
