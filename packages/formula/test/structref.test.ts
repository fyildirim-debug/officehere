import { describe, expect, it } from "vitest";
import { FormulaEngine } from "../src/engine";
import { fromLocalFormula, toLocalFormula } from "../src/locale";
import { parseStructInner, renderStructRef } from "../src/structref";
import { qualifyStructRefs, renameTableInFormula, shiftFormula } from "../src/transform";

function setup() {
  const e = new FormulaEngine();
  const s = e.addSheet("Sayfa1");
  const rows: (string | number)[][] = [
    ["Ürün", "Adet", "Tutar"],
    ["Çay", 2, 100],
    ["Şeker", 3, 250],
    ["Un", 1, 50],
  ];
  rows.forEach((r, i) => r.forEach((v, j) => e.setValue(s, i, j, v)));
  e.setTables([{ name: "Tablo1", sheet: s, r1: 0, c1: 0, r2: 3, c2: 2, headerRow: true, totalRow: false, columns: ["Ürün", "Adet", "Tutar"] }]);
  return { e, s };
}

describe("yapısal başvurular", () => {
  it("iç sözdizimi ayrıştırılır", () => {
    expect(parseStructInner("Tutar")).toEqual({ items: [], col1: "Tutar" });
    expect(parseStructInner("@Tutar")).toEqual({ items: ["#This Row"], col1: "Tutar" });
    expect(parseStructInner("[#Tümü];[Tutar]")).toEqual({ items: ["#All"], col1: "Tutar" });
    expect(parseStructInner("[#This Row],[Adet]:[Tutar]")).toEqual({ items: ["#This Row"], col1: "Adet", col2: "Tutar" });
    expect(parseStructInner("#Başlıklar")).toEqual({ items: ["#Headers"] });
    expect(parseStructInner("@[Birim Fiyat]")).toEqual({ items: ["#This Row"], col1: "Birim Fiyat" });
  });

  it("Türkçe ⇄ kanonik dönüşüm", () => {
    expect(fromLocalFormula("=TOPLA(Tablo1[[#Tümü];[Tutar]])", "tr")).toBe("=SUM(Tablo1[[#All],[Tutar]])");
    expect(fromLocalFormula("=Tablo1[@Tutar]*2", "tr")).toBe("=Tablo1[[#This Row],[Tutar]]*2");
    expect(toLocalFormula("=Tablo1[[#This Row],[Tutar]]*2", "tr")).toBe("=Tablo1[@Tutar]*2");
    expect(toLocalFormula("=SUM(Table1[[#All],[Tutar]])", "tr")).toBe("=TOPLA(Table1[[#Tümü];[Tutar]])");
    expect(toLocalFormula("=SUM(Satislar[Satış])", "tr")).toBe("=TOPLA(Satislar[Satış])");
    expect(renderStructRef({ table: "T", items: ["#This Row"], col1: "Birim Fiyat" }, "tr")).toBe("T[@Birim Fiyat]");
    expect(renderStructRef({ table: "T", items: ["#This Row"], col1: "A", col2: "B" }, "tr")).toBe("T[@[A]:[B]]");
  });

  it("değerlendirme: sütun, bu satır, #Tümü, tablo adı", () => {
    const { e, s } = setup();
    e.setFormula(s, 5, 0, "=SUM(Tablo1[Tutar])");
    e.setFormula(s, 5, 1, "=SUM(Tablo1[[#All],[Tutar]])");
    e.setFormula(s, 5, 2, "=ROWS(Tablo1)");
    e.setFormula(s, 1, 3, "=Tablo1[[#This Row],[Tutar]]*2");
    e.setFormula(s, 2, 3, "=[@Adet]*10");
    e.setFormula(s, 5, 3, "=COUNTA(Tablo1[#Headers])");
    e.recalc();
    expect(e.getValue(s, 5, 0)).toBe(400);
    expect(e.getValue(s, 5, 1)).toBe(400);
    expect(e.getValue(s, 5, 2)).toBe(3);
    expect(e.getValue(s, 1, 3)).toBe(200);
    expect(e.getValue(s, 5, 3)).toBe(3);
    // Tablosuz [@Adet]: D3 tablonun dışında → #BAŞV!
    expect(e.getValue(s, 2, 3)).toMatchObject({ code: "#REF!" });
  });

  it("tablo genişleyince ve değer değişince yeniden hesaplanır", () => {
    const { e, s } = setup();
    e.setFormula(s, 6, 0, "=SUM(Tablo1[Tutar])");
    e.recalc();
    e.setValue(s, 1, 2, 1000);
    e.recalc();
    expect(e.getValue(s, 6, 0)).toBe(1300);
    e.setValue(s, 4, 0, "Pirinç");
    e.setValue(s, 4, 2, 700);
    e.setTables([{ name: "Tablo1", sheet: s, r1: 0, c1: 0, r2: 4, c2: 2, headerRow: true, totalRow: false, columns: ["Ürün", "Adet", "Tutar"] }]);
    e.recalc();
    expect(e.getValue(s, 6, 0)).toBe(2000);
  });

  it("bilinmeyen sütun ve tablo #BAŞV! verir; kaydırma yapısal başvuruyu bozmaz", () => {
    const { e, s } = setup();
    e.setFormula(s, 6, 0, "=SUM(Tablo1[Yok])");
    e.setFormula(s, 6, 1, "=SUM(Yok[Tutar])");
    e.recalc();
    expect(e.getValue(s, 6, 0)).toMatchObject({ code: "#REF!" });
    expect(e.getValue(s, 6, 1)).toMatchObject({ code: "#REF!" });
    expect(shiftFormula("=Tablo1[[#This Row],[Tutar]]+A1", 1, 0)).toBe("=Tablo1[[#This Row],[Tutar]]+A2");
  });

  it("tablo adı değişimi ve nitelendirme", () => {
    expect(renameTableInFormula("=SUM(Tablo1[Tutar])+Tablo2[@X]", "Tablo1", "Satış")).toBe("=SUM(Satış[Tutar])+Tablo2[@X]");
    expect(qualifyStructRefs("=[[#This Row],[Adet]]*2", "Tablo1")).toBe("=Tablo1[[#This Row],[Adet]]*2");
    const { e, s } = setup();
    e.setFormula(s, 6, 0, "=SUM(Tablo1[Tutar])");
    e.renameTable("Tablo1", "Satış");
    e.recalc();
    expect(e.getFormula(s, 6, 0)).toBe("=SUM(Satış[Tutar])");
    expect(e.getValue(s, 6, 0)).toBe(400);
  });
});
