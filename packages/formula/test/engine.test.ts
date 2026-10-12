import { describe, expect, it } from "vitest";
import { FormulaEngine, shiftFormula, adjustFormula, renameSheetInFormula, toLocalFormula, fromLocalFormula, parseCellInput, cycleAbsolute, tokenize, renderRefToken, findRefSpans, moveRefsInFormula } from "../src";
import { book, err } from "./helpers";

describe("bağımlılık ve yeniden hesaplama", () => {
  it("zincirleme güncelleme", () => {
    const b = book({ A1: 1, A2: "=A1*2", A3: "=A2+A1", A4: "=SUM(A1:A3)" });
    expect(b.get("A4")).toBe(6);
    b.set("A1", 10);
    b.e.recalc();
    expect(b.get("A2")).toBe(20);
    expect(b.get("A3")).toBe(30);
    expect(b.get("A4")).toBe(60);
  });
  it("uzun zincir yığın taşması olmadan hesaplanır", () => {
    const e = new FormulaEngine();
    const s = e.addSheet("S");
    e.setValue(s, 0, 0, 1);
    for (let r = 1; r < 5000; r++) e.setFormula(s, r, 0, `=A${r}+1`);
    e.recalc();
    expect(e.getValue(s, 4999, 0)).toBe(5000);
    // Ters sırada (aşağıdan yukarı) bağımlılık
    const e2 = new FormulaEngine();
    const s2 = e2.addSheet("S");
    e2.setValue(s2, 4999, 0, 1);
    for (let r = 4998; r >= 0; r--) e2.setFormula(s2, r, 0, `=A${r + 2}+1`);
    e2.recalc();
    expect(e2.getValue(s2, 0, 0)).toBe(5000);
  });
  it("artımlı: yalnız etkilenenler hesaplanır", () => {
    const b = book({ A1: 1, B1: "=A1+1", C1: 5, D1: "=C1*2" });
    b.set("A1", 2);
    b.e.recalc();
    expect(b.get("B1")).toBe(3);
    expect(b.get("D1")).toBe(10);
  });
  it("tüm sütun başvurusu bağımlılığı", () => {
    const b = book({ B1: "=SUM(A:A)" });
    b.set("A100", 5);
    b.e.recalc();
    expect(b.get("B1")).toBe(5);
    b.set("A50000", 5);
    b.e.recalc();
    expect(b.get("B1")).toBe(10);
  });
  it("sayfalar arası başvuru", () => {
    const b = book({ A1: 3 });
    const s2 = b.e.addSheet("Sayfa 2");
    b.set("A1", "=Sayfa1!A1*2", s2);
    b.set("B1", "='Sayfa 2'!A1+1");
    b.e.recalc();
    expect(b.get("A1", s2)).toBe(6);
    expect(b.get("B1")).toBe(7);
    b.set("A1", 4);
    b.e.recalc();
    expect(b.get("B1")).toBe(9);
  });
  it("döngüsel başvuru algılanır", () => {
    const b = book({ A1: "=B1+1", B1: "=A1+1" });
    expect(b.e.getCircularRefs().length).toBeGreaterThan(0);
    expect(typeof b.get("A1")).toBe("number");
    const c = book({ A1: "=A1+1" });
    expect(c.get("A1")).toBe(0);
    expect(c.e.getCircularRefs().length).toBe(1);
  });
  it("uçucu fonksiyonlar her hesapta güncellenir", () => {
    let t = 1;
    const e = new FormulaEngine({ now: () => t });
    const s = e.addSheet("S");
    e.setFormula(s, 0, 0, "=NOW()");
    e.recalc();
    expect(e.getValue(s, 0, 0)).toBe(1);
    t = 2;
    e.setValue(s, 5, 5, 1);
    e.recalc();
    expect(e.getValue(s, 0, 0)).toBe(2);
  });
  it("DOLAYLI ile dinamik başvuru", () => {
    const b = book({ A1: "B1", B1: 5, C1: "=INDIRECT(A1)" });
    expect(b.get("C1")).toBe(5);
    b.set("B1", 7);
    b.e.recalc();
    expect(b.get("C1")).toBe(7);
  });
  it("el ile hesaplama modu", () => {
    const b = book({ A1: 1, B1: "=A1" });
    b.e.autoCalc = false;
    b.set("A1", 2);
    expect(b.get("B1")).toBe(1);
    b.e.recalc();
    expect(b.get("B1")).toBe(2);
  });
  it("adlandırılmış aralıklar", () => {
    const b = book({ A1: 1, A2: 2, A3: 3 });
    b.e.defineName("Veri", "=Sayfa1!$A$1:$A$3");
    b.e.defineName("Oran", "=0.5");
    b.set("B1", "=SUM(Veri)*Oran");
    b.e.recalc();
    expect(b.get("B1")).toBe(3);
    b.set("A3", 13);
    b.e.recalc();
    expect(b.get("B1")).toBe(8);
    b.e.removeName("Oran");
    b.e.recalc();
    expect(err(b.get("B1"))).toBe("#NAME?");
  });
  it("arama önbelleği doğru sonuç verir", () => {
    const cells: Record<string, string | number> = {};
    for (let i = 1; i <= 200; i++) {
      cells[`A${i}`] = `k${i}`;
      cells[`B${i}`] = i * 2;
    }
    for (let i = 1; i <= 50; i++) cells[`D${i}`] = `=VLOOKUP("k${i * 3}",A:B,2,FALSE)`;
    const b = book(cells);
    expect(b.get("D10")).toBe(60);
    b.set("B30", 999);
    b.e.recalc();
    expect(b.get("D10")).toBe(999);
  });
});

describe("yapısal değişiklikler", () => {
  it("satır ekleme başvuruları kaydırır ve aralıkları genişletir", () => {
    const b = book({ A1: 1, A2: 2, A3: 3, B1: "=SUM(A1:A3)", B5: "=A3*2", C1: "=$A$2" });
    b.e.insertRows(b.s, 1, 2);
    b.e.recalc();
    expect(b.e.getFormula(b.s, 0, 1)).toBe("=SUM(A1:A5)");
    expect(b.e.getFormula(b.s, 6, 1)).toBe("=A5*2");
    expect(b.e.getFormula(b.s, 0, 2)).toBe("=$A$4");
    expect(b.get("B1")).toBe(6);
    expect(b.get("B7")).toBe(6);
  });
  it("satır silme: silinen başvurular #REF!", () => {
    const b = book({ A1: 1, A2: 2, A3: 3, B1: "=A2", C1: "=SUM(A1:A3)", B4: "=A3" });
    b.e.deleteRows(b.s, 1, 1);
    b.e.recalc();
    expect(b.e.getFormula(b.s, 0, 1)).toBe("=#REF!");
    expect(err(b.get("B1"))).toBe("#REF!");
    expect(b.e.getFormula(b.s, 0, 2)).toBe("=SUM(A1:A2)");
    expect(b.get("C1")).toBe(4);
    expect(b.e.getFormula(b.s, 2, 1)).toBe("=A2");
  });
  it("sütun ekleme/silme", () => {
    const b = book({ A1: 1, B1: 2, C1: "=A1+B1", D1: "=SUM(A1:B1)" });
    b.e.insertCols(b.s, 1, 1);
    b.e.recalc();
    expect(b.e.getFormula(b.s, 0, 3)).toBe("=A1+C1");
    expect(b.e.getFormula(b.s, 0, 4)).toBe("=SUM(A1:C1)");
    b.e.deleteCols(b.s, 0, 1);
    b.e.recalc();
    expect(b.e.getFormula(b.s, 0, 2)).toBe("=#REF!+B1");
    expect(b.e.getFormula(b.s, 0, 3)).toBe("=SUM(A1:B1)");
  });
  it("başka sayfadaki başvurular da güncellenir", () => {
    const b = book({ A5: 9 });
    const s2 = b.e.addSheet("Diğer");
    b.set("A1", "=Sayfa1!A5", s2);
    b.e.insertRows(b.s, 0, 1);
    b.e.recalc();
    expect(b.e.getFormula(s2, 0, 0)).toBe("=Sayfa1!A6");
    expect(b.get("A1", s2)).toBe(9);
  });
  it("sayfa yeniden adlandırma ve silme", () => {
    const b = book({ A1: 2 });
    const s2 = b.e.addSheet("Veri");
    b.set("A1", 5, s2);
    b.set("B1", "=Veri!A1*A1");
    b.e.recalc();
    expect(b.get("B1")).toBe(10);
    b.e.renameSheet(s2, "Yeni Veri");
    b.e.recalc();
    expect(b.e.getFormula(b.s, 0, 1)).toBe("='Yeni Veri'!A1*A1");
    expect(b.get("B1")).toBe(10);
    b.e.removeSheet(s2);
    b.e.recalc();
    expect(b.e.getFormula(b.s, 0, 1)).toBe("='Yeni Veri'!#REF!*A1");
    expect(err(b.get("B1"))).toBe("#REF!");
  });
  it("kes-yapıştır başvuruları taşır", () => {
    const b = book({ A1: 5, B1: "=A1*2", C1: "=SUM(A1:A1)" });
    b.e.moveRange(b.s, 0, 0, 0, 0, b.s, 4, 0);
    b.e.recalc();
    expect(b.e.getFormula(b.s, 0, 1)).toBe("=A5*2");
    expect(b.get("B1")).toBe(10);
    expect(b.get("A5")).toBe(5);
  });
  it("hücreleri aşağı kaydır", () => {
    const b = book({ A1: 1, A2: 2, B1: "=A2" });
    b.e.shiftCells(b.s, 0, 0, 0, 0, "down");
    b.e.recalc();
    expect(b.get("A2")).toBe(1);
    expect(b.get("A3")).toBe(2);
    expect(b.e.getFormula(b.s, 0, 1)).toBe("=A3");
  });
  it("anlık görüntü ve geri yükleme", () => {
    const b = book({ A1: 1, B1: "=A1+1" });
    const snap = b.e.exportInputs();
    b.e.deleteRows(b.s, 0, 1);
    b.e.recalc();
    b.e.importInputs(snap);
    b.e.recalc();
    expect(b.get("B1")).toBe(2);
    expect(b.e.getFormula(b.s, 0, 1)).toBe("=A1+1");
  });
});

describe("başvuru dönüşümleri", () => {
  it("göreli kaydırma", () => {
    expect(shiftFormula("=A1+$B$2+C$3+$D4", 1, 1)).toBe("=B2+$B$2+D$3+$D5");
    expect(shiftFormula("=SUM(A1:B2)", 2, 0)).toBe("=SUM(A3:B4)");
    expect(shiftFormula("=SUM(A:A)", 5, 1)).toBe("=SUM(B:B)");
    expect(shiftFormula("=SUM(1:1)", 1, 3)).toBe("=SUM(2:2)");
    expect(shiftFormula("=A1", -1, 0)).toBe("=#REF!");
    expect(shiftFormula("=Sayfa2!A1 + 1", 1, 0)).toBe("=Sayfa2!A2 + 1");
    expect(shiftFormula('="A1"&A1', 1, 0)).toBe('="A1"&A2');
  });
  it("ekleme/silme dönüşümü", () => {
    expect(adjustFormula("=A5", "S", { sheet: "S", axis: "row", at: 2, count: 3 })).toBe("=A8");
    expect(adjustFormula("=A1", "S", { sheet: "S", axis: "row", at: 2, count: 3 })).toBe("=A1");
    expect(adjustFormula("=SUM(A2:A4)", "S", { sheet: "S", axis: "row", at: 1, count: -2 })).toBe("=SUM(A2:A2)");
    expect(adjustFormula("=SUM(A2:A3)", "S", { sheet: "S", axis: "row", at: 1, count: -2 })).toBe("=SUM(#REF!)");
    expect(adjustFormula("=X!A5", "S", { sheet: "S", axis: "row", at: 0, count: 1 })).toBe("=X!A5");
    expect(adjustFormula("=SUM(B:D)", "S", { sheet: "S", axis: "col", at: 2, count: -1 })).toBe("=SUM(B:C)");
  });
  it("sayfa adı", () => {
    expect(renameSheetInFormula("=Eski!A1+'Eski'!B2", "Eski", "Yeni Ad")).toBe("='Yeni Ad'!A1+'Yeni Ad'!B2");
  });
  it("F4 döngüsü", () => {
    const r = tokenize("A1")[0].ref!;
    const r1 = cycleAbsolute(r);
    expect(renderRefToken(r1)).toBe("$A$1");
    const r2 = cycleAbsolute(r1);
    expect(renderRefToken(r2)).toBe("A$1");
    const r3 = cycleAbsolute(r2);
    expect(renderRefToken(r3)).toBe("$A1");
    expect(renderRefToken(cycleAbsolute(r3))).toBe("A1");
  });
  it("başvuru aralıkları (renkli vurgu)", () => {
    const spans = findRefSpans("=A1+SUM(B2:C3)");
    expect(spans.map((s) => s.start)).toEqual([1, 8]);
    expect(spans[1].end).toBe(13);
  });
  it("taşıma dönüşümü", () => {
    const mv = { srcSheet: "S", r1: 0, c1: 0, r2: 1, c2: 1, dRow: 5, dCol: 0, dstSheet: "S" };
    expect(moveRefsInFormula("=A1+B2+C3+SUM(A1:B2)", "S", mv)).toBe("=A6+B7+C3+SUM(A6:B7)");
  });
});

describe("yerelleştirme", () => {
  it("kanonik → Türkçe", () => {
    expect(toLocalFormula("=SUM(A1:A3,1.5)")).toBe("=TOPLA(A1:A3;1,5)");
    expect(toLocalFormula('=IF(A1>0,TRUE,"x")')).toBe('=EĞER(A1>0;DOĞRU;"x")');
    expect(toLocalFormula("=IFERROR(VLOOKUP(A1,B:C,2,FALSE),#N/A)")).toBe("=EĞERHATA(DÜŞEYARA(A1;B:C;2;YANLIŞ);#YOK)");
    expect(toLocalFormula("={1,2;3,4}")).toBe("={1\\2;3\\4}");
    expect(toLocalFormula("=COUNTIFS(A:A,\"a,b\")")).toBe('=ÇOKEĞERSAY(A:A;"a,b")');
    expect(toLocalFormula("=Sheet1!#REF!+1")).toBe("=Sheet1!#BAŞV!+1");
    expect(toLocalFormula("=_xlfn.XLOOKUP(1,A:A,B:B)")).toBe("=ÇAPRAZARA(1;A:A;B:B)");
  });
  it("Türkçe → kanonik", () => {
    expect(fromLocalFormula("=TOPLA(A1:A3;1,5)")).toBe("=SUM(A1:A3,1.5)");
    expect(fromLocalFormula("=eğer(a1>0;doğru;yanlış)")).toBe("=IF(a1>0,TRUE,FALSE)");
    expect(fromLocalFormula("=SUM(1;2)")).toBe("=SUM(1,2)");
    expect(fromLocalFormula("=BAĞ_DEĞ_SAY(A:A)+YAZIM.DÜZENİ(B1)")).toBe("=COUNT(A:A)+PROPER(B1)");
    expect(fromLocalFormula("=min(1;2)")).toBe("=MIN(1,2)");
    expect(fromLocalFormula("={1\\2;3\\4}")).toBe("={1,2;3,4}");
    expect(fromLocalFormula("=#YOK")).toBe("=#N/A");
  });
});

describe("hücre girişi ayrıştırma", () => {
  it("sayılar ve biçimler", () => {
    expect(parseCellInput("12")).toEqual({ kind: "value", value: 12 });
    expect(parseCellInput("12,5")).toEqual({ kind: "value", value: 12.5 });
    expect(parseCellInput("1.234,5")).toEqual({ kind: "value", value: 1234.5, numFmt: "#,##0.0" });
    expect(parseCellInput("-3")).toEqual({ kind: "value", value: -3 });
    expect(parseCellInput("(3)")).toEqual({ kind: "value", value: -3 });
    expect(parseCellInput("%15")).toMatchObject({ value: 0.15, numFmt: "0%" });
    expect(parseCellInput("12,5%")).toMatchObject({ value: 0.125, numFmt: "0.0%" });
    expect(parseCellInput("₺1.500")).toMatchObject({ value: 1500, numFmt: '#,##0 "₺"' });
    expect(parseCellInput("99,90 TL")).toMatchObject({ value: 99.9, numFmt: '#,##0.00 "₺"' });
    expect(parseCellInput("1,5E+3")).toMatchObject({ value: 1500 });
    expect(parseCellInput("0 1/2")).toMatchObject({ value: 0.5, numFmt: "# ?/?" });
  });
  it("tarih ve saat", () => {
    expect(parseCellInput("15.03.2024")).toMatchObject({ value: 45366, numFmt: "dd.mm.yyyy" });
    expect(parseCellInput("2024-03-15")).toMatchObject({ value: 45366 });
    expect(parseCellInput("14:30")).toMatchObject({ numFmt: "h:mm" });
    expect((parseCellInput("14:30").value as number)).toBeCloseTo(14.5 / 24);
    expect(parseCellInput("15.03.2024 14:30")).toMatchObject({ numFmt: "dd.mm.yyyy h:mm" });
    expect(parseCellInput("15 Mart 2024")).toMatchObject({ value: 45366 });
  });
  it("metin, mantıksal, hata, formül", () => {
    expect(parseCellInput("merhaba")).toEqual({ kind: "value", value: "merhaba" });
    expect(parseCellInput("'0123")).toEqual({ kind: "value", value: "0123" });
    expect(parseCellInput("0123", { numFmt: "@" })).toEqual({ kind: "value", value: "0123" });
    expect(parseCellInput("doğru")).toEqual({ kind: "value", value: true });
    expect(parseCellInput("YANLIŞ")).toEqual({ kind: "value", value: false });
    expect((parseCellInput("#YOK").value as { code: string }).code).toBe("#N/A");
    expect(parseCellInput("=TOPLA(A1;2,5)")).toEqual({ kind: "formula", formula: "=SUM(A1,2.5)" });
    expect(parseCellInput("")).toEqual({ kind: "empty" });
    expect(parseCellInput("Ürün 1")).toEqual({ kind: "value", value: "Ürün 1" });
  });
});
