/** 5. tur QA: LAMBDA çağrıları, TÜR, SAYIDEĞERİ, T.TEST düzeltmeleri (sonuçlar masaüstü hesap tablosuyla aynı). */
import { describe, expect, it } from "vitest";
import { isError, parse, CALL_FN } from "../src";
import { book, calc, calcTr, err } from "./helpers";

describe("LAMBDA çağrıları", () => {
  it("doğrudan çağrı: =LAMBDA(x;x*2)(4)", () => {
    expect(calc("=LAMBDA(x,x*2)(4)")).toBe(8);
    expect(calc("=LAMBDA(x,y,x*y)(3,4)")).toBe(12);
    expect(calcTr("=LAMBDA(x;x*2)(4)")).toBe(8);
    // Bağımsız değişken sayısı uymazsa #DEĞER!
    expect(err(calc("=LAMBDA(x,x*2)(4,5)"))).toBe("#VALUE!");
    const ast = parse("LAMBDA(x,x+1)(1)");
    expect(ast.type === "func" && ast.name).toBe(CALL_FN);
  });

  it("Ad Yöneticisi'nde tanımlı LAMBDA işlev gibi çağrılır ve ad değişince yeniden hesaplanır", () => {
    const b = book({ A1: 3 });
    b.e.defineName("KARE", "=LAMBDA(x,x*x)");
    b.e.defineName("vergi", "=LAMBDA(t,t*0.2)");
    b.set("B1", "=KARE(A1)");
    b.set("B2", "=vergi(100)");
    b.set("B3", "=MAP({1,2},KARE)");
    b.e.recalc();
    expect(b.get("B1")).toBe(9);
    expect(b.get("B2")).toBe(20);
    expect(b.get("B3")).toBe(1);
    b.e.defineName("KARE", "=LAMBDA(x,x*x*x)");
    b.e.recalc();
    expect(b.get("B1")).toBe(27);
    // Hücre değişince de
    b.set("A1", 2);
    b.e.recalc();
    expect(b.get("B1")).toBe(8);
  });

  it("LAMBDA olmayan ad çağrılırsa #DEĞER!, bilinmeyen ad #AD?", () => {
    const b = book();
    b.e.defineName("Oran", "=0.18");
    b.set("A1", "=Oran(2)");
    b.set("A2", "=YOKBOYLE(2)");
    b.e.recalc();
    expect(err(b.get("A1"))).toBe("#VALUE!");
    expect(err(b.get("A2"))).toBe("#NAME?");
  });

  it("LET değişkenine bağlı LAMBDA", () => {
    expect(calc("=LET(f,LAMBDA(x,x+1),f(2))")).toBe(3);
  });
});

describe("bilgi ve metin işlevleri", () => {
  it("TÜR: tek öğeli dizi sabiti de dizidir (64)", () => {
    expect(calc("=TYPE({1})")).toBe(64);
    expect(calc("=TYPE({1,2})")).toBe(64);
    expect(calc("=TYPE(1)")).toBe(1);
    expect(calc("=TYPE(\"a\")")).toBe(2);
  });

  it("SAYIDEĞERİ: ondalık ayırıcı verilince yerel grup ayırıcısıyla çakışmaz", () => {
    expect(calcTr('=SAYIDEĞERİ("1.5";".")')).toBe(1.5);
    expect(calcTr('=SAYIDEĞERİ("1.234,5")')).toBe(1234.5);
    expect(calcTr('=SAYIDEĞERİ("1,234.5";".";",")')).toBe(1234.5);
    expect(calcTr('=SAYIDEĞERİ("12%")')).toBe(0.12);
  });
});

describe("T.TEST", () => {
  it("eşleştirilmiş testte farkların varyansı sıfırsa #SAYI/0!", () => {
    expect(err(calc("=T.TEST({1,2,3},{2,3,4},2,1)"))).toBe("#DIV/0!");
    expect(err(calc("=TTEST({5,5,5},{5,5,5},2,2)"))).toBe("#DIV/0!");
    const v = calc("=T.TEST({1,2,4},{2,3,7},2,1)");
    expect(isError(v)).toBe(false);
    // Farklar -1, -1, -3 → t = -2,5 (çift kuyruk, 2 serbestlik derecesi)
    expect(v as number).toBeCloseTo(0.12965, 4);
  });
});

describe("Türkçe büyük/küçük harf duyarsız joker eşleşme", () => {
  it("EĞERSAY/KAÇINCI/MBUL i/İ ve ı/I çiftlerini eşler", () => {
    const cells = { A1: "İSTANBUL", A2: "Isparta", A3: "IŞIK", A4: "istanbul" };
    expect(calcTr('=EĞERSAY(A1:A4;"ist*")', cells)).toBe(2);
    expect(calcTr('=EĞERSAY(A1:A4;"ış*")', cells)).toBe(1);
    expect(calcTr('=KAÇINCI("?ŞIK";A1:A4;0)', cells)).toBe(3);
    expect(calcTr('=MBUL("t?n";"İSTANBUL")')).toBe(3);
    // Büyük/küçük harf duyarlı BUL etkilenmez
    expect(err(calcTr('=BUL("i";"İSTANBUL")'))).toBe("#VALUE!");
  });
});

describe("metin eşleyici (filtre, koşullu biçim, Bul ortak)", () => {
  it("eşittir/ile başlar/ile biter/içerir: joker ve Türkçe büyük/küçük harf", async () => {
    const { textMatcher } = await import("../src");
    expect(textMatcher("izmir", "equals")("İzmir")).toBe(true);
    expect(textMatcher("ısparta", "equals")("ISPARTA")).toBe(true);
    expect(textMatcher("ış", "begins")("IŞIK")).toBe(true);
    expect(textMatcher("bul", "ends")("İSTANBUL")).toBe(true);
    expect(textMatcher("STA", "contains")("istanbul")).toBe(true);
    expect(textMatcher("i?mir", "equals")("İZMİR")).toBe(true);
    expect(textMatcher("a*", "equals", false)("abc")).toBe(false);
    expect(textMatcher("a*", "equals", false)("a*")).toBe(true);
    expect(textMatcher("1.5", "contains")("x1.5y")).toBe(true);
    expect(textMatcher("1.5", "contains")("x125y")).toBe(false);
  });
});
