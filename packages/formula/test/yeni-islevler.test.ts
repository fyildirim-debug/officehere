/** Yeni işlevler: veritabanı (VSEÇTOPLA…), mühendislik (ONDAİKİ…, ÇEVİR), düzenli ifade, EKSİKMİ, ISO.TAVAN. Beklenen sonuçlar masaüstü hesap tablosunun belgelenmiş örnekleriyle aynıdır. */
import { describe, expect, it } from "vitest";
import { FUNCTIONS, isArray, toLocalFormula, fromLocalFormula } from "../src";
import { book, calc, calcTr, err } from "./helpers";

/** Ağaç veritabanı (A5:E11) ve ölçüt aralığı (A1:F3). */
function trees() {
  const b = book({
    A1: "Ağaç", B1: "Boy", C1: "Yaş", D1: "Verim", E1: "Kâr", F1: "Boy",
    A2: "=Elma", B2: ">10", F2: "<16",
    A3: "=Armut",
    A5: "Ağaç", B5: "Boy", C5: "Yaş", D5: "Verim", E5: "Kâr",
    A6: "Elma", B6: 18, C6: 20, D6: 14, E6: 105,
    A7: "Armut", B7: 12, C7: 12, D7: 10, E7: 96,
    A8: "Kiraz", B8: 13, C8: 14, D8: 9, E8: 105,
    A9: "Elma", B9: 14, C9: 15, D9: 10, E9: 75,
    A10: "Armut", B10: 9, C10: 8, D10: 8, E10: 76.8,
    A11: "Elma", B11: 8, C11: 9, D11: 6, E11: 45,
  });
  // Ölçüt hücreleri formülle yazılmış gibi metin: "=Elma" (tam eşleşme)
  b.set("A2", "'=Elma");
  b.set("A3", "'=Armut");
  return b;
}

function evalIn(b: ReturnType<typeof book>, f: string) {
  b.set("H1", f);
  b.e.recalc();
  return b.get("H1");
}

describe("veritabanı işlevleri", () => {
  it("belgelenmiş örnek sonuçları", () => {
    const b = trees();
    // "'=Elma" metni kesme işaretiyle girildi; motorda değer "=Elma" olmalı
    b.e.setValue(b.s, 1, 0, "=Elma");
    b.e.setValue(b.s, 2, 0, "=Armut");
    expect(evalIn(b, '=DCOUNT(A5:E11,"Yaş",A1:F2)')).toBe(1);
    expect(evalIn(b, '=DCOUNTA(A5:E11,"Kâr",A1:F2)')).toBe(1);
    expect(evalIn(b, '=DMAX(A5:E11,"Kâr",A1:A3)')).toBe(105);
    expect(evalIn(b, '=DMIN(A5:E11,"Kâr",A1:B2)')).toBe(75);
    expect(evalIn(b, '=DSUM(A5:E11,"Kâr",A1:A2)')).toBe(225);
    expect(evalIn(b, '=DSUM(A5:E11,"Kâr",A1:F2)')).toBe(75);
    expect(evalIn(b, '=DPRODUCT(A5:E11,"Verim",A1:F2)')).toBe(10);
    expect(evalIn(b, '=DAVERAGE(A5:E11,"Verim",A1:B2)')).toBe(12);
    expect(evalIn(b, "=DAVERAGE(A5:E11,3,A5:E11)")).toBe(13);
    expect(evalIn(b, '=DSTDEV(A5:E11,"Verim",A1:A3)')).toBeCloseTo(2.96647939483827, 12);
    expect(evalIn(b, '=DSTDEVP(A5:E11,"Verim",A1:A3)')).toBeCloseTo(2.65329983228432, 12);
    expect(evalIn(b, '=DVAR(A5:E11,"Verim",A1:A3)')).toBeCloseTo(8.8, 12);
    expect(evalIn(b, '=DVARP(A5:E11,"Verim",A1:A3)')).toBeCloseTo(7.04, 12);
    expect(err(evalIn(b, '=DGET(A5:E11,"Verim",A1:A3)'))).toBe("#NUM!");
  });

  it("işlecsiz metin 'ile başlar', joker ve sayı ölçütleri; alan boşsa kayıt sayısı", () => {
    const b = trees();
    b.e.setValue(b.s, 1, 0, "Ki");
    b.e.setValue(b.s, 1, 1, null);
    b.e.setValue(b.s, 1, 5, null);
    expect(evalIn(b, '=DGET(A5:E11,"Kâr",A1:A2)')).toBe(105);
    b.e.setValue(b.s, 1, 0, "*rm*");
    expect(evalIn(b, '=DSUM(A5:E11,"Kâr",A1:A2)')).toBeCloseTo(172.8, 10);
    b.e.setValue(b.s, 1, 0, null);
    b.e.setValue(b.s, 1, 2, 12);
    expect(evalIn(b, "=DCOUNT(A5:E11,,C1:C2)")).toBe(1);
    expect(evalIn(b, '=DCOUNTA(A5:E11,"Ağaç",C1:C2)')).toBe(1);
    // Boş ölçüt satırı tüm kayıtları seçer
    b.e.setValue(b.s, 1, 2, null);
    expect(evalIn(b, '=DSUM(A5:E11,"Verim",A1:A2)')).toBe(57);
    // Olmayan alan #DEĞER!, eşleşme yoksa VAL #DEĞER!, ortalama #SAYI/0!
    expect(err(evalIn(b, '=DSUM(A5:E11,"Yok",A1:A2)'))).toBe("#VALUE!");
    b.e.setValue(b.s, 1, 0, "Muz");
    expect(err(evalIn(b, '=DGET(A5:E11,"Kâr",A1:A2)'))).toBe("#VALUE!");
    expect(err(evalIn(b, '=DAVERAGE(A5:E11,"Kâr",A1:A2)'))).toBe("#DIV/0!");
    expect(evalIn(b, '=DMAX(A5:E11,"Kâr",A1:A2)')).toBe(0);
  });

  it("Türkçe adlar ve kategori", () => {
    expect(fromLocalFormula("=VSEÇTOPLA(A1:B3;2;D1:D2)", "tr")).toBe("=DSUM(A1:B3,2,D1:D2)");
    expect(toLocalFormula("=DGET(A1:B3,2,D1:D2)+DCOUNTA(A1:B3,,D1:D2)", "tr")).toBe("=VAL(A1:B3;2;D1:D2)+VSEÇSAYDOLU(A1:B3;;D1:D2)");
    expect([...FUNCTIONS.values()].filter((d) => d.category === "database").length).toBe(12);
  });
});

describe("mühendislik işlevleri", () => {
  it("taban dönüşümleri (negatifler 10 basamaklı tümleyen)", () => {
    const cases: [string, unknown][] = [
      ["=DEC2BIN(9)", "1001"], ["=DEC2BIN(9,4)", "1001"], ["=DEC2BIN(-100)", "1110011100"], ["=DEC2BIN(9,2)", "#NUM!"], ["=DEC2BIN(512)", "#NUM!"],
      ["=DEC2HEX(100,4)", "0064"], ["=DEC2HEX(-54)", "FFFFFFFFCA"], ["=DEC2HEX(28)", "1C"], ["=DEC2OCT(58,3)", "072"], ["=DEC2OCT(-100)", "7777777634"],
      ["=BIN2DEC(1100100)", 100], ["=BIN2DEC(1111111111)", -1], ["=BIN2HEX(11111011,4)", "00FB"], ["=BIN2HEX(1110)", "E"], ["=BIN2HEX(1111111111)", "FFFFFFFFFF"],
      ["=BIN2OCT(1001,3)", "011"], ["=BIN2OCT(1100100)", "144"], ["=BIN2OCT(1111111111)", "7777777777"], ["=BIN2DEC(102)", "#NUM!"],
      ['=HEX2BIN("F",8)', "00001111"], ['=HEX2BIN("B7")', "10110111"], ['=HEX2BIN("FFFFFFFFFF")', "1111111111"],
      ['=HEX2DEC("A5")', 165], ['=HEX2DEC("FFFFFFFF5B")', -165], ['=HEX2DEC("3DA408B9")', 1034160313],
      ['=HEX2OCT("F",3)', "017"], ['=HEX2OCT("3B4E")', "35516"], ['=HEX2OCT("FFFFFFFF00")', "7777777400"],
      ["=OCT2BIN(3,3)", "011"], ["=OCT2BIN(7777777000)", "1000000000"], ["=OCT2DEC(54)", 44], ["=OCT2DEC(7777777533)", -165],
      ["=OCT2HEX(100,4)", "0040"], ["=OCT2HEX(7777777533)", "FFFFFFFF5B"],
    ];
    for (const [f, exp] of cases) {
      const v = calc(f);
      expect(err(v) ?? v, f).toBe(exp);
    }
    expect(calcTr("=ONDAİKİ(9;8)")).toBe("00001001");
    expect(calcTr('=ONALTIONDA("FF")')).toBe(255);
  });

  it("bit işlemleri, DELTA ve BESINIR", () => {
    expect(calc("=BITAND(1,5)")).toBe(1);
    expect(calc("=BITAND(13,25)")).toBe(9);
    expect(calc("=BITOR(23,10)")).toBe(31);
    expect(calc("=BITXOR(5,3)")).toBe(6);
    expect(calc("=BITLSHIFT(4,2)")).toBe(16);
    expect(calc("=BITRSHIFT(13,2)")).toBe(3);
    expect(calc("=BITLSHIFT(13,-2)")).toBe(3);
    expect(err(calc("=BITAND(-1,5)"))).toBe("#NUM!");
    expect(err(calc("=BITAND(1.5,5)"))).toBe("#NUM!");
    expect(calc("=DELTA(5,4)")).toBe(0);
    expect(calc("=DELTA(5,5)")).toBe(1);
    expect(calc("=DELTA(0.5)")).toBe(0);
    expect(calc("=GESTEP(5,4)")).toBe(1);
    expect(calc("=GESTEP(-4,-5)")).toBe(1);
    expect(calc("=GESTEP(-1)")).toBe(0);
  });

  it("ÇEVİR (CONVERT): kütle, sıcaklık, uzunluk, alan, bilgi, önekler", () => {
    expect(calc('=CONVERT(1,"lbm","kg")')).toBeCloseTo(0.45359237, 12);
    expect(calc('=CONVERT(68,"F","C")')).toBeCloseTo(20, 12);
    expect(calc('=CONVERT(0,"C","K")')).toBeCloseTo(273.15, 12);
    expect(err(calc('=CONVERT(2.5,"ft","sec")'))).toBe("#N/A");
    expect(calc('=CONVERT(CONVERT(100,"ft","m"),"ft","m")')).toBeCloseTo(9.290304, 12);
    expect(calc('=CONVERT(1,"mi","km")')).toBeCloseTo(1.609344, 12);
    expect(calc('=CONVERT(1,"gal","l")')).toBeCloseTo(3.785411784, 12);
    expect(calc('=CONVERT(1,"kbyte","bit")')).toBe(8000);
    expect(calc('=CONVERT(1,"kibyte","bit")')).toBe(8192);
    expect(calc('=CONVERT(100,"km2","ha")')).toBeCloseTo(10000, 9);
    expect(calc('=CONVERT(1,"hr","mn")')).toBe(60);
    expect(calc('=CONVERT(1,"m/s","kn")')).toBeCloseTo(1.94384449244061, 12);
    expect(calcTr('=ÇEVİR(1;"in";"cm")')).toBeCloseTo(2.54, 12);
    expect(err(calc('=CONVERT(1,"xyz","m")'))).toBe("#N/A");
  });
});

describe("düzenli ifade ve metinleştirme işlevleri", () => {
  it("REGEXTEST", () => {
    expect(calc('=REGEXTEST("abc123","[0-9]+")')).toBe(true);
    expect(calc('=REGEXTEST("ABC","abc")')).toBe(false);
    expect(calc('=REGEXTEST("ABC","abc",1)')).toBe(true);
    expect(calc('=REGEXTEST("Şeker","^Ş")')).toBe(true);
    expect(err(calc('=REGEXTEST("a","(")'))).toBe("#VALUE!");
  });

  it("REGEXEXTRACT: ilk eşleşme, tüm eşleşmeler (alt alta), yakalama grupları (yan yana)", () => {
    expect(calc('=REGEXEXTRACT("abc123def456","[0-9]+")')).toBe("123");
    expect(err(calc('=REGEXEXTRACT("abc","[0-9]+")'))).toBe("#N/A");
    const b = book();
    b.set("A1", '=REGEXEXTRACT("abc123def456","[0-9]+",1)');
    b.set("C1", '=REGEXEXTRACT("Ali Veli","([A-Za-z]+) ([A-Za-z]+)",2)');
    b.e.recalc();
    expect(b.get("A1")).toBe("123");
    expect(b.get("A2")).toBe("456");
    expect(b.get("C1")).toBe("Ali");
    expect(b.get("D1")).toBe("Veli");
  });

  it("REGEXREPLACE: tümü, n. ve sondan n. eşleşme, grup başvurusu", () => {
    expect(calc('=REGEXREPLACE("a1b2c3","[0-9]","#")')).toBe("a#b#c#");
    expect(calc('=REGEXREPLACE("a1b2c3","[0-9]","#",2)')).toBe("a1b#c3");
    expect(calc('=REGEXREPLACE("a1b2c3","[0-9]","#",-1)')).toBe("a1b2c#");
    expect(calc('=REGEXREPLACE("Ali Veli","([A-Za-z]+) ([A-Za-z]+)","$2, $1")')).toBe("Veli, Ali");
    expect(calc('=REGEXREPLACE("Şeker Çay","\\s","-")')).toBe("Şeker-Çay");
    expect(calc('=REGEXREPLACE("AbAb","a","x",0,1)')).toBe("xbxb");
  });

  it("DİZİDENMETNE / DEĞERDENMETNE (Türkçe ayırıcılarla)", () => {
    expect(calc('=ARRAYTOTEXT({1,"a";TRUE,1.5})')).toBe("1; a; DOĞRU; 1,5");
    expect(calc('=ARRAYTOTEXT({1,"a";TRUE,1.5},1)')).toBe('{1\\"a";DOĞRU\\1,5}');
    expect(calc('=VALUETOTEXT("a",1)')).toBe('"a"');
    expect(calc("=VALUETOTEXT(1.5)")).toBe("1,5");
    expect(err(calc("=ARRAYTOTEXT({1},2)"))).toBe("#VALUE!");
  });
});

describe("EKSİKMİ ve ISO.TAVAN", () => {
  it("boş bırakılan LAMBDA parametresi", () => {
    expect(calc("=LAMBDA(a,b,IF(ISOMITTED(b),a,a+b))(1,)")).toBe(1);
    expect(calc("=LAMBDA(a,b,IF(ISOMITTED(b),a,a+b))(1,2)")).toBe(3);
    expect(calc("=LAMBDA(a,b,IF(ISOMITTED(b),a,a+b))(1)")).toBe(1);
    expect(calcTr("=LAMBDA(a;b;EĞER(EKSİKMİ(b);a;a+b))(5;)")).toBe(5);
  });

  it("ISO.TAVAN işaretten bağımsız yukarı yuvarlar", () => {
    expect(calc("=ISO.CEILING(-4.3)")).toBe(-4);
    expect(calc("=ISO.CEILING(4.3,2)")).toBe(6);
    expect(calc("=ISO.CEILING(-4.3,-2)")).toBe(-4);
    expect(calc("=ISO.CEILING(4.3,0)")).toBe(0);
  });

  it("dizi sonuçları dizi olarak döner", () => {
    const b = book();
    b.set("A1", '=VALUETOTEXT({1,"x"},1)');
    b.e.recalc();
    expect(b.get("B1")).toBe('"x"');
    expect(isArray(1)).toBe(false);
  });
});

describe("regresyon: DOT, LOT, EĞİLİM, BÜYÜME", () => {
  // Belgelerdeki ofis binası örneği: y = değer, x1..x4 = alan, ofis, giriş, yaş
  const data = [[142000, 2310, 2, 2, 20], [144000, 2333, 2, 2, 12], [151000, 2356, 3, 1.5, 33], [150000, 2379, 3, 2, 43], [139000, 2402, 2, 3, 53], [169000, 2425, 4, 2, 23], [126000, 2448, 2, 1.5, 99], [142900, 2471, 2, 2, 34], [163000, 2494, 3, 3, 23], [169000, 2517, 4, 4, 55], [149000, 2540, 2, 3, 22]];
  function sheet() {
    const cells: Record<string, number> = {};
    data.forEach((r, i) => r.forEach((v, j) => (cells[`${"ABCDE"[j]}${i + 1}`] = v)));
    return book(cells);
  }

  it("DOT: çoklu regresyon katsayıları ve istatistikler", () => {
    const b = sheet();
    b.set("H1", "=LINEST(A1:A11,B1:E11,TRUE,TRUE)");
    b.e.recalc();
    const row = (r: number) => ["H", "I", "J", "K", "L"].map((c) => b.get(`${c}${r}`));
    const close = (got: unknown[], exp: number[]) => exp.forEach((e, i) => expect(got[i] as number).toBeCloseTo(e, 5));
    close(row(1), [-234.2371645, 2553.21066, 12529.76817, 27.64138737, 52317.83051]);
    close(row(2), [13.26801148, 530.6691519, 400.0668382, 5.429374042, 12237.3616]);
    close(row(3).slice(0, 2), [0.996747993, 970.5784629]);
    expect(err(b.get("J3"))).toBe("#N/A");
    close(row(4).slice(0, 2), [459.7536742, 6]);
    expect((b.get("H5") as number) / 1732393319).toBeCloseTo(1, 9);
    expect(b.get("I5") as number).toBeCloseTo(5652135.316, 2);
  });

  it("DOT basit örnek, sabitsiz ve doğrusal bağımlı sütun", () => {
    const b = book();
    b.set("A1", "=LINEST({1,9,5,7},{0,4,2,3})");
    b.set("A3", "=LINEST({2,4,6},{1,2,3},FALSE)");
    b.set("A5", "=LINEST({1,2,4},{1;1;1})");
    b.e.recalc();
    expect(b.get("A1")).toBeCloseTo(2, 12);
    expect(b.get("B1")).toBeCloseTo(1, 12);
    expect(b.get("A3")).toBeCloseTo(2, 12);
    expect(b.get("B3")).toBe(0);
    expect(calcTr("=İNDİS(DOT({1;9;5;7};{0;4;2;3});1)")).toBeCloseTo(2, 12);
  });

  it("EĞİLİM ve BÜYÜME tahminleri; LOT üstel katsayıları", () => {
    const b = sheet();
    b.set("H1", "=TREND(A1:A11,B1:B11,B1:B3)");
    b.set("J1", "=GROWTH(A1:A11,B1:B11,B1:B2)");
    b.set("L1", "=LOGEST(A1:A11,B1:B11)");
    b.e.recalc();
    expect(b.get("H1")).toBeCloseTo(143000, 6);
    expect(b.get("H3")).toBeCloseTo(145614.5454545, 4);
    expect(b.get("J1")).toBeCloseTo(142935.9391934, 4);
    expect(b.get("J2")).toBeCloseTo(144130.0964155, 4);
    // Birim x başına büyüme oranı: (BÜYÜME(B2)/BÜYÜME(B1))^(1/23)
    expect(b.get("L1")).toBeCloseTo(1.0003617953, 9);
    expect(err(calc("=GROWTH({1,-2,3})"))).toBe("#NUM!");
    expect(calcTr("=EĞİLİM({1;2;3};{1;2;3};4)")).toBeCloseTo(4, 12);
    expect(calcTr("=BÜYÜME({2;4;8};{1;2;3};4)")).toBeCloseTo(16, 10);
  });
});
