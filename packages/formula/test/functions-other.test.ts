import { describe, expect, it } from "vitest";
import { calc, calcTr, err, book, NOW } from "./helpers";

describe("mantıksal", () => {
  it("EĞER ailesi", () => {
    expect(calc('=IF(1>0,"evet","hayır")')).toBe("evet");
    expect(calc('=IF(0,"evet")')).toBe(false);
    expect(calc("=IF(TRUE,A1)")).toBe(0);
    expect(calc("=IF(1,1/0,2)") instanceof Object).toBe(true);
    expect(calc("=IF(0,1/0,2)")).toBe(2);
    expect(calc('=IF("x",1,2)') instanceof Object).toBe(true);
    expect(calc("=SUM(IF({1,0,1},{1,2,3},10))")).toBe(14);
    expect(calc('=IFS(1>2,"a",2>1,"b")')).toBe("b");
    expect(err(calc("=IFS(FALSE,1)"))).toBe("#N/A");
    expect(calc("=IFERROR(1/0,99)")).toBe(99);
    expect(calc("=IFERROR(5,99)")).toBe(5);
    expect(calc('=IFNA(NA(),"yok")')).toBe("yok");
    expect(err(calc('=IFNA(1/0,"yok")'))).toBe("#DIV/0!");
    expect(calc('=SWITCH(2,1,"bir",2,"iki","diğer")')).toBe("iki");
    expect(calc('=SWITCH(9,1,"bir","diğer")')).toBe("diğer");
    expect(err(calc('=SWITCH(9,1,"bir")'))).toBe("#N/A");
    expect(calc("=SUM(IFERROR({1,2}/{0,1},0))")).toBe(2);
  });
  it("VE/YADA/DEĞİL/ÖZELVEYA", () => {
    expect(calc("=AND(TRUE,1)")).toBe(true);
    expect(calc("=AND(TRUE,0)")).toBe(false);
    expect(calc("=OR(FALSE,0,1)")).toBe(true);
    expect(calc("=AND(A1:A3)", { A1: true, A2: "x", A3: true })).toBe(true);
    expect(err(calc("=AND(A1:A2)"))).toBe("#VALUE!");
    expect(calc("=NOT(TRUE)")).toBe(false);
    expect(calc("=XOR(TRUE,TRUE,TRUE)")).toBe(true);
    expect(calc("=TRUE()")).toBe(true);
  });
  it("LET ve LAMBDA", () => {
    expect(calc("=LET(x,2,y,x*3,x+y)")).toBe(8);
    expect(calc("=LET(f,LAMBDA(a,b,a*b),f(3,4))")).toBe(12);
    expect(calc("=SUM(MAP({1,2,3},LAMBDA(v,v*v)))")).toBe(14);
    expect(calc("=REDUCE(0,{1,2,3},LAMBDA(a,v,a+v))")).toBe(6);
    expect(calc("=SUM(BYROW({1,2;3,4},LAMBDA(r,SUM(r))))")).toBe(10);
  });
});

describe("metin", () => {
  it("birleştirme ve parçalar", () => {
    expect(calc('=CONCATENATE("a","b",1)')).toBe("ab1");
    expect(calc("=CONCAT(A1:B1)", { A1: "x", B1: 2 })).toBe("x2");
    expect(calc('=TEXTJOIN("-",TRUE,A1:A3)', { A1: "a", A3: "c" })).toBe("a-c");
    expect(calc('=TEXTJOIN("-",FALSE,A1:A3)', { A1: "a", A3: "c" })).toBe("a--c");
    expect(calc('=LEFT("Merhaba",3)')).toBe("Mer");
    expect(calc('=LEFT("Merhaba")')).toBe("M");
    expect(calc('=RIGHT("Merhaba",2)')).toBe("ba");
    expect(calc('=MID("Merhaba",2,3)')).toBe("erh");
    expect(err(calc('=MID("a",0,1)'))).toBe("#VALUE!");
    expect(calc('=LEN("Çağrı")')).toBe(5);
    expect(calc("=LEN(123.5)")).toBe(5);
  });
  it("Türkçe büyük/küçük harf", () => {
    expect(calc('=UPPER("istanbul ılık")')).toBe("İSTANBUL ILIK");
    expect(calc('=LOWER("İSTANBUL IŞIK")')).toBe("istanbul ışık");
    expect(calc('=PROPER("iSTANBUL ılgaz dağı")')).toBe("İstanbul Ilgaz Dağı");
    expect(calc('=PROPER("ali-veli 2b")')).toBe("Ali-Veli 2B");
  });
  it("arama ve değiştirme", () => {
    expect(calc('=TRIM("  a   b  ")')).toBe("a b");
    expect(calc('=SUBSTITUTE("a-b-c","-","+")')).toBe("a+b+c");
    expect(calc('=SUBSTITUTE("a-b-c","-","+",2)')).toBe("a-b+c");
    expect(calc('=REPLACE("abcdef",2,3,"X")')).toBe("aXef");
    expect(calc('=FIND("b","abcb")')).toBe(2);
    expect(calc('=FIND("b","abcb",3)')).toBe(4);
    expect(err(calc('=FIND("B","abc")'))).toBe("#VALUE!");
    expect(calc('=SEARCH("B","abc")')).toBe(2);
    expect(calc('=SEARCH("?c","abc")')).toBe(2);
    expect(calc('=SEARCH("İ","kitap")')).toBe(2);
    expect(calc('=EXACT("a","A")')).toBe(false);
    expect(calc('=REPT("ab",3)')).toBe("ababab");
    expect(calc("=CHAR(65)")).toBe("A");
    expect(calc("=CHAR(208)")).toBe("Ğ");
    expect(calc('=CODE("Ş")')).toBe(222);
    expect(calc('=UNICODE("€")')).toBe(8364);
    expect(calc('=T("a")&T(1)')).toBe("a");
    expect(calc('=CLEAN("a"&CHAR(9)&"b")')).toBe("ab");
  });
  it("METNEÇEVİR, SAYIYAÇEVİR, SAYIDÜZENLE", () => {
    expect(calc('=TEXT(1234.567,"#,##0.00")')).toBe("1.234,57");
    expect(calc('=TEXT(1234.567,"#.##0,00")')).toBe("1.234,57");
    expect(calc('=TEXT(0.256,"0%")')).toBe("26%");
    expect(calc('=TEXT(45366,"dd.mm.yyyy")')).toBe("15.03.2024");
    expect(calc('=TEXT(45366,"gg.aa.yyyy")')).toBe("15.03.2024");
    expect(calc('=TEXT(45366,"dddd")')).toBe("Cuma");
    expect(calc('=TEXT(45366,"mmmm")')).toBe("Mart");
    expect(calc('=TEXT(0.75,"hh:mm")')).toBe("18:00");
    expect(calc('=VALUE("1.234,5")')).toBe(1234.5);
    expect(calc('=VALUE("12%")')).toBe(0.12);
    expect(err(calc('=VALUE("abc")'))).toBe("#VALUE!");
    expect(calc('=NUMBERVALUE("1.234,5",",",".")')).toBe(1234.5);
    expect(calc("=FIXED(1234.567,1)")).toBe("1.234,6");
    expect(calc("=FIXED(1234.567,1,TRUE)")).toBe("1234,6");
    expect(calc("=DOLLAR(1234.5)")).toBe("₺1.234,50");
  });
  it("yeni metin fonksiyonları", () => {
    expect(calc('=TEXTBEFORE("ad.soyad@x.com","@")')).toBe("ad.soyad");
    expect(calc('=TEXTAFTER("a-b-c","-",2)')).toBe("c");
    expect(calc('=TEXTAFTER("a-b-c","-",-1)')).toBe("c");
    expect(err(calc('=TEXTAFTER("abc","-")'))).toBe("#N/A");
    expect(calc('=COLUMNS(TEXTSPLIT("a,b,c",","))')).toBe(3);
    expect(calc('=ROWS(TEXTSPLIT("a,b;c,d",",",";"))')).toBe(2);
  });
});

const table = {
  A1: "Elma", B1: 10, C1: "kırmızı",
  A2: "Armut", B2: 20, C2: "yeşil",
  A3: "Muz", B3: 30, C3: "sarı",
  E1: 1, E2: 3, E3: 5, E4: 7,
  F1: "a", F2: "b", F3: "c", F4: "d",
};

describe("arama ve başvuru", () => {
  it("DÜŞEYARA / YATAYARA", () => {
    expect(calc('=VLOOKUP("armut",A1:C3,2,FALSE)', table)).toBe(20);
    expect(calc('=VLOOKUP("Mu*",A1:C3,3,FALSE)', table)).toBe("sarı");
    expect(err(calc('=VLOOKUP("Kiraz",A1:C3,2,FALSE)', table))).toBe("#N/A");
    expect(err(calc('=VLOOKUP("Elma",A1:C3,4,FALSE)', table))).toBe("#REF!");
    expect(calc("=VLOOKUP(4,E1:F4,2)", table)).toBe("b");
    expect(calc("=VLOOKUP(100,E1:F4,2,TRUE)", table)).toBe("d");
    expect(err(calc("=VLOOKUP(0,E1:F4,2)", table))).toBe("#N/A");
    expect(calc("=VLOOKUP(5,E:F,2,0)", table)).toBe("c");
    expect(calc('=HLOOKUP("Armut",{"Elma","Armut";1,2},2,FALSE)')).toBe(2);
  });
  it("İNDİS / KAÇINCI", () => {
    expect(calc("=INDEX(A1:C3,2,3)", table)).toBe("yeşil");
    expect(calc("=INDEX(E1:E4,3)", table)).toBe(5);
    expect(calc("=INDEX({1,2,3},2)")).toBe(2);
    expect(calc("=SUM(INDEX(B1:C3,0,1))", table)).toBe(60);
    expect(err(calc("=INDEX(A1:C3,5,1)", table))).toBe("#REF!");
    expect(calc("=SUM(B1:INDEX(B1:B3,2))", table)).toBe(30);
    expect(calc('=MATCH("Muz",A1:A3,0)', table)).toBe(3);
    expect(calc("=MATCH(6,E1:E4)", table)).toBe(3);
    expect(calc("=MATCH(6,{7,5,3,1},-1)")).toBe(1);
    expect(err(calc('=MATCH("x",A1:A3,0)', table))).toBe("#N/A");
    expect(calc('=INDEX(B1:B3,MATCH("Armut",A1:A3,0))', table)).toBe(20);
  });
  it("ÇAPRAZARA / ÇAPRAZEŞLEŞTİR", () => {
    expect(calc('=XLOOKUP("Muz",A1:A3,B1:B3)', table)).toBe(30);
    expect(calc('=XLOOKUP("Kiraz",A1:A3,B1:B3,"yok")', table)).toBe("yok");
    expect(err(calc('=XLOOKUP("Kiraz",A1:A3,B1:B3)', table))).toBe("#N/A");
    expect(calc("=XLOOKUP(4,E1:E4,F1:F4,,-1)", table)).toBe("b");
    expect(calc("=XLOOKUP(4,E1:E4,F1:F4,,1)", table)).toBe("c");
    expect(calc('=XLOOKUP("A*",A1:A3,B1:B3,,2)', table)).toBe(20);
    expect(calc("=XLOOKUP(5,E1:E4,F1:F4,,0,-1)", table)).toBe("c");
    expect(calc("=XLOOKUP(5,E1:E4,F1:F4,,0,2)", table)).toBe("c");
    expect(calc('=COLUMNS(XLOOKUP("Armut",A1:A3,B1:C3))', table)).toBe(2);
    expect(calc('=XMATCH("Muz",A1:A3)', table)).toBe(3);
    expect(calc("=XMATCH(6,E1:E4,1)", table)).toBe(4);
  });
  it("ELEMAN, ARA, KAYDIR, DOLAYLI", () => {
    expect(calc('=CHOOSE(2,"a","b","c")')).toBe("b");
    expect(err(calc('=CHOOSE(4,"a","b","c")'))).toBe("#VALUE!");
    expect(calc("=LOOKUP(4,E1:E4,F1:F4)", table)).toBe("b");
    expect(calc("=LOOKUP(6,{1,3,5,7})")).toBe(5);
    expect(calc("=SUM(OFFSET(B1,1,0,2,1))", table)).toBe(50);
    expect(calc("=OFFSET(A1,2,1)", table)).toBe(30);
    expect(err(calc("=OFFSET(A1,-1,0)", table))).toBe("#REF!");
    expect(calc('=INDIRECT("B2")', table)).toBe(20);
    expect(calc('=SUM(INDIRECT("B1:B3"))', table)).toBe(60);
    expect(calc('=INDIRECT("R2C2",FALSE)', table)).toBe(20);
    expect(err(calc('=INDIRECT("??")', table))).toBe("#REF!");
  });
  it("SATIR, SÜTUN, ADRES, DEVRİK", () => {
    expect(calc("=ROW(C5)")).toBe(5);
    expect(calc("=COLUMN(C5)")).toBe(3);
    expect(calc("=ROW()")).toBe(1000);
    expect(calc("=COLUMN()")).toBe(26);
    expect(calc("=ROWS(A1:C3)")).toBe(3);
    expect(calc("=COLUMNS(A:C)")).toBe(3);
    expect(calc("=ROWS(A:A)")).toBe(1048576);
    expect(calc("=ADDRESS(2,3)")).toBe("$C$2");
    expect(calc("=ADDRESS(2,3,4)")).toBe("C2");
    expect(calc('=ADDRESS(2,3,1,TRUE,"Sayfa 1")')).toBe("'Sayfa 1'!$C$2");
    expect(calc("=ADDRESS(2,3,1,FALSE)")).toBe("R2C3");
    expect(calc("=INDEX(TRANSPOSE({1,2,3}),3,1)")).toBe(3);
    expect(calc("=SUM(ROW(A1:A3))")).toBe(6);
  });
});

describe("tarih ve saat", () => {
  it("TARİH ve bileşenler", () => {
    expect(calc("=DATE(2024,1,1)")).toBe(45292);
    expect(calc("=DATE(1900,1,1)")).toBe(1);
    expect(calc("=DATE(1900,2,29)")).toBe(60);
    expect(calc("=DATE(1900,3,1)")).toBe(61);
    expect(calc("=DATE(2024,13,1)")).toBe(calc("=DATE(2025,1,1)"));
    expect(calc("=DATE(2024,3,0)")).toBe(calc("=DATE(2024,2,29)"));
    expect(calc("=DATE(99,1,1)")).toBe(calc("=DATE(1999,1,1)"));
    expect(calc("=YEAR(45292)")).toBe(2024);
    expect(calc("=MONTH(45292)")).toBe(1);
    expect(calc("=DAY(60)")).toBe(29);
    expect(calc("=MONTH(60)")).toBe(2);
    expect(calc('=YEAR("15.03.2024")')).toBe(2024);
    expect(calc("=TIME(12,30,0)")).toBeCloseTo(0.520833, 5);
    expect(calc("=HOUR(0.75)")).toBe(18);
    expect(calc("=MINUTE(TIME(1,2,3))")).toBe(2);
    expect(calc("=SECOND(TIME(1,2,3))")).toBe(3);
    expect(calc("=TODAY()")).toBe(Math.floor(NOW));
    expect(calc("=NOW()")).toBe(NOW);
    expect(err(calc("=YEAR(-1)"))).toBe("#NUM!");
  });
  it("hafta günleri ve hafta numarası", () => {
    expect(calc("=WEEKDAY(45292)")).toBe(2);
    expect(calc("=WEEKDAY(45292,2)")).toBe(1);
    expect(calc("=WEEKDAY(45292,3)")).toBe(0);
    expect(calc("=WEEKDAY(1)")).toBe(1);
    expect(calc("=WEEKNUM(DATE(2024,1,1))")).toBe(1);
    expect(calc("=WEEKNUM(DATE(2024,3,15))")).toBe(11);
    expect(calc("=WEEKNUM(DATE(2024,3,15),2)")).toBe(11);
    expect(calc("=ISOWEEKNUM(DATE(2021,1,3))")).toBe(53);
  });
  it("tarih aritmetiği", () => {
    expect(calc("=EDATE(DATE(2024,1,31),1)")).toBe(calc("=DATE(2024,2,29)"));
    expect(calc("=EOMONTH(DATE(2024,1,15),1)")).toBe(calc("=DATE(2024,2,29)"));
    expect(calc("=EOMONTH(DATE(2024,1,15),-1)")).toBe(calc("=DATE(2023,12,31)"));
    expect(calc('=DATEDIF(DATE(2020,5,15),DATE(2024,3,10),"Y")')).toBe(3);
    expect(calc('=DATEDIF(DATE(2020,5,15),DATE(2024,3,10),"M")')).toBe(45);
    expect(calc('=DATEDIF(DATE(2020,5,15),DATE(2024,3,10),"D")')).toBe(1395);
    expect(calc('=DATEDIF(DATE(2020,5,15),DATE(2024,3,10),"MD")')).toBe(24);
    expect(calc('=DATEDIF(DATE(2020,5,15),DATE(2024,3,10),"YM")')).toBe(9);
    expect(err(calc('=DATEDIF(DATE(2024,1,1),DATE(2020,1,1),"Y")'))).toBe("#NUM!");
    expect(calc("=DAYS(DATE(2024,3,1),DATE(2024,2,1))")).toBe(29);
    expect(calc("=DAYS360(DATE(2024,1,30),DATE(2024,2,29))")).toBe(29);
    expect(calc("=DAYS360(DATE(2024,1,31),DATE(2024,3,31))")).toBe(60);
    expect(calc("=NETWORKDAYS(DATE(2024,3,1),DATE(2024,3,31))")).toBe(21);
    expect(calc("=NETWORKDAYS(DATE(2024,3,1),DATE(2024,3,31),DATE(2024,3,18))")).toBe(20);
    expect(calc("=NETWORKDAYS.INTL(DATE(2024,3,1),DATE(2024,3,31),11)")).toBe(26);
    expect(calc("=WORKDAY(DATE(2024,3,15),1)")).toBe(calc("=DATE(2024,3,18)"));
    expect(calc("=WORKDAY(DATE(2024,3,18),-1)")).toBe(calc("=DATE(2024,3,15)"));
    expect(calc('=WORKDAY.INTL(DATE(2024,3,15),1,"0000011")')).toBe(calc("=DATE(2024,3,18)"));
    expect(calc('=DATEVALUE("15.03.2024")')).toBe(45366);
    expect(calc('=DATEVALUE("15 Mart 2024")')).toBe(45366);
    expect(calc('=TIMEVALUE("18:00")')).toBe(0.75);
    expect(calc("=YEARFRAC(DATE(2024,1,1),DATE(2024,7,1))")).toBe(0.5);
    expect(calc("=YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),3)")).toBeCloseTo(182 / 365);
  });
});

describe("bilgi", () => {
  it("E... fonksiyonları", () => {
    const c = { A1: 1, A2: "x", A3: true, A5: "=1/0", A6: "=NA()" };
    expect(calc("=ISBLANK(A4)", c)).toBe(true);
    expect(calc("=ISBLANK(A1)", c)).toBe(false);
    expect(calc("=ISNUMBER(A1)", c)).toBe(true);
    expect(calc("=ISTEXT(A2)", c)).toBe(true);
    expect(calc("=ISNONTEXT(A1)", c)).toBe(true);
    expect(calc("=ISLOGICAL(A3)", c)).toBe(true);
    expect(calc("=ISERROR(A5)", c)).toBe(true);
    expect(calc("=ISERR(A6)", c)).toBe(false);
    expect(calc("=ISNA(A6)", c)).toBe(true);
    expect(calc("=ISEVEN(4)")).toBe(true);
    expect(calc("=ISODD(-3)")).toBe(true);
    expect(calc("=ISREF(A1)")).toBe(true);
    expect(calc("=ISREF(1)")).toBe(false);
    expect(calc("=ISFORMULA(A5)", c)).toBe(true);
    expect(calc("=FORMULATEXT(A5)", c)).toBe("=1/0");
    expect(calc("=TYPE(1)")).toBe(1);
    expect(calc('=TYPE("a")')).toBe(2);
    expect(calc("=TYPE(TRUE)")).toBe(4);
    expect(calc("=TYPE(1/0)")).toBe(16);
    expect(calc("=TYPE({1,2})")).toBe(64);
    expect(calc("=ERROR.TYPE(1/0)")).toBe(2);
    expect(err(calc("=ERROR.TYPE(1)"))).toBe("#N/A");
    expect(calc("=N(TRUE)")).toBe(1);
    expect(calc('=N("a")')).toBe(0);
    expect(calc("=SHEET()")).toBe(1);
    expect(calc("=SHEETS()")).toBe(1);
    expect(calc('=CELL("address",B3)')).toBe("$B$3");
    expect(calc('=CELL("row",B3)')).toBe(3);
  });
});

describe("finansal", () => {
  it("annüite", () => {
    expect(calc("=PMT(0.08/12,10,10000)")).toBeCloseTo(-1037.03, 2);
    expect(calc("=PMT(0.08/12,10,10000,0,1)")).toBeCloseTo(-1030.16, 2);
    expect(calc("=PMT(0,10,1000)")).toBe(-100);
    expect(calc("=FV(0.06/12,10,-200,-500,1)")).toBeCloseTo(2581.4, 1);
    expect(calc("=PV(0.08/12,12*20,500)")).toBeCloseTo(-59777.15, 2);
    expect(calc("=NPER(0.12/12,-100,-1000,10000,1)")).toBeCloseTo(59.6739, 3);
    expect(calc("=RATE(4*12,-200,8000)")).toBeCloseTo(0.0077, 4);
    expect(calc("=IPMT(0.1/12,1,36,8000)")).toBeCloseTo(-66.67, 2);
    expect(calc("=PPMT(0.1/12,1,24,2000)")).toBeCloseTo(-75.62, 2);
    expect(calc("=CUMIPMT(0.09/12,30*12,125000,13,24,0)")).toBeCloseTo(-11135.23, 2);
  });
  it("nakit akışı", () => {
    expect(calc("=NPV(0.1,-10000,3000,4200,6800)")).toBeCloseTo(1188.44, 2);
    expect(calc("=IRR({-70000,12000,15000,18000,21000,26000})")).toBeCloseTo(0.0866, 4);
    expect(err(calc("=IRR({1,2,3})"))).toBe("#NUM!");
    expect(calc("=XNPV(0.09,{-10000,2750,4250,3250,2750},{39448,39508,39751,39859,39904})")).toBeCloseTo(2086.65, 2);
    expect(calc("=XIRR({-10000,2750,4250,3250,2750},{39448,39508,39751,39859,39904})")).toBeCloseTo(0.373363, 5);
    expect(calc("=MIRR({-120000,39000,30000,21000,37000,46000},0.1,0.12)")).toBeCloseTo(0.126094, 5);
  });
  it("amortisman ve faiz", () => {
    expect(calc("=SLN(30000,7500,10)")).toBe(2250);
    expect(calc("=SYD(30000,7500,10,1)")).toBeCloseTo(4090.91, 2);
    expect(calc("=DDB(2400,300,10,1)")).toBe(480);
    expect(calc("=DDB(2400,300,10,2)")).toBe(384);
    expect(calc("=DB(1000000,100000,6,1,7)")).toBeCloseTo(186083.33, 2);
    expect(calc("=DB(1000000,100000,6,2,7)")).toBeCloseTo(259639.42, 2);
    expect(calc("=EFFECT(0.0525,4)")).toBeCloseTo(0.053543, 6);
    expect(calc("=NOMINAL(0.053543,4)")).toBeCloseTo(0.0525, 5);
  });
});

describe("dinamik diziler ve taşma", () => {
  it("SIRALA/FİLTRE/BENZERSİZ/SIRALI", () => {
    const b = book({ A1: 3, A2: 1, A3: 2, B1: "c", B2: "a", B3: "b", D1: "=SORT(A1:A3)", E1: "=SORT(A1:B3,2,-1)", G1: "=FILTER(A1:A3,A1:A3>1)", H1: "=UNIQUE({1;2;2;3;1})", I1: "=SEQUENCE(2,3,10,5)" });
    expect([b.get("D1"), b.get("D2"), b.get("D3")]).toEqual([1, 2, 3]);
    expect([b.get("E1"), b.get("F1"), b.get("E3"), b.get("F3")]).toEqual([3, "c", 1, "a"]);
    expect([b.get("G1"), b.get("G2"), b.get("G3")]).toEqual([3, 2, null]);
    expect([b.get("H1"), b.get("H2"), b.get("H3"), b.get("H4")]).toEqual([1, 2, 3, null]);
    expect([b.get("I1"), b.get("J1"), b.get("K1"), b.get("I2"), b.get("K2")]).toEqual([10, 15, 20, 25, 35]);
    expect(b.e.getSpill(b.s, 1, 3)).toEqual({ anchorRow: 0, anchorCol: 3, rows: 3, cols: 1 });
  });
  it("#SPILL! engelleme ve kaldırma", () => {
    const b = book({ A1: "=SEQUENCE(3)", A3: "engel" });
    expect(err(b.get("A1"))).toBe("#SPILL!");
    b.set("A3", null);
    b.e.recalc();
    expect(b.get("A1")).toBe(1);
    expect(b.get("A3")).toBe(3);
    b.set("A2", 99);
    b.e.recalc();
    expect(err(b.get("A1"))).toBe("#SPILL!");
  });
  it("taşan alana başvuru ve bağımlılık", () => {
    const b = book({ A1: 2, B1: "=SEQUENCE(A1)", C1: "=SUM(B1:B5)" });
    expect(b.get("C1")).toBe(3);
    b.set("A1", 4);
    b.e.recalc();
    expect(b.get("C1")).toBe(10);
    expect(b.get("B4")).toBe(4);
    b.set("A1", 1);
    b.e.recalc();
    expect(b.get("B2")).toBe(null);
    expect(b.get("C1")).toBe(1);
  });
  it("aralık formülü taşar", () => {
    const b = book({ A1: 1, A2: 2, B1: "=A1:A2*10" });
    expect([b.get("B1"), b.get("B2")]).toEqual([10, 20]);
    const c = book({ A1: 1, B1: "=A1:A3" });
    expect([c.get("B1"), c.get("B2"), c.get("B3")]).toEqual([1, 0, 0]);
  });
  it("yığma ve seçme", () => {
    expect(calc("=SUM(VSTACK({1,2},{3,4}))")).toBe(10);
    expect(calc("=COLUMNS(HSTACK({1,2},{3}))")).toBe(3);
    expect(calc("=SUM(TAKE({1;2;3},2))")).toBe(3);
    expect(calc("=SUM(DROP({1;2;3},-1))")).toBe(3);
    expect(calc("=SUM(CHOOSECOLS({1,2,3},1,-1))")).toBe(4);
    expect(calc("=ROWS(TOCOL({1,2;3,4}))")).toBe(4);
    expect(calc("=ROWS(WRAPROWS({1,2,3,4,5},2))")).toBe(3);
    expect(calc("=SUM(SORTBY({1;2;3},{3;1;2}))")).toBe(6);
    expect(calc("=INDEX(SORTBY({\"a\";\"b\";\"c\"},{3;1;2}),1)")).toBe("b");
    expect(err(calc("=FILTER({1,2},{FALSE,FALSE})"))).toBe("#CALC!");
    expect(calc('=FILTER({1,2},{FALSE,FALSE},"boş")')).toBe("boş");
  });
});

describe("Türkçe formüller", () => {
  it("Türkçe adlarla hesap", () => {
    expect(calcTr("=TOPLA(1;2;3)")).toBe(6);
    expect(calcTr("=ORTALAMA(2;4)")).toBe(3);
    expect(calcTr('=EĞER(1>0;"evet";"hayır")')).toBe("evet");
    expect(calcTr("=YUVARLA(2,345;2)")).toBe(2.35);
    expect(calcTr('=BÜYÜKHARF("istanbul")')).toBe("İSTANBUL");
    expect(calcTr('=DÜŞEYARA("Muz";A1:C3;2;YANLIŞ)', table)).toBe(30);
    expect(calcTr("=EĞERHATA(1/0;0)")).toBe(0);
    expect(calcTr("=BAĞ_DEĞ_SAY(E1:E4)", table)).toBe(4);
    expect(calcTr("=BAĞ_DEĞ_DOLU_SAY(A1:A3)", table)).toBe(3);
    expect(calcTr('=EĞERSAY(E1:E4;">2")', table)).toBe(3);
    expect(calcTr("=YAZIM.DÜZENİ(\"ali veli\")")).toBe("Ali Veli");
    expect(calcTr("=topla(1;2)")).toBe(3);
    expect(calcTr("=SUM(1;2)")).toBe(3);
    expect(calcTr("=VE(DOĞRU;YANLIŞ)")).toBe(false);
    expect(calcTr("=TOPLA({1\\2;3\\4})")).toBe(10);
    expect(calcTr("=MİN(3;1;2)+MAK(1;5)")).toBe(6);
    expect(calcTr("=İNDİS(E1:E4;KAÇINCI(5;E1:E4;0))", table)).toBe(5);
  });
});
