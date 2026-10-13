/** 97-2003 çalışma kitabı (.xls, BIFF8) okuyucusu. */
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { dateToSerial } from "@officehere/formula";
import { readBiffStream, readXls } from "../src/io/xls/reader";
import { decodeFormula, type XlsFormulaContext } from "../src/io/xls/formula";
import { rk } from "../src/io/xls/records";
import { openBook } from "../src/io/formats";
import { displayText } from "../src/model/display";

const fixture = (name: string) => {
  const b = readFileSync(new URL(`../src/io/__fixtures__/${name}`, import.meta.url));
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer;
};

describe("XLS okuma (LibreOffice örneği)", () => {
  it("sayfalar, değerler, formüller ve paylaşılan formüller", () => {
    const wb = readXls(fixture("ornek.xls"));
    expect(wb.sheets.map((s) => s.name)).toEqual(["Satışlar", "Özet", "Uzun", "Gizli"]);
    expect(wb.sheets[3].hidden).toBe(true);
    const [s, o, u] = wb.sheets;
    expect(wb.getValue(s, 0, 0)).toBe("Ürün");
    expect(wb.getValue(s, 3, 6)).toBe("ığüşöç ĞÜŞİÖÇ");
    // D2:D4 paylaşılan formül (SHRFMLA) her hücrede kendi satırıyla
    expect(wb.getFormula(s, 1, 3)).toBe("=B2*C2");
    expect(wb.getFormula(s, 3, 3)).toBe("=B4*C4");
    expect(wb.getValue(s, 4, 3)).toBe(234.25);
    expect(wb.getFormula(s, 5, 2)).toBe('=IF(D5>100,"Büyük","Küçük")');
    expect(wb.getFormula(s, 5, 4)).toBe('=VLOOKUP("Şeker",A2:C4,3,FALSE)');
    expect(wb.getFormula(s, 6, 4)).toBe("=IFERROR(1/0,-1)");
    expect(wb.getFormula(s, 6, 6)).toBe('=UPPER(LEFT(G2,2))&MID("abc",2,1)&RIGHT("xyz",1)');
    expect(wb.getFormula(s, 7, 1)).toBe("=$B$2+Özet!B1");
    expect(wb.getValue(s, 7, 1)).toBe(244.25);
    expect(wb.getFormula(o, 0, 1)).toBe("=Satışlar!D5");
    expect(wb.getFormula(o, 4, 0)).toBe("=NOT(TRUE)");
    expect(wb.engine.getNames()).toContainEqual(expect.objectContaining({ name: "Toplam_Adet", formula: "=Satışlar!$B$2:$B$4" }));
    // 360 satırlık sayfa: SST CONTINUE kayıtlarına bölünmüş dizgiler
    expect(wb.getValue(u, 0, 0)).toBe("Row 1 plain ascii text value 13");
    expect(wb.getValue(u, 299, 0)).toBe("Satır 300 — çğıöşü ÇĞİÖŞÜ uzun metin 2100");
    expect(wb.getValue(u, 359, 0)).toBe("Satır 360 — çğıöşü ÇĞİÖŞÜ uzun metin 2520");
    expect(wb.getFormula(u, 358, 2)).toBe("=B359*2");
    expect(wb.getValue(u, 358, 2)).toBe(1077);
  });
  it("biçimler, tarih, birleştirme, boyutlar", () => {
    const wb = readXls(fixture("ornek.xls"));
    const s = wb.sheets[0];
    expect(wb.styleAt(s, 0, 0)).toMatchObject({ bold: true, fill: "#1F4E79", color: "#FFFFFF", hAlign: "center", bBottom: { style: "thin" } });
    expect(wb.styleAt(s, 4, 0)).toMatchObject({ italic: true, underline: "single", size: 14, color: "#C00000" });
    expect(wb.styleAt(s, 4, 3).bBottom?.style).toBe("double");
    expect(wb.styleAt(s, 9, 0)).toMatchObject({ fill: "#FFFF00", wrap: true, vAlign: "middle" });
    expect(wb.getValue(s, 1, 4)).toBe(dateToSerial(2024, 3, 15));
    expect(displayText(wb, s, 1, 4)).toBe("15.03.2024");
    expect(displayText(wb, s, 1, 5)).toBe("25,00%");
    expect(displayText(wb, s, 1, 2)).toBe("2,50");
    expect(displayText(wb, wb.sheets[1], 2, 0)).toBe("1.234.567,891");
    expect(s.merges).toEqual([{ r1: 9, r2: 10, c1: 0, c2: 2 }]);
    expect(s.rows.size(9)).toBe(40);
    expect(s.cols.size(0)).toBeGreaterThan(120);
  });
  it("1904 tarih sistemi 1900 sistemine çevrilir", () => {
    const wb = readXls(fixture("tarih1904.xls"));
    const s = wb.sheets[0];
    expect(wb.getValue(s, 0, 0)).toBe(dateToSerial(2024, 3, 15));
    expect(wb.getValue(s, 1, 0)).toBe(5);
    expect(wb.getValue(s, 2, 0)).toBe(dateToSerial(2024, 3, 20));
  });
  it("şifreli ve bozuk dosyalar anlaşılır hatayla reddedilir", () => {
    expect(() => readXls(fixture("sifreli.xls"))).toThrow(/parola/);
    expect(() => readXls(new Uint8Array(100))).toThrow(/97-2003/);
    const cut = new Uint8Array(fixture("ornek.xls")).slice(0, 4096);
    expect(() => readXls(cut)).toThrow();
  });
  it("biçim tanıma: uzantı yanlış olsa da içerikten .xls anlaşılır", async () => {
    const o = await openBook("yanlis.xlsx", fixture("ornek.xls"));
    expect(o.format).toBe("xls");
    expect(o.notice).toMatch(/Uyumluluk modu/);
  });
});

/** BIFF kaydı: tür + uzunluk + veri. */
function rec(type: number, data: number[]): number[] {
  return [type & 0xff, type >> 8, data.length & 0xff, data.length >> 8, ...data];
}
const le16 = (n: number) => [n & 0xff, (n >> 8) & 0xff];
const le32 = (n: number) => [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >>> 24) & 0xff];
const ascii = (s: string) => [...s].map((c) => c.charCodeAt(0));
const utf16 = (s: string) => [...s].flatMap((c) => le16(c.charCodeAt(0)));

describe("SST ve CONTINUE", () => {
  it("kayıt sınırında bölünen dizgi: her parçada seçenek baytı yeniden okunur", () => {
    const bof = (dt: number) => rec(0x0809, [...le16(0x0600), ...le16(dt), ...new Array(12).fill(0)]);
    // 1. dizgi sıkıştırılmış; 2. dizgi SST'de sıkıştırılmış başlar, CONTINUE'da 16 bit sürer; 3. dizgi CONTINUE'da başlar
    const s2a = "Uzun metin ";
    const s2b = "çğıöşü";
    const sst = rec(0x00fc, [...le32(3), ...le32(3), ...le16(7), 0, ...ascii("Merhaba"), ...le16(s2a.length + s2b.length), 0, ...ascii(s2a)]);
    const cont = rec(0x003c, [1, ...utf16(s2b), ...le16(5), 1, ...utf16("İzmir")]);
    const globals1 = [...bof(5)];
    // BOUNDSHEET: lbPlyPos, durum, tür, ad (cch=2, bayrak=0, "S1"); toplam 4 + 10 bayt
    const sheetStart = globals1.length + 14 + sst.length + cont.length + 4;
    const bound = rec(0x0085, [...le32(sheetStart), 0, 0, 2, 0, ...ascii("S1")]);
    const eof = rec(0x000a, []);
    const lbl = (r: number, i: number) => rec(0x00fd, [...le16(r), ...le16(0), ...le16(15), ...le32(i)]);
    const stream = [...globals1, ...bound, ...sst, ...cont, ...eof, ...bof(0x10), ...lbl(0, 0), ...lbl(1, 1), ...lbl(2, 2), ...eof];
    const wb = readBiffStream(new Uint8Array(stream));
    const s = wb.sheets[0];
    expect(s.name).toBe("S1");
    expect(wb.getValue(s, 0, 0)).toBe("Merhaba");
    expect(wb.getValue(s, 1, 0)).toBe("Uzun metin çğıöşü");
    expect(wb.getValue(s, 2, 0)).toBe("İzmir");
  });
  it("RK sayıları", () => {
    expect(rk((150 << 2) | 2)).toBe(150);
    expect(rk((12345 << 2) | 3)).toBe(123.45);
    // 1.0 IEEE: üst 32 bit 0x3FF00000
    expect(rk(0x3ff00000)).toBe(1);
    expect(rk(((-7 << 2) | 2) >>> 0)).toBe(-7);
  });
});

describe("BIFF5 (95 biçimi) temel destek", () => {
  it("sayfa adları, kod sayfalı metinler, sayılar ve formül önbellek değerleri", () => {
    const bof = (dt: number) => rec(0x0809, [...le16(0x0500), ...le16(dt), 0, 0, 0, 0]);
    // Windows-1254: "Şehir" = DE 65 68 69 72, "Muğla" = 4D 75 F0 6C 61
    const name = [0x56, 0x65, 0x72, 0x69]; // "Veri"
    const g1 = [...bof(5), ...rec(0x0042, le16(1254))];
    const boundLen = 4 + 4 + 2 + 1 + name.length;
    const eof = rec(0x000a, []);
    const sheetStart = g1.length + boundLen + eof.length;
    const bound = rec(0x0085, [...le32(sheetStart), 0, 0, name.length, ...name]);
    const label = (r: number, b: number[]) => rec(0x0204, [...le16(r), ...le16(0), ...le16(15), ...le16(b.length), ...b]);
    const num = rec(0x0203, [...le16(2), ...le16(0), ...le16(15), 0, 0, 0, 0, 0, 0, 0x45, 0x40]);
    // FORMULA: önbellek 7 (BIFF5 belirteçleri çevrilmez, değer kalır)
    const formula = rec(0x0006, [...le16(3), ...le16(0), ...le16(15), 0, 0, 0, 0, 0, 0, 0x1c, 0x40, 0, 0, 0, 0, 0, 0, ...le16(3), 0x1e, 7, 0]);
    const stream = [...g1, ...bound, ...eof, ...bof(0x10), ...label(0, [0xde, 0x65, 0x68, 0x69, 0x72]), ...label(1, [0x4d, 0x75, 0xf0, 0x6c, 0x61]), ...num, ...formula, ...eof];
    const wb = readBiffStream(new Uint8Array(stream));
    const s = wb.sheets[0];
    expect(s.name).toBe("Veri");
    expect(wb.getValue(s, 0, 0)).toBe("Şehir");
    expect(wb.getValue(s, 1, 0)).toBe("Muğla");
    expect(wb.getValue(s, 2, 0)).toBe(42);
    expect(wb.getValue(s, 3, 0)).toBe(7);
  });
});

describe("ptg → A1 formül çevirisi", () => {
  const ctx: XlsFormulaContext = { sheetNames: ["Veri", "Özet Sayfa"], xti: [{ sup: 0, first: 0, last: 0 }, { sup: 0, first: 1, last: 1 }, { sup: 1, first: 0, last: 0 }], supbooks: [{ kind: "self", externNames: [] }, { kind: "external", externNames: [] }], names: [{ name: "Oran", builtin: false }] };
  const f = (bytes: number[], base = { row: 0, col: 0 }) => decodeFormula(new Uint8Array(bytes), new Uint8Array(0), ctx, base);
  const int = (n: number) => [0x1e, ...le16(n)];
  const ref = (r: number, c: number, rel = true) => [0x44, ...le16(r), ...le16(c | (rel ? 0xc000 : 0))];
  it("işleçler ve öncelik", () => {
    expect(f([...int(1), ...int(2), 0x03, ...int(3), 0x05])).toBe("=(1+2)*3");
    expect(f([...int(1), ...int(2), ...int(3), 0x05, 0x03])).toBe("=1+2*3");
    expect(f([...int(2), 0x13, ...int(2), 0x07])).toBe("=-2^2");
    expect(f([...int(1), ...int(2), 0x03, 0x15, ...int(3), 0x05])).toBe("=(1+2)*3");
    expect(f([...int(50), 0x14])).toBe("=50%");
    expect(f([...ref(0, 0), ...ref(0, 1, false), 0x08])).toBe("=A1&$B$1");
  });
  it("işlevler, metin, ad, 3B başvuru", () => {
    expect(f([...ref(0, 0), ...ref(0, 1), 0x42, 2, ...le16(4)])).toBe("=SUM(A1,B1)");
    expect(f([0x25, ...le16(0), ...le16(2), ...le16(0xc000), ...le16(0xc000), 0x19, 0x10, 0, 0])).toBe("=SUM(A1:A3)");
    expect(f([0x1f, 0, 0, 0, 0, 0, 0, 0xf4, 0x3f, ...int(1), 0x21, ...le16(27)])).toBe("=ROUND(1.25,1)");
    expect(f([0x17, 3, 0, ...ascii("a\"c")])).toBe('="a""c"');
    expect(f([0x23, ...le32(1), ...int(2), 0x05])).toBe("=Oran*2");
    expect(f([0x3a, ...le16(1), ...le16(4), ...le16(0x8002)])).toBe("='Özet Sayfa'!$C5");
    expect(f([0x3b, ...le16(0), ...le16(0), ...le16(0xffff), ...le16(0), ...le16(0)])).toBe("=Veri!$A:$A");
    // Göreli paylaşılan başvuru (ptgRefN): -1 sütun, aynı satır
    expect(f([0x4c, ...le16(0), ...le16(0xc0ff), ...int(2), 0x05], { row: 9, col: 2 })).toBe("=B10*2");
  });
  it("çevrilemeyenler undefined döner", () => {
    expect(f([0x3a, ...le16(2), ...le16(0), ...le16(0)])).toBeUndefined();
    expect(f([0x01, 0, 0, 0, 0])).toBeUndefined();
    expect(f([0x42, 1, ...le16(0x8001)])).toBeUndefined();
    expect(f([0x03])).toBeUndefined();
  });
});
