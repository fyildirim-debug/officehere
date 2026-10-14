/** Metin biçimleri (CSV/TSV), kodlama ve ayırıcı tanıma, biçim imzası ve şablonlar. */
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import JSZip from "jszip";
import { decodeText, detectDelimiter, readCsv, writeCsv } from "../src/io/csv";
import { openBook, saveFormatOf, sniffFormat, writableAs, writeBook } from "../src/io/formats";
import { WorkbookModel } from "../src/model/workbook";

const fixture = (name: string) => {
  const b = readFileSync(new URL(`../src/io/__fixtures__/${name}`, import.meta.url));
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer;
};
const bytes = (s: string) => new TextEncoder().encode(s);

describe("CSV/TSV: ayırıcı ve kodlama tanıma", () => {
  it("virgül, noktalı virgül ve sekme", () => {
    expect(detectDelimiter("a,b,c\n1,2,3\n")).toBe(",");
    expect(detectDelimiter("a;b;c\n1;2;3\n")).toBe(";");
    expect(detectDelimiter("a\tb\tc\n1\t2\t3\n")).toBe("\t");
    expect(detectDelimiter('"x;y",b\n"1;2",3\n')).toBe(",");
  });
  it("UTF-8 BOM, UTF-16 BOM, BOM'suz UTF-16 ve Windows-1254", () => {
    expect(decodeText(new Uint8Array([0xef, 0xbb, 0xbf, ...bytes("Çay;Şeker")]))).toBe("Çay;Şeker");
    const u16le = (s: string) => [...s].flatMap((c) => [c.charCodeAt(0) & 0xff, c.charCodeAt(0) >> 8]);
    expect(decodeText(new Uint8Array([0xff, 0xfe, ...u16le("İzmir;1")]))).toBe("İzmir;1");
    expect(decodeText(new Uint8Array(u16le("Ad;Tutar\nMuğla;5")))).toBe("Ad;Tutar\nMuğla;5");
    // Windows-1254: ğ=0xF0, ş=0xFE, İ=0xDD, ı=0xFD
    expect(decodeText(new Uint8Array([0x44, 0xf0, 0x3b, 0xfe, 0x3b, 0xdd, 0x3b, 0xfd]))).toBe("Dğ;ş;İ;ı");
  });
  it("TSV okuma/yazma ve sayı yerel ayarı tahmini", () => {
    const wb = readCsv(bytes("Ad\tTutar\tOran\nÇay\t1234.5\t0.25\nŞeker\t12.75\t1\n"), "veri", { delimiter: "\t" });
    const s = wb.sheets[0];
    expect(s.name).toBe("veri");
    expect(wb.getValue(s, 1, 1)).toBe(1234.5);
    expect(wb.getValue(s, 2, 1)).toBe(12.75);
    const tr = readCsv(bytes("Ad\tTutar\nÇay\t1.234,5\n"), "x", { delimiter: "\t" });
    expect(tr.getValue(tr.sheets[0], 1, 1)).toBe(1234.5);
    const out = writeCsv(wb, s, { delimiter: "\t", bom: false });
    expect(out.split("\r\n")[0]).toBe("Ad\tTutar\tOran");
    const back = readCsv(bytes(out), "x", { delimiter: "\t" });
    expect(back.getValue(back.sheets[0], 1, 0)).toBe("Çay");
  });
  it("CSV açılırken ayırıcı korunur; yalnız etkin sayfa yazılır", async () => {
    const o = await openBook("satis.csv", bytes("Ad,Tutar\nÇay,12.5\n").buffer as ArrayBuffer);
    expect(o.format).toBe("csv");
    expect(o.wb.getValue(o.wb.sheets[0], 1, 1)).toBe(12.5);
    const wb = new WorkbookModel();
    wb.setInputRaw(wb.sheets[0], 0, 0, "bir");
    const b = wb.addSheet("İki");
    wb.setInputRaw(b, 0, 0, "iki");
    const csv = (await writeBook(wb, wb.sheets[0], "csv", ",")) as string;
    expect(csv).toContain("bir");
    expect(csv).not.toContain("iki");
  });
});

describe("Biçim tanıma (içerik imzası)", () => {
  it("OOXML türleri, ODF, CFB ve metin", async () => {
    expect(await sniffFormat("a.xlsx", fixture("sablon.xltx"))).toBe("xltx");
    expect(await sniffFormat("a.xlsx", fixture("sablon.xlsm"))).toBe("xlsm");
    expect(await sniffFormat("a.bin", fixture("sablon.xltm"))).toBe("xltm");
    expect(await sniffFormat("a.xlsx", fixture("ornek.ods"))).toBe("ods");
    expect(await sniffFormat("a.ods", fixture("ornek.xls"))).toBe("xls");
    expect(await sniffFormat("a.csv", bytes("a;b").buffer as ArrayBuffer)).toBe("csv");
    expect(await sniffFormat("a.tsv", bytes("a\tb").buffer as ArrayBuffer)).toBe("tsv");
    expect(await sniffFormat("a.txt", bytes("<!DOCTYPE html><table></table>").buffer as ArrayBuffer)).toBe("html");
  });
  it("başka uygulamanın dosyası ve bozuk ZIP anlaşılır hata verir", async () => {
    const z = new JSZip();
    z.file("[Content_Types].xml", '<Types><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
    await expect(sniffFormat("x.xlsx", (await z.generateAsync({ type: "arraybuffer" })) as ArrayBuffer)).rejects.toThrow(/Belge/);
    const bad = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 1, 2, 3, 4, 5, 6]);
    await expect(sniffFormat("x.xlsx", bad.buffer)).rejects.toThrow(/bozuk/);
  });
  it("şablon ve makrolu kitaplar: makro korunur ama çalıştırılmaz, şablon adsız açılır", async () => {
    const m = await openBook("sablon.xlsm", fixture("sablon.xlsm"));
    expect(m.template).toBe(false);
    expect(m.notice).toMatch(/makrolar kaydederken korunur, ancak OfficeHere.da çalıştırılmaz/);
    expect(m.wb.getValue(m.wb.sheets[0], 1, 1)).toBe(30);
    const t = await openBook("sablon.xltx", fixture("sablon.xltx"));
    expect(t.template).toBe(true);
    expect(t.wb.getFormula(t.wb.sheets[0], 1, 1)).toBe("=A2*10");
    const tm = await openBook("sablon.xltm", fixture("sablon.xltm"));
    expect(tm.template).toBe(true);
    expect(tm.notice).toMatch(/makro/);
  });
  it("kaydetme biçimi kararları", () => {
    expect(writableAs("xls")).toBeNull();
    // Makrolu kitap kendi biçiminde kaydedilir (makrolar korunur, çalıştırılmaz)
    expect(writableAs("xlsm")).toBe("xlsm");
    expect(saveFormatOf("a.XLSM")).toBe("xlsm");
    expect(writableAs("ods")).toBe("ods");
    expect(writableAs("txt")).toBe("tsv");
    expect(saveFormatOf("a.ODS")).toBe("ods");
    expect(saveFormatOf("a.xls")).toBeNull();
  });
});
