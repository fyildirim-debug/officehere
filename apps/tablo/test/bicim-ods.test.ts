/** OpenDocument Hesap Tablosu (.ods/.ots) okuma, yazma ve gidiş-dönüş. */
import { describe, expect, it } from "vitest";
import { existsSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { execFileSync } from "node:child_process";
import JSZip from "jszip";
import { cellKey, dateToSerial } from "@officehere/formula";
import { readOds } from "../src/io/ods/reader";
import { writeOds } from "../src/io/ods/writer";
import { a1FormulaToOdf, odfFormulaToA1 } from "../src/io/ods/formula";
import { codeToDataStyles, dataStyleToCode } from "../src/io/ods/numfmt";
import { child, kids, parseXml } from "../src/io/ods/xml";
import { openBook, sniffFormat } from "../src/io/formats";
import { readXlsx } from "../src/io/xlsx";
import { displayText } from "../src/model/display";
import { WorkbookModel } from "../src/model/workbook";

const fixture = (name: string) => {
  const b = readFileSync(new URL(`../src/io/__fixtures__/${name}`, import.meta.url));
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer;
};

describe("OpenFormula çevirisi", () => {
  it("başvurular, ayırıcılar ve işlev adları", () => {
    expect(odfFormulaToA1("of:=SUM([.A1:.B2])")).toBe("=SUM(A1:B2)");
    expect(odfFormulaToA1("of:=[.$B$2]+[$Sayfa2.C3]")).toBe("=$B$2+Sayfa2!C3");
    expect(odfFormulaToA1("of:=[$'Yeni Sayfa'.A1:.B3]")).toBe("='Yeni Sayfa'!A1:B3");
    expect(odfFormulaToA1('of:=IF([.D5]>100;"Büyük";"Küçük")')).toBe('=IF(D5>100,"Büyük","Küçük")');
    expect(odfFormulaToA1("of:=SUM({1;2|3;4})")).toBe("=SUM({1,2;3,4})");
    expect(odfFormulaToA1("of:=SUM([.A:.A])")).toBe("=SUM(A:A)");
    expect(odfFormulaToA1("of:=SUM([.1:.3])")).toBe("=SUM(1:3)");
    expect(odfFormulaToA1("of:=COM.MICROSOFT.IFS([.A1]>1;\"a\";TRUE();\"b\")")).toBe('=IFS(A1>1,"a",TRUE(),"b")');
    expect(odfFormulaToA1("of:=LEGACY.NORMSDIST(1)")).toBe("=NORMSDIST(1)");
    expect(odfFormulaToA1("of:=SUM(([.A1]~[.B2]))")).toBe("=SUM((A1,B2))");
    expect(odfFormulaToA1("of:=VLOOKUP(\"x\";[.A2:.C4];3;FALSE())")).toBe('=VLOOKUP("x",A2:C4,3,FALSE())');
    expect(odfFormulaToA1('of:="a;b"&"[.A1]"')).toBe('="a;b"&"[.A1]"');
    expect(odfFormulaToA1("msoxl:=SUM(A1:A3)")).toBe("=SUM(A1:A3)");
  });
  it("çevrilemeyenler: bilinmeyen işlev, 3B aralık, dış dosya", () => {
    expect(odfFormulaToA1("of:=ORG.OPENOFFICE.WEEKS([.A1];[.A2];0)")).toBeUndefined();
    expect(odfFormulaToA1("of:=SUM([$Diger.A1:$Diger2.A1])")).toBeUndefined();
    expect(odfFormulaToA1("of:=['file:///x.ods'#$Sayfa1.A1]")).toBeUndefined();
    expect(odfFormulaToA1("of:=[.A1")).toBeUndefined();
  });
  it("A1 → OpenFormula", () => {
    expect(a1FormulaToOdf("=SUM(A1:B2)*2")).toBe("of:=SUM([.A1:.B2])*2");
    expect(a1FormulaToOdf("='Yeni Sayfa'!$A$1+Özet!B1")).toBe("of:=[$'Yeni Sayfa'.$A$1]+[$Özet.B1]");
    expect(a1FormulaToOdf('=IF(A1>0,"evet, doğru",FALSE)')).toBe('of:=IF([.A1]>0;"evet, doğru";FALSE())');
    expect(a1FormulaToOdf("=SUM({1,2;3,4})")).toBe("of:=SUM({1;2|3;4})");
    expect(a1FormulaToOdf("=IFS(A1>1,1,TRUE,2)")).toBe("of:=COM.MICROSOFT.IFS([.A1]>1;1;TRUE();2)");
    expect(a1FormulaToOdf("=SUM(A:A)")).toBe("of:=SUM([.A:.A])");
    expect(a1FormulaToOdf("=SUM(Tablo1[Tutar])")).toBeUndefined();
  });
});

describe("Sayı biçimi ⇄ veri stili", () => {
  const roundTrip = (code: string) => {
    const xml = `<r xmlns:number="n" xmlns:style="s">${codeToDataStyles("N1", code)}</r>`;
    const map = new Map(kids(parseXml(xml).children[0] as never).map((e) => [e.attrs["style:name"], e]));
    return dataStyleToCode("N1", map);
  };
  it("yaygın biçimler", () => {
    expect(roundTrip("#,##0.00")).toBe("#,##0.00");
    expect(roundTrip("0%")).toBe("0%");
    expect(roundTrip("0.00%")).toBe("0.00%");
    expect(roundTrip("dd.mm.yyyy")).toBe("dd.mm.yyyy");
    expect(roundTrip("hh:mm:ss")).toBe("hh:mm:ss");
    expect(roundTrip("0.00E+00")).toBe("0.00E+00");
    expect(roundTrip("@")).toBe("@");
    expect(roundTrip("#,##0.00;[Red]-#,##0.00")).toBe("#,##0.00;[Red]-#,##0.00");
    expect(roundTrip('#,##0.00 "₺"')).toBe('#,##0.00" ₺"');
    expect(roundTrip("d mmmm yyyy dddd")).toBe("d mmmm yyyy dddd");
  });
});

describe("ODS okuma (LibreOffice örneği)", () => {
  it("sayfalar, değerler, formüller, biçimler", async () => {
    const wb = await readOds(fixture("ornek.ods"));
    expect(wb.sheets.map((s) => s.name)).toEqual(["Satışlar", "Özet", "Uzun", "Gizli"]);
    expect(wb.sheets[3].hidden).toBe(true);
    const [s, o, u] = wb.sheets;
    expect(wb.getValue(s, 0, 0)).toBe("Ürün");
    expect(wb.getValue(s, 3, 6)).toBe("ığüşöç ĞÜŞİÖÇ");
    expect(wb.getFormula(s, 1, 3)).toBe("=B2*C2");
    expect(wb.getValue(s, 4, 3)).toBe(234.25);
    expect(wb.getFormula(s, 4, 3)).toBe("=SUM(D2:D4)");
    expect(wb.getFormula(s, 5, 2)).toBe('=IF(D5>100,"Büyük","Küçük")');
    expect(wb.getValue(s, 5, 2)).toBe("Büyük");
    expect(wb.getValue(s, 6, 4)).toBe(-1);
    expect(wb.getFormula(o, 0, 1)).toBe("=Satışlar!D5");
    expect(wb.getValue(s, 1, 4)).toBe(dateToSerial(2024, 3, 15));
    expect(displayText(wb, s, 1, 4)).toBe("15.03.2024");
    expect(displayText(wb, s, 1, 5)).toBe("25,00%");
    expect(displayText(wb, s, 1, 2)).toBe("2,50");
    expect(displayText(wb, s, 7, 0)).toBe("10:30");
    expect(wb.styleAt(s, 0, 0)).toMatchObject({ bold: true, fill: "#1F4E79", color: "#FFFFFF", hAlign: "center" });
    expect(wb.styleAt(s, 4, 0)).toMatchObject({ italic: true, underline: "single", size: 14, color: "#C00000" });
    expect(wb.styleAt(s, 4, 3).bBottom?.style).toBe("double");
    expect(s.merges).toEqual([{ r1: 9, c1: 0, r2: 10, c2: 2 }]);
    expect(s.rows.size(9)).toBe(40);
    expect(s.cols.size(0)).toBeGreaterThan(120);
    expect(wb.engine.getNames().map((n) => n.name)).toContain("Toplam_Adet");
    expect(wb.getValue(u, 299, 0)).toBe("Satır 300 — çğıöşü ÇĞİÖŞÜ uzun metin 2100");
    expect(wb.getValue(u, 359, 2)).toBe(1080);
  });
  it("şablon (.ots) adsız kitap olarak açılır", async () => {
    expect(await sniffFormat("x.bin", fixture("sablon.ots"))).toBe("ots");
    const o = await openBook("sablon.ots", fixture("sablon.ots"));
    expect(o.template).toBe(true);
    expect(o.wb.getValue(o.wb.sheets[0], 0, 0)).toBe("Fatura");
    expect(o.wb.getValue(o.wb.sheets[0], 1, 1)).toBe(30);
  });
  it("1904 tarih sistemli dosyada tarihler doğru", async () => {
    const wb = await readOds(fixture("tarih1904.ods"));
    expect(wb.getValue(wb.sheets[0], 0, 0)).toBe(dateToSerial(2024, 3, 15));
    expect(wb.getValue(wb.sheets[0], 2, 0)).toBe(dateToSerial(2024, 3, 20));
  });
});

/** Elle content.xml ile ODS paketi. */
async function odsWith(body: string, extra = ""): Promise<Uint8Array> {
  const z = new JSZip();
  z.file("mimetype", "application/vnd.oasis.opendocument.spreadsheet", { compression: "STORE" });
  z.file(
    "content.xml",
    `<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">${extra}<office:body><office:spreadsheet>${body}</office:spreadsheet></office:body></office:document-content>`,
  );
  return z.generateAsync({ type: "uint8array" });
}

describe("ODS sınırları ve bozuk dosyalar", () => {
  it("dev yinelemeler (1M satır, 16384 sütun) belleği şişirmez", async () => {
    const styles = `<office:automatic-styles><style:style style:name="ce1" style:family="table-cell"><style:table-cell-properties fo:background-color="#ffff00"/></style:style></office:automatic-styles>`;
    const body =
      `<table:table table:name="Büyük"><table:table-column table:number-columns-repeated="16384"/>` +
      `<table:table-row table:number-rows-repeated="3"><table:table-cell office:value-type="float" office:value="7"><text:p>7</text:p></table:table-cell><table:table-cell table:number-columns-repeated="16383"/></table:table-row>` +
      `<table:table-row><table:table-cell table:style-name="ce1" table:number-columns-repeated="5"/><table:table-cell table:number-columns-repeated="16379"/></table:table-row>` +
      `<table:table-row table:number-rows-repeated="1048570"><table:table-cell table:style-name="ce1" table:number-columns-repeated="16384"/></table:table-row>` +
      `<table:table-row table:number-rows-repeated="10"><table:table-cell table:number-columns-repeated="16384"/></table:table-row></table:table>`;
    const t = Date.now();
    const wb = await readOds(await odsWith(body, styles));
    expect(Date.now() - t).toBeLessThan(5000);
    const s = wb.sheets[0];
    expect([0, 1, 2].map((r) => wb.getValue(s, r, 0))).toEqual([7, 7, 7]);
    expect(wb.getValue(s, 3, 0)).toBeNull();
    expect(wb.styleAt(s, 3, 4).fill).toBe("#FFFF00");
    // Yalnız stilli dev kuyruk açılmaz
    expect(s.styles.size).toBeLessThan(10);
    expect(wb.dataExtent(s).rows).toBe(3);
  });
  it("bozuk ve şifreli paketler anlaşılır hatayla reddedilir", async () => {
    await expect(readOds(new Uint8Array([1, 2, 3, 4, 5]))).rejects.toThrow(/OpenDocument/);
    const z = new JSZip();
    z.file("mimetype", "application/vnd.oasis.opendocument.spreadsheet");
    z.file("META-INF/manifest.xml", '<manifest:manifest><manifest:file-entry manifest:full-path="content.xml"><manifest:encryption-data/></manifest:file-entry></manifest:manifest>');
    z.file("content.xml", "xx");
    await expect(readOds(await z.generateAsync({ type: "uint8array" }))).rejects.toThrow(/parola/);
    await expect(readOds(await odsWith("<table:table table:name='a'><table:table-row>"))).rejects.toThrow();
  });
});

function sample(): WorkbookModel {
  const wb = new WorkbookModel();
  const s = wb.sheets[0];
  wb.renameSheet(0, "Satış Özeti");
  wb.setInputRaw(s, 0, 0, "Ürün");
  wb.setInputRaw(s, 0, 1, "Tutar");
  wb.setInputRaw(s, 1, 0, "Çay  ve  şeker");
  wb.setInputRaw(s, 1, 1, 1234.5);
  wb.setInputRaw(s, 2, 1, 0.25);
  wb.setInputRaw(s, 3, 1, { formula: "=SUM(B2:B3)" });
  wb.setInputRaw(s, 4, 1, { formula: '=IF(B4>100,"büyük","küçük")' });
  wb.setInputRaw(s, 5, 1, dateToSerial(2024, 3, 15));
  wb.setInputRaw(s, 6, 1, true);
  wb.setInputRaw(s, 7, 0, "satır1\nsatır2");
  s.styles.set(cellKey(0, 0), wb.styles.intern({ bold: true, fill: "#1F4E79", color: "#FFFFFF", hAlign: "center", bBottom: { style: "thin", color: "#000000" } }));
  s.styles.set(cellKey(1, 1), wb.styles.intern({ numFmt: "#,##0.00", italic: true }));
  s.styles.set(cellKey(2, 1), wb.styles.intern({ numFmt: "0%" }));
  s.styles.set(cellKey(5, 1), wb.styles.intern({ numFmt: "dd.mm.yyyy" }));
  s.merges.push({ r1: 9, c1: 0, r2: 10, c2: 2 });
  wb.setInputRaw(s, 9, 0, "Birleşik");
  s.cols.set(0, 150);
  s.rows.set(9, 40, true);
  s.notes.set(cellKey(0, 1), { text: "Not metni", author: "Furkan" });
  s.links.set(cellKey(0, 0), "https://ornek.com/");
  const g = wb.addSheet("Gizli");
  g.hidden = true;
  wb.setInputRaw(g, 0, 0, { formula: "='Satış Özeti'!B4*2" });
  wb.engine.defineName("Tutarlar", "='Satış Özeti'!$B$2:$B$3");
  wb.engine.recalc();
  return wb;
}

describe("ODS yazma", () => {
  it("geçerli paket: mimetype ilk ve sıkıştırılmamış, manifest var", async () => {
    const bytes = await writeOds(sample());
    // İlk yerel dosya başlığı: ad "mimetype", sıkıştırma yöntemi 0
    expect(String.fromCharCode(...bytes.subarray(30, 38))).toBe("mimetype");
    expect(bytes[8] | (bytes[9] << 8)).toBe(0);
    expect(String.fromCharCode(...bytes.subarray(38, 38 + 46))).toBe("application/vnd.oasis.opendocument.spreadsheet");
    const z = await JSZip.loadAsync(bytes);
    for (const p of ["META-INF/manifest.xml", "content.xml", "styles.xml", "meta.xml", "settings.xml"]) expect(z.file(p)).toBeTruthy();
    const content = await z.file("content.xml")!.async("string");
    expect(content).toContain('table:formula="of:=SUM([.B2:.B3])"');
    const root = parseXml(content);
    expect(child(root, "office:document-content")).toBeTruthy();
  });
  it("gidiş-dönüş: değerler, formüller, stiller, birleştirmeler, boyutlar, notlar, adlar", async () => {
    const wb = await readOds(await writeOds(sample()));
    const [s, g] = wb.sheets;
    expect(wb.sheets.map((x) => x.name)).toEqual(["Satış Özeti", "Gizli"]);
    expect(g.hidden).toBe(true);
    expect(wb.getValue(s, 1, 0)).toBe("Çay  ve  şeker");
    expect(wb.getValue(s, 7, 0)).toBe("satır1\nsatır2");
    expect(wb.getValue(s, 1, 1)).toBe(1234.5);
    expect(wb.getFormula(s, 3, 1)).toBe("=SUM(B2:B3)");
    expect(wb.getValue(s, 4, 1)).toBe("büyük");
    expect(wb.getValue(s, 5, 1)).toBe(dateToSerial(2024, 3, 15));
    expect(wb.getValue(s, 6, 1)).toBe(true);
    expect(wb.getFormula(g, 0, 0)).toBe("='Satış Özeti'!B4*2");
    expect(wb.getValue(g, 0, 0)).toBe(2469.5);
    expect(wb.styleAt(s, 0, 0)).toMatchObject({ bold: true, fill: "#1F4E79", color: "#FFFFFF", hAlign: "center", bBottom: { style: "thin" } });
    expect(wb.styleAt(s, 1, 1)).toMatchObject({ numFmt: "#,##0.00", italic: true });
    expect(displayText(wb, s, 2, 1)).toBe("25%");
    expect(displayText(wb, s, 5, 1)).toBe("15.03.2024");
    expect(s.merges).toEqual([{ r1: 9, c1: 0, r2: 10, c2: 2 }]);
    expect(s.cols.size(0)).toBe(150);
    expect(s.rows.size(9)).toBe(40);
    expect(s.notes.get(cellKey(0, 1))).toEqual({ text: "Not metni", author: "Furkan" });
    expect(s.links.get(cellKey(0, 0))).toBe("https://ornek.com/");
    expect(wb.engine.getNames()).toContainEqual(expect.objectContaining({ name: "Tutarlar", formula: "='Satış Özeti'!$B$2:$B$3" }));
  });
  it("LibreOffice örneği okunup yeniden yazılınca bilgiler korunur", async () => {
    const a = await readOds(fixture("ornek.ods"));
    const b = await readOds(await writeOds(a));
    for (let k = 0; k < a.sheets.length; k++) {
      const sa = a.sheets[k];
      const sb = b.sheets[k];
      expect(sb.name).toBe(sa.name);
      for (let r = 0; r < 12; r++)
        for (let c = 0; c < 7; c++) {
          expect(b.getFormula(sb, r, c)).toBe(a.getFormula(sa, r, c));
          expect(displayText(b, sb, r, c)).toBe(displayText(a, sa, r, c));
        }
    }
  });
});

const soffice = ["/usr/bin/soffice", "/usr/local/bin/soffice"].find((p) => existsSync(p));

describe.skipIf(!soffice)("LibreOffice doğrulaması", () => {
  it("yazılan .ods LibreOffice ile .xlsx'e çevrilir; değer ve formüller doğru", async () => {
    const dir = mkdtempSync(join(tmpdir(), "oh-ods-"));
    const file = join(dir, "yazilan.ods");
    writeFileSync(file, await writeOds(sample()));
    execFileSync("timeout", ["120", soffice!, "--headless", `-env:UserInstallation=file://${dir}/profil`, "--convert-to", "xlsx", file, "--outdir", dir], { stdio: "ignore" });
    const x = readFileSync(join(dir, "yazilan.xlsx"));
    const wb = await readXlsx(x.buffer.slice(x.byteOffset, x.byteOffset + x.byteLength) as ArrayBuffer);
    const s = wb.sheets[0];
    expect(wb.sheets.map((v) => v.name)).toEqual(["Satış Özeti", "Gizli"]);
    expect(wb.getFormula(s, 3, 1)).toBe("=SUM(B2:B3)");
    expect(wb.getValue(s, 3, 1)).toBe(1234.75);
    expect(wb.getValue(s, 1, 0)).toBe("Çay  ve  şeker");
    expect(wb.getValue(s, 5, 1)).toBe(dateToSerial(2024, 3, 15));
    expect(wb.styleAt(s, 0, 0)).toMatchObject({ bold: true, fill: "#1F4E79" });
    expect(s.merges).toEqual([{ r1: 9, c1: 0, r2: 10, c2: 2 }]);
    expect(wb.getFormula(wb.sheets[1], 0, 0)).toBe("='Satış Özeti'!B4*2");
  }, 120_000);
});
