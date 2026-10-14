// XLSX birleştirmeleri exceljs'i atlar (Faz 0): kendi taramamızla okunur, mergeCells'e doğrudan yazılır.
// Sınır durumları (tek hücrelik, sayfa kenarı, çakışan/bozuk aralık) ve 50.000 birleştirmeli dosya.
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { readXlsx, writeXlsx } from "../xlsx";
import { prepareXlsx, repairMerges } from "../xlsxGuard";
import { WorkbookModel } from "../../model/workbook";
import type { Rect } from "../../model/types";

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
/** Stil 1: kalın + dört kenarda ince kenarlık. */
const STYLES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
  `<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>` +
  `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
  `<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom><diagonal/></border></borders>` +
  `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
  `<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/></cellXfs></styleSheet>`;

const ab = (u: Uint8Array): ArrayBuffer => u.buffer.slice(u.byteOffset, u.byteOffset + u.byteLength) as ArrayBuffer;
const R = (r1: number, c1: number, r2: number, c2: number): Rect => ({ r1, c1, r2, c2 });
const key = (m: Rect) => `${m.r1},${m.c1},${m.r2},${m.c2}`;
const sameSet = (a: Rect[], b: Rect[]) => {
  const x = new Set(a.map(key));
  return x.size === b.length && b.every((m) => x.has(key(m)));
};
/** Süre sınırı: O(n²) birleştirme denetimi dakikalar sürerdi. */
const quick = async <T>(fn: () => Promise<T>, ms = 20_000): Promise<T> => {
  const t = performance.now();
  const r = await fn();
  expect(performance.now() - t).toBeLessThan(ms);
  return r;
};

/** Tek sayfalık ("Veri") en küçük XLSX paketi. */
async function paket(sheetData: string, mergeCells: string): Promise<ArrayBuffer> {
  const z = new JSZip();
  z.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
  );
  z.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG_REL}"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  z.file("xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook ${NS}><sheets><sheet name="Veri" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  z.file(
    "xl/_rels/workbook.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG_REL}"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/></Relationships>`,
  );
  z.file("xl/styles.xml", STYLES);
  z.file("xl/worksheets/sheet1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet ${NS}><sheetData>${sheetData}</sheetData>${mergeCells}</worksheet>`);
  return ab(await z.generateAsync({ type: "uint8array", compression: "DEFLATE" }));
}

const block = (refs: string[]) => `<mergeCells count="${refs.length}">${refs.map((r) => `<mergeCell ref="${r}"/>`).join("")}</mergeCells>`;
const str = (ref: string, s: string, style = 0) => `<c r="${ref}"${style ? ` s="${style}"` : ""} t="inlineStr"><is><t>${s}</t></is></c>`;
const sheetXml = async (bytes: Uint8Array) => (await JSZip.loadAsync(bytes)).file("xl/worksheets/sheet1.xml")!.async("string");
const writtenMerges = (xml: string) => [...xml.matchAll(/<mergeCell ref="([^"]+)"\/>/g)].map((m) => m[1]);
const MERGE_WARN = "geçersiz ya da çakışan birleştirmeler";

describe("XLSX birleştirmeleri (exceljs'siz)", { timeout: 120_000 }, () => {
  it("exceljs'e verilen kopyada mergeCells yok; birleştirmeler kendi taramamızla gelir", async () => {
    const data = await paket(`<row r="1">${str("A1", "Başlık")}</row>`, block(["A1:C1", "B3:B6"]));
    const pkg = await prepareXlsx(data);
    expect(pkg.merges.get("Veri")).toEqual([R(0, 0, 0, 2), R(2, 1, 5, 1)]);
    const excel = await (await JSZip.loadAsync(pkg.excel)).file("xl/worksheets/sheet1.xml")!.async("string");
    expect(excel).not.toContain("mergeCell");
    expect(excel).toContain("Başlık");
    // Kendi okuyucularımızın paketi (onarılmış özgün) birleştirmeleri korur
    const own = await (await JSZip.loadAsync(pkg.data)).file("xl/worksheets/sheet1.xml")!.async("string");
    expect(own).toContain('<mergeCell ref="A1:C1"/>');
  });

  it("repairMerges: geçerli liste dosyadaki sırayla; sağlam blok olduğu gibi kalır", () => {
    const xml = `<worksheet><sheetData/>${block(["D4:E5", "A1:B1"])}<pageMargins/></worksheet>`;
    const r = repairMerges(xml)!;
    expect(r.xml).toBe(xml);
    expect(r.merges).toEqual([R(3, 3, 4, 4), R(0, 0, 0, 1)]);
    expect(r.dropped).toBe(0);
    expect(r.excel).toBe("<worksheet><sheetData/><pageMargins/></worksheet>");
    expect(repairMerges("<worksheet><sheetData/></worksheet>")).toBeNull();
    // Önekli ad alanı
    const p = repairMerges(`<x:worksheet><x:sheetData/><x:mergeCells count="1"><x:mergeCell ref="A1:A2"/></x:mergeCells></x:worksheet>`)!;
    expect(p.merges).toEqual([R(0, 0, 1, 0)]);
    expect(p.excel).toBe("<x:worksheet><x:sheetData/></x:worksheet>");
  });

  it("1 hücrelik birleştirme sessizce yok sayılır ve yazılmaz", async () => {
    const wb = await readXlsx(await paket(`<row r="1">${str("A1", "x")}</row>`, block(["A1:A1", "B2:C2", "D4"])));
    const s = wb.sheets[0];
    expect(s.merges).toEqual([R(1, 1, 1, 2)]);
    expect(wb.loadWarnings).not.toContain(MERGE_WARN);
    // Modelde tek hücrelik birleştirme olsa da dosyaya yazılmaz
    s.merges.push(R(5, 5, 5, 5));
    const xml = await sheetXml(await writeXlsx(wb));
    expect(writtenMerges(xml)).toEqual(["B2:C2"]);
    expect(xml).toContain('<mergeCells count="1">');
  });

  it("sayfa sınırındaki birleştirmeler gidiş-dönüşte korunur; sınırı aşan aralık uyarıyla atılır", async () => {
    const rows = `<row r="1">${str("XFC1", "sağ üst")}</row><row r="1048575">${str("A1048575", "sol alt")}${str("XFC1048575", "köşe")}</row>`;
    const data = await paket(rows, block(["A1048575:B1048576", "XFC1:XFD2", "XFC1048575:XFD1048576", "C1048576:C1048577", "XFD1:XFE1"]));
    const wb = await quick(() => readXlsx(data));
    const s = wb.sheets[0];
    const expected = [R(1048574, 0, 1048575, 1), R(0, 16382, 1, 16383), R(1048574, 16382, 1048575, 16383)];
    expect(s.merges).toEqual(expected);
    expect(wb.loadWarnings).toContain(MERGE_WARN);
    expect(wb.getValue(s, 1048574, 0)).toBe("sol alt");
    expect(wb.getValue(s, 0, 16382)).toBe("sağ üst");
    const out = await quick(() => writeXlsx(wb));
    expect(writtenMerges(await sheetXml(out))).toEqual(["A1048575:B1048576", "XFC1:XFD2", "XFC1048575:XFD1048576"]);
    const back = await quick(() => readXlsx(ab(out)));
    expect(sameSet(back.sheets[0].merges, expected)).toBe(true);
    expect(back.getValue(back.sheets[0], 1048574, 16382)).toBe("köşe");
  });

  it("çakışan ve bozuk aralıklar: takılmaz, geçerliler kalır, kayıp uyarılır", async () => {
    const refs = ["A1:B2", "B2:C3", "D1:E1", "", "A1:", "ZZZZ1:B2", "A0:B1", "C9:A7", ...Array.from({ length: 20_000 }, () => "H1:I2")];
    const bad = block(refs).replace('<mergeCell ref=""/>', "<mergeCell/>");
    const data = await paket(`<row r="1">${str("A1", "a")}</row>`, bad);
    const wb = await quick(() => readXlsx(data));
    const s = wb.sheets[0];
    expect(s.merges).toEqual([R(0, 0, 1, 1), R(0, 3, 0, 4), R(6, 0, 8, 2), R(0, 7, 1, 8)]);
    expect(wb.loadWarnings).toContain(MERGE_WARN);
    const back = await quick(async () => readXlsx(ab(await writeXlsx(wb))));
    expect(sameSet(back.sheets[0].merges, s.merges)).toBe(true);
    expect(back.loadWarnings).not.toContain(MERGE_WARN);
  });

  it("birden çok mergeCells bloğu toplanır; dev birleştirme küçüklere yenilir, değerler kaybolmaz", async () => {
    const rows = Array.from({ length: 6 }, (_, i) => `<row r="${i + 1}">${str(`A${i + 1}`, `a${i}`)}${str(`D${i + 1}`, `d${i}`)}</row>`).join("");
    const data = await paket(rows, block(["A1:XFD1048576"]) + block(["C2:E4", "B6:C6"]));
    const pkg = await prepareXlsx(data);
    const excel = await (await JSZip.loadAsync(pkg.excel)).file("xl/worksheets/sheet1.xml")!.async("string");
    expect(excel).not.toContain("mergeCell");
    const wb = await quick(() => readXlsx(data));
    const s = wb.sheets[0];
    expect(s.merges).toEqual([R(1, 2, 3, 4), R(5, 1, 5, 2)]);
    expect(wb.loadWarnings).toContain("çok büyük birleştirilmiş alanlar");
    const back = await readXlsx(ab(await writeXlsx(wb)));
    for (let r = 0; r < 6; r++) expect(back.getValue(back.sheets[0], r, 0)).toBe(`a${r}`);
    // D2:D4 C2:E4'ün gizli hücreleri: değerleri (exceljs'teki gibi) okunmaz
    expect(back.getValue(back.sheets[0], 0, 3)).toBe("d0");
    expect(back.getValue(back.sheets[0], 4, 3)).toBe("d4");
    expect(sameSet(back.sheets[0].merges, s.merges)).toBe(true);
  });

  it("kapanmamış mergeCells bloğu dosyanın tamamını düşürmez", async () => {
    const data = await paket(`<row r="1">${str("A1", "a")}</row>`, '<mergeCells count="2"><mergeCell ref="A1:B1"/><mergeCell ref="C3:D4"/>');
    const wb = await quick(() => readXlsx(data));
    expect(wb.getValue(wb.sheets[0], 0, 0)).toBe("a");
    expect(wb.sheets[0].merges).toEqual([R(0, 0, 0, 1), R(2, 2, 3, 3)]);
    expect(wb.loadWarnings).toContain(MERGE_WARN);
  });

  it("gizli hücrelerin değeri okunmaz (exceljs'teki gibi); dev birleştirmede korunur", async () => {
    const rows = `<row r="1">${str("A1", "ana")}${str("B1", "gizli")}</row><row r="5">${str("A5", "dev içinde")}</row>`;
    const wb = await readXlsx(await paket(rows, block(["A1:B1", "A3:A200002"])));
    const s = wb.sheets[0];
    expect(wb.getValue(s, 0, 0)).toBe("ana");
    expect(wb.getValue(s, 0, 1)).toBeNull();
    expect(wb.getValue(s, 4, 0)).toBe("dev içinde");
    expect(s.merges).toEqual([R(0, 0, 0, 1), R(2, 0, 200001, 0)]);
  });

  it("yazmada gizli hücreler ana hücrenin stilini alır ve değer taşımaz", async () => {
    const wb = new WorkbookModel(true);
    const s = wb.addSheet("Veri");
    const st = wb.styles.intern({ bold: true, bTop: { style: "thin", color: "#000000" }, bBottom: { style: "thin", color: "#000000" }, bLeft: { style: "thin", color: "#000000" }, bRight: { style: "thin", color: "#000000" } });
    wb.setInputRaw(s, 0, 0, "Başlık");
    wb.setStyleRaw(s, 0, 0, st);
    wb.setInputRaw(s, 0, 1, "gizli");
    s.merges.push(R(0, 0, 1, 2));
    const out = await writeXlsx(wb);
    const z = await JSZip.loadAsync(out);
    const xml = await z.file("xl/worksheets/sheet1.xml")!.async("string");
    const sOf = (ref: string) => new RegExp(`<c r="${ref}"[^>]*?\\ss="(\\d+)"`).exec(xml)?.[1];
    expect(sOf("A1")).toBeTruthy();
    for (const ref of ["B1", "C1", "A2", "B2", "C2"]) expect(sOf(ref)).toBe(sOf("A1"));
    expect((await z.file("xl/sharedStrings.xml")?.async("string")) ?? "").not.toContain("gizli");
    expect(writtenMerges(xml)).toEqual(["A1:C2"]);
    const back = await readXlsx(ab(out));
    expect(back.getValue(back.sheets[0], 0, 0)).toBe("Başlık");
    expect(back.getValue(back.sheets[0], 0, 1)).toBeNull();
  });

  it("modeldeki çakışan birleştirmeler yazılmaz (sol üstü önde olan kalır)", async () => {
    const wb = new WorkbookModel(true);
    const s = wb.addSheet("Veri");
    s.merges.push(R(1, 1, 2, 2), R(0, 0, 1, 1), R(0, 4, 0, 5));
    expect(writtenMerges(await sheetXml(await writeXlsx(wb)))).toEqual(["A1:B2", "E1:F1"]);
  });

  it("50.000 küçük birleştirme: açma ve yazma hızlı, gidiş-dönüşte aynı küme", async () => {
    // POI 57893-many-merges.xlsx biçiminde: her satırda iki hücrelik birleştirme, dönüşümlü A:B ve B:C
    const rows: string[] = [];
    const refs: string[] = [];
    for (let r = 1; r <= 50_000; r++) {
      const [a, b] = r % 2 ? ["A", "B"] : ["B", "C"];
      rows.push(`<row r="${r}"><c r="${a}${r}" s="1"${r % 100 ? "" : ' t="inlineStr"'}>${r % 100 ? "" : `<is><t>s${r}</t></is>`}</c><c r="${b}${r}" s="1"/></row>`);
      refs.push(`${a}${r}:${b}${r}`);
    }
    // Dosyadaki sıra karışık (Excel/POI sıralı yazmaz)
    const mixed = [...refs.slice(25_000), ...refs.slice(0, 25_000)];
    const data = await paket(rows.join(""), block(mixed));
    const wb = await quick(() => readXlsx(data));
    const s = wb.sheets[0];
    expect(s.merges).toHaveLength(50_000);
    expect(wb.getValue(s, 99, 1)).toBe("s100");
    expect(wb.loadWarnings).toEqual([]);
    const out = await quick(() => writeXlsx(wb));
    const back = await quick(() => readXlsx(ab(out)));
    expect(sameSet(back.sheets[0].merges, s.merges)).toBe(true);
    expect(back.getValue(back.sheets[0], 99, 1)).toBe("s100");
    expect(back.styleAt(back.sheets[0], 1, 2).bRight?.style).toBe("thin");
  });
});
