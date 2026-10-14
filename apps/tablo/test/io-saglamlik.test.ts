/**
 * Tablo dosya G/Ç sağlamlığı (Aşama 1): tablolu (ListObject) XLSX'in kaydedip
 * yeniden açılması, ODS stil kayması, bozuk/kötü niyetli girdilerde takılma ve
 * sessiz kayıp, ham teknik hataların Türkçe iletiye çevrilmesi.
 * Fikstürler: lo-tablo.xlsx ve lo-stil.ods LibreOffice ile üretildi (ölçüm derlemi);
 * "Excel benzeri" paketler testte elle kurulur.
 */
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import JSZip from "jszip";
import { cellKey, keyCol, keyRow } from "@officehere/formula";
import { addXlfnPrefixes, readXlsx, writeXlsx } from "../src/io/xlsx";
import { readOds } from "../src/io/ods/reader";
import { readCsv } from "../src/io/csv";
import { writeOds } from "../src/io/ods/writer";
import { openBook } from "../src/io/formats";
import { FormatError } from "../src/io/bookUtil";
import { classifyError, OPEN_MESSAGES, toOpenError } from "../src/io/openErrors";
import { checkPackageLimits, collapseWhitespace, repairMerges, repairRowNumbers, repairTruncatedSheet, WsCollapser, XLSX_LIMITS } from "../src/io/xlsxGuard";
import { WorkbookModel, type SheetModel } from "../src/model/workbook";
import type { CellStyle } from "../src/model/styles";

const fixture = (name: string): ArrayBuffer => {
  const b = readFileSync(new URL(`./fixtures/${name}`, import.meta.url));
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer;
};
/** Ağır G/Ç testleri: yük altındaki CI'da da zaman aşımına düşmesin (takılma eşiği ayrıca ölçülür). */
const SLOW = { timeout: 60_000 };
const ab = (u: Uint8Array): ArrayBuffer => u.buffer.slice(u.byteOffset, u.byteOffset + u.byteLength) as ArrayBuffer;

/** Paketteki bir parçayı değiştirir (silmek için null döndür). */
async function mutate(data: ArrayBuffer, edits: Record<string, (xml: string) => string | null>): Promise<ArrayBuffer> {
  const z = await JSZip.loadAsync(data);
  for (const [path, fn] of Object.entries(edits)) {
    const xml = (await z.file(path)?.async("string")) ?? "";
    const next = fn(xml);
    if (next === null) z.remove(path);
    else z.file(path, next);
  }
  return ab(await z.generateAsync({ type: "uint8array", compression: "DEFLATE" }));
}

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** En küçük XLSX paketi (Excel'in yazdığı biçime benzer). */
async function makeXlsx(sheetXml: string, opts: { table?: string; styles?: string } = {}): Promise<ArrayBuffer> {
  const z = new JSZip();
  z.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${opts.table ? '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' : ""}</Types>`,
  );
  z.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  z.file("xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook ${NS}><sheets><sheet name="Veri" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  z.file(
    "xl/_rels/workbook.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/></Relationships>`,
  );
  z.file(
    "xl/styles.xml",
    opts.styles ??
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><dxfs count="1"><dxf><numFmt numFmtId="2" formatCode="0.00"/></dxf></dxfs></styleSheet>`,
  );
  z.file("xl/worksheets/sheet1.xml", sheetXml);
  if (opts.table) {
    z.file("xl/worksheets/_rels/sheet1.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/table" Target="../tables/table1.xml"/></Relationships>`);
    z.file("xl/tables/table1.xml", opts.table);
  }
  return ab(await z.generateAsync({ type: "uint8array", compression: "DEFLATE" }));
}

/** Excel'in yazdığı gibi: başlık satırı (headerRowCount yazılmaz), ortada hesaplanmış sütun, son sütunda filtre. */
function excelLikeTable(): Promise<ArrayBuffer> {
  const rows = [`<row r="1"><c r="A1" t="inlineStr"><is><t>Ad</t></is></c><c r="B1" t="inlineStr"><is><t>Adet</t></is></c><c r="C1" t="inlineStr"><is><t>Tutar</t></is></c><c r="D1" t="inlineStr"><is><t>Fiyat</t></is></c></row>`];
  for (let i = 0; i < 5; i++) {
    const r = i + 2;
    rows.push(
      `<row r="${r}"><c r="A${r}" t="inlineStr"><is><t>Ürün ${i}</t></is></c><c r="B${r}"><v>${i + 1}</v></c><c r="C${r}"><f>Tablo1[[#This Row],[Adet]]*Tablo1[[#This Row],[Fiyat]]</f><v>${(i + 1) * 5}</v></c><c r="D${r}"><v>5</v></c></row>`,
    );
  }
  const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet ${NS}><dimension ref="A1:D6"/><sheetData>${rows.join("")}</sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>`;
  const table =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Tablo1" displayName="Tablo1" ref="A1:D6" totalsRowShown="0">` +
    `<autoFilter ref="A1:D6"><filterColumn colId="3"><filters><filter val="5"/></filters></filterColumn></autoFilter>` +
    `<tableColumns count="4"><tableColumn id="1" name="Ad"/><tableColumn id="2" name="Adet"/><tableColumn id="3" name="Tutar" dataDxfId="0"><calculatedColumnFormula>Tablo1[[#This Row],[Adet]]*Tablo1[[#This Row],[Fiyat]]</calculatedColumnFormula></tableColumn><tableColumn id="4" name="Fiyat"/></tableColumns>` +
    `<tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`;
  return makeXlsx(sheet, { table });
}

const tableXml = async (bytes: Uint8Array) => (await JSZip.loadAsync(bytes)).file("xl/tables/table1.xml")!.async("string");
const countCells = (wb: WorkbookModel, s: SheetModel) => {
  let n = 0;
  wb.engine.forEachCell(s.id, () => n++);
  return n;
};

/* ======================= 1. Tablolu XLSX: kaydet → yeniden aç ======================= */

describe("Tablolu (ListObject) XLSX gidiş-dönüş", SLOW, () => {
  it("LibreOffice üretimi: hücre değiştir → kaydet → yeniden açılır, 8 sütun korunur (iki tur)", async () => {
    const wb = await readXlsx(fixture("lo-tablo.xlsx"));
    const s = wb.sheets[0];
    expect(s.tables).toHaveLength(1);
    const t = s.tables[0];
    expect(t.columns).toEqual(["Ürün", "Şehir", "Adet", "Fiyat", "Tarih", "Tutar", "KDV", "Durum"]);
    wb.setInputRaw(s, 4, 2, 99);
    wb.engine.recalc();
    const out = await writeXlsx(wb);
    // Excel/LibreOffice uyumlu tablo XML'i: boş filterColumn yok, hesaplanmış sütunlar var
    const x = await tableXml(out);
    expect(x).not.toMatch(/hiddenButton/);
    expect(x).toMatch(/<autoFilter ref="A2:H52"\/>/);
    expect(x.match(/<calculatedColumnFormula>/g)).toHaveLength(3);
    expect(x).toMatch(/<tableColumns count="8">/);
    const back = await readXlsx(ab(out));
    expect(back.loadWarnings).toEqual([]);
    expect(back.sheets[0].tables[0].columns).toEqual(t.columns);
    expect(back.getValue(back.sheets[0], 4, 2)).toBe(99);
    // İkinci tur
    const back2 = await readXlsx(ab(await writeXlsx(back)));
    expect(back2.sheets[0].tables[0].columns).toHaveLength(8);
  });

  it("Excel benzeri: ortadaki hesaplanmış sütun ve filtreli son sütun — 4 sütun, başlık ve ölçüt korunur", async () => {
    const wb = await readXlsx(await excelLikeTable());
    expect(wb.loadWarnings).toEqual([]);
    const t = wb.sheets[0].tables[0];
    expect(t.name).toBe("Tablo1");
    expect(t.columns).toEqual(["Ad", "Adet", "Tutar", "Fiyat"]);
    // headerRowCount yazılmamış: Excel'in varsayılanı 1 (başlık var)
    expect(t.headerRow).toBe(true);
    expect(t.filterButton).toBe(true);
    expect(t.style).toBe("TableStyleMedium9");
    expect(t.filter?.columns[3]).toEqual({ kind: "values", values: ["5"], blanks: false });
    expect(wb.getValue(wb.sheets[0], 3, 2)).toBe(15);
    const out = await writeXlsx(wb);
    const x = await tableXml(out);
    expect(x).toMatch(/<autoFilter ref="A1:D6"><filterColumn colId="3"><filters><filter val="5"\/><\/filters><\/filterColumn><\/autoFilter>/);
    expect(x.match(/<calculatedColumnFormula>/g)).toHaveLength(1);
    const back = await readXlsx(ab(out));
    expect(back.loadWarnings).toEqual([]);
    expect(back.sheets[0].tables[0].columns).toEqual(["Ad", "Adet", "Tutar", "Fiyat"]);
    expect(back.sheets[0].tables[0].filter?.columns[3]).toEqual({ kind: "values", values: ["5"], blanks: false });
  });

  it("exceljs'i çökerten biçim (her sütunda filterColumn + hesaplanmış sütun) yine açılır", async () => {
    const data = await mutate(await excelLikeTable(), {
      "xl/tables/table1.xml": (x) => x.replace(/<autoFilter[\s\S]*?<\/autoFilter>/, `<autoFilter ref="A1:D6">${[0, 1, 2, 3].map((i) => `<filterColumn colId="${i}" hiddenButton="0"/>`).join("")}</autoFilter>`),
    });
    const wb = await readXlsx(data);
    expect(wb.sheets[0].tables[0].columns).toHaveLength(4);
  });

  it("yapısal başvurunun @ kısaltması dosyaya uzun biçimde yazılır", () => {
    expect(addXlfnPrefixes("Tablo1[@Adet]*Tablo1[@[Birim Fiyat]]")).toBe("Tablo1[[#This Row],[Adet]]*Tablo1[[#This Row],[Birim Fiyat]]");
    expect(addXlfnPrefixes("SUM(Tablo1[Adet])")).toBe("SUM(Tablo1[Adet])");
  });

  it("başlıksız ya da filtre düğmesiz tablo autoFilter yazmaz", async () => {
    const wb = await readXlsx(await excelLikeTable());
    const t = wb.sheets[0].tables[0];
    t.filterButton = false;
    t.filter = undefined;
    expect(await tableXml(await writeXlsx(wb))).not.toMatch(/autoFilter/);
    const back = await readXlsx(ab(await writeXlsx(wb)));
    expect(back.sheets[0].tables[0].filterButton).toBe(false);
  });
});

/* ======================= 2. ODS stil kayması ======================= */

const STYLE_KEYS: (keyof CellStyle)[] = ["font", "size", "bold", "italic", "underline", "color", "fill", "hAlign", "vAlign", "wrap", "numFmt", "bTop", "bBottom", "bLeft", "bRight"];

/** İki kitapta hücre başına etkin stil farkları (birleşik alanın gizli hücreleri hariç). */
function styleDiffs(a: WorkbookModel, b: WorkbookModel): string[] {
  const out: string[] = [];
  a.sheets.forEach((sa, i) => {
    const sb = b.sheets.find((s) => s.name === sa.name) ?? b.sheets[i];
    const keys = new Set<number>();
    for (const [w, s] of [
      [a, sa],
      [b, sb],
    ] as const) {
      w.engine.forEachCell(s.id, (r, c) => keys.add(cellKey(r, c)));
      for (const k of s.styles.keys()) keys.add(k);
    }
    for (const k of keys) {
      const r = keyRow(k);
      const c = keyCol(k);
      const m = sa.merges.find((x) => r >= x.r1 && r <= x.r2 && c >= x.c1 && c <= x.c2);
      if (m && (m.r1 !== r || m.c1 !== c)) continue;
      const xa = a.styleAt(sa, r, c);
      const xb = b.styleAt(sb, r, c);
      for (const key of STYLE_KEYS) if (JSON.stringify(xa[key] ?? null) !== JSON.stringify(xb[key] ?? null)) out.push(`${sa.name}!R${r + 1}C${c + 1}.${key}: ${JSON.stringify(xa[key])} → ${JSON.stringify(xb[key])}`);
    }
  });
  return out;
}

describe("ODS stilleri", SLOW, () => {
  it("LibreOffice ODS → ODS: hücre başına etkin stil aynı kalır", async () => {
    const a = await readOds(fixture("lo-stil.ods"));
    const b = await readOds(await writeOds(a));
    expect(styleDiffs(a, b)).toEqual([]);
  });

  it("açıkça \"Default\" stilli hücre, biçimli sütunda biçimsiz kalır (LibreOffice'teki gibi)", async () => {
    const wb = await readOds(fixture("lo-stil.ods"));
    const s = wb.sheets[0];
    // D54: sütun varsayılanı "#,##0.00 TL" + kenarlık; hücre table:style-name="Default"
    expect(wb.styleAt(s, 2, 3).numFmt).toBe('#,##0.00" TL"');
    expect(wb.styleAt(s, 53, 3).numFmt).toBeUndefined();
    expect(wb.styleAt(s, 53, 3).bTop).toBeUndefined();
    // Yinelenen boş hücre her sütunun kendi varsayılanını alır (ilk sütununkini değil)
    expect(wb.styleAt(s, 53, 9).bTop).toBeUndefined();
  });

  it("ODS → XLSX → okuma: etkin stiller korunur (biçimsiz hücre stilli sütunda)", async () => {
    const a = await readOds(fixture("lo-stil.ods"));
    const b = await readXlsx(ab(await writeXlsx(a)));
    expect(styleDiffs(a, b).filter((d) => !/\.(font|size)/.test(d))).toEqual([]);
  });

  it("yinelenen hücreler sütun öbeklerine bölünür; satır/sütun varsayılanı doğru uygulanır", async () => {
    const content =
      `<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.3">` +
      `<office:automatic-styles><style:style style:name="kalin" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-weight="bold"/></style:style><style:style style:name="egik" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-style="italic"/></style:style></office:automatic-styles>` +
      `<office:body><office:spreadsheet><table:table table:name="S"><table:table-column table:default-cell-style-name="kalin"/><table:table-column table:default-cell-style-name="egik"/><table:table-column table:number-columns-repeated="2" table:default-cell-style-name="Default"/>` +
      `<table:table-row><table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell><table:table-cell table:number-columns-repeated="3"/></table:table-row>` +
      `<table:table-row><table:table-cell table:number-columns-repeated="2" office:value-type="float" office:value="2"><text:p>2</text:p></table:table-cell><table:table-cell table:style-name="Default" office:value-type="float" office:value="3"><text:p>3</text:p></table:table-cell></table:table-row>` +
      `<table:table-row><table:table-cell table:style-name="Default" table:number-columns-repeated="2"/><table:table-cell office:value-type="float" office:value="4"><text:p>4</text:p></table:table-cell></table:table-row>` +
      `</table:table></office:spreadsheet></office:body></office:document-content>`;
    const z = new JSZip();
    z.file("mimetype", "application/vnd.oasis.opendocument.spreadsheet");
    z.file("content.xml", content);
    const wb = await readOds(await z.generateAsync({ type: "uint8array" }));
    const s = wb.sheets[0];
    expect(wb.styleAt(s, 0, 0).bold).toBe(true);
    // Yinelenen içerikli hücre (2 sütun): A kalın, B eğik
    expect(wb.styleAt(s, 1, 0).bold).toBe(true);
    expect(wb.styleAt(s, 1, 1).italic).toBe(true);
    expect(wb.styleAt(s, 1, 1).bold).toBeUndefined();
    // Açık "Default": sütun stili uygulanmaz
    expect(wb.styleAt(s, 2, 0).bold).toBeUndefined();
    expect(wb.styleAt(s, 2, 1).italic).toBeUndefined();
    // Yazılıp yeniden okununca aynı
    const back = await readOds(await writeOds(wb));
    expect(styleDiffs(wb, back)).toEqual([]);
  });
});

/* ======================= 3. Takılma: dev öznitelikler, birleştirmeler, bombalar ======================= */

/** Takılma denetimi: düzeltme öncesi bu girdiler 60 sn'yi aşıyordu; eşik yük altındaki makinelere göre geniş. */
const quick = async <T>(fn: () => Promise<T>, ms = 20_000): Promise<T> => {
  const t = performance.now();
  const r = await fn();
  expect(performance.now() - t).toBeLessThan(ms);
  return r;
};

describe("Takılmaya karşı girdi sınırları", SLOW, () => {
  it("satır numarası 2^31-1 (ve NaN/0/-1) takılmaz; hücreler sırayla yerleşir", async () => {
    for (const bad of ["2147483647", "NaN", "0", "-1"]) {
      const data = await mutate(fixture("lo-tablo.xlsx"), { "xl/worksheets/sheet1.xml": (x) => x.replace(/<row r="5"/, `<row r="${bad}"`).replace(/<c r="A5"/, `<c r="${bad === "NaN" ? "NaN" : "A5"}"`) });
      const wb = await quick(() => readXlsx(data));
      expect(wb.loadWarnings).toContain("satır/hücre numaraları (onarıldı)");
      expect(wb.getValue(wb.sheets[0], 4, 1)).toBeTruthy();
      await quick(() => writeXlsx(wb));
    }
  });

  it("A1:XFD1048576 birleştirmesi takılmaz; kayıt da takılmaz", async () => {
    const data = await mutate(fixture("lo-tablo.xlsx"), { "xl/worksheets/sheet3.xml": (x) => x.replace("</sheetData>", '</sheetData><mergeCells count="1"><mergeCell ref="A1:XFD1048576"/></mergeCells>') });
    const wb = await quick(() => readXlsx(data));
    const s = wb.sheets[2];
    for (const m of s.merges) expect((m.r2 - m.r1 + 1) * (m.c2 - m.c1 + 1)).toBeLessThanOrEqual(2_000_000);
    expect(wb.loadWarnings.some((w) => /birleştiril/.test(w))).toBe(true);
    const back = await quick(async () => readXlsx(ab(await writeXlsx(wb))));
    expect(countCells(back, back.sheets[2])).toBe(countCells(wb, s));
  });

  it("çakışan birleştirmeler dosyanın tamamını reddettirmez", async () => {
    const r = repairMerges(`<worksheet><sheetData/><mergeCells count="3"><mergeCell ref="A1:B2"/><mergeCell ref="B2:C3"/><mergeCell ref="D1:E1"/></mergeCells></worksheet>`)!;
    expect(r.xml).toContain('<mergeCell ref="A1:B2"/>');
    expect(r.xml).toContain('<mergeCell ref="D1:E1"/>');
    expect(r.xml).not.toContain("B2:C3");
    expect(r.dropped).toBe(1);
  });

  it("dev sütun aralığı/genişliği kırpılır", async () => {
    const data = await mutate(fixture("lo-tablo.xlsx"), { "xl/worksheets/sheet2.xml": (x) => x.replace(/<col [^>]*\/>/, '<col min="1" max="2147483647" width="2147483647" customWidth="1"/>') });
    const wb = await quick(() => readXlsx(data));
    expect(wb.sheets[1].cols.size(0)).toBeLessThan(2000);
  });

  it("ODS: dev yayılma (number-rows-spanned=2^31) kullanılan alana kırpılır, xlsx kaydı takılmaz", async () => {
    const z = await JSZip.loadAsync(fixture("lo-stil.ods"));
    const c = await z.file("content.xml")!.async("string");
    z.file("content.xml", c.replace(/table:number-columns-spanned="8" table:number-rows-spanned="1"/, 'table:number-columns-spanned="8" table:number-rows-spanned="2147483647"'));
    const wb = await quick(async () => readOds(await z.generateAsync({ type: "uint8array" })));
    const big = wb.sheets[0].merges.find((m) => m.r1 === 0 && m.c1 === 0)!;
    expect(big.r2).toBeLessThan(1000);
    expect(wb.loadWarnings.some((w) => /birleştiril/.test(w))).toBe(true);
    await quick(() => writeXlsx(wb));
  });

  it("sıkıştırma bombası: sayfa XML'ine eklenmiş dev boşluk ayıklanır, hızla açılır", async () => {
    const pad = " ".repeat(XLSX_LIMITS.bombPart + 1024 * 1024);
    const data = await mutate(fixture("lo-tablo.xlsx"), { "xl/worksheets/sheet2.xml": (x) => x.replace(/>\s*$/, ">" + pad) });
    const wb = await quick(() => readXlsx(data));
    expect(countCells(wb, wb.sheets[1])).toBeGreaterThan(10);
    expect(collapseWhitespace(`<a>${" ".repeat(5000)}<b/><t>${" ".repeat(2000)}</t></a>${" ".repeat(3000)}`)).toBe(`<a><b/><t>${" ".repeat(2000)}</t></a>`);
  });

  it("boşluk ayıklayıcı parça sınırlarında da aynı sonucu verir", () => {
    const xml = `<a>${" ".repeat(3000)}<b>x</b>${" ".repeat(10)}<t>${" ".repeat(1500)}</t>\n<c/>${" ".repeat(5000)}`;
    const whole = collapseWhitespace(xml);
    expect(whole).toBe(`<a><b>x</b>${" ".repeat(10)}<t>${" ".repeat(1500)}</t>\n<c/>`);
    for (const size of [1, 7, 100, 1023, 1024, 2048]) {
      const c = new WsCollapser();
      for (let i = 0; i < xml.length; i += size) c.push(xml.slice(i, i + size));
      // Parçalı okumada metin düğümündeki boşluk en çok 1024 karakterle korunur; diğerleri aynı
      expect(c.end().replace(/<t> +<\/t>/, "<t></t>")).toBe(whole.replace(/<t> +<\/t>/, "<t></t>"));
    }
  });

  it("ODS sıkıştırma bombası (content.xml sonuna dev boşluk) hızla açılır", async () => {
    const z = await JSZip.loadAsync(fixture("lo-stil.ods"));
    const c = await z.file("content.xml")!.async("string");
    z.file("content.xml", c + " ".repeat(XLSX_LIMITS.bombPart + 1024 * 1024));
    const wb = await quick(async () => readOds(await z.generateAsync({ type: "uint8array", compression: "DEFLATE" })));
    expect(countCells(wb, wb.sheets[0])).toBeGreaterThan(400);
  });

  it("paket sınırları: bildirilen açılmış toplam 1 GB'ı aşarsa ya da boyutlar tutarsızsa anlaşılır hata", () => {
    const fake = (files: Record<string, [number, number]>) => ({ files: Object.fromEntries(Object.entries(files).map(([n, [u, c]]) => [n, { name: n, dir: false, _data: { uncompressedSize: u, compressedSize: c } }])) }) as unknown as JSZip;
    const MB = 1024 * 1024;
    expect(() => checkPackageLimits(fake({ "a.xml": [300 * MB, 3 * MB], "b.xml": [300 * MB, 3 * MB], "c.xml": [300 * MB, 3 * MB], "d.xml": [300 * MB, 3 * MB] }))).toThrow(/açılamayacak kadar büyük/);
    expect(() => checkPackageLimits(fake({ "a.xml": [10, 50_000] }))).toThrow(/tutarsız/);
    expect(checkPackageLimits(fake({ "a.xml": [100 * MB, 200 * 1024] })).has("a.xml")).toBe(true);
  });

  it("CSV: 16.384 sütunu aşan satır sonraki satırın üzerine yazmaz; uyarılır", () => {
    const line = Array.from({ length: 20_000 }, (_, i) => String(i)).join(";");
    const wb = readCsv(new TextEncoder().encode(`${line}\nson;satır\n`), "Liste");
    const s = wb.sheets[0];
    expect(wb.getValue(s, 0, 16383)).toBe(16383);
    expect(wb.getValue(s, 1, 0)).toBe("son");
    expect(wb.getValue(s, 1, 3615)).toBeNull();
    expect(wb.loadWarnings.join(" ")).toMatch(/sınırını aşan/);
  });

  it("satır numarası onarımı yalnız sorun varsa yeniden yazar", () => {
    expect(repairRowNumbers('<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>')).toBeNull();
    const r = repairRowNumbers('<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2147483647"><c r="2147483647"><v>2</v></c><c><v>3</v></c></row></sheetData></worksheet>')!;
    expect(r.xml).toContain('<row r="2"><c r="A2"><v>2</v></c><c r="B2"><v>3</v></c></row>');
  });
});

/* ======================= 4. Sessiz kayıp: bozuk yapı uyarılı açılır ======================= */

describe("Bozuk yapı sessizce boş açılmaz", SLOW, () => {
  const notice = async (data: ArrayBuffer) => openBook("x.xlsx", data);

  it("workbook.xml yoksa sayfalar kurtarılır, ileti ve uyarı verilir", async () => {
    const r = await notice(await mutate(fixture("lo-tablo.xlsx"), { "xl/workbook.xml": () => null }));
    expect(r.wb.sheets).toHaveLength(3);
    expect(countCells(r.wb, r.wb.sheets[0])).toBeGreaterThan(400);
    expect(r.notice).toMatch(/kurtarıldı/);
    expect(r.wb.loadWarnings.length).toBeGreaterThan(0);
  });

  it("ilişki hedefleri bozuksa (ya da ilişki dosyası yoksa) sayfalar sırayla eşleşir", async () => {
    for (const edit of [(x: string) => x.replace(/Target="[^"]+"/g, 'Target="../../yok/olmayan.xml"'), () => null]) {
      const r = await notice(await mutate(fixture("lo-tablo.xlsx"), { "xl/_rels/workbook.xml.rels": edit }));
      expect(r.wb.sheets.map((s) => s.name)).toEqual(["Satışlar", "Özet", "Zincir"]);
      expect(countCells(r.wb, r.wb.sheets[0])).toBeGreaterThan(400);
      expect(r.notice).toMatch(/kurtarıldı/);
    }
  });

  it("döngüsel ilişki workbook.xml'i sayfa sanıp taşımaz", async () => {
    const r = await notice(await mutate(fixture("lo-tablo.xlsx"), { "xl/_rels/workbook.xml.rels": (x) => x.replace('Target="worksheets/sheet2.xml"', 'Target="workbook.xml"') }));
    expect(r.wb.sheets.map((s) => s.name)).toEqual(["Satışlar", "Özet", "Zincir"]);
    expect(countCells(r.wb, r.wb.sheets[1])).toBeGreaterThan(10);
  });

  it("sayfa parçası yoksa sayfa boş ama adıyla açılır ve uyarılır", async () => {
    const r = await notice(await mutate(fixture("lo-tablo.xlsx"), { "xl/worksheets/sheet1.xml": () => null }));
    expect(r.wb.sheets.map((s) => s.name)).toEqual(["Satışlar", "Özet", "Zincir"]);
    expect(r.wb.loadWarnings.join(" ")).toMatch(/Satışlar/);
    expect(r.notice).toBeTruthy();
  });

  it("paylaşılan metinler yoksa metin hücreleri dizin sayısı olarak görünmez; uyarılır", async () => {
    const r = await notice(await mutate(fixture("lo-tablo.xlsx"), { "xl/sharedStrings.xml": () => null }));
    expect(r.wb.loadWarnings.join(" ")).toMatch(/metin/);
    expect(typeof r.wb.getValue(r.wb.sheets[0], 2, 0)).not.toBe("number");
    expect(r.notice).toBeTruthy();
  });

  it("stil bölümü yoksa (koşullu biçimler dahil) açılır ve uyarılır", async () => {
    const wb = await readXlsx(await mutate(fixture("lo-tablo.xlsx"), { "xl/styles.xml": () => null }));
    expect(countCells(wb, wb.sheets[0])).toBeGreaterThan(400);
    expect(wb.loadWarnings.join(" ")).toMatch(/biçim/);
  });

  it("yarıda kesilmiş sayfa son tam satıra kadar kurtarılır", async () => {
    const r = await notice(await mutate(fixture("lo-tablo.xlsx"), { "xl/worksheets/sheet1.xml": (x) => x.slice(0, x.length >> 1) }));
    const n = countCells(r.wb, r.wb.sheets[0]);
    expect(n).toBeGreaterThan(50);
    expect(r.notice).toMatch(/yarıda kesilmiş/);
    expect(repairTruncatedSheet('<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v')).toBe('<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>');
  });

  it("silinmiş not/çizim/tablo parçaları uyarı verir; sayfa ilişkileri silinse de açılır", async () => {
    for (const [path, word] of [
      ["xl/comments1.xml", /notlar/],
      ["xl/charts/chart1.xml", /grafik/],
      ["xl/tables/table1.xml", /tablolar/],
      ["xl/worksheets/_rels/sheet1.xml.rels", /./],
    ] as const) {
      const wb = await readXlsx(await mutate(fixture("lo-tablo.xlsx"), { [path]: () => null }));
      expect(countCells(wb, wb.sheets[0])).toBeGreaterThan(400);
      expect(wb.loadWarnings.join(" ")).toMatch(word);
    }
  });

  it("tanımsız XML varlığı (&lol;) açılışı engellemez", async () => {
    const wb = await readXlsx(await mutate(fixture("lo-tablo.xlsx"), { "xl/sharedStrings.xml": (x) => x.replace("<t", "<t>&lol;</t></si><si><t") }));
    expect(wb.loadWarnings.join(" ")).toMatch(/varlık/);
  });

  it("ODS: yarıda kesilmiş içerik okunabilen bölümüyle açılır", async () => {
    const z = await JSZip.loadAsync(fixture("lo-stil.ods"));
    const c = await z.file("content.xml")!.async("string");
    z.file("content.xml", c.slice(0, c.length >> 1));
    const r = await openBook("x.ods", ab(await z.generateAsync({ type: "uint8array" })));
    expect(countCells(r.wb, r.wb.sheets[0])).toBeGreaterThan(50);
    expect(r.notice).toMatch(/yarıda kesilmiş/);
  });
});

/* ======================= 5. Ham teknik hatalar → Türkçe ileti + ayrıntı ======================= */

describe("Açma hatalarının sınıflandırılması", SLOW, () => {
  it("kitaplık iletileri sınıflandırılır", () => {
    const cases: [string, keyof typeof OPEN_MESSAGES][] = [
      ["Corrupted zip: can't find end of central directory", "corruptZip"],
      ["Bug : uncompressed data size mismatch", "corruptZip"],
      ["End of data reached (data length = 5, asked index = 9). Corrupted zip ?", "corruptZip"],
      ["invalid stored block lengths", "corruptZip"],
      ["2:16273: documents may contain only one root.", "corruptXml"],
      ["1:999: undefined entity.", "corruptXml"],
      ["[xmldom error]\telement parse error", "corruptXml"],
      ["Cannot set properties of undefined (setting 'sheetNo')", "corrupt"],
      ["Cannot read properties of undefined (reading 'Target')", "corrupt"],
      ["Invalid row number in model", "corrupt"],
      ["Invalid string length", "tooLarge"],
      ["File is encrypted", "encrypted"],
    ];
    for (const [m, k] of cases) expect([m, classifyError(new Error(m))]).toEqual([m, k]);
    expect(classifyError(new RangeError("Maximum call stack size exceeded"))).toBe("tooDeep");
    const e = toOpenError(new TypeError("Cannot read properties of undefined (reading 'Target')"));
    expect(e).toBeInstanceOf(FormatError);
    expect(e.message).toBe(OPEN_MESSAGES.corrupt);
    expect(e.detail).toBe("TypeError: Cannot read properties of undefined (reading 'Target')");
  });

  it("kesik ZIP: Türkçe ileti, teknik ayrıntı ayrı alanda", async () => {
    const b = new Uint8Array(fixture("lo-tablo.xlsx"));
    const cut = ab(b.slice(0, b.length >> 1));
    const err = (await openBook("kesik.xlsx", cut).catch((e) => e)) as FormatError;
    expect(err.name).toBe("FormatError");
    expect(err.message).toMatch(/^Dosya bozuk/);
    expect(err.message).not.toMatch(/zip|Corrupted|undefined/);
  });

  it("bozuk sıkıştırılmış veri tek sayfayı boşaltır, dosyanın tamamını değil", async () => {
    const z = await JSZip.loadAsync(fixture("lo-tablo.xlsx"));
    const raw = new Uint8Array(await z.generateAsync({ type: "uint8array", compression: "DEFLATE" }));
    // sheet2.xml'in sıkıştırılmış verisinin ortasında bayt boz
    const name = new TextEncoder().encode("xl/worksheets/sheet2.xml");
    let at = -1;
    for (let i = 0; i < raw.length - name.length && at < 0; i++) if (raw[i] === 0x50 && raw[i + 1] === 0x4b && raw[i + 2] === 3 && raw[i + 3] === 4 && name.every((v, j) => raw[i + 30 + j] === v)) at = i;
    expect(at).toBeGreaterThan(-1);
    const dataStart = at + 30 + name.length + (raw[at + 28] | (raw[at + 29] << 8));
    for (let i = 10; i < 40; i++) raw[dataStart + i] ^= 0x5a;
    const r = await openBook("x.xlsx", ab(raw)).catch((e) => e as FormatError);
    if (r instanceof Error) expect(r.message).toMatch(/^Dosya bozuk/);
    else {
      expect(r.wb.sheets).toHaveLength(3);
      expect(countCells(r.wb, r.wb.sheets[0])).toBeGreaterThan(400);
      expect(r.wb.loadWarnings.length).toBeGreaterThan(0);
    }
  });
});

/* ======================= 6-7. Sayı biçimi kaçışları, büyük dosya yazımı ======================= */

describe("Sayı biçimleri ve yazma verimi", SLOW, () => {
  it("ters bölü kaçışlı sayı biçimleri özgün haliyle korunur", async () => {
    const styles =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="2"><numFmt numFmtId="164" formatCode="0\\ \\k\\m"/><numFmt numFmtId="165" formatCode="dd\\.mm\\.yyyy"/></numFmts>` +
      `<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
      `<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>`;
    const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet ${NS}><sheetData><row r="1"><c r="A1" s="1"><v>12</v></c><c r="B1" s="2"><v>45000</v></c></row></sheetData></worksheet>`;
    const wb = await readXlsx(await makeXlsx(sheet, { styles }));
    const s = wb.sheets[0];
    expect(wb.styleAt(s, 0, 0).numFmt).toBe("0\\ \\k\\m");
    expect(wb.styleAt(s, 0, 1).numFmt).toBe("dd\\.mm\\.yyyy");
    const back = await readXlsx(ab(await writeXlsx(wb)));
    expect(back.styleAt(back.sheets[0], 0, 0).numFmt).toBe("0\\ \\k\\m");
  });

  it("aynı stilli çok hücre az sayıda xf üretir; dev birleştirme doğrudan yazılır", async () => {
    const wb = new WorkbookModel(true);
    const s = wb.addSheet("Veri");
    const ids = [wb.styles.intern({ bold: true }), wb.styles.intern({ numFmt: "0.00" }), wb.styles.intern({ fill: "#FFEEAA" })];
    for (let r = 0; r < 4000; r++)
      for (let c = 0; c < 5; c++) {
        wb.setInputRaw(s, r, c, r * 10 + c);
        wb.setStyleRaw(s, r, c, ids[(r + c) % 3]);
      }
    s.merges.push({ r1: 5000, c1: 0, r2: 25000, c2: 9 });
    const out = await quick(() => writeXlsx(wb), 30_000);
    const z = await JSZip.loadAsync(out);
    const st = await z.file("xl/styles.xml")!.async("string");
    expect(Number(/<cellXfs count="(\d+)"/.exec(st)![1])).toBeLessThan(10);
    expect(await z.file("xl/worksheets/sheet1.xml")!.async("string")).toContain('<mergeCell ref="A5001:J25001"/>');
    const back = await readXlsx(ab(out));
    expect(back.sheets[0].merges).toContainEqual({ r1: 5000, c1: 0, r2: 25000, c2: 9 });
  });
});
