/** 5. tur QA: doldurma (Ctrl), özel yapıştır, not/köprü/doğrulama kopyalama, Yinelenenleri Kaldır, dosya önekleri ve dinamik diziler. */
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { cellKey, isError } from "@officehere/formula";
import { Controller } from "../src/controller";
import * as D from "../src/actions/data";
import * as E from "../src/actions/edit";
import { fillTipText } from "../src/grid/gridUtil";
import { rect } from "../src/model/rect";
import type { CellInput } from "../src/model/workbook";
import { addXlfnPrefixes, normalizeFileFormula } from "../src/io/xlsxFormula";
import { readXlsx, writeXlsx } from "../src/io/xlsx";
import { a1FormulaToOdf, odfFormulaToA1 } from "../src/io/ods/formula";
import { WorkbookModel } from "../src/model/workbook";
import { findAll } from "../src/model/find";
import { CondFormatter } from "../src/model/condfmt";

function put(ctl: Controller, cells: [number, number, CellInput][]) {
  ctl.ops.setInputs(ctl.sheet, cells.map(([row, col, input]) => ({ row, col, input })), "x");
}
const val = (ctl: Controller, row: number, col: number) => ctl.wb.getValue(ctl.sheet, row, col);

describe("doldurma tutamacı: Ctrl varsayılanı tersine çevirir", () => {
  it("tek sayı Ctrl ile birer artar (ipucu = sonuç), yukarı doğru azalır", () => {
    const ctl = new Controller();
    put(ctl, [[4, 0, 5]]);
    expect(fillTipText(ctl, rect(4, 0), rect(4, 0, 7, 0), true)).toBe("8");
    ctl.ops.fill(ctl.sheet, rect(4, 0), rect(4, 0, 7, 0), "alt");
    expect([5, 6, 7].map((r) => val(ctl, r, 0))).toEqual([6, 7, 8]);
    ctl.ops.fill(ctl.sheet, rect(4, 0), rect(2, 0, 4, 0), "alt");
    expect([2, 3].map((r) => val(ctl, r, 0))).toEqual([3, 4]);
    // Ctrl'siz tek sayı kopyalanır
    put(ctl, [[0, 2, 5]]);
    ctl.ops.fill(ctl.sheet, rect(0, 2), rect(0, 2, 2, 2), "series");
    expect(val(ctl, 2, 2)).toBe(5);
  });

  it("tarih, metin+sayı ve iki sayılı seri Ctrl ile kopyalanır", () => {
    const ctl = new Controller();
    put(ctl, [
      [0, 0, 45292],
      [0, 1, "Ürün 1"],
      [0, 2, 2],
      [1, 2, 4],
    ]);
    ctl.ops.applyStyle(ctl.sheet, [rect(0, 0)], { numFmt: "dd.mm.yyyy" });
    expect(fillTipText(ctl, rect(0, 0), rect(0, 0, 3, 0), true)).toBe("01.01.2024");
    ctl.ops.fill(ctl.sheet, rect(0, 0), rect(0, 0, 3, 0), "alt");
    expect(val(ctl, 3, 0)).toBe(45292);
    ctl.ops.fill(ctl.sheet, rect(0, 1), rect(0, 1, 2, 1), "alt");
    expect(val(ctl, 2, 1)).toBe("Ürün 1");
    ctl.ops.fill(ctl.sheet, rect(0, 2, 1, 2), rect(0, 2, 3, 2), "alt");
    expect([2, 3].map((r) => val(ctl, r, 2))).toEqual([2, 4]);
  });
});

describe("özel yapıştır ve pano", () => {
  function setup() {
    const ctl = new Controller();
    put(ctl, [
      [0, 0, 1234.5],
      [1, 0, { formula: "=A1*2" }],
    ]);
    ctl.ops.applyStyle(ctl.sheet, [rect(0, 0, 1, 0)], { numFmt: "#,##0.00", bold: true, fill: "#FFFF00" });
    ctl.sheet.notes.set(cellKey(0, 0), { text: "Kaynak notu" });
    ctl.sheet.links.set(cellKey(1, 0), "https://ornek.org/");
    return ctl;
  }
  const copyRange = (ctl: Controller, r: ReturnType<typeof rect>, cut = false) => {
    ctl.selectRange(r);
    E.copySelection(ctl, cut);
  };
  const pasteAt = (ctl: Controller, r: ReturnType<typeof rect>, opts?: Parameters<typeof E.pasteData>[3]) => {
    ctl.selectRange(r);
    E.pasteData(ctl, "", undefined, opts);
  };

  it("Değerler ve sayı biçimleri yalnız sayı biçimini getirir; hedefin dolgusu korunur", () => {
    const ctl = setup();
    ctl.ops.applyStyle(ctl.sheet, [rect(0, 3, 1, 3)], { fill: "#C6EFCE", italic: true });
    copyRange(ctl, rect(0, 0, 1, 0));
    pasteAt(ctl, rect(0, 3), { what: "valuesAndFormats" });
    const st = ctl.wb.styleAt(ctl.sheet, 0, 3);
    expect(st.numFmt).toBe("#,##0.00");
    expect(st.fill).toBe("#C6EFCE");
    expect(st.italic).toBe(true);
    expect(st.bold).toBeUndefined();
    expect(ctl.wb.getFormula(ctl.sheet, 1, 3)).toBeUndefined();
    expect(val(ctl, 1, 3)).toBe(2469);
    pasteAt(ctl, rect(0, 5), { what: "formulasAndFormats" });
    expect(ctl.wb.getFormula(ctl.sheet, 1, 5)).toBe("=F1*2");
    expect(ctl.wb.styleAt(ctl.sheet, 1, 5).fill).toBeUndefined();
  });

  it("Tümünü yapıştır notu ve köprüyü de kopyalar; geri alınca kalkar", () => {
    const ctl = setup();
    copyRange(ctl, rect(0, 0, 1, 0));
    pasteAt(ctl, rect(5, 2));
    expect(ctl.sheet.notes.get(cellKey(5, 2))?.text).toBe("Kaynak notu");
    expect(ctl.sheet.links.get(cellKey(6, 2))).toBe("https://ornek.org/");
    expect(ctl.sheet.notes.get(cellKey(0, 0))?.text).toBe("Kaynak notu");
    ctl.ops.undo();
    expect(ctl.sheet.notes.has(cellKey(5, 2))).toBe(false);
    expect(ctl.sheet.links.has(cellKey(6, 2))).toBe(false);
  });

  it("Kes-yapıştır notu ve köprüyü taşır", () => {
    const ctl = setup();
    copyRange(ctl, rect(0, 0, 1, 0), true);
    pasteAt(ctl, rect(10, 1));
    expect(ctl.sheet.notes.get(cellKey(10, 1))?.text).toBe("Kaynak notu");
    expect(ctl.sheet.links.get(cellKey(11, 1))).toBe("https://ornek.org/");
    expect(ctl.sheet.notes.has(cellKey(0, 0))).toBe(false);
    expect(ctl.sheet.links.has(cellKey(1, 0))).toBe(false);
  });

  it("Yalnız notlar ve doğrulama yapıştırılabilir; Tümü doğrulamayı da taşır", () => {
    const ctl = setup();
    ctl.sheet.validations.push({ id: "v1", ranges: [rect(0, 0, 1, 0)], rule: { type: "list", formula1: "Evet;Hayır" } });
    copyRange(ctl, rect(0, 0, 1, 0));
    pasteAt(ctl, rect(0, 7), { what: "comments" });
    expect(ctl.sheet.notes.get(cellKey(0, 7))?.text).toBe("Kaynak notu");
    expect(val(ctl, 0, 7)).toBeNull();
    pasteAt(ctl, rect(0, 8), { what: "validation" });
    expect(val(ctl, 0, 8)).toBeNull();
    const at = (r: number, c: number) => ctl.sheet.validations.find((v) => v.ranges.some((x) => r >= x.r1 && r <= x.r2 && c >= x.c1 && c <= x.c2));
    expect(at(1, 8)?.rule.formula1).toBe("Evet;Hayır");
    pasteAt(ctl, rect(4, 9));
    expect(at(5, 9)?.rule.type).toBe("list");
    // Doğrulamasız kaynak, hedefteki doğrulamayı kaldırır
    put(ctl, [[20, 0, "x"]]);
    copyRange(ctl, rect(20, 0));
    pasteAt(ctl, rect(1, 8));
    expect(at(1, 8)).toBeUndefined();
    expect(at(0, 8)).toBeDefined();
    ctl.ops.undo();
    expect(at(1, 8)).toBeDefined();
  });

  it("Bağlantıyı Yapıştır: tek hücrede mutlak, aralıkta göreli; başka sayfadan sayfa adlı", () => {
    const ctl = setup();
    copyRange(ctl, rect(0, 0));
    pasteAt(ctl, rect(0, 4), { what: "link" });
    expect(ctl.wb.getFormula(ctl.sheet, 0, 4)).toBe("=$A$1");
    copyRange(ctl, rect(0, 0, 1, 0));
    pasteAt(ctl, rect(3, 4), { what: "link" });
    expect(ctl.wb.getFormula(ctl.sheet, 4, 4)).toBe("=A2");
    expect(val(ctl, 4, 4)).toBe(2469);
    const other = ctl.ops.addSheet();
    ctl.ops.renameSheet(ctl.wb.sheets.indexOf(other), "Özet Ş");
    ctl.wb.activeIndex = ctl.wb.sheets.indexOf(other);
    pasteAt(ctl, rect(0, 0), { what: "link" });
    expect(ctl.wb.getFormula(other, 0, 0)).toBe("=Sayfa1!A1");
    expect(ctl.wb.getValue(other, 1, 0)).toBe(2469);
  });
});

describe("Yinelenenleri Kaldır", () => {
  it("yukarı kayan satırlardaki formüller kendi satırına başvurmaya devam eder", () => {
    const ctl = new Controller();
    const rows: (string | number)[][] = [
      ["Ad", "Adet", "Tutar"],
      ["a", 1, "=B2*10"],
      ["a", 1, "=B3*10"],
      ["b", 2, "=B4*10"],
      ["c", 3, "=B5*10"],
    ];
    rows.forEach((r, i) => r.forEach((v, j) => ctl.ops.enterText(ctl.sheet, [{ row: i, col: j }], String(v))));
    ctl.sheet.notes.set(cellKey(4, 0), { text: "c notu" });
    const res = D.removeDuplicates(ctl, rect(0, 0, 4, 2), [0, 1], true);
    expect(res).toEqual({ removed: 1, kept: 3 });
    expect([1, 2, 3].map((r) => val(ctl, r, 0))).toEqual(["a", "b", "c"]);
    expect([1, 2, 3].map((r) => val(ctl, r, 2))).toEqual([10, 20, 30]);
    expect(ctl.wb.getFormula(ctl.sheet, 2, 2)).toBe("=B3*10");
    expect(val(ctl, 4, 0)).toBeNull();
    expect(ctl.wb.getFormula(ctl.sheet, 4, 2)).toBeUndefined();
    expect(ctl.sheet.notes.get(cellKey(3, 0))?.text).toBe("c notu");
    // Tek adımda geri alınır
    ctl.ops.undo();
    expect([1, 2, 3, 4].map((r) => val(ctl, r, 2))).toEqual([10, 10, 20, 30]);
    expect(ctl.wb.getFormula(ctl.sheet, 3, 2)).toBe("=B4*10");
  });
});

describe("XLSX: işlev önekleri ve dinamik diziler", () => {
  it("FİLTRE/SIRALA _xlfn._xlws., LET/LAMBDA değişkenleri _xlpm. önekiyle yazılır", () => {
    expect(addXlfnPrefixes("SUM(FILTER(A1:A3,B1:B3))+INDEX(SORT(A1:A3),1)")).toBe("SUM(_xlfn._xlws.FILTER(A1:A3,B1:B3))+INDEX(_xlfn._xlws.SORT(A1:A3),1)");
    expect(addXlfnPrefixes("LET(x,2,y,x+1,x*y)")).toBe("_xlfn.LET(_xlpm.x,2,_xlpm.y,_xlpm.x+1,_xlpm.x*_xlpm.y)");
    expect(addXlfnPrefixes("MAP(A1:A2,LAMBDA(v,v*2))")).toBe("_xlfn.MAP(A1:A2,_xlfn.LAMBDA(_xlpm.v,_xlpm.v*2))");
    expect(addXlfnPrefixes("LET(f,LAMBDA(x,x+1),f(2))")).toBe("_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1),_xlpm.f(2))");
    // Kapsam dışındaki aynı adlı tanımlı ad öneksiz kalır
    expect(addXlfnPrefixes("LAMBDA(x,x)(1)+x")).toBe("_xlfn.LAMBDA(_xlpm.x,_xlpm.x)(1)+x");
    expect(normalizeFileFormula("_xlfn.LET(_xlpm.x,2,_xlfn._xlws.SORT(_xlpm.x))")).toBe("LET(x,2,SORT(x))");
  });

  it("taşan formül kaydet-aç sonrası yeniden taşar (#TAŞMA! olmaz); cm üst verisi yazılır", async () => {
    const wb = new WorkbookModel(true);
    const s = wb.addSheet("Sayfa1");
    wb.setInputRaw(s, 0, 0, { formula: "=SEQUENCE(2,2)" });
    wb.setInputRaw(s, 4, 0, { formula: "=MAP({1,2},LAMBDA(v,v*2))" });
    wb.setInputRaw(s, 6, 0, { formula: '=XLOOKUP(2,{1,2},{10,20})' });
    wb.setInputRaw(s, 7, 0, { formula: "=SUM(1,2)" });
    wb.engine.recalc();
    const bytes = await writeXlsx(wb);
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
    expect(xml).toMatch(/<c r="A1"[^>]* cm="1"[^>]*><f t="array" ref="A1:B2">_xlfn\.SEQUENCE\(2,2\)<\/f>/);
    expect(xml).toMatch(/<c r="A7"[^>]* cm="1"[^>]*><f t="array" ref="A7">_xlfn\.XLOOKUP/);
    expect(xml).not.toMatch(/<c r="A8"[^>]* cm="1"/);
    expect(zip.file("xl/metadata.xml")).toBeTruthy();
    expect(await zip.file("xl/metadata.xml")!.async("string")).toContain('name="XLDAPR"');
    expect(await zip.file("[Content_Types].xml")!.async("string")).toContain('PartName="/xl/metadata.xml"');
    expect(await zip.file("xl/_rels/workbook.xml.rels")!.async("string")).toContain('relationships/sheetMetadata" Target="metadata.xml"');
    const back = await readXlsx(bytes.buffer);
    const b = back.sheets[0];
    expect(back.getValue(b, 0, 0)).toBe(1);
    expect(back.getValue(b, 1, 1)).toBe(4);
    expect(back.getValue(b, 4, 0)).toBe(2);
    expect(back.getValue(b, 4, 1)).toBe(4);
    expect(back.getValue(b, 6, 0)).toBe(20);
    // Taşma alanı sabit değil: kaynak değişince sonuç küçülür
    back.setInputRaw(b, 0, 0, { formula: "=SEQUENCE(1,1)" });
    back.engine.recalc();
    expect(back.getValue(b, 1, 1)).toBeNull();
    expect(isError(back.getValue(b, 0, 0))).toBe(false);
  });
});

describe("ODS: OpenFormula işlev adları (LibreOffice ile doğrulandı)", () => {
  it("OpenFormula'nın kendi işlevleri öneksiz, FORMÜLMETNİ FORMULA, eski TINV/FTEST öneksiz", () => {
    expect(a1FormulaToOdf("=XOR(TRUE,FALSE)+IFNA(1,2)+DAYS(2,1)+ISOWEEKNUM(1)")).toBe("of:=XOR(TRUE();FALSE())+IFNA(1;2)+DAYS(2;1)+ISOWEEKNUM(1)");
    expect(a1FormulaToOdf("=FORMULATEXT(A1)")).toBe("of:=FORMULA([.A1])");
    expect(a1FormulaToOdf("=TINV(0.5,2)+FTEST(A1:A2,B1:B2)+FDIST(1,2,3)")).toBe("of:=TINV(0.5;2)+FTEST([.A1:.A2];[.B1:.B2])+LEGACY.FDIST(1;2;3)");
    expect(a1FormulaToOdf("=NETWORKDAYS.INTL(1,9,1)+STDEV.S(1,2)")).toBe("of:=COM.MICROSOFT.NETWORKDAYS.INTL(1;9;1)+COM.MICROSOFT.STDEV.S(1;2)");
  });

  it("okurken öneksiz FDIST/FINV/CHISQDIST sol kuyrukludur; eksik kümülatif DOĞRU olur", () => {
    expect(odfFormulaToA1("of:=FDIST([.A1];2;3;TRUE())")).toBe("=F.DIST(A1,2,3,TRUE())");
    expect(odfFormulaToA1("of:=FDIST([.A1];2;3)")).toBe("=F.DIST(A1,2,3,TRUE)");
    expect(odfFormulaToA1("of:=LEGACY.FDIST([.A1];2;3)")).toBe("=FDIST(A1,2,3)");
    expect(odfFormulaToA1("of:=FINV(0.5;2;3)")).toBe("=F.INV(0.5,2,3)");
    expect(odfFormulaToA1("of:=CHISQDIST(1;2)")).toBe("=CHISQ.DIST(1,2,TRUE)");
    expect(odfFormulaToA1("of:=LEGACY.CHIDIST(1;2)")).toBe("=CHIDIST(1,2)");
    expect(odfFormulaToA1("of:=FORMULA([.B2])")).toBe("=FORMULATEXT(B2)");
    expect(odfFormulaToA1("of:=COM.MICROSOFT.NETWORKDAYS.INTL(1;9;1)")).toBe("=NETWORKDAYS.INTL(1,9,1)");
  });
});

describe("Bul: Türkçe büyük/küçük harf", () => {
  it("istanbul İSTANBUL'u, ışık IŞIK'ı bulur; büyük/küçük harf eşleştir açıkken bulmaz", () => {
    const ctl = new Controller();
    ["İSTANBUL", "Isparta", "IŞIK", "istanbul"].forEach((t, i) => ctl.ops.enterText(ctl.sheet, [{ row: i, col: 0 }], t));
    const find = (query: string, matchCase = false) => findAll(ctl.wb, [ctl.sheet], { query, matchCase, wholeCell: false, lookIn: "values", byColumns: false }).map((h) => h.text);
    expect(find("istanbul")).toEqual(["İSTANBUL", "istanbul"]);
    expect(find("ışık")).toEqual(["IŞIK"]);
    expect(find("istanbul", true)).toEqual(["istanbul"]);
    expect(find("i?tanbul")).toEqual(["İSTANBUL", "istanbul"]);
  });
});

describe("sütun ekleme", () => {
  it("yeni sütun soldaki sütunun hücre biçimlerini alır (geri alınabilir)", () => {
    const ctl = new Controller();
    ctl.ops.enterText(ctl.sheet, [{ row: 0, col: 0 }], "Başlık");
    ctl.ops.applyStyle(ctl.sheet, [rect(0, 0)], { bold: true, fill: "#FFFF00" });
    ctl.ops.insertCols(ctl.sheet, 1, 2);
    expect(ctl.wb.styleAt(ctl.sheet, 0, 2)).toMatchObject({ bold: true, fill: "#FFFF00" });
    expect(val(ctl, 0, 1)).toBeNull();
    ctl.ops.undo();
    expect(ctl.wb.styleAt(ctl.sheet, 0, 1).bold).toBeUndefined();
  });
});

describe("filtre ve koşullu biçim: Türkçe büyük/küçük harf", () => {
  it("özel filtre ölçütleri ve metin içerir kuralı i/İ, ı/I çiftlerini eşler, joker kabul eder", () => {
    const ctl = new Controller();
    put(ctl, [[0, 0, "Şehir"], [1, 0, "İzmir"], [2, 0, "ISPARTA"], [3, 0, "Ankara"], [4, 0, "izmit"]]);
    const keep = (op1: string, v1: string) => [1, 2, 3, 4].filter((r) => D.criteriaMatch({ kind: "custom", op1, v1 }, val(ctl, r, 0), String(val(ctl, r, 0)), undefined, []));
    expect(keep("eq", "izmir")).toEqual([1]);
    expect(keep("begins", "ıs")).toEqual([2]);
    expect(keep("contains", "İZ")).toEqual([1, 4]);
    expect(keep("eq", "iz*")).toEqual([1, 4]);
    expect(keep("ne", "ankara")).toEqual([1, 2, 4]);
    ctl.sheet.condFormats.push({ id: "c", type: "containsText", text: "izm", ranges: [rect(1, 0, 4, 0)], style: { fill: "#FF0000" } });
    const cf = new CondFormatter(ctl.wb);
    expect([1, 2, 3, 4].map((r) => !!cf.evaluate(ctl.sheet, r, 0)?.style)).toEqual([true, false, false, true]);
  });
});
