/** Metin modeli düzenleme işlemlerinin birim testleri (veri kaybı olmamalı). */
import { describe, expect, it } from "vitest";
import { para, textBody } from "./factory";
import {
  applyRunProps,
  applyRunPropsAll,
  caret,
  deleteRange,
  insertParagraphs,
  insertText,
  mergeParagraphs,
  nextPos,
  paraText,
  prevPos,
  propsAt,
  removeBackward,
  removeForward,
  sliceBody,
  splitParagraph,
  textInRange,
  wordBoundary,
  type Sel,
} from "./text";
import type { Paragraph, RunProps, TextBody } from "./types";

const texts = (b: TextBody) => b.paragraphs.map(paraText);
const sel = (ap: number, ao: number, fp: number, fo: number): Sel => ({ anchor: { p: ap, o: ao }, focus: { p: fp, o: fo } });
const P = (runs: [string, RunProps?][], level = 0, endProps?: RunProps): Paragraph => ({ runs: runs.map(([text, props]) => ({ text, props: props ?? {} })), level, props: {}, endProps });
const body = (...paras: Paragraph[]): TextBody => ({ paragraphs: paras });
/** Gövdedeki tüm çalımları [metin, kalın mı] olarak döker. */
const runsOf = (b: TextBody, i: number) => b.paragraphs[i].runs.map((r) => [r.text, !!r.props.bold]);

describe("deleteRange — paragraf içi", () => {
  it("tek paragrafta orta kısmı siler", () => {
    expect(texts(deleteRange(textBody("Merhaba dünya"), { p: 0, o: 3 }, { p: 0, o: 8 }))).toEqual(["Merdünya"]);
  });
  it("ters sıralı uçları kabul eder", () => {
    expect(texts(deleteRange(textBody("abcdef"), { p: 0, o: 4 }, { p: 0, o: 1 }))).toEqual(["aef"]);
  });
  it("çalım sınırını aşan silme her iki çalımın biçimini korur", () => {
    const b = body(P([["Kalın", { bold: true }], [" normal"]]));
    const r = deleteRange(b, { p: 0, o: 3 }, { p: 0, o: 8 });
    expect(runsOf(r, 0)).toEqual([["Kal", true], ["rmal", false]]);
  });
  it("tüm metni silince paragraf sonu biçimi silinen metnin biçimini alır", () => {
    const b = body(P([["X", { bold: true, size: 40 }]]));
    const r = deleteRange(b, { p: 0, o: 0 }, { p: 0, o: 1 });
    expect(r.paragraphs[0].runs).toEqual([]);
    expect(r.paragraphs[0].endProps).toMatchObject({ bold: true, size: 40 });
  });
  it("alan (slayt numarası) çalımı bölünmez, tümüyle silinir", () => {
    const b = body(P([["No "], ["12", { field: "slidenum" }], [" son"]]));
    const r = deleteRange(b, { p: 0, o: 4 }, { p: 0, o: 5 });
    expect(texts(r)).toEqual(["No  son"]);
  });
});

describe("deleteRange — paragraflar arası (K1)", () => {
  it("ikinci paragrafın başında Backspace: AlphaBeta", () => {
    const b = textBody("Alpha\nBeta");
    expect(texts(deleteRange(b, { p: 0, o: 5 }, { p: 1, o: 0 }))).toEqual(["AlphaBeta"]);
  });
  it("sondaki boş paragrafta Backspace önceki maddeleri korur", () => {
    const b = textBody("Madde 1\nMadde 2\n");
    const r = deleteRange(b, prevPos(b, { p: 2, o: 0 }), { p: 2, o: 0 });
    expect(texts(r)).toEqual(["Madde 1", "Madde 2"]);
  });
  it("çok paragraflı seçim ilk paragrafın başını ve son paragrafın sonunu korur", () => {
    const b = textBody("Birinci satır\nİkinci\nÜçüncü satır");
    const r = deleteRange(b, { p: 0, o: 4 }, { p: 2, o: 6 });
    expect(texts(r)).toEqual(["Biri satır"]);
  });
  it("biçimli çalımlar birleşmede korunur", () => {
    const b = body(P([["Ka", { bold: true }], ["lın"]]), P([["İkinci", { italic: true }]]));
    const r = deleteRange(b, { p: 0, o: 5 }, { p: 1, o: 0 });
    expect(r.paragraphs).toHaveLength(1);
    expect(r.paragraphs[0].runs.map((x) => x.text)).toEqual(["Ka", "lın", "İkinci"]);
    expect(r.paragraphs[0].runs[0].props.bold).toBe(true);
    expect(r.paragraphs[0].runs[2].props.italic).toBe(true);
  });
  it("ilk paragrafın bir çalımın ortasından başlayan silme", () => {
    const b = body(P([["abc", { bold: true }], ["def"]]), P([["ghi"], ["jkl", { bold: true }]]));
    const r = deleteRange(b, { p: 0, o: 4 }, { p: 1, o: 4 });
    expect(runsOf(r, 0)).toEqual([["abc", true], ["d", false], ["kl", true]]);
  });
  it("birleşen paragraf ilk paragrafın düzeyini ve paragraf biçimini alır", () => {
    const b = body({ ...P([["Bir"]], 1), props: { align: "ctr" } }, P([["İki"]], 0));
    const r = deleteRange(b, { p: 0, o: 3 }, { p: 1, o: 0 });
    expect(r.paragraphs[0]).toMatchObject({ level: 1, props: { align: "ctr" } });
  });
  it("boş paragraflar arasında silme", () => {
    const b = textBody("\n\n");
    expect(deleteRange(b, { p: 0, o: 0 }, { p: 1, o: 0 }).paragraphs).toHaveLength(2);
    expect(deleteRange(b, { p: 0, o: 0 }, { p: 2, o: 0 }).paragraphs).toHaveLength(1);
  });
  it("tüm metni silmek tek boş paragraf bırakır ve ilk biçimi saklar", () => {
    const b = body(P([["A", { bold: true }]]), P([["B"]]));
    const r = deleteRange(b, { p: 0, o: 0 }, { p: 1, o: 1 });
    expect(r.paragraphs).toHaveLength(1);
    expect(r.paragraphs[0].runs).toEqual([]);
    expect(r.paragraphs[0].endProps?.bold).toBe(true);
  });
  it("paragrafın sonundaki alan çalımı ile birleşme alanı korur", () => {
    const b = body(P([["Slayt "], ["3", { field: "slidenum" }]]), P([["devam"]]));
    const r = deleteRange(b, { p: 0, o: 7 }, { p: 1, o: 0 });
    expect(r.paragraphs[0].runs.map((x) => x.text)).toEqual(["Slayt ", "3", "devam"]);
  });
});

describe("removeBackward / removeForward (Backspace ve Delete)", () => {
  it("Backspace paragraf başında birleştirir, imleç birleşme noktasına gider", () => {
    const r = removeBackward(textBody("Alpha\nBeta"), caret(1, 0));
    expect(texts(r.body)).toEqual(["AlphaBeta"]);
    expect(r.caret).toEqual({ p: 0, o: 5 });
  });
  it("Delete paragraf sonunda sonraki paragrafı birleştirir", () => {
    const r = removeForward(textBody("Alpha\nBeta"), caret(0, 5));
    expect(texts(r.body)).toEqual(["AlphaBeta"]);
    expect(r.caret).toEqual({ p: 0, o: 5 });
  });
  it("Backspace girintili paragrafın başında önce düzeyi düşürür", () => {
    const r = removeBackward(textBody("Üst\n\tAlt"), caret(1, 0));
    expect(texts(r.body)).toEqual(["Üst", "Alt"]);
    expect(r.body.paragraphs[1].level).toBe(0);
  });
  it("metnin başında Backspace ve sonunda Delete hiçbir şey yapmaz", () => {
    const b = textBody("Tek");
    expect(removeBackward(b, caret(0, 0)).body).toBe(b);
    expect(removeForward(b, caret(0, 3)).body).toBe(b);
  });
  it("seçim varsa yalnız seçim silinir", () => {
    const r = removeBackward(textBody("Madde 1\nMadde 2"), sel(1, 3, 0, 5));
    expect(texts(r.body)).toEqual(["Maddede 2"]);
    expect(r.caret).toEqual({ p: 0, o: 5 });
  });
  it("kelime silme sınırları", () => {
    const b = textBody("bir iki üç");
    expect(wordBoundary(b, { p: 0, o: 10 }, -1)).toEqual({ p: 0, o: 8 });
    expect(wordBoundary(b, { p: 0, o: 0 }, 1)).toEqual({ p: 0, o: 4 });
    expect(nextPos(textBody("a\nb"), { p: 0, o: 1 })).toEqual({ p: 1, o: 0 });
  });
});

describe("mergeParagraphs", () => {
  it("iki paragrafı biçimleriyle birleştirir", () => {
    const b = body(P([["Bir", { bold: true }]]), P([["İki"]]), P([["Üç"]]));
    const r = mergeParagraphs(b, 0);
    expect(texts(r)).toEqual(["Birİki", "Üç"]);
    expect(runsOf(r, 0)).toEqual([["Bir", true], ["İki", false]]);
  });
  it("son paragrafta değişiklik yapmaz", () => {
    const b = textBody("a\nb");
    expect(mergeParagraphs(b, 1)).toBe(b);
  });
});

describe("insertText", () => {
  it("imlece yazar ve imleci ilerletir", () => {
    const r = insertText(textBody("Merhaba"), caret(0, 3), "XYZ");
    expect(texts(r.body)).toEqual(["MerXYZhaba"]);
    expect(r.caret).toEqual({ p: 0, o: 6 });
  });
  it("çok paragraflı seçimin üzerine yazmak ilk ve son paragrafın dışını korur", () => {
    const r = insertText(textBody("Madde 1\nMadde 2\nMadde 3"), sel(0, 2, 2, 5), "Z");
    expect(texts(r.body)).toEqual(["MaZ 3"]);
    expect(r.caret).toEqual({ p: 0, o: 3 });
  });
  it("yazılan metin imleçteki çalımın biçimini alır", () => {
    const b = body(P([["Kalın", { bold: true }], ["düz"]]));
    const r = insertText(b, caret(0, 3), "x");
    expect(runsOf(r.body, 0)).toEqual([["Kalxın", true], ["düz", false]]);
  });
  it("boş paragrafa yazınca paragraf sonu biçimi kullanılır", () => {
    const b = body(P([], 0, { italic: true }));
    const r = insertText(b, caret(0, 0), "a");
    expect(r.body.paragraphs[0].runs[0].props.italic).toBe(true);
  });
  it("çok satırlı metin paragraflara bölünür, son satır kuyruğu taşır", () => {
    const r = insertText(textBody("AB"), caret(0, 1), "1\n2\n3");
    expect(texts(r.body)).toEqual(["A1", "2", "3B"]);
    expect(r.caret).toEqual({ p: 2, o: 1 });
  });
  it("satır sonu (\\v) aynı paragrafta kalır", () => {
    const r = insertText(textBody("AB"), caret(0, 1), "\v");
    expect(r.body.paragraphs).toHaveLength(1);
    expect(textInRange(r.body, { p: 0, o: 0 }, { p: 0, o: 3 })).toBe("A\vB");
  });
  it("alanın içine yazılmaz, alanın sonrasına yazılır", () => {
    const b = body(P([["S"], ["12", { field: "slidenum" }]]));
    const r = insertText(b, caret(0, 2), "x");
    expect(r.body.paragraphs[0].runs.map((x) => x.text)).toEqual(["S", "12", "x"]);
  });
});

describe("splitParagraph (Enter)", () => {
  it("paragrafı imleçte böler, düzey ve biçim yeni paragrafa geçer", () => {
    const b = body({ ...P([["Önce", { bold: true }], ["Sonra"]], 2), props: { align: "r" } });
    const r = splitParagraph(b, caret(0, 4));
    expect(texts(r.body)).toEqual(["Önce", "Sonra"]);
    expect(r.body.paragraphs[1]).toMatchObject({ level: 2, props: { align: "r" } });
    expect(r.caret).toEqual({ p: 1, o: 0 });
  });
  it("seçim varken Enter seçimi silip böler", () => {
    const r = splitParagraph(textBody("Madde 1\nMadde 2"), sel(0, 5, 1, 5));
    expect(texts(r.body)).toEqual(["Madde", " 2"]);
  });
  it("paragraf sonunda Enter boş paragraf ekler ve biçimi taşır", () => {
    const r = splitParagraph(body(P([["X", { size: 30 }]])), caret(0, 1));
    expect(texts(r.body)).toEqual(["X", ""]);
    expect(r.body.paragraphs[1].endProps?.size).toBe(30);
  });
});

describe("yapıştırma (insertParagraphs) ve kopyalama (sliceBody)", () => {
  const src = body(P([["Bir", { bold: true }], ["inci"]]), P([["İkinci"]], 1), P([["Üçüncü"]]));
  it("kopyalanan aralık paragraf sınırını doğru keser", () => {
    const sl = sliceBody(src, { p: 0, o: 2 }, { p: 1, o: 3 });
    expect(sl.map(paraText)).toEqual(["rinci", "İki"]);
    expect(textInRange(src, { p: 0, o: 2 }, { p: 1, o: 3 })).toBe("rinci\nİki");
  });
  it("tek paragraflık zengin yapıştırma biçimi korur", () => {
    const r = insertParagraphs(textBody("AB"), caret(0, 1), [P([["k", { bold: true }]])]);
    expect(runsOf(r.body, 0)).toEqual([["A", false], ["k", true], ["B", false]]);
    expect(r.caret).toEqual({ p: 0, o: 2 });
  });
  it("çok paragraflı yapıştırma baş ve kuyruğu korur", () => {
    const sl = sliceBody(src, { p: 0, o: 0 }, { p: 2, o: 6 });
    const r = insertParagraphs(textBody("[]"), caret(0, 1), sl);
    expect(texts(r.body)).toEqual(["[Birinci", "İkinci", "Üçüncü]"]);
    expect(r.body.paragraphs[1].level).toBe(1);
    expect(r.caret).toEqual({ p: 2, o: 6 });
  });
  it("seçimin üzerine çok paragraflı yapıştırma seçimi değiştirir", () => {
    const sl = sliceBody(src, { p: 0, o: 0 }, { p: 1, o: 6 });
    const r = insertParagraphs(textBody("Madde 1\nMadde 2"), sel(0, 5, 1, 5), sl);
    expect(texts(r.body)).toEqual(["MaddeBirinci", "İkinci 2"]);
  });
});

describe("biçim uygulama (applyRunProps)", () => {
  it("paragraflar arası aralığa kalın uygular, dışını değiştirmez", () => {
    const r = applyRunProps(textBody("abcd\nefgh"), { p: 0, o: 2 }, { p: 1, o: 2 }, (p) => ({ ...p, bold: true }));
    expect(runsOf(r, 0)).toEqual([["ab", false], ["cd", true]]);
    expect(runsOf(r, 1)).toEqual([["ef", true], ["gh", false]]);
    expect(texts(r)).toEqual(["abcd", "efgh"]);
  });
  it("boş paragrafın sonu biçimi de güncellenir", () => {
    const r = applyRunPropsAll(textBody("a\n\nb"), (p) => ({ ...p, size: 12 }));
    expect(r.paragraphs[1].endProps?.size).toBe(12);
    expect(propsAt(r, { p: 1, o: 0 }).size).toBe(12);
  });
  it("biçimi kaldırmak çalımları birleştirir", () => {
    const b = body(P([["a", { bold: true }], ["b"]]));
    const r = applyRunPropsAll(b, (p) => ({ ...p, bold: undefined }));
    expect(r.paragraphs[0].runs).toHaveLength(1);
    expect(paraText(r.paragraphs[0])).toBe("ab");
  });
  it("factory para() boş paragrafı biçimle kurar", () => {
    expect(para("", { bold: true }).endProps?.bold).toBe(true);
  });
});
