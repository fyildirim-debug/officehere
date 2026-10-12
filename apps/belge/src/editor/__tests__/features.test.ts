// Otomatik Düzelt kuralları, yazdırma seçenekleri, kaynakça biçimleri, adres mektup verisi ve tablo sıralama yardımcıları.
import { describe, expect, it } from "vitest";
import { correctBefore, defaultAutoCorrect, isSentenceStart, smartQuote } from "../autocorrectRules";
import { defaultPrintOptions, parsePageRange, printSheets, selectedPages } from "../../print/printOptions";
import { bibEntry, citationText, orderSources, parseAuthors, sourceTag, type Source } from "../../model/sources";
import { fillRecord, mergeDocument, parseCsv, rowsToList } from "../../mailmerge/data";
import { compareKeys, parseDate, parseNumber } from "../tableTools";
import { normalizeTabs } from "../tabStops";
import { listStyleLevel, fieldNode } from "../../io/docx/partReader";

const ac = defaultAutoCorrect();
/** `before` metnine düzeltmeyi uygular. */
const fix = (before: string, o = ac, cell = false) => {
  const c = correctBefore(before, o, cell);
  return c ? before.slice(0, c.start) + c.text + before.slice(c.end) : before;
};

describe("Otomatik Düzelt", () => {
  it("değiştirme listesi", () => {
    expect(fix("Telif (c)")).toBe("Telif ©");
    expect(fix("Marka (tm)")).toBe("Marka ™");
    expect(fix("Bekle...")).toBe("Bekle…");
    expect(fix("Güldüm :)")).toBe("Güldüm ☺");
    expect(fix("Sonra -->")).toBe("Sonra →");
    expect(fix("Bunu herşey")).toBe("Bunu her şey");
    expect(fix("Bu malesef")).toBe("Bu maalesef");
    // Büyük harf biçimi korunur.
    expect(fix("Yanlız")).toBe("Yalnız");
  });

  it("kesirler ve tireler", () => {
    expect(fix("Yarım 1/2")).toBe("Yarım ½");
    expect(fix("Çeyrek 3/4")).toBe("Çeyrek ¾");
    expect(fix("Bir -- iki")).toBe("Bir – iki");
    expect(fix("Ankara--İstanbul")).toBe("Ankara—İstanbul");
  });

  it("cümle başı büyük harf (Türkçe i/İ)", () => {
    expect(fix("istanbul")).toBe("İstanbul");
    expect(fix("Geldi. ılık")).toBe("Geldi. Ilık");
    expect(fix("Bitti! şimdi")).toBe("Bitti! Şimdi");
    // Sıra sayısı ve kısaltmalar cümleyi bitirmez.
    expect(fix("Okulun 3. sınıfı")).toBe("Okulun 3. sınıfı");
    expect(fix("Bkz. vb. şeyler")).toBe("Bkz. vb. şeyler");
    expect(isSentenceStart("Merhaba. ", false, ac)).toBe(true);
    expect(isSentenceStart("Dr. ", false, ac)).toBe(false);
    expect(fix("hücre", { ...ac, capFirstCell: false }, true)).toBe("hücre");
    expect(fix("hücre", ac, true)).toBe("Hücre");
  });

  it("İKi BÜyük harf", () => {
    expect(fix("Bu İStanbul")).toBe("Bu İstanbul");
    expect(fix("Bu ANkara")).toBe("Bu Ankara");
    // Kısaltma (tümü büyük) değişmez.
    expect(fix("Bu TBMM")).toBe("Bu TBMM");
    expect(fix("Bu İKi", { ...ac, twoInitialCaps: false })).toBe("Bu İKi");
  });

  it("adresler ve kapalı seçenekler", () => {
    expect(fix("bkz. www.ornek.com")).toBe("bkz. www.ornek.com");
    expect(fix("Telif (c)", { ...ac, replaceText: false })).toBe("Telif (c)");
    expect(correctBefore("Zaten Doğru", ac)).toBeNull();
  });

  it("akıllı tırnaklar", () => {
    expect(smartQuote(undefined, '"')).toBe("“");
    expect(smartQuote(" ", '"')).toBe("“");
    expect(smartQuote("a", '"')).toBe("”");
    expect(smartQuote("(", "'")).toBe("‘");
    expect(smartQuote("k", "'")).toBe("’");
  });
});

describe("Yazdırma seçenekleri", () => {
  it("sayfa aralığı ayrıştırma", () => {
    expect(parsePageRange("1-3, 5", 10)).toEqual([0, 1, 2, 4]);
    expect(parsePageRange("5, 2; 2", 10)).toEqual([1, 4]);
    expect(parsePageRange("8-", 10)).toEqual([7, 8, 9]);
    expect(parsePageRange("9-20", 10)).toEqual([8, 9]);
    expect(typeof parsePageRange("3-1", 10)).toBe("string");
    expect(typeof parsePageRange("a", 10)).toBe("string");
    expect(typeof parsePageRange("", 10)).toBe("string");
    expect(typeof parsePageRange("12", 10)).toBe("string");
  });

  it("seçili sayfalar", () => {
    const o = defaultPrintOptions();
    expect(selectedPages(o, 3, 1)).toEqual([0, 1, 2]);
    expect(selectedPages({ ...o, range: "current" }, 3, 1)).toEqual([1]);
    expect(selectedPages({ ...o, range: "custom", pages: "2-3" }, 3, 0)).toEqual([1, 2]);
  });

  it("kopyalar, harmanlama, yaprak başına sayfa ve elle iki taraflı", () => {
    const o = defaultPrintOptions();
    expect(printSheets([0, 1], { ...o, copies: 2, collate: true })).toEqual([[0], [1], [0], [1]]);
    expect(printSheets([0, 1], { ...o, copies: 2, collate: false })).toEqual([[0], [0], [1], [1]]);
    expect(printSheets([0, 1, 2], { ...o, perSheet: 2 })).toEqual([[0, 1], [2]]);
    const manual = { ...o, duplex: "manual" as const };
    expect(printSheets([0, 1, 2, 3, 4], manual, 1)).toEqual([[0], [2], [4]]);
    expect(printSheets([0, 1, 2, 3, 4], manual, 2)).toEqual([[3], [1]]);
  });
});

const book: Source = { tag: "Yilmaz2020", type: "book", author: "Yılmaz, Ayşe; Demir, Can", title: "Türkçe Yazım", year: "2020", publisher: "Dil Yayınları", city: "Ankara" };
const article: Source = {
  tag: "Kaya2019",
  type: "article",
  author: "Kaya, Ali; Er, Su; Ak, Nur",
  title: "Bir İnceleme",
  year: "2019",
  publisher: "Dil Dergisi",
  volume: "12",
  issue: "3",
  pages: "10-20",
};

describe("Kaynakça", () => {
  it("yazar ayrıştırma ve etiket", () => {
    expect(parseAuthors("Yılmaz, Ayşe; Ali Demir")).toEqual([
      { last: "Yılmaz", first: "Ayşe" },
      { last: "Demir", first: "Ali" },
    ]);
    expect(sourceTag({ author: "Yılmaz, Ayşe", title: "X", year: "2020" }, [])).toBe("Yılmaz2020");
    expect(sourceTag({ author: "Yılmaz, Ayşe", title: "X", year: "2020" }, ["Yılmaz2020"])).toBe("Yılmaz2020_2");
  });

  it("APA ve IEEE alıntıları", () => {
    expect(citationText(book, "apa", 0)).toBe("(Yılmaz ve Demir, 2020)");
    expect(citationText(article, "apa", 1, "12")).toBe("(Kaya vd., 2019, s. 12)");
    expect(citationText(book, "ieee", 0)).toBe("[1]");
    expect(citationText(undefined, "apa", 0)).toBe("(Kaynak yok)");
  });

  it("kaynakça girdileri ve sıralama", () => {
    const apa = bibEntry(book, "apa", 0);
    expect(apa.before + apa.italic + apa.after).toBe("Yılmaz, A. ve Demir, C. (2020). Türkçe Yazım. Dil Yayınları.");
    const ieee = bibEntry(article, "ieee", 1);
    expect(ieee.before + ieee.italic + ieee.after).toBe("[2] A. Kaya, S. Er ve N. Ak, “Bir İnceleme,” Dil Dergisi, c. 12, sy. 3, ss. 10-20, 2019.");
    expect(orderSources([book, article], "apa", []).map((s) => s.tag)).toEqual(["Kaya2019", "Yilmaz2020"]);
    expect(orderSources([book, article], "ieee", ["Yilmaz2020", "Kaya2019"]).map((s) => s.tag)).toEqual(["Yilmaz2020", "Kaya2019"]);
  });
});

describe("Adres mektup verisi", () => {
  it("CSV: ayraç tahmini, tırnaklı alanlar, BOM", () => {
    expect(parseCsv('﻿Ad;Şehir\n"Yılmaz; Ayşe";İzmir\r\nAli;"An""kara"\n')).toEqual([
      ["Ad", "Şehir"],
      ["Yılmaz; Ayşe", "İzmir"],
      ["Ali", 'An"kara'],
    ]);
    expect(parseCsv("a,b\n1,2")).toEqual([
      ["a", "b"],
      ["1", "2"],
    ]);
    expect(parseCsv("a\tb\n1\t2\n\n")).toEqual([
      ["a", "b"],
      ["1", "2"],
    ]);
  });

  it("alıcı listesi ve birleştirme", () => {
    const list = rowsToList("alicilar.csv", [
      ["Ad", "Ad", " "],
      ["Ayşe", "Y", "x"],
    ]);
    expect(list.fields).toEqual(["Ad", "Ad_2", "Alan3"]);
    expect(list.records[0]).toEqual({ Ad: "Ayşe", Ad_2: "Y", Alan3: "x" });
    const content = {
      type: "doc",
      content: [
        {
          type: "section",
          content: [
            {
              type: "paragraph",
              content: [
                { type: "text", text: "Sayın " },
                { type: "mergeField", attrs: { name: "Ad" }, marks: [{ type: "bold" }] },
              ],
            },
          ],
        },
      ],
    };
    expect(fillRecord(content, { Ad: "Ayşe" }).content?.[0].content?.[0].content).toEqual([
      { type: "text", text: "Sayın " },
      { type: "text", text: "Ayşe", marks: [{ type: "bold" }] },
    ]);
    const merged = mergeDocument(content, [{ Ad: "Ayşe" }, { Ad: "Ali" }]);
    expect(merged.content).toHaveLength(2);
    expect(merged.content?.[1].attrs?.kind).toBe("nextPage");
    expect(mergeDocument(content, []).content).toHaveLength(1);
  });
});

describe("Tablo sıralama anahtarları", () => {
  it("Türkçe sayı ve tarih", () => {
    expect(parseNumber("1.234,5")).toBe(1234.5);
    expect(parseNumber("%12")).toBe(12);
    expect(parseNumber("-3")).toBe(-3);
    expect(parseNumber("1.000")).toBe(1000);
    expect(Number.isNaN(parseNumber("yok"))).toBe(true);
    expect(parseDate("15.12.2025")).toBe(Date.UTC(2025, 11, 15));
    expect(parseDate("1/2/24")).toBe(Date.UTC(2024, 1, 1));
  });

  it("karşılaştırma", () => {
    expect(["Zeynep", "çiğdem", "Ali", "İsmail", "ılgaz"].sort((a, b) => compareKeys(a, b, "text"))).toEqual(["Ali", "çiğdem", "ılgaz", "İsmail", "Zeynep"]);
    expect(["10", "9", "100"].sort((a, b) => compareKeys(a, b, "number"))).toEqual(["9", "10", "100"]);
    expect(["01.01.2024", "31.12.2023"].sort((a, b) => compareKeys(a, b, "date"))).toEqual(["31.12.2023", "01.01.2024"]);
    expect(compareKeys("", "5", "number")).toBeLessThan(0);
  });
});

describe("Sekme durakları ve DOCX yardımcıları", () => {
  it("sekme duraklarını sıralar ve yinelenenleri ayıklar", () => {
    expect(
      normalizeTabs([
        { pos: 100, align: "left", leader: "none" },
        { pos: 50, align: "right", leader: "dot" },
        { pos: 100, align: "center", leader: "none" },
      ]).map((t) => t.pos),
    ).toEqual([50, 100]);
  });

  it("liste stili düzeyi ve alan kodları", () => {
    expect(listStyleLevel(null, "List Bullet 2")).toBe(1);
    expect(listStyleLevel(null, "List Number")).toBe(0);
    expect(listStyleLevel(2, "List Bullet 3")).toBe(2);
    expect(fieldNode(" SEQ Şekil \\* ARABIC ", "3")).toMatchObject({ type: "seqField", attrs: { label: "Şekil", num: 3 } });
    expect(fieldNode(" REF _Ref123 \\h ", "Şekil 1")).toMatchObject({ type: "crossRef", attrs: { ref: "_Ref123", text: "Şekil 1" } });
    expect(fieldNode(" PAGEREF _Ref9 \\h ", "4")).toMatchObject({ type: "crossRef", attrs: { ref: "_Ref9", kind: "page" } });
    expect(fieldNode(" CITATION Kaya2019 \\l 1055 ", "(Kaya vd., 2019)")).toMatchObject({ type: "citation", attrs: { tag: "Kaya2019" } });
    expect(fieldNode(' MERGEFIELD "Ad" ', "«Ad»")).toMatchObject({ type: "mergeField", attrs: { name: "Ad" } });
    expect(fieldNode(" DATE ", "1")).toBeNull();
  });
});
