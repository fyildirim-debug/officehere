// @vitest-environment jsdom
// Çekirdek düzenleme mantığı: büyük/küçük harf, arama, stiller, izlenen değişiklikler, komutlar.
import { afterEach, describe, expect, it } from "vitest";
import { Editor, type JSONContent } from "@tiptap/core";
import { normalizeDoc } from "../../model/document";
import { TextSelection } from "@tiptap/pm/state";
import { schemaExtensions } from "../extensions";
import { Search, adaptCase, findMatches, foldTr, getSearch, replaceAll, setSearch } from "../search";
import { TrackChanges, listChanges, resolveChanges, setTracking } from "../trackChanges";
import { applyStyle, changeCase, changeIndent, clearFormatting, countWords, docStats, setLineSpacing, sortParagraphs, stepFontSize, transformCase, effectiveFontSize } from "../commands";
import { resolveStyles, styleSheetCss, STYLE_SETS } from "../../model/styles";
import { levelDef, sigFromLevel } from "../../io/docx/numbering";

let editor: Editor | null = null;
afterEach(() => {
  editor?.destroy();
  editor = null;
});

function make(content: object) {
  editor = new Editor({ extensions: [...schemaExtensions(), Search, TrackChanges], content: normalizeDoc(content as JSONContent) });
  return editor;
}

/** Tek bölümlü belgenin blokları. */
const blocks = (json: JSONContent) => json.content![0].content!;

const para = (text: string) => ({ type: "paragraph", content: text ? [{ type: "text", text }] : undefined });
const doc = (...c: object[]) => ({ type: "doc", content: c });
const styles = resolveStyles("default", "OfficeHere");

describe("Türkçe büyük/küçük harf", () => {
  it("BÜYÜK HARF ve küçük harf (i/İ, ı/I)", () => {
    expect(transformCase("istanbul ılık", "upper")).toBe("İSTANBUL ILIK");
    expect(transformCase("İSTANBUL ILIK", "lower")).toBe("istanbul ılık");
  });
  it("Her Sözcüğün İlk Harfi Büyük ve Cümle düzeni", () => {
    expect(transformCase("izmir ve ığdır", "title")).toBe("İzmir Ve Iğdır");
    expect(transformCase("BU BİR CÜMLE. İKİNCİ cümle!", "sentence")).toBe("Bu bir cümle. İkinci cümle!");
  });
  it("bÜYÜK/kÜÇÜK hARF dEĞİŞTİR", () => {
    expect(transformCase("Iğdır", "toggle")).toBe("ıĞDIR");
  });
  it("editörde işaretleri koruyarak değiştirir", () => {
    const ed = make(doc({ type: "paragraph", content: [{ type: "text", text: "kalın ", marks: [{ type: "bold" }] }, { type: "text", text: "ince" }] }));
    ed.commands.selectAll();
    changeCase(ed, "upper");
    const json = ed.getJSON();
    expect(blocks(json)[0].content![0]).toEqual({ type: "text", text: "KALIN ", marks: [{ type: "bold" }] });
    expect((blocks(json)[0].content![1] as { text?: string }).text).toBe("İNCE");
  });
});

describe("arama", () => {
  it("Türkçe katlama", () => {
    expect(foldTr("İSTANBUL")).toBe(foldTr("istanbul"));
    expect(foldTr("Işık")).toBe(foldTr("ışık"));
  });
  it("eşleşmeler, tam sözcük ve büyük/küçük harf", () => {
    const ed = make(doc(para("Kedi kediler KEDİ. kedi"), para("bir kedi")));
    expect(findMatches(ed.state.doc, { query: "kedi", caseSensitive: false, wholeWord: false }).length).toBe(5);
    expect(findMatches(ed.state.doc, { query: "kedi", caseSensitive: false, wholeWord: true }).length).toBe(4);
    expect(findMatches(ed.state.doc, { query: "kedi", caseSensitive: true, wholeWord: false }).length).toBe(3);
  });
  it("Tümünü Değiştir", () => {
    const ed = make(doc(para("elma armut elma"), para("Elma")));
    setSearch(ed.view, { query: "elma", caseSensitive: false, wholeWord: true });
    expect(getSearch(ed.state)?.matches.length).toBe(3);
    expect(replaceAll(ed.view, "ayva")).toBe(3);
    // Word gibi: büyük/küçük harf eşleştirme kapalıyken yeni metin bulunanın biçimine uyar.
    expect(ed.state.doc.textContent).toBe("ayva armut ayvaAyva");
  });
  it("değiştirmede harf biçimi uyarlaması", () => {
    expect(adaptCase("ELMA", "ayva", false)).toBe("AYVA");
    expect(adaptCase("Elma", "ayva", false)).toBe("Ayva");
    expect(adaptCase("elma", "ayva", false)).toBe("ayva");
    expect(adaptCase("Elma", "ayva", true)).toBe("ayva");
    expect(adaptCase("ELMA", "Ayva", false)).toBe("Ayva");
    expect(adaptCase("İNCİR", "ıhlamur", false)).toBe("IHLAMUR");
  });
});

describe("stiller ve paragraf komutları", () => {
  it("stil kümeleri CSS üretir", () => {
    for (const s of STYLE_SETS) {
      const css = styleSheetCss(resolveStyles(s.id, "OfficeHere"));
      expect(css).toContain(".bl-doc h1");
      expect(css).toContain('[data-style="Title"]');
    }
    expect(styles.Normal.props.size).toBe(11);
    expect(styles.Normal.props.line).toBe(1.08);
    expect(styles.Normal.props.spaceAfter).toBe(8);
  });
  it("Başlık uygulama ve Normal'e dönüş", () => {
    const ed = make(doc(para("Metin")));
    applyStyle(ed, styles, "Heading2");
    expect(blocks(ed.getJSON())[0]).toMatchObject({ type: "heading", attrs: { level: 2 } });
    applyStyle(ed, styles, "Title");
    expect(blocks(ed.getJSON())[0]).toMatchObject({ type: "paragraph", attrs: { styleId: "Title" } });
    applyStyle(ed, styles, "Normal");
    expect(blocks(ed.getJSON())[0].attrs?.styleId).toBeNull();
  });
  it("karakter stili açılıp kapanır", () => {
    const ed = make(doc(para("vurgulu")));
    ed.commands.selectAll();
    applyStyle(ed, styles, "Emphasis");
    expect(JSON.stringify(ed.getJSON())).toContain('"charStyle"');
    applyStyle(ed, styles, "Emphasis");
    expect(JSON.stringify(ed.getJSON())).not.toContain('"charStyle"');
  });
  it("girinti, satır aralığı, temizleme", () => {
    const ed = make(doc(para("a")));
    changeIndent(ed, 1);
    changeIndent(ed, 1);
    expect(blocks(ed.getJSON())[0].attrs?.indentLeft).toBeCloseTo(70.8);
    changeIndent(ed, -1);
    expect(blocks(ed.getJSON())[0].attrs?.indentLeft).toBeCloseTo(35.4);
    setLineSpacing(ed, 1.5);
    expect(blocks(ed.getJSON())[0].attrs?.lineHeight).toBe(1.5);
    clearFormatting(ed);
    expect(blocks(ed.getJSON())[0].attrs?.indentLeft).toBeNull();
    expect(blocks(ed.getJSON())[0].attrs?.lineHeight).toBeNull();
  });
  it("yazı tipini büyült/küçült boyut listesini izler", () => {
    const ed = make(doc(para("abc")));
    ed.commands.selectAll();
    stepFontSize(ed, styles, 1);
    expect(effectiveFontSize(ed, styles)).toBe(12);
    stepFontSize(ed, styles, 1);
    expect(effectiveFontSize(ed, styles)).toBe(14);
    stepFontSize(ed, styles, -1, true);
    expect(effectiveFontSize(ed, styles)).toBe(13);
  });
  it("paragrafları Türkçe sıralar", () => {
    const ed = make(doc(para("Çilek"), para("Armut"), para("Üzüm"), para("Ceviz")));
    ed.commands.selectAll();
    sortParagraphs(ed, false);
    expect(blocks(ed.getJSON()).map((p) => (p.content?.[0] as { text?: string } | undefined)?.text)).toEqual(["Armut", "Ceviz", "Çilek", "Üzüm"]);
  });
});

describe("sayım", () => {
  it("sözcük ve karakter", () => {
    expect(countWords("Merhaba dünya, bu bir deneme.")).toBe(5);
    expect(countWords("Türkiye'nin başkenti Ankara'dır")).toBe(3);
    const ed = make(doc(para("İki sözcük"), para(""), para("üç")));
    const s = docStats(ed.state.doc);
    expect(s.words).toBe(3);
    expect(s.paragraphs).toBe(2);
    expect(s.charsNoSpaces).toBe(11);
  });
});

describe("Değişiklikleri İzle", () => {
  it("yazılan metin eklenmiş, silinen metin silinmiş olarak işaretlenir", () => {
    const ed = make(doc(para("merhaba dünya")));
    setTracking(ed.view, true, "Test");
    // "dünya" sonuna ekle
    ed.view.dispatch(ed.state.tr.insertText("!", ed.state.doc.content.size - 2));
    // "merhaba" sil (belge > bölüm > paragraf: metin 2'de başlar)
    ed.view.dispatch(ed.state.tr.delete(2, 9));
    const changes = listChanges(ed.state.doc);
    expect(changes.map((c) => [c.kind, c.text])).toEqual([
      ["deletion", "merhaba"],
      ["insertion", "!"],
    ]);
    // Metin belgede kalır
    expect(ed.state.doc.textContent).toBe("merhaba dünya!");
    // Tümünü kabul et
    ed.view.dispatch(resolveChanges(ed.state, 0, ed.state.doc.content.size, true));
    expect(ed.state.doc.textContent).toBe(" dünya!");
    expect(listChanges(ed.state.doc).length).toBe(0);
  });
  it("reddetme eklenenleri kaldırır, silinenleri geri getirir", () => {
    const ed = make(doc(para("abc")));
    setTracking(ed.view, true, "Test");
    ed.view.dispatch(ed.state.tr.setSelection(TextSelection.create(ed.state.doc, 3, 4)));
    ed.view.dispatch(ed.state.tr.insertText("X"));
    expect(ed.state.doc.textContent).toBe("abXc");
    ed.view.dispatch(resolveChanges(ed.state, 0, ed.state.doc.content.size, false));
    expect(ed.state.doc.textContent).toBe("abc");
  });
});

describe("Değişiklikleri İzle: biçim değişiklikleri", () => {
  it("kalın yapma izlenir; reddetme önceki biçimi geri yükler", () => {
    const ed = make(doc({ type: "paragraph", content: [{ type: "text", text: "ab", marks: [{ type: "italic" }] }, { type: "text", text: "cd" }] }));
    setTracking(ed.view, true, "Test");
    ed.view.dispatch(ed.state.tr.addMark(2, 6, ed.state.schema.marks.bold.create()));
    const changes = listChanges(ed.state.doc);
    expect(changes.map((c) => [c.kind, c.text, c.author])).toEqual([["formatChange", "abcd", "Test"]]);
    // Metin değişmez; yalnız biçim kaydı eklenir.
    expect(ed.state.doc.textContent).toBe("abcd");
    ed.view.dispatch(resolveChanges(ed.state, 0, ed.state.doc.content.size, false));
    const c = blocks(ed.getJSON())[0].content!;
    expect(c).toEqual([
      { type: "text", text: "ab", marks: [{ type: "italic" }] },
      { type: "text", text: "cd" },
    ]);
  });
  it("kabul etme yeni biçimi tutar; geri alınan biçim kaydı siler", () => {
    const ed = make(doc(para("abcd")));
    setTracking(ed.view, true, "Test");
    const bold = ed.state.schema.marks.bold.create();
    ed.view.dispatch(ed.state.tr.addMark(2, 4, bold));
    ed.view.dispatch(resolveChanges(ed.state, 0, ed.state.doc.content.size, true));
    expect(listChanges(ed.state.doc).length).toBe(0);
    expect(blocks(ed.getJSON())[0].content![0]).toEqual({ type: "text", text: "ab", marks: [{ type: "bold" }] });
    // Kalın yapıp yeniden kaldırınca değişiklik kalmaz.
    ed.view.dispatch(ed.state.tr.addMark(4, 6, bold));
    expect(listChanges(ed.state.doc).length).toBe(1);
    ed.view.dispatch(ed.state.tr.removeMark(4, 6, bold));
    expect(listChanges(ed.state.doc).length).toBe(0);
  });
});

describe("numaralandırma eşlemesi", () => {
  it("liste stilleri gidiş-dönüş", () => {
    for (const style of [null, "decimal-paren", "upper-roman", "upper-alpha", "lower-alpha-paren", "lower-alpha", "lower-roman"]) {
      const d = levelDef({ type: "ordered", style, multilevel: null }, 0);
      expect(sigFromLevel(d.numFmt, d.lvlText, 0).style).toBe(style);
    }
    for (const style of [null, "square", "diamond", "arrow", "check", "dash"]) {
      const d = levelDef({ type: "bullet", style, multilevel: null }, 0);
      expect(sigFromLevel(d.numFmt, d.lvlText, 0).style).toBe(style);
    }
    // Word'ün Symbol yazı tipi madde işareti
    expect(sigFromLevel("bullet", "", 0)).toEqual({ type: "bullet", style: null, multilevel: null });
  });
});

describe("liste uygulama", () => {
  it("liste ve paragrafa yayılan seçimde hepsi numaralı listeye döner (Tümünü Seç)", async () => {
    const { toggleList } = await import("../listApply");
    const ed = make(doc({ type: "bulletList", content: [{ type: "listItem", content: [para("Elma")] }, { type: "listItem", content: [para("Armut")] }] }, para("Liste dışı"), para("Son")));
    ed.commands.selectAll();
    toggleList(ed, "orderedList");
    const b = blocks(ed.getJSON());
    expect(b.length).toBe(1);
    expect(b[0].type).toBe("orderedList");
    expect(b[0].content!.length).toBe(4);
    // Hepsi numaralıyken yeniden uygulamak listeyi kaldırır.
    ed.commands.selectAll();
    toggleList(ed, "orderedList");
    expect(blocks(ed.getJSON()).map((x) => x.type)).toEqual(["paragraph", "paragraph", "paragraph", "paragraph"]);
  });
});

describe("Hücreleri Böl", () => {
  const cell = (text: string, attrs: Record<string, unknown> = {}) => ({ type: "tableCell", attrs, content: [para(text)] });
  const row = (...cells: object[]) => ({ type: "tableRow", content: cells });
  const grid = (ed: Editor) => {
    const t = blocks(ed.getJSON()).find((b) => b.type === "table")!;
    return t.content!.map((r) => (r.content ?? []).map((c) => `${(c.content ?? []).map((p) => (p.content ?? []).map((x) => x.text).join("")).join("/") || "_"}${c.attrs!.colspan > 1 ? `c${c.attrs!.colspan}` : ""}${c.attrs!.rowspan > 1 ? `r${c.attrs!.rowspan}` : ""}`).join("|"));
  };
  it("birleştirilmemiş hücre 3 sütuna bölünür; diğer satırdaki hücre genişler", async () => {
    const { splitCellGrid } = await import("../tableSplit");
    const ed = make(doc({ type: "table", content: [row(cell("A1", { colwidth: [120] }), cell("B1", { colwidth: [60] })), row(cell("A2", { colwidth: [120] }), cell("B2", { colwidth: [60] }))] }, para("")));
    // A1 içine imleç
    ed.commands.setTextSelection(5);
    expect(splitCellGrid(ed.state, ed.view.dispatch, 3, 1)).toBe(true);
    expect(grid(ed)).toEqual(["A1|_|_|B1", "A2c3|B2"]);
    const t = blocks(ed.getJSON())[0];
    expect(t.content![0].content![0].attrs!.colwidth).toEqual([40]);
    expect(t.content![1].content![0].attrs!.colwidth).toEqual([40, 40, 40]);
  });
  it("2 satıra bölme yeni satır ekler; komşu hücre satır kapsamı artar", async () => {
    const { splitCellGrid } = await import("../tableSplit");
    const ed = make(doc({ type: "table", content: [row(cell("A1"), cell("B1")), row(cell("A2"), cell("B2"))] }, para("")));
    ed.commands.setTextSelection(5);
    expect(splitCellGrid(ed.state, ed.view.dispatch, 1, 2)).toBe(true);
    expect(grid(ed)).toEqual(["A1|B1r2", "_", "A2|B2"]);
  });
  it("birleştirilmiş hücre geri bölünür (2×2)", async () => {
    const { splitCellGrid } = await import("../tableSplit");
    const ed = make(doc({ type: "table", content: [row(cell("X", { colspan: 2, rowspan: 2 }), cell("C1")), row(cell("C2"))] }, para("")));
    ed.commands.setTextSelection(5);
    expect(splitCellGrid(ed.state, ed.view.dispatch, 2, 2)).toBe(true);
    expect(grid(ed)).toEqual(["X|_|C1", "_|_|C2"]);
  });
});
