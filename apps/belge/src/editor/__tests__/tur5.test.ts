// @vitest-environment jsdom
// 5. tur (QA + rakip özellikleri): numaralandırma denetimi, yapıştırmada bölüm sarmalı, büyük/küçük harf
// döngüsü ve diğer düzenleyici davranışları.
import { afterEach, describe, expect, it } from "vitest";
import { Editor, type JSONContent } from "@tiptap/core";
import { TextSelection } from "@tiptap/pm/state";
import { closeHistory } from "@tiptap/pm/history";
import { normalizeDoc } from "../../model/document";
import { schemaExtensions } from "../extensions";
import { canContinueNumbering, continueNumbering, orderedListAt, restartNumbering } from "../listNumbering";
import { docStats, nextCaseMode } from "../commands";
import { listifySlice, textToSlice, unwrapSectionSlice } from "../paste";
import { formatPicture, isTimePicture, parsePicture, pictureOfInstr } from "../dateFormat";

let editor: Editor | null = null;
afterEach(() => {
  editor?.destroy();
  editor = null;
});
const make = (content: JSONContent) => (editor = new Editor({ extensions: schemaExtensions(), content: normalizeDoc(content) }));
const p = (text: string): JSONContent => ({ type: "paragraph", content: text ? [{ type: "text", text }] : [] });
const doc = (...c: JSONContent[]): JSONContent => ({ type: "doc", content: [{ type: "section", content: c }] });
const ol = (items: string[], attrs: Record<string, unknown> = {}): JSONContent => ({ type: "orderedList", attrs, content: items.map((t) => ({ type: "listItem", content: [p(t)] })) });

/** Metni içeren paragrafın başına imleç koyar. */
function cursorAt(ed: Editor, text: string) {
  let at = -1;
  ed.state.doc.descendants((n, pos) => {
    if (at < 0 && n.isTextblock && n.textContent === text) at = pos + 1;
    return at < 0;
  });
  if (at < 0) throw new Error(`bulunamadı: ${text}`);
  ed.view.dispatch(ed.state.tr.setSelection(TextSelection.create(ed.state.doc, at)));
}

/** Bölümdeki numaralı listelerin [başlangıç, öğe sayısı] çiftleri. */
const lists = (ed: Editor) => {
  const out: [number, number][] = [];
  ed.state.doc.descendants((n) => {
    if (n.type.name === "orderedList") out.push([Number(n.attrs.start) || 1, n.childCount]);
    return true;
  });
  return out;
};

describe("numaralandırma denetimi", () => {
  it("ortadaki öğede Yeniden Başla listeyi o öğeden böler ve 1'den başlatır", () => {
    const ed = make(doc(ol(["Bir", "İki", "Üç", "Dört"])));
    cursorAt(ed, "Üç");
    expect(restartNumbering(ed)).toBe(true);
    expect(lists(ed)).toEqual([
      [1, 2],
      [1, 2],
    ]);
    // Geri al tek adımda eski listeyi getirir.
    ed.commands.undo();
    expect(lists(ed)).toEqual([[1, 4]]);
  });

  it("ilk öğede Yeniden Başla yalnız başlangıç değerini ayarlar; değer verilebilir", () => {
    const ed = make(doc(ol(["Bir", "İki"], { start: 5 })));
    cursorAt(ed, "Bir");
    restartNumbering(ed);
    expect(lists(ed)).toEqual([[1, 2]]);
    cursorAt(ed, "İki");
    restartNumbering(ed, 10);
    expect(lists(ed)).toEqual([
      [1, 1],
      [10, 1],
    ]);
  });

  it("Devam Et: araya paragraf giren liste önceki listenin son numarasından sürer", () => {
    const ed = make(doc(ol(["Bir", "İki", "Üç"]), p("Ara metin"), ol(["Dört", "Beş"])));
    cursorAt(ed, "Beş");
    expect(canContinueNumbering(ed.state)).toBe(true);
    expect(continueNumbering(ed)).toBe(true);
    expect(lists(ed)).toEqual([
      [1, 3],
      [4, 2],
    ]);
    // Zaten devam ediyorsa komut etkin değildir.
    expect(canContinueNumbering(ed.state)).toBe(false);
  });

  it("Devam Et: bitişik (yeniden başlatılmış) liste öncekiyle birleşir", () => {
    const ed = make(doc(ol(["Bir", "İki", "Üç"])));
    cursorAt(ed, "Üç");
    restartNumbering(ed);
    expect(lists(ed)).toEqual([
      [1, 2],
      [1, 1],
    ]);
    cursorAt(ed, "Üç");
    continueNumbering(ed);
    expect(lists(ed)).toEqual([[1, 3]]);
  });

  it("önceki liste yoksa ya da imleç listede değilse komutlar uygulanmaz", () => {
    const ed = make(doc(p("Düz"), ol(["Bir"])));
    cursorAt(ed, "Bir");
    expect(canContinueNumbering(ed.state)).toBe(false);
    expect(continueNumbering(ed)).toBe(false);
    cursorAt(ed, "Düz");
    expect(orderedListAt(ed.state)).toBeNull();
    expect(restartNumbering(ed)).toBe(false);
  });

  it("iç içe liste: aynı düzeydeki önceki alt listeden devam eder", () => {
    const nested = (a: string, subs: string[]): JSONContent => ({ type: "listItem", content: [p(a), ol(subs)] });
    const ed = make(doc({ type: "orderedList", content: [nested("A", ["a1", "a2"]), nested("B", ["b1"])] }));
    cursorAt(ed, "b1");
    continueNumbering(ed);
    expect(lists(ed)).toEqual([
      [1, 2],
      [1, 2],
      [3, 1],
    ]);
  });
});

describe("yapıştırma: tek bölümün içeriği", () => {
  it("Tümünü Seç ile kopyalanan bölüm sarmalı açılır, kenar paragraflar birleşir", () => {
    const ed = make(doc(p("Birinci"), p("İkinci")));
    const all = ed.state.doc.slice(0, ed.state.doc.content.size);
    expect(all.content.firstChild!.type.name).toBe("section");
    const un = unwrapSectionSlice(all);
    expect(un.content.firstChild!.type.name).toBe("paragraph");
    expect([un.openStart, un.openEnd]).toEqual([1, 1]);
    // "Bi|rinci" konumuna yapıştırma yeni bölüm (sayfa sonu) eklemez.
    cursorAt(ed, "Birinci");
    const at = ed.state.selection.from + 2;
    ed.view.dispatch(ed.state.tr.replaceRange(at, at, un));
    expect(ed.state.doc.childCount).toBe(1);
    expect(ed.state.doc.textBetween(0, ed.state.doc.content.size, "|")).toBe("BiBirinci|İkincirinci|İkinci");
  });

  it("birden çok bölüme yayılan kopya bölüm sonlarını korur", () => {
    const ed = make({ type: "doc", content: [{ type: "section", content: [p("A")] }, { type: "section", content: [p("B")] }] });
    const all = ed.state.doc.slice(0, ed.state.doc.content.size);
    expect(unwrapSectionSlice(all)).toBe(all);
  });
});

describe("yapıştırma: liste öğesine çok satırlı metin", () => {
  it("her satır ayrı liste öğesi olur; ilk satır imleçteki öğeyle birleşir", () => {
    const ed = make(doc({ type: "bulletList", content: [{ type: "listItem", content: [p("İlk")] }] }));
    cursorAt(ed, "İlk");
    const at = ed.state.selection.from + 3;
    ed.view.dispatch(ed.state.tr.setSelection(TextSelection.create(ed.state.doc, at)));
    const slice = listifySlice(textToSlice(ed.state.schema, "a\nb\nc"), ed.state.selection.$from);
    ed.view.dispatch(ed.state.tr.replaceSelection(slice));
    const items: string[] = [];
    ed.state.doc.descendants((n) => {
      if (n.type.name === "listItem") items.push(n.textContent);
      return true;
    });
    expect(items).toEqual(["İlka", "b", "c"]);
    // Liste dışında ya da tek paragrafta dokunulmaz.
    const one = textToSlice(ed.state.schema, "tek");
    expect(listifySlice(one, ed.state.selection.$from)).toBe(one);
  });
});

describe("Shift+F3 döngüsü", () => {
  it("BÜYÜK → küçük → Her Sözcük; tam cümlede Cümle düzeni", () => {
    expect(nextCaseMode("merhaba dünya")).toBe("title");
    expect(nextCaseMode("Merhaba Dünya")).toBe("upper");
    expect(nextCaseMode("MERHABA DÜNYA")).toBe("lower");
    expect(nextCaseMode("merhaba dünya. iyi günler.")).toBe("sentence");
    expect(nextCaseMode("İSTANBUL")).toBe("lower");
    expect(nextCaseMode("ılık")).toBe("title");
  });
});

describe("Sözcük Sayısı", () => {
  it("seçim için tüm istatistikler; dipnot metni isteğe bağlı sayılır", () => {
    const ed = make(doc({ type: "paragraph", content: [{ type: "text", text: "Bir iki üç" }, { type: "note", attrs: { id: "n1", kind: "footnote", text: "Dipnot sözcükleri" } }] }, p("Dört beş"), p("")));
    const all = docStats(ed.state.doc);
    expect(all).toMatchObject({ words: 5, paragraphs: 2 });
    expect(docStats(ed.state.doc, null, true)).toMatchObject({ words: 7, paragraphs: 3 });
    // "iki üç" + "Dört" seçimi
    let a = 0;
    let b = 0;
    ed.state.doc.descendants((n, pos) => {
      if (n.isText && n.text === "Bir iki üç") a = pos + 4;
      if (n.isText && n.text === "Dört beş") b = pos + 4;
      return true;
    });
    const sel = docStats(ed.state.doc, { from: a, to: b });
    expect(sel).toMatchObject({ words: 3, paragraphs: 2, charsNoSpaces: 9 });
  });
});

describe("tarih biçim resimleri", () => {
  const d = new Date(2026, 1, 3, 9, 5, 7); // 3 Şubat 2026 Salı 09:05:07
  it("Türkçe ay/gün adları ve belirteçler", () => {
    expect(formatPicture("dd.MM.yyyy", d)).toBe("03.02.2026");
    expect(formatPicture("d MMMM yyyy dddd", d)).toBe("3 Şubat 2026 Salı");
    expect(formatPicture("d MMM yy", d)).toBe("3 Şub 26");
    expect(formatPicture("HH:mm:ss", d)).toBe("09:05:07");
    expect(formatPicture("h:mm AM/PM", new Date(2026, 1, 3, 15, 0))).toBe("3:00 ÖS");
    expect(formatPicture("'Tarih:' d.M.yyyy", d)).toBe("Tarih: 3.2.2026");
  });
  it("saat resmi ve alan yönergesi", () => {
    expect(isTimePicture("HH:mm")).toBe(true);
    expect(isTimePicture("dd.MM.yyyy HH:mm")).toBe(false);
    expect(pictureOfInstr(' DATE \\@ "d MMMM yyyy" \\* MERGEFORMAT')).toBe("d MMMM yyyy");
    expect(pictureOfInstr("DATE \\@ dd.MM.yyyy")).toBe("dd.MM.yyyy");
    expect(pictureOfInstr("DATE")).toBeNull();
    // LibreOffice'in yazdığı biçim: \@ ile tırnak bitişik, boşluklar kaçışlı.
    expect(pictureOfInstr(' DATE \\@"d\\ MMMM\\ yyyy\\ dddd" ')).toBe("d MMMM yyyy dddd");
    expect(parsePicture("yyyy-MM-dd").map((x) => x.kind)).toEqual(["year", "lit", "month", "lit", "day"]);
  });
});

describe("Büyük İlk Harf komutları", () => {
  it("uygula, kenar boşluğu kipine geç, kaldır; geri al", async () => {
    const { setDropCap, dropCapTarget, normalizeDropCap, firstGrapheme } = await import("../dropCap");
    const ed = make(doc(p("Şimdi metin"), ol(["Liste"])));
    cursorAt(ed, "Şimdi metin");
    expect(setDropCap(ed, { mode: "drop", lines: 3, distance: 0, font: null })).toBe(true);
    expect(ed.state.doc.firstChild!.firstChild!.attrs.dropCap).toEqual({ mode: "drop", lines: 3, distance: 0, font: null });
    setDropCap(ed, { mode: "margin", lines: 20, distance: -4, font: "Georgia" });
    // Satır sayısı 1-10, uzaklık ≥ 0 aralığına çekilir.
    expect(ed.state.doc.firstChild!.firstChild!.attrs.dropCap).toEqual({ mode: "margin", lines: 10, distance: 0, font: "Georgia" });
    ed.view.dispatch(closeHistory(ed.state.tr));
    setDropCap(ed, null);
    expect(ed.state.doc.firstChild!.firstChild!.attrs.dropCap).toBeNull();
    ed.commands.undo();
    expect(ed.state.doc.firstChild!.firstChild!.attrs.dropCap?.mode).toBe("margin");
    // Liste öğesinde uygulanmaz.
    cursorAt(ed, "Liste");
    expect(dropCapTarget(ed.state)).toBeNull();
    expect(setDropCap(ed, { mode: "drop", lines: 3, distance: 0, font: null })).toBe(false);
    expect(normalizeDropCap({ mode: "yan" })).toBeNull();
    expect(firstGrapheme("  Şimdi")).toBe("  Ş");
    expect(firstGrapheme("école")).toBe("é");
  });

  it("HTML'e yazılır ve geri okunur (kopyala-yapıştır)", () => {
    const ed = make(doc({ type: "paragraph", attrs: { dropCap: { mode: "drop", lines: 4, distance: 3, font: "Georgia" } }, content: [{ type: "text", text: "Harf" }] }));
    const html = ed.getHTML();
    expect(html).toContain('data-dropcap="drop"');
    expect(html).toMatch(/--dc-lines: ?4/);
    const ed2 = make({ type: "doc", content: [{ type: "section", content: [p("x")] }] });
    ed2.commands.setContent(html);
    expect(ed2.state.doc.firstChild!.firstChild!.attrs.dropCap).toEqual({ mode: "drop", lines: 4, distance: 3, font: "Georgia" });
  });
});

describe("Sesli Oku", () => {
  it("okunacak birimler: boş paragraflar atlanır, silinmiş metin okunmaz, konum eşlemesi doğru", async () => {
    const { readUnits, wordEnd, sortVoices } = await import("../readAloud");
    const ed = make(
      doc(
        { type: "paragraph", content: [{ type: "text", text: "Bir " }, { type: "text", text: "silinen ", marks: [{ type: "deletion", attrs: { author: "A", date: "" } }] }, { type: "text", text: "iki" }] },
        p(""),
        p("Üç"),
      ),
    );
    const units = readUnits(ed.state.doc);
    expect(units.map((u) => u.text)).toEqual(["Bir iki", "Üç"]);
    const u = units[0];
    expect(ed.state.doc.textBetween(u.map[4], u.map[6] + 1)).toBe("iki");
    expect(wordEnd("Merhaba dünya", 8)).toBe(13);
    expect(wordEnd("İstanbul'da güzel", 0)).toBe(11);
    const v = (name: string, lang: string) => ({ name, lang }) as SpeechSynthesisVoice;
    expect(sortVoices([v("Zed", "en-US"), v("Yelda", "tr-TR"), v("Anna", "de-DE")]).map((x) => x.name)).toEqual(["Yelda", "Anna", "Zed"]);
  });

  it("oynatıcı paragrafları sırayla okur, sözcüğü vurgular, sonunda vurguyu kaldırır", async () => {
    const { ReadAloudPlayer, ReadAloudHighlight, readAloudKey } = await import("../readAloud");
    const spoken: { text: string; u: { onboundary?: (e: unknown) => void; onend?: () => void } }[] = [];
    const fake = { speak: (u: { text: string }) => spoken.push({ text: u.text, u: u as never }), cancel: () => {}, pause: () => {}, resume: () => {}, getVoices: () => [] };
    Object.defineProperty(window, "speechSynthesis", { value: fake, configurable: true });
    (window as unknown as { SpeechSynthesisUtterance: unknown }).SpeechSynthesisUtterance = class {
      constructor(public text: string) {}
    };
    const ed = (editor = new Editor({ extensions: [...schemaExtensions(), ReadAloudHighlight], content: normalizeDoc(doc(p("Birinci cümle."), p("İkinci cümle."))) }));
    const player = new ReadAloudPlayer(ed.view);
    player.start(0);
    expect(spoken.map((s) => s.text)).toEqual(["Birinci cümle."]);
    spoken[0].u.onboundary?.({ name: "word", charIndex: 8, charLength: 6 });
    const hl = readAloudKey.getState(ed.state)!.hl;
    expect(ed.state.doc.textBetween(hl.word!.from, hl.word!.to)).toBe("cümle.");
    spoken[0].u.onend?.();
    expect(spoken.map((s) => s.text)).toEqual(["Birinci cümle.", "İkinci cümle."]);
    expect(player.state).toMatchObject({ status: "playing", index: 1, count: 2 });
    spoken[1].u.onend?.();
    expect(player.state.status).toBe("idle");
    expect(readAloudKey.getState(ed.state)!.hl).toEqual({ para: null, word: null });
  });
});
