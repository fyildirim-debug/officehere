// @vitest-environment jsdom
// 4. tur: Birden Çok Sayfa satır planı, Türkçe arama (i/İ, ı/I), Yinele (F4), Otomatik Düzelt (bitişik
// simge), alt/üst simge dışlaması, kaydırmalı resim yerleşimi, bölüm üst/alt bilgisi bağlantısı.
import { afterEach, describe, expect, it } from "vitest";
import { Editor, type JSONContent } from "@tiptap/core";
import { TextSelection } from "@tiptap/pm/state";
import { normalizeDoc, defaultSettings } from "../../model/document";
import { schemaExtensions } from "../extensions";
import { Search, findMatches, foldTr, getSearch, replaceAll, setSearch } from "../search";
import { evenTileRows, planTileRows, tileCell, tileColumns } from "../tileRows";
import { Repeat, repeatOrRedo } from "../repeat";
import { correctBefore, defaultAutoCorrect } from "../autocorrectRules";
import { anchoredPlacement, wrapStyle } from "../imageWrap";
import { hfContent, hfLinked, hfOwner, withHfContent } from "../sectionHf";
import { fieldNowText } from "../nodes";

let editor: Editor | null = null;
afterEach(() => {
  editor?.destroy();
  editor = null;
});

function make(content: object, extra: object[] = []) {
  editor = new Editor({ extensions: [...schemaExtensions(), Search, ...(extra as never[])], content: normalizeDoc(content as JSONContent) });
  return editor;
}
const para = (text: string) => ({ type: "paragraph", content: text ? [{ type: "text", text }] : undefined });
const doc = (...c: object[]) => ({ type: "doc", content: c });

describe("Birden Çok Sayfa: satır planı", () => {
  it("her sınır satır başı olabiliyorsa eşit satırlar", () => {
    expect(planTileRows(7, 3, () => true)).toEqual([0, 3, 6]);
    expect(tileColumns([0, 3, 6], 7)).toBe(3);
  });
  it("satır başı olamayan sınırda satır erken biter, gerekirse uzar", () => {
    // 3. ve 4. sayfa tablo satırıyla başlıyor: satır 2. sayfadan sonra biter.
    const can = (k: number) => k !== 3 && k !== 4;
    expect(planTileRows(8, 3, can)).toEqual([0, 2, 5]);
    // Uzun çok sütunlu bölüm (1..5 satır başı olamaz): satır uzar.
    const can2 = (k: number) => k === 6;
    const rows = planTileRows(8, 3, can2);
    expect(rows).toEqual([0, 6]);
    expect(tileColumns(rows, 8)).toBe(6);
  });
  it("tek sayfa sınırı yoksa tek satır, hücre konumları", () => {
    expect(planTileRows(5, 2, () => false)).toEqual([0]);
    expect(tileCell([0, 2, 5], 4)).toEqual({ row: 1, col: 2 });
    expect(tileCell([0, 2, 5], 5)).toEqual({ row: 2, col: 0 });
    expect(evenTileRows(5, 2)).toEqual([0, 2, 4]);
  });
});

describe("Bul/Değiştir: Türkçe i/İ ve ı/I", () => {
  const text = "İstanbul ve istanbul; ISI ve ısı ve isi; Iğdır ve ığdır; iğdır.";
  const count = (q: string) => findMatches(make(doc(para(text))).state.doc, { query: q, caseSensitive: false, wholeWord: false }).length;
  it("katlama: İ→i, I→ı; i ile ı ayrı", () => {
    expect(foldTr("İSTANBUL")).toBe("istanbul");
    expect(foldTr("ISI")).toBe("ısı");
    expect(foldTr("isi")).not.toBe(foldTr("ısı"));
  });
  it("Word gibi eşleşme sayıları", () => {
    expect(count("istanbul")).toBe(2);
    expect(count("İSTANBUL")).toBe(2);
    expect(count("iğdır")).toBe(1);
    expect(count("ığdır")).toBe(2);
    expect(count("ısı")).toBe(2);
    expect(count("isi")).toBe(1);
  });
  it("sayaç, vurgu ve Tümünü Değiştir aynı eşleşmeleri kullanır", () => {
    const ed = make(doc(para(text)));
    setSearch(ed.view, { query: "istanbul", caseSensitive: false, wholeWord: false });
    const st = getSearch(ed.state)!;
    expect(st.matches.length).toBe(2);
    expect(ed.view.dom.querySelectorAll(".bl-hit").length).toBe(2);
    expect(replaceAll(ed.view, "ankara")).toBe(2);
    expect(ed.state.doc.textContent).toContain("Ankara ve ankara");
  });
});

describe("Yinele (F4 / Ctrl+Y)", () => {
  const sel = (ed: Editor, from: number, to: number) => ed.view.dispatch(ed.state.tr.setSelection(TextSelection.create(ed.state.doc, from, to)));
  const posOf = (ed: Editor, s: string) => {
    let at = -1;
    ed.state.doc.descendants((n, p) => {
      if (at < 0 && n.isText && n.text!.includes(s)) at = p + n.text!.indexOf(s);
    });
    return at;
  };
  it("karakter biçimini başka seçime yeniden uygular", () => {
    const ed = make(doc(para("birinci satır"), para("ikinci satır")), [Repeat]);
    const a = posOf(ed, "birinci");
    sel(ed, a, a + 7);
    ed.commands.toggleBold();
    const b = posOf(ed, "ikinci");
    sel(ed, b, b + 6);
    expect(repeatOrRedo(ed)).toBe(true);
    const html = ed.getHTML();
    expect(html).toContain("<strong>birinci</strong>");
    expect(html).toContain("<strong>ikinci</strong>");
  });
  it("paragraf biçimini (hizalama) yeniden uygular; geri alınmış adım önce yinelenir", () => {
    const ed = make(doc(para("bir"), para("iki")), [Repeat]);
    sel(ed, 3, 3);
    ed.commands.setTextAlign("center");
    sel(ed, posOf(ed, "iki") + 1, posOf(ed, "iki") + 1);
    repeatOrRedo(ed);
    const aligns: unknown[] = [];
    ed.state.doc.descendants((n) => (n.isTextblock ? (aligns.push(n.attrs.textAlign), false) : true));
    expect(aligns).toEqual(["center", "center"]);
    ed.commands.undo();
    expect(repeatOrRedo(ed)).toBe(true);
    const again: unknown[] = [];
    ed.state.doc.descendants((n) => (n.isTextblock ? (again.push(n.attrs.textAlign), false) : true));
    expect(again).toEqual(["center", "center"]);
  });
  it("yazılan metni yineler", () => {
    const ed = make(doc(para("x")), [Repeat]);
    sel(ed, 3, 3);
    for (const ch of "abc") ed.view.dispatch(ed.state.tr.insertText(ch));
    repeatOrRedo(ed);
    expect(ed.state.doc.textContent).toBe("xabcabc");
  });
});

describe("Otomatik Düzelt ve simgeler", () => {
  it("sözcüğe bitişik simge girdileri", () => {
    const o = defaultAutoCorrect();
    expect(correctBefore("Marka(c)", o)?.text).toBe("Marka©");
    expect(correctBefore("Ürün(tm)", o)?.text).toBe("Ürün™");
    expect(correctBefore("Bekle...", o)?.text).toBe("Bekle…");
    expect(correctBefore("Marka(c).", o)?.text).toBe("Marka©.");
  });
  it("alt simge üst simgeyi kaldırır (birbirini dışlar)", () => {
    const ed = make(doc(para("x2")));
    ed.commands.selectAll();
    ed.commands.toggleSuperscript();
    ed.commands.toggleSubscript();
    const marks = new Set<string>();
    ed.state.doc.descendants((n) => (n.marks.forEach((m) => marks.add(m.type.name)), true));
    expect([...marks]).toEqual(["subscript"]);
  });
  it("tarih/saat alanı Türkçe biçimde", () => {
    const d = new Date(2026, 9, 12, 9, 5);
    expect(fieldNowText("DATE", d)).toBe("12.10.2026");
    expect(fieldNowText("TIME", d)).toBe("09:05");
    expect(fieldNowText("PAGE", d)).toBeNull();
  });
});

describe("Kaydırmalı resim yerleşimi", () => {
  it("bırakılan konuma göre yaslanma kenarı ve uzaklıklar", () => {
    const para = { left: 100, top: 200, width: 600 };
    expect(anchoredPlacement("square", { left: 150, top: 230, width: 100 }, para)).toEqual({ floatSide: "left", posX: 50, posY: 30 });
    expect(anchoredPlacement("square", { left: 560, top: 190, width: 100 }, para)).toEqual({ floatSide: "right", posX: 40, posY: 0 });
    expect(anchoredPlacement("topBottom", { left: 350, top: 210, width: 100 }, para).floatSide).toBe("center");
  });
  it("kare kaydırmada uzaklık kenar boşluğuna, üst ve altta yaslamaya dönüşür", () => {
    expect(wrapStyle({ wrap: "square", floatSide: "left", posX: 40, posY: 12 })).toMatchObject({ float: "left", margin: "12px 12px 6px 40px" });
    expect(wrapStyle({ wrap: "square", floatSide: "right", posX: 30, posY: 0 })).toMatchObject({ float: "right", margin: "0px 30px 6px 12px" });
    expect(wrapStyle({ wrap: "topBottom", floatSide: "right", posX: 0, posY: 0 })).toMatchObject({ justifyContent: "flex-end" });
  });
});

describe("Bölüm üst/alt bilgisi: Öncekine Bağla", () => {
  const H = (t: string): JSONContent => ({ type: "doc", content: [{ type: "paragraph", content: [{ type: "text", text: t }] }] });
  it("bağlı bölüm öncekini gösterir, bağı kaldırılan kendi içeriğini", () => {
    const s = { ...defaultSettings(), header: H("Genel"), sectionHf: { s2: { header: H("Bölüm 3") } } };
    const ids = [null, null, "s2", null];
    expect(hfOwner(ids, s, 1, "header")).toBe(-1);
    expect(hfContent(ids, s, 1, "header")).toEqual(H("Genel"));
    expect(hfContent(ids, s, 3, "header")).toEqual(H("Bölüm 3"));
    expect(hfLinked(ids, s, 3, "header")).toBe(true);
    expect(hfLinked(ids, s, 2, "header")).toBe(false);
    expect(hfLinked(ids, s, 0, "header")).toBe(false);
    // Bağlı bölümde düzenleme sahibine yazılır.
    const s2 = withHfContent(s, ids, 3, "header", H("Yeni"));
    expect(s2.sectionHf!.s2.header).toEqual(H("Yeni"));
    expect(s2.header).toEqual(H("Genel"));
    // Bölüm sonu silinince ilk bölüme geçen kendi üst bilgisi kullanılır.
    expect(hfContent(["s2", null], s, 0, "header")).toEqual(H("Bölüm 3"));
    expect(hfContent(["s2", null], s, 1, "header")).toEqual(H("Bölüm 3"));
  });
});
