// @vitest-environment jsdom
// Markdown: CommonMark çekirdeği + GFM (tablo, üstü çizili, görev listesi, dipnot) okuma/yazma ve gidiş-dönüş.
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { readMarkdown, validateContent, writeMarkdown } from "../index";
import { FIX, all, docOf, markAttr, marksOf, norm, p, richDoc, t, text } from "./bicim-yardimci";

const md = (s: string, opts = {}) => readMarkdown(s, opts);

describe("Markdown okuma", () => {
  it("örnek dosya: başlıklar, vurgular, listeler, alıntı, kod, tablo, dipnot, resimler", async () => {
    const d = await md(readFileSync(join(FIX, "ornek.md"), "utf8"));
    validateContent(d.content);
    const c = d.content;
    expect(d.settings.properties.title).toBe("Markdown Örneği");
    expect(all(c, "heading").map((h) => [h.attrs?.level, text(h)])).toEqual([
      [1, "Ana Başlık İĞÜŞÖÇ"],
      [2, "Alt Çizgili Başlık"],
    ]);
    expect(marksOf(c, "kalın")).toEqual(["bold"]);
    expect(marksOf(c, "italik")).toEqual(["italic"]);
    expect(marksOf(c, "ikisi birden").sort()).toEqual(["bold", "italic"]);
    expect(markAttr(c, "kod parçası", "textStyle")?.fontFamily).toBe("Courier New");
    expect(marksOf(c, "üstü çizili")).toEqual(["strike"]);
    expect(markAttr(c, "bir bağlantı", "link")?.href).toBe("https://ornek.com.tr/");
    expect(markAttr(c, "OfficeHere", "link")?.href).toBe("https://github.com/fyildirim-debug/officehere");
    expect(markAttr(c, "https://officehere.dev", "link")?.href).toBe("https://officehere.dev");
    expect(all(c, "hardBreak")).toHaveLength(1);
    const bl = all(c, "bulletList")[0];
    expect(all(bl, "listItem").map((li) => text(li.content![0]))).toEqual(["Birinci madde", "İkinci madde", "İç madde", "İç madde kalın", "☒ Tamamlanan görev", "☐ Bekleyen görev"]);
    expect(all(all(c, "orderedList")[0], "listItem").map(text)).toEqual(["Bir", "İki", "Üç"]);
    expect(all(c, "paragraph").find((x) => x.attrs?.styleId === "Quote") && text(all(c, "paragraph").find((x) => x.attrs?.styleId === "Quote")!)).toBe("Alıntılanan paragraf ikinci satır.");
    const code = all(c, "paragraph").filter((x) => x.attrs?.shading === "#F2F2F2");
    expect(code.map(text)).toEqual(['const selam = "Merhaba";', "console.log(selam);"]);
    const table = all(c, "table")[0];
    expect(table.content!.map((r) => r.content!.map(text))).toEqual([
      ["Ad", "Şehir", "Puan"],
      ["Ayşe", "İzmir", "90"],
      ["Ömer", "Çorum", "85"],
    ]);
    expect(table.content![0].content![1].content![0].attrs?.textAlign).toBe("center");
    expect(table.content![0].content![2].content![0].attrs?.textAlign).toBe("right");
    const note = all(c, "note")[0];
    expect(note.attrs?.text).toBe("Dipnot metni vurgulu.");
    const imgs = all(c, "image");
    expect(imgs).toHaveLength(1);
    expect(imgs[0].attrs).toMatchObject({ width: 16, height: 12, alt: "Kırmızı kare" });
    // Çözülemeyen göreli resim yolu alternatif metne düşer.
    expect(text(c)).toContain("Bulunamayan resim");
    expect(all(c, "horizontalRule")).toHaveLength(1);
  });

  it("göreli resim yolu çözücüyle okunur", async () => {
    const png = Uint8Array.from(atob("iVBORw0KGgoAAAANSUhEUgAAABAAAAAMCAIAAADkharWAAAAFUlEQVR4nGM4oaFBEmIY1TCqATsCALRr0gH04WUkAAAAAElFTkSuQmCC"), (ch) => ch.charCodeAt(0));
    const d = await md("![a](resimler/k%C4%B1rm%C4%B1z%C4%B1.png)", { resolveImage: async (path: string) => (path === "resimler/kırmızı.png" ? png : null) });
    expect(all(d.content, "image")[0].attrs?.src).toMatch(/^data:image\/png;base64,/);
  });

  it("CommonMark ayrıntıları: kaçışlar, varlıklar, iç içe vurgu, kod parçası, girintili kod", async () => {
    const d = await md("\\*yıldız\\* &amp; &#305; *a **b** c* `` a`b `` __alt__\n\n    girintili kod\n");
    const c = d.content;
    expect(text(all(c, "paragraph")[0])).toBe("*yıldız* & ı a b c a`b alt");
    expect(marksOf(c, "b").sort()).toEqual(["bold", "italic"]);
    expect(markAttr(c, "a`b", "textStyle")?.fontFamily).toBe("Courier New");
    expect(marksOf(c, "alt")).toEqual(["bold"]);
    expect(text(all(c, "paragraph")[1])).toBe("girintili kod");
  });

  it("Windows-1254 kodlu dosya da okunur (Türkçe)", async () => {
    const { openDocument } = await import("../index");
    const bytes = new Uint8Array([0x23, 0x20, 0xdd, 0xfe, 0x0a]);
    const r = await openDocument("eski.md", bytes.buffer);
    expect(r.format).toBe("md");
    expect(text(all(r.doc.content, "heading")[0])).toBe("İş");
  });
});

describe("Markdown yazma", () => {
  it("biçimler, listeler, tablo, kod bloğu ve dipnot", () => {
    const d = docOf(
      { type: "heading", attrs: { level: 2 }, content: [t("Başlık")] },
      p([t("a "), t("kalın", [{ type: "bold" }]), t(" "), t("eğik", [{ type: "italic" }]), t(" "), t("kod", [{ type: "textStyle", attrs: { fontFamily: "Consolas" } }]), t(" "), t("bağ", [{ type: "link", attrs: { href: "https://a.b/c d" } }]), { type: "note", attrs: { kind: "footnote", text: "Not", id: "x" } }]),
      p("Alıntı", { styleId: "Quote" }),
      { type: "orderedList", attrs: { start: 3 }, content: [{ type: "listItem", content: [p("üç"), { type: "bulletList", content: [{ type: "listItem", content: [p("☒ iç görev")] }] }] }] },
      { type: "table", content: [{ type: "tableRow", content: [{ type: "tableCell", content: [p("A|B", { textAlign: "center" })] }, { type: "tableCell", content: [p("C")] }] }, { type: "tableRow", content: [{ type: "tableCell", content: [p("1")] }, { type: "tableCell", content: [p("2")] }] }] },
      p([t("x = 1;", [{ type: "textStyle", attrs: { fontFamily: "Courier New" } }])], { shading: "#F2F2F2" }),
      p("* yıldızla başlayan # metin"),
    );
    const out = writeMarkdown(d);
    expect(out).toContain("## Başlık");
    expect(out).toContain("a **kalın** *eğik* `kod` [bağ](<https://a.b/c d>)[^1]");
    expect(out).toContain("> Alıntı");
    expect(out).toContain("3. üç\n   - [x] iç görev");
    expect(out).toContain("| A\\|B | C |\n|:---:|---|\n| 1 | 2 |");
    expect(out).toContain("```\nx = 1;\n```");
    expect(out).toContain("\\* yıldızla başlayan # metin");
    expect(out).toContain("[^1]: Not");
  });

  it("gidiş-dönüş: Belge → Markdown → Belge (Markdown'un taşıyabildiği biçimler)", async () => {
    const src = richDoc();
    const out = writeMarkdown(src);
    const back = await md(out);
    validateContent(back.content);
    // Markdown'da renk, yazı tipi, hizalama, hücre birleştirme ve sayfa sonu yoktur: metin ve temel biçimler korunur.
    expect(text(back.content).replace(/\s+/g, " ")).toBe(text(src.content).replace(/\s+/g, " "));
    for (const [needle, mark] of [["kalın", "bold"], ["italik", "italic"], ["altı çizili", "underline"], ["üstü çizili", "strike"], ["bağlantı", "link"]]) expect(marksOf(back.content, needle)).toContain(mark);
    expect(all(back.content, "heading").map(text)).toEqual(["Rapor Başlığı İĞÜŞÖÇ", "Tablo"]);
    expect(all(back.content, "listItem").length).toBe(5);
    expect(all(back.content, "image")).toHaveLength(1);
    expect(all(back.content, "note")[0].attrs?.text).toBe("Dipnot metni");
    // İkinci tur değişmez (kararlı çıktı).
    expect(writeMarkdown(back)).toBe(out);
  });

  it("basit belgede tam gidiş-dönüş", async () => {
    const src = docOf({ type: "heading", attrs: { level: 1 }, content: [t("Başlık")] }, p([t("Merhaba "), t("dünya", [{ type: "bold" }])]), { type: "bulletList", content: [{ type: "listItem", content: [p("bir")] }, { type: "listItem", content: [p("iki")] }] });
    const back = await md(writeMarkdown(src));
    expect(norm(back.content)).toEqual(norm(src.content));
  });
});
