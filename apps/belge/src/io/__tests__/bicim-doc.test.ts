// @vitest-environment jsdom
// Eski ikili .doc (97-2003): LibreOffice örneğini okuma ve şifreli/eski/bozuk dosyaların reddi.
import { describe, expect, it } from "vitest";
import { readDoc, validateContent } from "../index";
import { all, fixture, markAttr, marksOf, text } from "./bicim-yardimci";

/** WordDocument akışının başını (FIB, 512 bayt sınırında A5EC imzası) dosyada bulur. */
function fibOffset(b: Uint8Array): number {
  for (let o = 512; o + 4 < b.length; o += 512) if (b[o] === 0xec && b[o + 1] === 0xa5) return o;
  throw new Error("FIB bulunamadı");
}

describe(".doc okuma (LibreOffice örneği)", () => {
  const d = readDoc(fixture("lo-ornek.doc"));
  const c = d.content;

  it("metin, paragraflar ve Türkçe karakterler", () => {
    validateContent(c);
    expect(text(c)).toContain("Türkçe: İĞÜŞÖÇ ığüşöç.");
    expect(text(c)).toContain("sekme\tsonrası");
    expect(text(c)).toContain("üç  boşluk");
    expect(all(c, "hardBreak").length).toBeGreaterThanOrEqual(1);
  });

  it("başlıklar (stil sayfası) ve karakter biçimleri", () => {
    expect(all(c, "heading").map((h) => [h.attrs?.level, text(h)])).toEqual([
      [1, "Giriş ve Amaç"],
      [2, "Tablo ve Resim"],
    ]);
    expect(marksOf(c, "kalın")).toContain("bold");
    expect(marksOf(c, "italik")).toContain("italic");
    expect(marksOf(c, "altı çizili")).toContain("underline");
    expect(marksOf(c, "üstü çizili")).toContain("strike");
    expect(marksOf(c, "2")).toContain("superscript");
    expect(markAttr(c, "kırmızı", "textStyle")?.color).toBe("#C00000");
    expect(markAttr(c, "büyük yazı tipi", "textStyle")).toMatchObject({ fontFamily: "Liberation Sans", fontSize: "14pt" });
    expect(markAttr(c, "ornek.com.tr", "link")?.href).toBe("https://ornek.com.tr/");
  });

  it("listeler, tablo (birleştirme ve dolgu), resim, dipnot", () => {
    expect(all(all(c, "orderedList")[0], "listItem").map(text)).toEqual(["Elma", "Armut", "Çilek"]);
    const bl = all(c, "bulletList")[0];
    expect(all(bl, "listItem").map((li) => text(li.content![0]))).toEqual(["Birinci madde", "İkinci madde", "İç madde"]);
    const table = all(c, "table")[0];
    expect(table.content!.map((r) => r.content!.map(text))).toEqual([["Ad", "Şehir", "Not"], ["Ayşe", "İzmir", "Dikey birleşik"], ["Yatay birleşik hücre"]]);
    expect(table.content![1].content![2].attrs?.rowspan).toBe(2);
    expect(table.content![2].content![0].attrs?.colspan).toBe(2);
    expect(table.content![0].content![0].attrs?.background).toBe("#DAE9F7");
    const img = all(c, "image")[0];
    expect(img.attrs?.src).toMatch(/^data:image\/png;base64,/);
    expect(img.attrs).toMatchObject({ width: 61, height: 45 });
    expect(all(c, "note")[0].attrs).toMatchObject({ kind: "footnote", text: "Dipnot açıklaması burada." });
  });

  it("hizalama, girinti, sayfa sonu, üst/alt bilgi ve sayfa yapısı", () => {
    expect(all(c, "paragraph").find((x) => text(x) === "Sağa hizalı paragraf")?.attrs?.textAlign).toBe("right");
    expect(all(c, "paragraph").find((x) => text(x).startsWith("İki yana"))?.attrs).toMatchObject({ textAlign: "justify", indentLeft: 28.35, indentFirst: 14.15 });
    expect(all(c, "paragraph").find((x) => text(x) === "Yeni sayfadaki paragraf.")?.attrs?.pageBreakBefore).toBe(true);
    expect(text(d.settings.header!)).toBe("Üst bilgi — Sayfa ");
    expect(JSON.stringify(d.settings.header)).toContain('"field":"PAGE"');
    expect(text(d.settings.footer!)).toBe("Alt bilgi: Gizlilik İçerir");
    expect(d.settings.page).toMatchObject({ width: 11906, height: 16838, orientation: "portrait" });
    expect(d.settings.page.margin.left).toBe(1417);
  });
});

describe(".doc reddi", () => {
  it("şifreli belge anlaşılır hatayla reddedilir", () => {
    const b = fixture("lo-ornek.doc").slice();
    const o = fibOffset(b);
    b[o + 0x0a] |= 0x01; // fEncrypted (0x0100)
    b[o + 0x0b] |= 0x01;
    expect(() => readDoc(b)).toThrow(/şifrelenmiş/);
  });

  it("97 öncesi sürüm (6.0/95) reddedilir", () => {
    const b = fixture("lo-ornek.doc").slice();
    const o = fibOffset(b);
    b[o + 2] = 0x65;
    b[o + 3] = 0x00;
    expect(() => readDoc(b)).toThrow(/97 öncesi/);
  });

  it("bileşik dosya olmayan ya da kesik veri reddedilir", () => {
    expect(() => readDoc(new TextEncoder().encode("bu bir doc değil"))).toThrow(/\.doc/);
    const b = fixture("lo-ornek.doc");
    expect(() => readDoc(b.subarray(0, 1536))).toThrow(/\.doc/);
  });

  it("arayüz iletilerinde ürün adı geçmez", () => {
    const msgs: string[] = [];
    for (const f of [() => readDoc(new Uint8Array(10)), () => readDoc(fixture("lo-ornek.doc").subarray(0, 1536))]) {
      try {
        f();
      } catch (e) {
        msgs.push((e as Error).message);
      }
    }
    expect(msgs.join(" ")).not.toMatch(/Word|Microsoft/);
  });
});
