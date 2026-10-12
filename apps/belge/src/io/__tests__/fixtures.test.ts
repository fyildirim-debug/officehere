// @vitest-environment jsdom
// LibreOffice'in ürettiği DOCX dosyalarını okuma testleri (dış üretici uyumluluğu).
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import type { JSONContent } from "@tiptap/core";
import { readDocx, validateContent } from "../index";

const load = (f: string) => readDocx(readFileSync(join(__dirname, "../__fixtures__", f)));

function all(n: JSONContent, type: string, out: JSONContent[] = []): JSONContent[] {
  if (n.type === type) out.push(n);
  for (const c of n.content ?? []) all(c, type, out);
  return out;
}

const text = (n: JSONContent): string => (n.type === "text" ? (n.text ?? "") : (n.content ?? []).map(text).join(""));

describe("LibreOffice DOCX okuma", () => {
  it("HTML'den dönüştürülmüş belge: başlıklar, listeler, tablo, bağlantı", async () => {
    const d = await load("libreoffice-html.docx");
    validateContent(d.content);
    const headings = all(d.content, "heading");
    expect(headings.map((h) => [h.attrs?.level, text(h)])).toEqual([
      [1, "Ana Başlık Bir"],
      [2, "İkinci Düzey"],
      [3, "Üçüncü Düzey"],
    ]);
    const bullets = all(d.content, "bulletList");
    expect(bullets.length).toBe(2);
    expect(all(bullets[0], "listItem").length).toBe(3);
    const ordered = all(d.content, "orderedList");
    expect(ordered.length).toBe(1);
    expect(all(ordered[0], "listItem").map(text)).toEqual(["Bir", "İki", "Üç"]);
    const table = all(d.content, "table")[0];
    const rows = table.content!;
    expect(rows.length).toBe(3);
    expect(rows[2].content![0].attrs?.colspan).toBe(2);
    const link = JSON.stringify(d.content).includes('"href":"https://ornek.com/"');
    expect(link).toBe(true);
    // Türkçe karakterler korunur.
    expect(JSON.stringify(d.content)).toContain("İĞÜŞÖÇ ığüşöç");
    // Kaynak yazı tipi farkları doğrudan biçim olarak korunur.
    expect(JSON.stringify(d.content)).toContain("Liberation Serif");
  });

  it("LibreOffice ile yeniden kaydedilmiş OfficeHere raporu", async () => {
    const d = await load("libreoffice-resaved-report.docx");
    validateContent(d.content);
    const paras = all(d.content, "paragraph");
    expect(paras.find((p) => p.attrs?.styleId === "Title")).toBeTruthy();
    expect(paras.find((p) => p.attrs?.styleId === "Subtitle")).toBeTruthy();
    expect(all(d.content, "heading").length).toBeGreaterThanOrEqual(5);
    const table = all(d.content, "table")[0];
    expect(table.attrs?.tableStyle).toBe("grid-blue");
    expect(table.content!.length).toBe(4);
    expect(all(d.content, "orderedList")[0].content!.length).toBe(3);
    // Alt bilgideki sayfa numarası alanları
    expect(JSON.stringify(d.settings.footer)).toContain("pageField");
  });
});
