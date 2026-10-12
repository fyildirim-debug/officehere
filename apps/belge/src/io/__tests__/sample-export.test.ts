// @vitest-environment jsdom
// BELGE_SAMPLE_OUT ayarlıysa örnek DOCX dosyaları üretir (LibreOffice / python-docx ile dış doğrulama için).
import { describe, it } from "vitest";
import { mkdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { writeDocx } from "../index";
import { TEMPLATES } from "../../templates/templates";
import { defaultSettings } from "../../model/document";

const OUT = process.env.BELGE_SAMPLE_OUT;
const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAADCAIAAAA7ljmRAAAAEElEQVR4nGM4IScHRww4OQD1xwwxB+utmQAAAABJRU5ErkJggg==";

describe.skipIf(!OUT)("örnek dışa aktarım", () => {
  it("şablonları ve karma belgeyi yazar", async () => {
    mkdirSync(OUT!, { recursive: true });
    for (const t of TEMPLATES) writeFileSync(join(OUT!, `${t.id}.docx`), await writeDocx(t.build()));
    const s = defaultSettings();
    s.watermark = { text: "TASLAK", color: "#C0C0C0", layout: "diagonal", font: "Aptos", semitransparent: true };
    s.header = { type: "doc", content: [{ type: "paragraph", attrs: { textAlign: "center" }, content: [{ type: "text", text: "OfficeHere Üst Bilgi" }] }] };
    s.footer = { type: "doc", content: [{ type: "paragraph", attrs: { textAlign: "center" }, content: [{ type: "text", text: "Sayfa " }, { type: "pageField", attrs: { field: "PAGE" } }] }] };
    s.pageBorder = { style: "single", color: "#156082", width: 1 };
    s.comments = { c0: { id: "c0", author: "Test", initials: "T", date: "2026-10-11T10:00:00Z", text: "Açıklama metni" } };
    const doc = {
      settings: s,
      content: {
        type: "doc",
        content: [
          { type: "toc", attrs: { title: "İçindekiler", entries: [{ level: 1, text: "Birinci Bölüm", page: 1, pos: 0 }] } },
          { type: "heading", attrs: { level: 1 }, content: [{ type: "text", text: "Birinci Bölüm" }] },
          {
            type: "paragraph",
            content: [
              { type: "text", text: "Kalın ", marks: [{ type: "bold" }] },
              { type: "text", text: "açıklamalı", marks: [{ type: "comment", attrs: { id: "c0" } }] },
              { type: "text", text: " eklenen", marks: [{ type: "insertion", attrs: { author: "Test", date: "2026-10-11T10:00:00Z" } }] },
              { type: "text", text: " silinen", marks: [{ type: "deletion", attrs: { author: "Test", date: "2026-10-11T10:00:00Z" } }] },
              { type: "image", attrs: { src: PNG, width: 40, height: 30 } },
              { type: "text", text: " bağlantı", marks: [{ type: "link", attrs: { href: "https://ornek.com" } }] },
              { type: "text", text: " biçimi değişen", marks: [{ type: "bold" }, { type: "formatChange", attrs: { author: "Test", date: "2026-10-11T10:00:00Z", old: JSON.stringify([{ type: "italic" }]) } }] },
            ],
          },
          { type: "bulletList", content: [{ type: "listItem", content: [{ type: "paragraph", content: [{ type: "text", text: "Madde" }] }] }] },
          {
            type: "paragraph",
            content: [
              { type: "text", text: "Dipnotlu cümle" },
              { type: "note", attrs: { kind: "footnote", id: "f1", text: "Bu bir dipnottur." } },
              { type: "text", text: " ve son notlu cümle" },
              { type: "note", attrs: { kind: "endnote", id: "e1", text: "Bu bir son nottur." } },
            ],
          },
          { type: "pageBreak" },
          { type: "paragraph", content: [{ type: "text", text: "İkinci sayfa" }] },
        ],
      },
    };
    writeFileSync(join(OUT!, "karma.docx"), await writeDocx(doc));

    // Bölümler: dikey (ilk sayfa farklı), yatay, iki sütunlu sürekli bölüm, tek/çift üst bilgiler.
    const s2 = defaultSettings();
    s2.page = { ...s2.page, titlePage: true };
    s2.header = { type: "doc", content: [{ type: "paragraph", attrs: { textAlign: "right" }, content: [{ type: "text", text: "Tek sayfa üst bilgisi" }] }] };
    s2.headerFirst = { type: "doc", content: [{ type: "paragraph", attrs: { textAlign: "center" }, content: [{ type: "text", text: "İlk sayfa üst bilgisi" }] }] };
    s2.headerEven = { type: "doc", content: [{ type: "paragraph", content: [{ type: "text", text: "Çift sayfa üst bilgisi" }] }] };
    s2.footer = { type: "doc", content: [{ type: "paragraph", attrs: { textAlign: "center" }, content: [{ type: "text", text: "Sayfa " }, { type: "pageField", attrs: { field: "PAGE" } }] }] };
    s2.oddEven = true;
    const lorem = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ";
    const para = (t: string) => ({ type: "paragraph", content: [{ type: "text", text: t }] });
    const land = { ...s2.page, width: 16838, height: 11906, orientation: "landscape" as const, titlePage: false };
    const cols = { ...s2.page, columns: 2, columnSpace: 708, columnLine: true };
    const doc2 = {
      settings: s2,
      content: {
        type: "doc",
        content: [
          { type: "section", attrs: { kind: "nextPage", page: null }, content: [{ type: "heading", attrs: { level: 1 }, content: [{ type: "text", text: "Dikey bölüm" }] }, { type: "textBox", attrs: { x: 300, y: 10, w: 260, h: 120, border: "#C00000", borderWidth: 1.5, fill: "#DAE9F7" }, content: [{ type: "paragraph", content: [{ type: "text", text: "Metin kutusu: " }, { type: "text", text: "kalın", marks: [{ type: "bold" }] }] }] }, para("Birinci sayfa."), { type: "pageBreak" }, para("İkinci sayfa (çift).")] },
          { type: "section", attrs: { kind: "nextPage", page: land }, content: [{ type: "heading", attrs: { level: 1 }, content: [{ type: "text", text: "Yatay bölüm" }] }, para(lorem.repeat(4))] },
          { type: "section", attrs: { kind: "continuous", page: { ...land, columns: 2, columnSpace: 708, columnLine: true } }, content: Array.from({ length: 6 }, (_, i) => para(`${i + 1}. ${lorem.repeat(2)}`)) },
          { type: "section", attrs: { kind: "nextPage", page: cols }, content: [para("Sütun bir."), { type: "columnBreak" }, para("Sütun iki.")] },
        ],
      },
    };
    writeFileSync(join(OUT!, "bolumler.docx"), await writeDocx(doc2));
  });
});
