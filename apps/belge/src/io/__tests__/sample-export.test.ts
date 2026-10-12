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
  });
});
