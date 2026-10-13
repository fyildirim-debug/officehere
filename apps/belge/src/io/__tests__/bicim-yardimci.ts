// Biçim testlerinin ortak yardımcıları: örnek dosya okuma, düğüm arama, karşılaştırma için normalleştirme.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import type { JSONContent } from "@tiptap/core";
import { Node as PMNode } from "@tiptap/pm/model";
import { belgeSchema } from "../index";
import { defaultSettings, normalizeDoc, type BelgeDocument } from "../../model/document";

export const FIX = join(__dirname, "../__fixtures__/bicimler");

export function fixture(name: string): Uint8Array {
  return new Uint8Array(readFileSync(join(FIX, name)));
}

export function fixtureBuffer(name: string): ArrayBuffer {
  const b = readFileSync(join(FIX, name));
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer;
}

/** Ağaçta türü verilen tüm düğümler. */
export function all(n: JSONContent, type: string, out: JSONContent[] = []): JSONContent[] {
  if (n.type === type) out.push(n);
  for (const c of n.content ?? []) all(c, type, out);
  return out;
}

export const text = (n: JSONContent): string => (n.type === "text" ? (n.text ?? "") : (n.content ?? []).map(text).join(""));

/** Metni içeren ilk metin düğümünün işaret türleri. */
export function marksOf(doc: JSONContent, needle: string): string[] {
  const t = all(doc, "text").find((x) => x.text?.includes(needle));
  return (t?.marks ?? []).map((m) => m.type);
}

export function markAttr(doc: JSONContent, needle: string, type: string): Record<string, unknown> | undefined {
  const t = all(doc, "text").find((x) => x.text?.includes(needle));
  return t?.marks?.find((m) => m.type === type)?.attrs;
}

/** Şema üzerinden normalleştirme (varsayılan öznitelikler, birleşen metinler); not kimlikleri karşılaştırma dışı. */
export function norm(json: JSONContent): JSONContent {
  const out = PMNode.fromJSON(belgeSchema(), normalizeDoc(json)).toJSON() as JSONContent;
  const walk = (n: JSONContent) => {
    if (n.type === "note" && n.attrs) n.attrs = { ...n.attrs, id: "" };
    if (n.type === "section" && n.attrs) n.attrs = { ...n.attrs, sid: null };
    n.content?.forEach(walk);
  };
  walk(out);
  return out;
}

export const t = (s: string, marks?: JSONContent["marks"]): JSONContent => ({ type: "text", text: s, ...(marks ? { marks } : {}) });
export const p = (content: JSONContent[] | string, attrs: Record<string, unknown> = {}): JSONContent => ({ type: "paragraph", attrs, content: typeof content === "string" ? [t(content)] : content });

export const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAMCAIAAADkharWAAAAFUlEQVR4nGM4oaFBEmIY1TCqATsCALRr0gH04WUkAAAAAElFTkSuQmCC";

export function docOf(...content: JSONContent[]): BelgeDocument {
  return { content: { type: "doc", content }, settings: defaultSettings() };
}

/** Biçimlerin ortak gidiş-dönüş belgesi: başlıklar, biçimler, listeler, tablo, resim, not, sayfa sonu. */
export function richDoc(): BelgeDocument {
  const d = docOf(
    { type: "heading", attrs: { level: 1 }, content: [t("Rapor Başlığı İĞÜŞÖÇ")] },
    p([
      t("Düz "),
      t("kalın", [{ type: "bold" }]),
      t(" "),
      t("italik", [{ type: "italic" }]),
      t(" "),
      t("altı çizili", [{ type: "underline", attrs: { type: "single", color: null } }]),
      t(" "),
      t("üstü çizili", [{ type: "strike" }]),
      t(" "),
      t("kırmızı Arial", [{ type: "textStyle", attrs: { color: "#C00000", fontFamily: "Arial", fontSize: "14pt" } }]),
      t(" "),
      t("bağlantı", [{ type: "link", attrs: { href: "https://ornek.com.tr/" } }]),
      t(" ığüşöç"),
      { type: "note", attrs: { kind: "footnote", text: "Dipnot metni", id: "n1" } },
    ]),
    p("Ortalanmış paragraf", { textAlign: "center" }),
    p([t("Satır"), { type: "hardBreak" }, t("sonu\tsekme")]),
    {
      type: "bulletList",
      content: [
        { type: "listItem", content: [p("Birinci"), { type: "bulletList", content: [{ type: "listItem", content: [p("İç madde")] }] }] },
        { type: "listItem", content: [p("İkinci")] },
      ],
    },
    { type: "orderedList", attrs: { start: 1 }, content: [{ type: "listItem", content: [p("Bir")] }, { type: "listItem", content: [p("İki")] }] },
    { type: "heading", attrs: { level: 2 }, content: [t("Tablo")] },
    {
      type: "table",
      content: [
        { type: "tableRow", content: [{ type: "tableCell", attrs: { colspan: 2, colwidth: [120, 120] }, content: [p("Birleşik")] }, { type: "tableCell", attrs: { rowspan: 2, colwidth: [100] }, content: [p("Dikey")] }] },
        { type: "tableRow", content: [{ type: "tableCell", attrs: { colwidth: [120] }, content: [p("A1")] }, { type: "tableCell", attrs: { colwidth: [120] }, content: [p("B1")] }] },
      ],
    },
    p([t("Resim: "), { type: "image", attrs: { src: PNG, width: 32, height: 24 } }]),
    { type: "pageBreak" },
    p("Yeni sayfa"),
  );
  d.settings.properties.title = "Gidiş Dönüş";
  return d;
}
