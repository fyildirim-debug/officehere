// Kapak sayfaları, üst/alt bilgi ve sayfa numarası hazır parçaları (OfficeHere'in kendi tasarımları).
import type { JSONContent } from "@tiptap/core";

const t = (text: string, marks?: JSONContent["marks"]): JSONContent => ({ type: "text", text, ...(marks ? { marks } : {}) });
const ts = (attrs: Record<string, unknown>) => ({ type: "textStyle", attrs });
const p = (content: JSONContent[] | string, attrs: Record<string, unknown> = {}): JSONContent => ({
  type: "paragraph",
  attrs,
  content: typeof content === "string" ? (content ? [t(content)] : undefined) : content,
});

export interface CoverDef {
  id: string;
  label: string;
  build: (title: string, date: string, author: string) => JSONContent[];
}

export const COVER_PAGES: CoverDef[] = [
  {
    id: "band",
    label: "Şerit",
    build: (title, date, author) => [
      p("", { spaceBefore: 120 }),
      p([t(title, [ts({ fontSize: "40pt", color: "#FFFFFF", fontFamily: "Aptos Display" })])], { shading: "#156082", spaceBefore: 0, spaceAfter: 0, indentLeft: 12, indentRight: 12, lineHeight: 1.1 }),
      p([t("Belge alt başlığı", [ts({ fontSize: "16pt", color: "#FFFFFF" })])], { shading: "#156082", spaceAfter: 0, indentLeft: 12, indentRight: 12 }),
      p("", { spaceBefore: 260 }),
      p([t(author, [ts({ fontSize: "14pt", color: "#0F4761" })]), { type: "hardBreak" }, t(date, [ts({ fontSize: "12pt", color: "#595959" })])], { textAlign: "right" }),
    ],
  },
  {
    id: "simple",
    label: "Yalın",
    build: (title, date, author) => [
      p("", { spaceBefore: 200 }),
      p([t(title, [ts({ fontSize: "36pt", fontFamily: "Aptos Display" })])], { textAlign: "center", spaceAfter: 6 }),
      { type: "horizontalRule" },
      p([t("Belge alt başlığı", [ts({ fontSize: "16pt", color: "#595959" })])], { textAlign: "center" }),
      p("", { spaceBefore: 300 }),
      p([t(author, [ts({ fontSize: "12pt" })])], { textAlign: "center", spaceAfter: 0 }),
      p([t(date, [ts({ fontSize: "12pt", color: "#595959" })])], { textAlign: "center" }),
    ],
  },
  {
    id: "side",
    label: "Kenar Çubuğu",
    build: (title, date, author) => [
      p([t(date, [ts({ fontSize: "12pt", color: "#E97132" })])], { spaceBefore: 40 }),
      p("", { spaceBefore: 150 }),
      p([t(title, [ts({ fontSize: "44pt", color: "#0E2841", fontFamily: "Aptos Display" })])], {
        borders: { left: { style: "thick", size: 6, color: "#E97132", space: 12 } },
        indentLeft: 18,
        lineHeight: 1,
        spaceAfter: 4,
      }),
      p([t("Belge alt başlığı", [ts({ fontSize: "18pt", color: "#595959" })])], { borders: { left: { style: "thick", size: 6, color: "#E97132", space: 12 } }, indentLeft: 18 }),
      p("", { spaceBefore: 260 }),
      p([t(author, [ts({ fontSize: "12pt", color: "#0E2841" }), { type: "bold" }])]),
    ],
  },
];

export interface HfPreset {
  id: string;
  label: string;
  content: JSONContent;
}

const doc = (...content: JSONContent[]): JSONContent => ({ type: "doc", content });
const field = (f: "PAGE" | "NUMPAGES"): JSONContent => ({ type: "pageField", attrs: { field: f } });

export const HEADER_PRESETS: HfPreset[] = [
  { id: "blank", label: "Boş", content: doc(p("[Buraya yazın]")) },
  { id: "center", label: "Ortalanmış Başlık", content: doc(p([t("[Belge başlığı]", [ts({ color: "#0F4761" })])], { textAlign: "center" })) },
  {
    id: "lined",
    label: "Çizgili",
    content: doc(p([t("[Belge başlığı]", [ts({ fontSize: "12pt", color: "#0F4761" })])], { borders: { bottom: { style: "single", size: 1, color: "#0F4761", space: 2 } } })),
  },
  { id: "right", label: "Sağa Hizalı", content: doc(p("[Yazar adı] | [Tarih]", { textAlign: "right" })) },
];

export const FOOTER_PRESETS: HfPreset[] = [
  { id: "blank", label: "Boş", content: doc(p("[Buraya yazın]")) },
  { id: "pageCenter", label: "Sayfa Numarası (Ortada)", content: doc(p([field("PAGE")], { textAlign: "center" })) },
  { id: "pageOf", label: "Sayfa X / Y", content: doc(p([t("Sayfa "), field("PAGE"), t(" / "), field("NUMPAGES")], { textAlign: "right" })) },
  {
    id: "lined",
    label: "Çizgili",
    content: doc(p([t("[Şirket adı]"), t("   •   "), field("PAGE")], { borders: { top: { style: "single", size: 0.5, color: "#7F7F7F", space: 2 } }, textAlign: "center" })),
  },
];

export function pageNumberDoc(align: "left" | "center" | "right", withTotal = false): JSONContent {
  return doc(p(withTotal ? [t("Sayfa "), field("PAGE"), t(" / "), field("NUMPAGES")] : [field("PAGE")], { textAlign: align }));
}

/** Üst/alt bilgiden sayfa numarası alanlarını kaldırır. */
export function stripPageFields(json: JSONContent | null): JSONContent | null {
  if (!json) return json;
  const walk = (n: JSONContent): JSONContent => ({ ...n, content: n.content?.filter((c) => c.type !== "pageField").map(walk) });
  return walk(json);
}

export const COMMON_SYMBOLS = ["€", "£", "¥", "₺", "©", "®", "™", "±", "≠", "≤", "≥", "÷", "×", "∞", "µ", "α", "β", "π", "Ω", "∑", "°", "§", "¶", "•", "…", "—", "«", "»", "√", "✓"];
