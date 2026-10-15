// Opak öğeler (geçiş katmanı, öğe düzeyi): OfficeHere'in düzenleyemediği .docx öğeleri (gömülü nesneler,
// denetimler, anlaşılmayan çizimler, alt belge, içe aktarılan içerik, bilinmeyen öğeler) özgün XML'leri ve
// ilişkileriyle düğüm özniteliklerinde taşınır. Belgede yer tutucu kart olarak görünür; .docx'e kaydederken
// özgün XML yazılır (ilişkiler ve benzersiz kimlikler yeniden kurulur, io/docx/partWriter.ts).
//
// Paragraf ve koşu özelliklerinin modelde karşılığı olmayan çocukları (dil, karakter ölçeği, w14 efektleri…)
// paragraf/başlık düğümünde ve textStyle işaretinde `ooxmlExtra` (XML dizesi) olarak taşınır.
//
// Pano: XML özniteliklere HTML'de yazılmaz; yalnız bu penceredeki kayıt defterinin anahtarı yazılır. Böylece
// kopyala/yapıştır/çoğalt öğeyi XML'iyle çoğaltır, dışarıdan yapıştırılan HTML ise belgeye XML sokamaz.
import { Extension, Node, mergeAttributes } from "@tiptap/core";

/** Opak öğenin okurken toplanan ilişkisi (@officehere/core/opc RawRef ile aynı biçim). */
export interface OpaqueRef {
  id: string;
  type: string;
  /** İç hedefte kaynak paketteki parça adı, dış hedefte adres. */
  target: string;
  external: boolean;
}

export type OpaqueKind = "ole" | "control" | "drawing" | "vml" | "hr" | "altChunk" | "subDoc" | "math" | "other";

export interface OpaqueAttrs {
  /** Özgün öğe (kendi ad alanı bildirimleriyle, tek kök). */
  xml: string;
  refs: OpaqueRef[];
  /** Kaynak paket kimliği (okunduğu belgenin `settings.kaynak`'ı). */
  src: string | null;
  kind: OpaqueKind;
  /** Kullanıcıya gösterilen ad ("Gömülü nesne (hesap tablosu)"). */
  label: string;
  /** Satır içi öğede yazılacağı düzey: koşu içinde (w:r) ya da paragraf içinde. */
  level?: "run" | "para";
  /** Tarayıcıda çizilebilen yedek görüntü (data: adresi; EMF/WMF yok). */
  preview: string | null;
  /** Görünür boyut (px; bilinmiyorsa null). */
  width: number | null;
  height: number | null;
}

/* ------------------------------------------------------------------ */
/* Pano kayıt defteri                                                    */
/* ------------------------------------------------------------------ */

const SESSION = Math.random().toString(36).slice(2, 10);
const MAX_ENTRIES = 4000;
let seq = 0;
const payloads = new Map<string, unknown>();
const byObject = new WeakMap<object, string>();
const byString = new Map<string, string>();

function keep(token: string, value: unknown): string {
  payloads.delete(token);
  payloads.set(token, value);
  if (payloads.size > MAX_ENTRIES) {
    const oldest = payloads.keys().next().value as string;
    payloads.delete(oldest);
  }
  return token;
}

/** Değeri kayıt defterine yazar ve HTML'e konacak anahtarı döner (aynı değer aynı anahtarı alır). */
export function clipToken(value: object | string): string {
  if (typeof value === "string") {
    let t = byString.get(value);
    if (!t) {
      t = `${SESSION}-${++seq}`;
      byString.set(value, t);
      if (byString.size > MAX_ENTRIES) byString.delete(byString.keys().next().value as string);
    }
    return keep(t, value);
  }
  let t = byObject.get(value);
  if (!t) {
    t = `${SESSION}-${++seq}`;
    byObject.set(value, t);
  }
  return keep(t, value);
}

/** Anahtarın değeri (başka pencereden ya da dışarıdan gelen anahtarda undefined). */
export function fromClip(token: string | null | undefined): unknown {
  return token ? payloads.get(token) : undefined;
}

/* ------------------------------------------------------------------ */
/* Yer tutucu metinler                                                   */
/* ------------------------------------------------------------------ */

const NOTE = "kaydederken korunur";

/** Kartın erişilebilir adı ve ipucu. */
export function opaqueTitle(label: string): string {
  return `Desteklenmeyen öğe: ${label} — ${NOTE}`;
}

/** .odt/.rtf/.md'ye kaydederken öğenin yerine yazılan metin. */
export function opaquePlaceholder(attrs: Record<string, unknown> | undefined): string {
  return `[Desteklenmeyen öğe: ${String(attrs?.label || "bilinmeyen öğe")}]`;
}

/** Belgede (gövde, üst/alt bilgi, notlar) opak öğe var mı? */
export function hasOpaque(json: unknown): boolean {
  return /"type":"opaque(?:Block|Inline)"/.test(JSON.stringify(json ?? null));
}

/* ------------------------------------------------------------------ */
/* Düğümler                                                              */
/* ------------------------------------------------------------------ */

const PX_MAX = 4000;
const px = (v: unknown): number | null => {
  const n = Number(v);
  return Number.isFinite(n) && n > 0 ? Math.min(PX_MAX, Math.round(n)) : null;
};

/** Ortak öznitelikler: HTML'e yazılmaz; yapıştırmada kayıt defterinden okunur. */
function opaqueAttributes(inline: boolean) {
  const fromToken = (name: keyof OpaqueAttrs, def: unknown) => ({
    default: def,
    rendered: false,
    parseHTML: (el: HTMLElement) => {
      const v = fromClip(el.getAttribute("data-opaque")) as Record<string, unknown> | undefined;
      return v ? (v[name] ?? def) : def;
    },
  });
  return {
    xml: fromToken("xml", ""),
    refs: fromToken("refs", []),
    src: fromToken("src", null),
    kind: fromToken("kind", "other"),
    label: fromToken("label", ""),
    ...(inline ? { level: fromToken("level", "run") } : {}),
    preview: fromToken("preview", null),
    width: fromToken("width", null),
    height: fromToken("height", null),
  };
}

/** Kart (düzenleyici, yazdırma ve okuma görünümü ortak; pano için kayıt defteri anahtarı). */
function cardSpec(attrs: Record<string, unknown>, inline: boolean, html: Record<string, unknown>) {
  const label = String(attrs.label || "bilinmeyen öğe");
  const title = opaqueTitle(label);
  const w = px(attrs.width);
  const h = px(attrs.height);
  const hr = attrs.kind === "hr";
  const style = inline ? [w && !hr ? `width:${w}px` : "", h ? `height:${h}px` : ""].filter(Boolean).join(";") : h ? `min-height:${Math.min(h, 600)}px` : "";
  const cls = `bl-opaque ${inline ? "is-inline" : "is-block"}${hr ? " is-hr" : ""}${attrs.preview ? " has-preview" : ""}${inline && !w && !hr ? " is-unsized" : ""}`;
  const outer = mergeAttributes(html, {
    "data-opaque": clipToken(attrs),
    class: cls,
    role: "img",
    "aria-label": title,
    title,
    ...(style ? { style } : {}),
  });
  const tag = inline ? "span" : "div";
  if (hr) return [tag, outer] as const;
  const body: unknown[] = [];
  if (attrs.preview) body.push(["img", { class: "bl-opaque-img", src: String(attrs.preview), alt: "", draggable: "false" }]);
  body.push([
    "span",
    { class: "bl-opaque-card" },
    ["span", { class: "bl-opaque-icon", "aria-hidden": "true" }, "⧉"],
    ["span", { class: "bl-opaque-text" }, `Desteklenmeyen öğe: ${label}`],
    ["span", { class: "bl-opaque-note" }, NOTE],
  ]);
  return [tag, outer, ...body] as const;
}

/** Blok düzeyi opak öğe (alt belge, içe aktarılan içerik, gövdede bilinmeyen öğe). */
export const OpaqueBlock = Node.create({
  name: "opaqueBlock",
  group: "block",
  atom: true,
  selectable: true,
  draggable: true,
  addAttributes() {
    return opaqueAttributes(false);
  },
  parseHTML() {
    // Bu pencerede kayıtlı olmayan anahtar (dış HTML ya da başka pencere) içeriğiyle birlikte atılır.
    return [
      { tag: "div[data-opaque]", priority: 70, getAttrs: (el) => (fromClip((el as HTMLElement).getAttribute("data-opaque")) ? null : false) },
      { tag: "div[data-opaque]", priority: 69, ignore: true },
    ];
  },
  renderHTML({ node, HTMLAttributes }) {
    return cardSpec(node.attrs, false, HTMLAttributes) as unknown as [string, Record<string, unknown>];
  },
  renderText() {
    return "";
  },
});

/** Satır içi opak öğe (gömülü nesne, denetim, anlaşılmayan çizim, bilinmeyen koşu çocuğu). */
export const OpaqueInline = Node.create({
  name: "opaqueInline",
  group: "inline",
  inline: true,
  atom: true,
  selectable: true,
  draggable: true,
  addAttributes() {
    return opaqueAttributes(true);
  },
  parseHTML() {
    return [
      { tag: "span[data-opaque]", priority: 70, getAttrs: (el) => (fromClip((el as HTMLElement).getAttribute("data-opaque")) ? null : false) },
      { tag: "span[data-opaque]", priority: 69, ignore: true },
    ];
  },
  renderHTML({ node, HTMLAttributes }) {
    return cardSpec(node.attrs, true, HTMLAttributes) as unknown as [string, Record<string, unknown>];
  },
  renderText() {
    return "";
  },
});

/**
 * Paragraf/başlık ve koşu (textStyle) ek özellikleri: modelin okumadığı pPr/rPr çocukları (XML dizesi).
 * Görünüme etkisi yok; panoda kayıt defteri anahtarıyla taşınır.
 */
export const OoxmlExtra = Extension.create({
  name: "ooxmlExtra",
  addGlobalAttributes() {
    const attribute = {
      default: null,
      parseHTML: (el: HTMLElement) => {
        const v = fromClip(el.getAttribute("data-ooxml"));
        return typeof v === "string" ? v : null;
      },
      renderHTML: (a: Record<string, unknown>) => (typeof a.ooxmlExtra === "string" && a.ooxmlExtra ? { "data-ooxml": clipToken(a.ooxmlExtra) } : {}),
    };
    return [
      { types: ["paragraph", "heading"], attributes: { ooxmlExtra: attribute } },
      { types: ["textStyle"], attributes: { ooxmlExtra: attribute } },
    ];
  },
});
