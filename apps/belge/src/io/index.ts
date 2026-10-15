// Dosya biçimleri: açma ve dışa aktarma girişi.
import { generateJSON, getSchema, type JSONContent } from "@tiptap/core";
import { Node as PMNode } from "@tiptap/pm/model";
import { defaultSettings, normalizeDoc, type BelgeDocument } from "../model/document";
import { schemaExtensions } from "../editor/extensions";
import { sanitizePastedHtml } from "../editor/paste";
import { readDocx } from "./docx/reader";
import { writeDocx } from "./docx/writer";
import { readFodt, readOdt } from "./odt/reader";
import { writeOdt } from "./odt/writer";
import { readRtf } from "./rtf/reader";
import { writeRtf } from "./rtf/writer";
import { readMarkdown } from "./markdown/reader";
import { writeMarkdown } from "./markdown/writer";
import { readDoc } from "./doc/reader";
import { decodeText, docToText, extractHtmlBody, htmlDocument, textToDoc } from "./text";
import { sniffImage, toBase64 } from "./common";
import { BelgeOpenError, LoadWarnings, classifyOpenError } from "./warnings";
import { DOCX_FILTER, FORMATS, OPEN_ACCEPT, OPEN_FILTERS, detectFormat, formatOfExt, isWritable, saveFilters, type FormatId } from "./formats";

export { writeDocx, readDocx, writeOdt, readOdt, readFodt, writeRtf, readRtf, writeMarkdown, readMarkdown, readDoc, docToText, htmlDocument };
export { BelgeOpenError, LoadWarnings, classifyOpenError } from "./warnings";
export { DOCX_FILTER, FORMATS, OPEN_ACCEPT, OPEN_FILTERS, detectFormat, formatOfExt, isWritable, saveFilters, type FormatId };

let schemaCache: ReturnType<typeof getSchema> | null = null;
export function belgeSchema() {
  if (!schemaCache) schemaCache = getSchema(schemaExtensions());
  return schemaCache;
}

/** JSON içeriği şemaya karşı doğrular; geçersizse hatayı fırlatır. */
export function validateContent(json: JSONContent): PMNode {
  const node = PMNode.fromJSON(belgeSchema(), normalizeDoc(json));
  node.check();
  return node;
}

export function htmlToDoc(html: string): JSONContent {
  const body = extractHtmlBody(html);
  // Açılan/dönüştürülen HTML de etkisizleştirilir (betik, olay öznitelikleri, javascript: adresleri).
  return normalizeDoc(generateJSON(typeof DOMParser === "undefined" ? body : sanitizePastedHtml(body), schemaExtensions()));
}

/** .doc açılınca ileti çubuğunda gösterilen uyumluluk bilgisi. */
export const COMPAT_NOTE = "Uyumluluk modu: Bu belge eski .doc (97-2003) biçiminde. Düzenleyebilirsiniz; kaydettiğinizde .docx olarak kaydedilir, özgün dosya değişmez.";

/** .fodt açılınca gösterilen bilgi (tek XML dosyası; kaydederken paketli biçim seçilir). */
export const FODT_NOTE = "Bu belge düz XML OpenDocument (.fodt) biçiminde. Düzenleyebilirsiniz; kaydettiğinizde yeni bir dosyaya (.docx ya da .odt) kaydedilir, özgün dosya değişmez.";

export interface OpenOptions {
  /** Markdown ve HTML'deki göreli resim yolunu çözer (masaüstünde belgenin klasöründen). */
  resolveImage?: (path: string) => Promise<Uint8Array | null>;
}

export interface OpenResult {
  doc: BelgeDocument;
  /** İleti çubuğunda gösterilecek metin: uyumluluk notu ve aşağıdaki uyarılar (yoksa tanımsız). */
  warning?: string;
  /**
   * Okunamayan ya da tam okunamayan, dolayısıyla kaydedilince kaybolacak içerik uyarıları (Türkçe, kullanıcıya
   * gösterilebilir). Boş değilse özgün dosyanın üzerine sessizce (otomatik) yazılmamalıdır.
   */
  warnings: string[];
  format: FormatId;
}

/** Yeni okuyucuların çıktısı şemaya uymazsa (beklenmeyen yapı) içerik düz metne düşürülür. */
function safe(doc: BelgeDocument, warnings: LoadWarnings): BelgeDocument {
  try {
    validateContent(doc.content);
    return doc;
  } catch (err) {
    warnings.add("flattened", (err as Error).message);
    return { content: textToDoc(docToText(doc.content, "\n")), settings: doc.settings };
  }
}

/**
 * Veri olmayan (göreli yol ya da web adresi) resimler: çözülebilenler gömülür; çözülemeyenler kaydedilince
 * kaybolacağından uyarılır.
 */
async function embedImages(doc: BelgeDocument, opts: OpenOptions, warnings: LoadWarnings) {
  const nodes: JSONContent[] = [];
  const walk = (n: JSONContent | null | undefined) => {
    if (!n) return;
    if (n.type === "image" && typeof n.attrs?.src === "string" && !n.attrs.src.startsWith("data:")) nodes.push(n);
    n.content?.forEach(walk);
  };
  walk(doc.content);
  for (const k of ["header", "footer", "headerFirst", "footerFirst", "headerEven", "footerEven"] as const) walk(doc.settings[k]);
  if (!nodes.length) return;
  const cache = new Map<string, string | null>();
  let web = 0;
  let missing = 0;
  for (const n of nodes.slice(0, 500)) {
    const src = String(n.attrs!.src);
    if (/^(https?:|\/\/)/i.test(src)) {
      web++;
      continue;
    }
    if (!cache.has(src)) {
      let url: string | null = null;
      try {
        const bytes = opts.resolveImage && !/^[a-z][a-z0-9+.-]*:/i.test(src) ? await opts.resolveImage(decodeURI(src.replace(/[?#].*$/, ""))) : null;
        const t = bytes ? sniffImage(bytes) : null;
        if (bytes && t) url = `data:${t.mime};base64,${toBase64(bytes)}`;
      } catch {
        url = null;
      }
      cache.set(src, url);
    }
    const url = cache.get(src);
    if (url) n.attrs = { ...n.attrs, src: url };
    else missing++;
  }
  if (missing) warnings.add("images", `çözülemeyen resim: ${missing}`);
  if (web) warnings.addText("Web adresinden bağlanan resimler belgeye gömülemedi; kaydederseniz bu resimler kaybolur.");
}

/** Teknik ayrıntı kullanıcı iletisinden ayrı olarak günlüğe yazılır. */
function logged(e: BelgeOpenError): BelgeOpenError {
  if (e.detail && typeof console !== "undefined") console.warn("[belge] açılamadı (ayrıntı):", e.detail);
  return e;
}

/** Okuyucu çağrısı: ham hatalar anlaşılır iletiye çevrilir. */
async function guarded<T>(format: string, fn: () => T | Promise<T>): Promise<T> {
  try {
    return await fn();
  } catch (err) {
    throw logged(classifyOpenError(err, format));
  }
}

/** Açmanın DOM gerektiren adımı (HTML/Markdown okuma, .docx yedek dönüştürücüsü) Worker'da yapılamaz. */
export class DomRequired extends Error {
  constructor() {
    super("Bu dosya ana iş parçacığında açılmalı.");
    this.name = "DomRequired";
  }
}

/** Açmanın ilk adımının (biçim tanıma ve okuma; Worker'da çalışabilir) sonucu. */
export interface ReadResult {
  doc: BelgeDocument;
  format: FormatId;
  /** İleti çubuğu notları (uyumluluk modu, düz XML). */
  notes: string[];
  warnings: LoadWarnings;
}

export interface ReadOptions extends Pick<OpenOptions, "resolveImage"> {
  /** DOM yok (Worker): HTML, Markdown ve .docx yedek dönüştürücüsü yerine `DomRequired` fırlatılır. */
  noDom?: boolean;
}

/**
 * Dosyayı biçimine göre açar (uzantı yanlışsa içerik imzası belirleyicidir). DOCX ayrıştırılamazsa mammoth ile
 * yedek dönüşüm yapılır. Okunamayan içerik `warnings` alanında bildirilir; açılamayan dosyada anlaşılır Türkçe
 * iletili `BelgeOpenError` fırlatılır (teknik ayrıntı `detail` alanında).
 */
export async function openDocument(name: string, data: ArrayBuffer, opts: OpenOptions = {}): Promise<OpenResult> {
  return finishOpen(await readDocument(name, data, opts), opts);
}

/** Açmanın son adımı (ana iş parçacığı): bağlı resimler gömülür, iletiler toplanır. */
export async function finishOpen({ doc, format, notes, warnings }: ReadResult, opts: OpenOptions = {}): Promise<OpenResult> {
  await embedImages(doc, opts, warnings);
  const list = warnings.messages();
  const shown = [...notes, ...list];
  if (warnings.details.length && typeof console !== "undefined") console.warn("[belge] açılış ayrıntıları:", warnings.details.join("; "));
  return { doc, format, warnings: list, ...(shown.length ? { warning: shown.join(" ") } : {}) };
}

/** Biçimi tanır ve okur (bağlı resimler ve iletiler `finishOpen`'da). */
export async function readDocument(name: string, data: ArrayBuffer, opts: ReadOptions = {}): Promise<ReadResult> {
  const format = await detectFormat(name, data);
  if (opts.noDom && (format === "md" || format === "html")) throw new DomRequired();
  const warnings = new LoadWarnings();
  const notes: string[] = [];
  let doc: BelgeDocument;
  switch (format) {
    case "docx":
      doc = await openDocx(data, warnings, opts.noDom);
      break;
    case "doc":
      notes.push(COMPAT_NOTE);
      doc = safe(await guarded(format, () => readDoc(data, warnings)), warnings);
      break;
    case "odt":
      doc = safe(await guarded(format, () => readOdt(data, warnings)), warnings);
      break;
    case "fodt":
      notes.push(FODT_NOTE);
      doc = safe(await guarded(format, () => readFodt(decodeText(data), warnings)), warnings);
      break;
    case "rtf":
      doc = safe(await guarded(format, () => readRtf(data, warnings)), warnings);
      break;
    case "md": {
      const md = await guarded(format, () =>
        readMarkdown(decodeText(data), {
          resolveImage: opts.resolveImage,
          htmlToBlocks: typeof DOMParser === "undefined" ? undefined : (html) => (htmlToDoc(html).content?.[0]?.content ?? []).filter((b) => b.content?.length || b.type !== "paragraph"),
        }),
      );
      doc = safe(md, warnings);
      break;
    }
    case "html":
      doc = await guarded(format, () => ({ content: htmlToDoc(decodeText(data)), settings: defaultSettings() }));
      break;
    default:
      doc = { content: textToDoc(decodeText(data)), settings: defaultSettings() };
  }
  return { doc, format, notes, warnings };
}

/** DOCX: kendi okuyucumuz; olmazsa (bozuk yapı) mammoth ile basitleştirilmiş dönüşüm. */
async function openDocx(data: ArrayBuffer, warnings: LoadWarnings, noDom = false): Promise<BelgeDocument> {
  try {
    const doc = await readDocx(data, warnings);
    validateContent(doc.content);
    return doc;
  } catch (err) {
    // Paket açılamıyorsa ya da dosya reddedildiyse yedek dönüşüm de başarısız olur: anlaşılır hata.
    const first = classifyOpenError(err, "docx");
    if (err instanceof BelgeOpenError || /ZIP/.test(first.message)) throw logged(first);
    // Yedek dönüştürücünün HTML'i ancak DOM'da belge modeline çevrilir.
    if (noDom) throw new DomRequired();
    try {
      const mammoth = await import("mammoth");
      const res = await mammoth.convertToHtml({ arrayBuffer: data });
      const content = htmlToDoc(res.value);
      validateContent(content);
      warnings.add("simplified", first.detail);
      return { content, settings: defaultSettings() };
    } catch (err2) {
      const second = classifyOpenError(err2, "docx");
      throw logged(new BelgeOpenError(first.message, [first.detail, `yedek dönüştürücü: ${second.detail}`].filter(Boolean).join(" | ")));
    }
  }
}
