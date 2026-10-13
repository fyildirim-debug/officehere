// Dosya biçimleri: açma ve dışa aktarma girişi.
import { generateJSON, getSchema, type JSONContent } from "@tiptap/core";
import { Node as PMNode } from "@tiptap/pm/model";
import { defaultSettings, normalizeDoc, type BelgeDocument } from "../model/document";
import { schemaExtensions } from "../editor/extensions";
import { readDocx } from "./docx/reader";
import { writeDocx } from "./docx/writer";
import { readOdt } from "./odt/reader";
import { writeOdt } from "./odt/writer";
import { readRtf } from "./rtf/reader";
import { writeRtf } from "./rtf/writer";
import { readMarkdown } from "./markdown/reader";
import { writeMarkdown } from "./markdown/writer";
import { readDoc } from "./doc/reader";
import { decodeText, docToText, extractHtmlBody, htmlDocument, textToDoc } from "./text";
import { DOCX_FILTER, FORMATS, OPEN_ACCEPT, OPEN_FILTERS, detectFormat, formatOfExt, isWritable, saveFilters, type FormatId } from "./formats";

export { writeDocx, readDocx, writeOdt, readOdt, writeRtf, readRtf, writeMarkdown, readMarkdown, readDoc, docToText, htmlDocument };
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
  return normalizeDoc(generateJSON(extractHtmlBody(html), schemaExtensions()));
}

/** .doc açılınca ileti çubuğunda gösterilen uyumluluk bilgisi. */
export const COMPAT_NOTE = "Uyumluluk modu: Bu belge eski .doc (97-2003) biçiminde. Düzenleyebilirsiniz; kaydettiğinizde .docx olarak kaydedilir, özgün dosya değişmez.";

export interface OpenOptions {
  /** Markdown'daki göreli resim yolunu çözer (masaüstünde belgenin klasöründen). */
  resolveImage?: (path: string) => Promise<Uint8Array | null>;
}

export interface OpenResult {
  doc: BelgeDocument;
  warning?: string;
  format: FormatId;
}

/** Yeni okuyucuların çıktısı şemaya uymazsa (beklenmeyen yapı) içerik düz metne düşürülür. */
function safe(doc: BelgeDocument, warning?: string): { doc: BelgeDocument; warning?: string } {
  try {
    validateContent(doc.content);
    return { doc, warning };
  } catch {
    return { doc: { content: textToDoc(docToText(doc.content, "\n")), settings: doc.settings }, warning: "Belgenin yapısı tam okunamadı; metin biçimsiz olarak açıldı." };
  }
}

/** Dosyayı biçimine göre açar (uzantı yanlışsa içerik imzası belirleyicidir). DOCX ayrıştırılamazsa mammoth ile yedek dönüşüm yapılır. */
export async function openDocument(name: string, data: ArrayBuffer, opts: OpenOptions = {}): Promise<OpenResult> {
  const format = await detectFormat(name, data);
  switch (format) {
    case "docx":
      try {
        const doc = await readDocx(data);
        validateContent(doc.content);
        return { doc, format };
      } catch (err) {
        const mammoth = await import("mammoth");
        const res = await mammoth.convertToHtml({ arrayBuffer: data });
        return { doc: { content: htmlToDoc(res.value), settings: defaultSettings() }, warning: `Belge basitleştirilmiş modda açıldı (${(err as Error).message}).`, format };
      }
    case "doc":
      return { ...safe(readDoc(data), COMPAT_NOTE), format };
    case "odt":
      return { ...safe(await readOdt(data)), format };
    case "rtf":
      return { ...safe(readRtf(data)), format };
    case "md": {
      const doc = await readMarkdown(decodeText(data), {
        resolveImage: opts.resolveImage,
        htmlToBlocks: typeof DOMParser === "undefined" ? undefined : (html) => (htmlToDoc(html).content?.[0]?.content ?? []).filter((b) => b.content?.length || b.type !== "paragraph"),
      });
      return { ...safe(doc), format };
    }
    case "html":
      return { doc: { content: htmlToDoc(decodeText(data)), settings: defaultSettings() }, format };
    default:
      return { doc: { content: textToDoc(decodeText(data)), settings: defaultSettings() }, format: "txt" };
  }
}
