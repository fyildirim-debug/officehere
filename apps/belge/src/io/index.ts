// Dosya biçimleri: açma ve dışa aktarma girişi.
import { generateJSON, getSchema, type JSONContent } from "@tiptap/core";
import { Node as PMNode } from "@tiptap/pm/model";
import { extensionOf } from "@officehere/core";
import { defaultSettings, type BelgeDocument } from "../model/document";
import { schemaExtensions } from "../editor/extensions";
import { readDocx } from "./docx/reader";
import { writeDocx } from "./docx/writer";
import { decodeText, docToText, extractHtmlBody, htmlDocument, textToDoc } from "./text";

export { writeDocx, readDocx, docToText, htmlDocument };

let schemaCache: ReturnType<typeof getSchema> | null = null;
export function belgeSchema() {
  if (!schemaCache) schemaCache = getSchema(schemaExtensions());
  return schemaCache;
}

/** JSON içeriği şemaya karşı doğrular; geçersizse hatayı fırlatır. */
export function validateContent(json: JSONContent): PMNode {
  const node = PMNode.fromJSON(belgeSchema(), json);
  node.check();
  return node;
}

export function htmlToDoc(html: string): JSONContent {
  return generateJSON(extractHtmlBody(html), schemaExtensions());
}

export const OPEN_FILTERS = [
  { name: "Tüm Belgeler", extensions: ["docx", "html", "htm", "txt"] },
  { name: "Belge (.docx)", extensions: ["docx"] },
  { name: "Web Sayfası", extensions: ["html", "htm"] },
  { name: "Düz Metin", extensions: ["txt"] },
];

export const DOCX_FILTER = [{ name: "Belge (.docx)", extensions: ["docx"] }];

/** Dosyayı uzantısına göre açar. DOCX ayrıştırılamazsa mammoth ile yedek dönüşüm yapılır. */
export async function openDocument(name: string, data: ArrayBuffer): Promise<{ doc: BelgeDocument; warning?: string }> {
  const ext = extensionOf(name);
  if (ext === "docx" || ext === "docm" || ext === "dotx") {
    try {
      const doc = await readDocx(data);
      validateContent(doc.content);
      return { doc };
    } catch (err) {
      const mammoth = await import("mammoth");
      const res = await mammoth.convertToHtml({ arrayBuffer: data });
      return { doc: { content: htmlToDoc(res.value), settings: defaultSettings() }, warning: `Belge basitleştirilmiş modda açıldı (${(err as Error).message}).` };
    }
  }
  const text = decodeText(data);
  if (ext === "html" || ext === "htm") return { doc: { content: htmlToDoc(text), settings: defaultSettings() } };
  return { doc: { content: textToDoc(text), settings: defaultSettings() } };
}
