// Zengin not içeriğinin statik HTML'i (Okuma Modu, HTML dışa aktarma).
import { generateHTML, type JSONContent } from "@tiptap/core";
import { headerFooterExtensions } from "../editor/extensions";

let ext: ReturnType<typeof headerFooterExtensions> | null = null;

export function renderNoteHtml(content: JSONContent[]): string {
  ext ??= headerFooterExtensions();
  try {
    return generateHTML({ type: "doc", content }, ext);
  } catch {
    return "";
  }
}
