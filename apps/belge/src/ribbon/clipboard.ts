// Pano işlemleri (Kes/Kopyala/Yapıştır ve yapıştırma seçenekleri).
import type { Editor } from "@tiptap/core";

export function execClipboard(editor: Editor, cmd: "cut" | "copy"): boolean {
  editor.commands.focus();
  try {
    return document.execCommand(cmd);
  } catch {
    return false;
  }
}

function textToContent(text: string) {
  const lines = text.replace(/\r\n?/g, "\n").split("\n");
  return lines.map((l) => ({ type: "paragraph", content: l ? [{ type: "text", text: l }] : undefined }));
}

/** Yapıştır. mode: keep = kaynak biçimlendirmesi, text = yalnız metin, merge = biçimleri birleştir. */
export async function pasteFromClipboard(editor: Editor, mode: "keep" | "merge" | "text", notify: (t: string) => void) {
  editor.commands.focus();
  try {
    if (mode !== "text" && navigator.clipboard?.read) {
      const items = await navigator.clipboard.read();
      for (const it of items) {
        if (it.types.includes("text/html")) {
          const html = await (await it.getType("text/html")).text();
          if (mode === "merge") {
            const tmp = document.createElement("div");
            tmp.innerHTML = html;
            tmp.querySelectorAll("[style]").forEach((el) => el.removeAttribute("style"));
            tmp.querySelectorAll("font").forEach((el) => el.replaceWith(...Array.from(el.childNodes)));
            editor.chain().focus().insertContent(tmp.innerHTML).run();
          } else editor.chain().focus().insertContent(html).run();
          return;
        }
        if (it.types.some((t) => t.startsWith("image/"))) {
          const type = it.types.find((t) => t.startsWith("image/"))!;
          const blob = await it.getType(type);
          const src = await new Promise<string>((res) => {
            const fr = new FileReader();
            fr.onload = () => res(String(fr.result));
            fr.readAsDataURL(blob);
          });
          editor.chain().focus().setImage({ src }).run();
          return;
        }
      }
    }
    const text = await navigator.clipboard.readText();
    if (!text) return;
    const content = textToContent(text);
    if (content.length === 1) editor.chain().focus().insertContent({ type: "text", text }).run();
    else editor.chain().focus().insertContent(content).run();
  } catch {
    notify("Tarayıcı panoya erişim izni vermedi. Yapıştırmak için Ctrl+V tuşlarını kullanın.");
  }
}
