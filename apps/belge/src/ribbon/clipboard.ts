// Pano işlemleri (Kes/Kopyala/Yapıştır ve yapıştırma seçenekleri).
import type { Editor } from "@tiptap/core";
import { cleanOfficeHtml, lastInternalCopy, sanitizePastedHtml } from "../editor/paste";

export function execClipboard(editor: Editor, cmd: "cut" | "copy"): boolean {
  editor.commands.focus();
  try {
    return document.execCommand(cmd);
  } catch {
    return false;
  }
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
          // Ctrl+V ile aynı yol: yapıştırma dönüştürücüleri (Office HTML, düz metin) uygulanır. HTML etkisiz bir
          // belgede ayrıştırılır (canlı belgede innerHTML ile <img onerror> gibi betikler çalışırdı) ve temizlenir.
          if (mode === "merge") {
            const tmp = new DOMParser().parseFromString(sanitizePastedHtml(cleanOfficeHtml(html)), "text/html").body;
            tmp.querySelectorAll("[style]").forEach((el) => el.removeAttribute("style"));
            tmp.querySelectorAll("font").forEach((el) => el.replaceWith(...Array.from(el.childNodes)));
            editor.view.pasteHTML(tmp.innerHTML);
          } else editor.view.pasteHTML(html);
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
    editor.view.pasteText(text);
  } catch {
    // Pano izni yok: Word'deki gibi bilgilendir; uygulama içinden kopyalanan son içerik yine de yapıştırılır.
    const slice = lastInternalCopy();
    if (slice && slice.size) {
      const { state } = editor;
      editor.view.dispatch(mode === "text" ? state.tr.insertText(slice.content.textBetween(0, slice.content.size, "\n")) : state.tr.replaceSelection(slice).scrollIntoView());
      notify("Tarayıcı panoya doğrudan erişime izin vermiyor; son kopyalanan içerik yapıştırıldı. Başka uygulamalardan yapıştırmak için Ctrl+V kullanın.");
      return;
    }
    notify("Tarayıcınız panoya doğrudan erişime izin vermiyor. Yapıştırmak için Ctrl+V tuşlarını kullanın.");
  }
}
