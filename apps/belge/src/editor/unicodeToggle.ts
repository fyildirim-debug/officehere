// Alt+X: imleçten önceki onaltılık kodu karaktere ya da karakteri koduna çevirir (seçim varsa seçimi).
// "20BA" → "₺", "U+00E7" → "ç"; "ç" → "00E7". Ekle › Simge iletişim kutusundaki karakter koduyla aynı gösterim.
import type { Editor } from "@tiptap/core";
import { TextSelection } from "@tiptap/pm/state";

/** Karakterin kodu (en az 4 haneli büyük harfli onaltılık). */
export function codeOf(ch: string): string {
  return ch.codePointAt(0)!.toString(16).toUpperCase().padStart(4, "0");
}

/**
 * Metnin sonundaki dönüşüm: sondaki onaltılık kod (en çok 6 hane, isteğe bağlı "U+") karaktere; değilse
 * son karakter koduna. remove: sondan silinecek karakter sayısı (UTF-16 birimi).
 */
export function unicodeToggle(before: string): { remove: number; insert: string } | null {
  if (!before) return null;
  const m = /(U\+)?([0-9A-Fa-f]{1,6})$/.exec(before);
  if (m) {
    // En uzun geçerli kod (ör. "x20BA" → "20BA"; 0x10FFFF üstü kodlarda daha kısa son ek denenir).
    for (let len = m[2].length; len > 0; len--) {
      const cp = parseInt(m[2].slice(-len), 16);
      if (cp < 0x20 || cp > 0x10ffff || (cp >= 0xd800 && cp <= 0xdfff) || (cp >= 0x7f && cp < 0xa0)) continue;
      return { remove: len + (len === m[2].length && m[1] ? 2 : 0), insert: String.fromCodePoint(cp) };
    }
  }
  const last = Array.from(before).pop()!;
  return { remove: last.length, insert: codeOf(last) };
}

/** Alt+X komutu. */
export function toggleUnicode(editor: Editor): boolean {
  const { state, view } = editor;
  const { from, to, empty, $from } = state.selection;
  if (!$from.parent.isTextblock) return false;
  if (!empty) {
    const sel = state.doc.textBetween(from, to, "\0", "\0");
    if (!sel || sel.includes("\0")) return false;
    // Seçim bir kodsa karaktere, değilse her karakter koduna çevrilir; sonuç seçili kalır.
    const r = unicodeToggle(sel);
    const text = r && r.remove === sel.length ? r.insert : Array.from(sel).map(codeOf).join(" ");
    const tr = state.tr.insertText(text, from, to);
    view.dispatch(tr.setSelection(TextSelection.create(tr.doc, from, from + text.length)));
    return true;
  }
  const before = $from.parent.textBetween(Math.max(0, $from.parentOffset - 8), $from.parentOffset, "\0", "\0");
  const r = unicodeToggle(before.slice(before.lastIndexOf("\0") + 1));
  if (!r) return false;
  view.dispatch(state.tr.insertText(r.insert, from - r.remove, from).scrollIntoView());
  return true;
}
