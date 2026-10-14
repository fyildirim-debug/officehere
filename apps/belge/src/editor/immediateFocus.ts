// Editöre odak, geçerli olay biter bitmez verilir.
import { Extension, commands as core } from "@tiptap/core";

/**
 * tiptap `focus` komutu odağı bir sonraki karede verir (React ile eşzamanlı odak sorun çıkarır).
 * Arada basılan tuş editöre değil hâlâ odaktaki şerit kutusuna gidiyordu: yazı tipini Enter ile
 * seçip hemen yazınca ilk harf kutuya düşüyor, iletişim kutusu kapanınca imleç belge başına
 * kayabiliyordu. Özgün komut aynen çalışır; ek olarak odak, komut zinciri gönderildikten hemen
 * sonra (mikro görevde, sonraki tuştan önce) verilir. Zincirin ortasında odak taşınmaz: odağı
 * kaybeden öğenin işleyicisi yarım kalmış işlemin üstüne yeni işlem gönderemez.
 */
export const ImmediateFocus = Extension.create({
  name: "immediateFocus",
  addCommands() {
    return {
      focus:
        (position, options) =>
        (props) => {
          const ok = core.focus(position, options)(props);
          const { editor, view, dispatch } = props;
          if (ok && dispatch && position !== false) {
            queueMicrotask(() => {
              if (!editor.isDestroyed && !view.hasFocus()) view.focus();
            });
          }
          return ok;
        },
    };
  },
});
