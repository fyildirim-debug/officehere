// Üst/alt bilgi düzenleyicisi: sayfanın üst ya da alt kenar boşluğunda açılan ikinci editör.
import { useEffect, useRef } from "react";
import { EditorContent, useEditor } from "@tiptap/react";
import type { Editor, JSONContent } from "@tiptap/core";
import { headerFooterExtensions } from "../editor/extensions";
import type { PageInfo } from "../editor/pagination";

interface Props {
  which: "header" | "footer";
  /** Alan etiketi (ör. "İlk Sayfa Üst Bilgisi -Bölüm 1-"). */
  label: string;
  page: PageInfo;
  /** Sayfanın sayfa kapsayıcısındaki konumu. */
  rect: { x: number; y: number };
  content: JSONContent | null;
  onChange: (json: JSONContent) => void;
  onEditor: (ed: Editor | null) => void;
  onClose: () => void;
}

export function HeaderFooterEditor({ which, label, page: g, rect, content, onChange, onEditor, onClose }: Props) {
  const changeRef = useRef(onChange);
  changeRef.current = onChange;
  const editor = useEditor(
    {
      extensions: headerFooterExtensions(),
      content: content ?? { type: "doc", content: [{ type: "paragraph" }] },
      // Odak aşağıdaki etkide hemen verilir (gecikmeli odakta ilk tuşlar kaybolmasın).
      autofocus: false,
      editorProps: { attributes: { class: "bl-hf-editor", spellcheck: "true", lang: "tr", "aria-label": label } },
      onUpdate: ({ editor: ed }) => changeRef.current(ed.getJSON()),
    },
    [which, label],
  );

  useEffect(() => {
    onEditor(editor);
    return () => onEditor(null);
  }, [editor, onEditor]);

  // Açılışta: hazır yer tutucu ("[Buraya yazın]") varsa seçilir, yazınca yerine geçer (Word'deki gibi);
  // yoksa imleç sona konur.
  useEffect(() => {
    let range: { from: number; to: number } | null = null;
    editor.state.doc.descendants((n, pos) => {
      if (range || !n.isText) return !range;
      const m = /\[[^\]]+\]/.exec(n.text ?? "");
      if (m) range = { from: pos + m.index, to: pos + m.index + m[0].length };
      return false;
    });
    if (range) editor.commands.setTextSelection(range);
    editor.commands.focus(range ? undefined : "end");
  }, [editor]);

  const left = rect.x + g.ml;
  const width = g.w - g.ml - g.mr;
  const style: React.CSSProperties =
    which === "header"
      ? { top: rect.y + g.headerY, left, width, minHeight: Math.max(24, g.mt - g.headerY - 4) }
      : { top: rect.y + g.h - g.mb + 4, left, width, height: Math.max(24, g.mb - g.footerY - 4), display: "flex", flexDirection: "column", justifyContent: "flex-end" };

  return (
    <div
      className={`bl-hf-edit bl-doc is-${which}`}
      style={style}
      onKeyDown={(e) => {
        if (e.key === "Escape") onClose();
      }}
    >
      <EditorContent editor={editor} />
    </div>
  );
}
