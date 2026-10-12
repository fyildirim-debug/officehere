// Üst/alt bilgi düzenleyicisi: sayfanın üst ya da alt kenar boşluğunda açılan ikinci editör.
import { useEffect, useRef } from "react";
import { EditorContent, useEditor } from "@tiptap/react";
import type { Editor, JSONContent } from "@tiptap/core";
import { headerFooterExtensions } from "../editor/extensions";
import type { PageGeom } from "./PageSheets";

interface Props {
  which: "header" | "footer";
  page: number;
  g: PageGeom;
  content: JSONContent | null;
  onChange: (json: JSONContent) => void;
  onEditor: (ed: Editor | null) => void;
  onClose: () => void;
}

export function HeaderFooterEditor({ which, page, g, content, onChange, onEditor, onClose }: Props) {
  const changeRef = useRef(onChange);
  changeRef.current = onChange;
  const editor = useEditor(
    {
      extensions: headerFooterExtensions(),
      content: content ?? { type: "doc", content: [{ type: "paragraph" }] },
      autofocus: "end",
      editorProps: { attributes: { class: "bl-hf-editor", spellcheck: "true", lang: "tr" } },
      onUpdate: ({ editor: ed }) => changeRef.current(ed.getJSON()),
    },
    [which],
  );

  useEffect(() => {
    onEditor(editor);
    return () => onEditor(null);
  }, [editor, onEditor]);

  const top = page * g.stride;
  const style: React.CSSProperties =
    which === "header"
      ? { top: top + g.headerY, left: g.ml, width: g.w - g.ml - g.mr, minHeight: Math.max(24, g.mt - g.headerY - 4) }
      : { top: top + g.h - g.mb + 4, left: g.ml, width: g.w - g.ml - g.mr, height: Math.max(24, g.mb - g.footerY - 4), display: "flex", flexDirection: "column", justifyContent: "flex-end" };

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
