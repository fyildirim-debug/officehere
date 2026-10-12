// Sayfa altındaki zengin dipnot metni: küçük bir düzenleyici (kalın, italik, bağlantı vb. şeritten uygulanır).
import { useEffect, useRef } from "react";
import { EditorContent, useEditor } from "@tiptap/react";
import { Editor, type JSONContent } from "@tiptap/core";
import { headerFooterExtensions } from "../editor/extensions";
import { noteDoc, noteEditorHost } from "../editor/notes";
import { getActiveNoteEditor, setActiveNoteEditor } from "../state/activeNote";

export interface FootnoteView {
  label: string;
  text: string;
  content: JSONContent[] | null;
}

const noteExtensions = () => headerFooterExtensions();

export function FootnoteText({ id, note, onChange }: { id: string; note: FootnoteView; onChange: (id: string, content: JSONContent) => void }) {
  const changeRef = useRef(onChange);
  changeRef.current = onChange;
  const editor = useEditor(
    {
      extensions: noteExtensions(),
      content: noteDoc(note),
      editorProps: { attributes: { class: "bl-fn-editor", "data-note-id": id, spellcheck: "true", lang: "tr", "aria-label": `Dipnot ${note.label}` } },
      onUpdate: ({ editor: ed }) => changeRef.current(id, ed.getJSON()),
      onFocus: ({ editor: ed }) => setActiveNoteEditor(ed),
    },
    [id],
  );
  // Ana belgede geri alma vb. ile içerik değişirse (düzenleyici odakta değilken) eşitle.
  useEffect(() => {
    if (!editor || editor.isDestroyed || editor.isFocused) return;
    const want = noteDoc(note);
    if (JSON.stringify(editor.getJSON().content) !== JSON.stringify(want.content)) editor.commands.setContent(want, { emitUpdate: false });
  }, [editor, note]);
  useEffect(
    () => () => {
      if (editor && getActiveNoteEditor() === editor) setActiveNoteEditor(null);
    },
    [editor],
  );
  return <EditorContent editor={editor} className="bl-fn-text" />;
}

// Son notlar alanı (belge sonundaki dekorasyon) için düzenleyici kurucu.
noteEditorHost.mount = (el, doc, onChange) => {
  const ed = new Editor({
    element: el,
    extensions: noteExtensions(),
    content: doc,
    editorProps: { attributes: { class: "bl-fn-editor", "data-note-id": el.dataset.noteId ?? "", spellcheck: "true", lang: "tr" } },
    onUpdate: ({ editor: e }) => onChange((e.getJSON().content ?? []) as JSONContent[]),
    onFocus: ({ editor: e }) => setActiveNoteEditor(e),
  });
  return () => {
    if (getActiveNoteEditor() === ed) setActiveNoteEditor(null);
    // Dekorasyon yenilenirken eşzamanlı yok etme ProseMirror güncellemesini bölmesin.
    window.setTimeout(() => ed.destroy(), 0);
  };
};
