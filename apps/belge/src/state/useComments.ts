// Açıklamalar (ekle / sil / tümünü sil) ve Değişiklikleri İzle anahtarı.
import { useCallback, useState } from "react";
import type { Editor } from "@tiptap/core";
import { getUserName } from "@officehere/ui";
import type { DocSettings } from "../model/document";
import type { ViewOptions } from "./context";
import { wordAt } from "../editor/commands";
import { isTracking, setTracking } from "../editor/trackChanges";
import { initialsOf } from "../io/docx/reader";

interface CommentDeps {
  editor: Editor;
  updateSettings: (fn: (s: DocSettings) => DocSettings) => void;
  setView: (patch: Partial<ViewOptions>) => void;
  notify: (text: string) => void;
}

export function useComments({ editor, updateSettings, setView, notify }: CommentDeps) {
  const [activeComment, setActiveComment] = useState<string | null>(null);

  const newComment = useCallback(() => {
    let { from, to } = editor.state.selection;
    if (from === to) {
      const w = wordAt(editor.state.doc, from);
      if (!w) return notify("Açıklama eklemek için metin seçin.");
      from = w.from;
      to = w.to;
    }
    const id = `c${Date.now().toString(36)}`;
    const author = getUserName();
    editor.chain().setTextSelection({ from, to }).setComment(id).run();
    updateSettings((s) => ({ ...s, comments: { ...s.comments, [id]: { id, author, initials: initialsOf(author), date: new Date().toISOString(), text: "" } } }));
    setActiveComment(id);
    setView({ showComments: true });
  }, [editor, notify, updateSettings, setView]);

  const deleteComment = useCallback(
    (id: string) => {
      editor.commands.unsetComment(id);
      updateSettings((s) => {
        const c = { ...s.comments };
        delete c[id];
        return { ...s, comments: c };
      });
      setActiveComment(null);
    },
    [editor, updateSettings],
  );

  const deleteAllComments = useCallback(() => {
    const type = editor.state.schema.marks.comment;
    editor.view.dispatch(editor.state.tr.removeMark(0, editor.state.doc.content.size, type));
    updateSettings((s) => ({ ...s, comments: {} }));
  }, [editor, updateSettings]);

  const toggleTrack = useCallback(() => {
    setTracking(editor.view, !isTracking(editor.state), getUserName());
  }, [editor]);

  return { activeComment, setActiveComment, newComment, deleteComment, deleteAllComments, toggleTrack };
}
