// Resim ekleme: dosya seçici, sürükle-bırak ve panodan yapıştırma. Eklenen resim seçili gelir ve
// Resim Biçimi sekmesi açılır (Word'deki gibi).
import type { MutableRefObject } from "react";
import type { Editor } from "@tiptap/core";
import { NodeSelection } from "@tiptap/pm/state";
import type { EditorView } from "@tiptap/pm/view";
import type { DocSettings } from "../model/document";
import { readImageFile, selectInsertedImage } from "../editor/image";
import { columnWidthAt } from "./pageSetup";

interface Deps {
  editor: Editor;
  active: Editor;
  settings: DocSettings;
  notify: (t: string, kind?: "info" | "warning" | "error" | "success") => void;
  setTab: (t: string) => void;
  insertRef: MutableRefObject<(f: File, pos: number | null, view: EditorView) => Promise<void>>;
}

export function useImageInsert({ editor, active, settings, notify, setTab, insertRef }: Deps) {
  /** Konumdaki bölümün sütun genişliği (resimler buna sığdırılır). */
  const colWidth = (pos: number) => columnWidthAt(editor, settings.page, pos);

  insertRef.current = async (file, pos, view) => {
    try {
      const img = await readImageFile(file, colWidth(pos ?? view.state.selection.from));
      const node = view.state.schema.nodes.image.create({ src: img.src, width: img.width, height: img.height, alt: file.name.replace(/\.[^.]+$/, "") });
      const at = Math.min(pos ?? view.state.selection.from, view.state.doc.content.size);
      const tr = view.state.tr.insert(at, node);
      if (tr.doc.nodeAt(at)?.type.name === "image") tr.setSelection(NodeSelection.create(tr.doc, at));
      view.dispatch(tr);
      setTab("picture");
    } catch (err) {
      notify((err as Error).message, "error");
    }
  };

  /** Ekle › Resimler › Bu Cihaz dosya girişi. */
  return async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    try {
      const img = await readImageFile(file, colWidth(editor.state.selection.from));
      active
        .chain()
        .focus()
        .setImage({ src: img.src, width: img.width, height: img.height, alt: file.name.replace(/\.[^.]+$/, "") })
        .run();
      selectInsertedImage(active);
      setTab("picture");
    } catch (err) {
      notify((err as Error).message, "error");
    }
  };
}
