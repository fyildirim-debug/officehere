// Resim ekleme: dosya seçici, sürükle-bırak ve panodan yapıştırma. Eklenen resim seçili gelir ve
// Resim Biçimi sekmesi açılır (Word'deki gibi). Aynı dosya girişi "Resmi Değiştir" için de kullanılır:
// seçilen dosya seçili resmin yerine geçer (genişlik, yerleşim ve biçim korunur).
import { useCallback, useEffect, useRef, type MutableRefObject, type RefObject } from "react";
import type { Editor } from "@tiptap/core";
import { NodeSelection } from "@tiptap/pm/state";
import type { EditorView } from "@tiptap/pm/view";
import type { DocSettings } from "../model/document";
import { readImageFile, selectInsertedImage } from "../editor/image";
import { selectedImage } from "../editor/imageWrap";
import { columnWidthAt } from "./pageSetup";

interface Deps {
  editor: Editor;
  active: Editor;
  settings: DocSettings;
  notify: (t: string, kind?: "info" | "warning" | "error" | "success") => void;
  setTab: (t: string) => void;
  insertRef: MutableRefObject<(f: File, pos: number | null, view: EditorView) => Promise<void>>;
  input: RefObject<HTMLInputElement | null>;
}

const baseName = (f: File) => f.name.replace(/\.[^.]+$/, "");

export function useImageInsert({ editor, active, settings, notify, setTab, insertRef, input }: Deps) {
  /** Konumdaki bölümün sütun genişliği (resimler buna sığdırılır). */
  const colWidth = (pos: number) => columnWidthAt(editor, settings.page, pos);
  /** Resmi Değiştir'in hedefi (dosya seçici açıkken). */
  const replaceTarget = useRef<{ editor: Editor; pos: number } | null>(null);

  insertRef.current = async (file, pos, view) => {
    try {
      const img = await readImageFile(file, colWidth(pos ?? view.state.selection.from));
      const node = view.state.schema.nodes.image.create({ src: img.src, width: img.width, height: img.height, alt: baseName(file) });
      const at = Math.min(pos ?? view.state.selection.from, view.state.doc.content.size);
      const tr = view.state.tr.insert(at, node);
      if (tr.doc.nodeAt(at)?.type.name === "image") tr.setSelection(NodeSelection.create(tr.doc, at));
      view.dispatch(tr);
      setTab("picture");
    } catch (err) {
      notify((err as Error).message, "error");
    }
  };

  // Dosya seçici iptal edilirse değiştirme hedefi unutulur (sonraki ekleme yeni resim ekler).
  useEffect(() => {
    const el = input.current;
    if (!el) return;
    const cancel = () => (replaceTarget.current = null);
    el.addEventListener("cancel", cancel);
    return () => el.removeEventListener("cancel", cancel);
  }, [input]);

  /** Dosya seçiciyi açar: yeni resim ekler ya da (seçili resim verilirse) onu değiştirir. */
  const pickImage = useCallback(
    (replace: Editor | null) => {
      const img = replace ? selectedImage(replace) : null;
      replaceTarget.current = replace && img ? { editor: replace, pos: img.pos } : null;
      input.current?.click();
    },
    [input],
  );

  const replace = async (file: File, target: { editor: Editor; pos: number }) => {
    const ed = target.editor;
    const old = ed.state.doc.nodeAt(target.pos);
    if (!old || old.type.name !== "image") return false;
    const img = await readImageFile(file, colWidth(target.pos));
    // Genişlik korunur, yükseklik yeni resmin oranına uyar; kırpma eski resme aittir.
    const width = (old.attrs.width as number | null) ?? img.width;
    const height = Math.max(1, Math.round((width * img.height) / Math.max(1, img.width)));
    const tr = ed.state.tr.setNodeMarkup(target.pos, undefined, { ...old.attrs, src: img.src, width, height, crop: null, alt: baseName(file) });
    tr.setSelection(NodeSelection.create(tr.doc, target.pos));
    ed.view.dispatch(tr);
    ed.view.focus();
    return true;
  };

  /** Ekle › Resimler › Bu Cihaz dosya girişi. */
  const onImageFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    const target = replaceTarget.current;
    replaceTarget.current = null;
    if (!file) return;
    try {
      if (target && (await replace(file, target))) return;
      const img = await readImageFile(file, colWidth(editor.state.selection.from));
      active.chain().focus().setImage({ src: img.src, width: img.width, height: img.height, alt: baseName(file) }).run();
      selectInsertedImage(active);
      setTab("picture");
    } catch (err) {
      notify((err as Error).message, "error");
    }
  };

  return { onImageFile, pickImage };
}
