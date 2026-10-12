// Kabuk düzeyi yan etkiler: pencereye dosya bırakma, odağın belgeye dönmesi, sürüm sayacı.
import { useEffect, useRef, type MutableRefObject } from "react";
import type { Editor } from "@tiptap/core";

const DOC_FILE = /\.(docx|docm|dotx|html?|txt)$/i;

/** Pencereye bırakılan belge dosyası açılır (editör dışına bırakılsa da). */
export function useWindowFileDrop(openFileRef: MutableRefObject<(f: File) => Promise<void>>) {
  useEffect(() => {
    const isFiles = (e: DragEvent) => Array.from(e.dataTransfer?.types ?? []).includes("Files");
    const over = (e: DragEvent) => {
      if (isFiles(e)) e.preventDefault();
    };
    const drop = (e: DragEvent) => {
      if (!isFiles(e) || (e.target as HTMLElement).closest?.(".ProseMirror")) return;
      const f = Array.from(e.dataTransfer?.files ?? []).find((x) => DOC_FILE.test(x.name));
      e.preventDefault();
      if (f) void openFileRef.current(f);
    };
    window.addEventListener("dragover", over);
    window.addEventListener("drop", drop);
    return () => {
      window.removeEventListener("dragover", over);
      window.removeEventListener("drop", drop);
    };
  }, [openFileRef]);
}

/**
 * Şerit/menü etkileşiminden ya da iletişim kutusu kapandıktan sonra odak boşta kalırsa şeridin
 * hedefindeki düzenleyiciye döner (Word'deki gibi).
 */
export function useFocusReturn(active: Editor, dialogOpen: boolean) {
  const activeRef = useRef(active);
  activeRef.current = active;
  useEffect(() => {
    const onUp = (e: PointerEvent) => {
      const t = e.target as HTMLElement;
      if (t.closest?.(".bl-body, .bl-hf-edit, input, textarea, select, .oh-dialog, .bl-nav, .bl-comments, .oh-titlebar, .oh-backstage, .bl-read, .bl-pane, .bl-crash, .bl-minibar")) return;
      window.setTimeout(() => {
        const ae = document.activeElement as HTMLElement | null;
        // Şerit sekmesine tıklamak da klavye odağını belgeden almaz.
        const idle = !ae || ae === document.body || !!ae.closest?.(".oh-ribbon-tabs [role=tab]");
        if (idle && !document.querySelector(".oh-dialog, [data-oh-popover]") && !activeRef.current.isDestroyed) activeRef.current.view.focus();
      }, 0);
    };
    // Menü öğesine fareyle basmak belgeden odağı almaz: komut sonrası ilk tuşlar belgeye gider.
    const onDown = (e: MouseEvent) => {
      const t = e.target as HTMLElement;
      if (t.closest?.('[data-oh-popover] [role="menuitem"]') && activeRef.current.view.hasFocus()) e.preventDefault();
    };
    document.addEventListener("pointerup", onUp, true);
    document.addEventListener("mousedown", onDown, true);
    return () => {
      document.removeEventListener("pointerup", onUp, true);
      document.removeEventListener("mousedown", onDown, true);
    };
  }, []);

  // İletişim kutusu kapanınca odak belgeye döner (kutuyu açan şerit düğmesinde kalmaz).
  const hadDialog = useRef(false);
  useEffect(() => {
    if (hadDialog.current && !dialogOpen) {
      window.setTimeout(() => {
        const ae = document.activeElement as HTMLElement | null;
        if (!ae || !ae.closest?.(".ProseMirror, input, textarea, select, .oh-dialog, .oh-backstage, .bl-fn-text, .bl-note-text, .bl-pane")) activeRef.current.view.focus();
      }, 0);
    }
    hadDialog.current = dialogOpen;
  }, [dialogOpen]);
}

/** Belge sürümü (gezinti, açıklamalar, sözcük sayısı için; seyreltilmiş). */
export function useVersionBump(editor: Editor, bump: () => void) {
  useEffect(() => {
    let t = 0;
    const on = () => {
      window.clearTimeout(t);
      t = window.setTimeout(bump, 200);
    };
    editor.on("update", on);
    return () => {
      editor.off("update", on);
      window.clearTimeout(t);
    };
  }, [editor, bump]);
}
