// Alan senkronizasyonu: kaynaklar/kaynakça stili alan bağlamına aktarılır; sayfalama değişince sayfa
// numarasına bağlı çapraz başvurular yenilenir. İçindekiler / Şekiller Tablosu güncelleme.
import { useCallback, useEffect } from "react";
import type { Editor } from "@tiptap/core";
import { TextSelection } from "@tiptap/pm/state";
import type { DocSettings } from "../model/document";
import type { PaginationController } from "../editor/pagination";
import { setFieldContext, updateFields } from "../editor/fields";
import { tocUpdate } from "../editor/references";

export function useFieldSync(editor: Editor, controller: PaginationController, settings: DocSettings) {
  useEffect(() => {
    setFieldContext(editor.view, { sources: settings.sources ?? [], bibStyle: settings.bibStyle ?? "apa" });
  }, [editor, settings.sources, settings.bibStyle]);
  useEffect(
    () =>
      controller.subscribe(() => {
        if (editor.isDestroyed) return;
        const tr = updateFields(editor.state, (p) => controller.pageOf(p));
        if (tr) editor.view.dispatch(tr);
      }),
    [editor, controller],
  );
  return useCallback(() => {
    const tr = tocUpdate(editor.state, (p) => controller.pageOf(p));
    if (tr) editor.view.dispatch(tr);
  }, [editor, controller]);
}

/** Üst/alt bilgi düzenlemesi kapanınca imleç gövdedeki eski konumuna döner; bağlamsal sekme seçilir. */
export function useHeaderFooterReturn(editor: Editor, hfEditing: "header" | "footer" | null, setTab: (fn: (t: string) => string) => void, hfReturn: { current: { from: number; to: number } | null }) {
  useEffect(() => {
    if (hfEditing) {
      setTab(() => "hf");
      hfReturn.current ??= { from: editor.state.selection.from, to: editor.state.selection.to };
      return;
    }
    setTab((t) => (t === "hf" ? "home" : t));
    const back = hfReturn.current;
    hfReturn.current = null;
    if (!back) return;
    const size = editor.state.doc.content.size;
    try {
      const { doc } = editor.state;
      editor.view.dispatch(editor.state.tr.setSelection(TextSelection.between(doc.resolve(Math.min(back.from, size)), doc.resolve(Math.min(back.to, size)))).setMeta("addToHistory", false));
    } catch {
      // konum geçersiz
    }
    requestAnimationFrame(() => editor.view.focus());
  }, [hfEditing, editor]); // eslint-disable-line react-hooks/exhaustive-deps
}
