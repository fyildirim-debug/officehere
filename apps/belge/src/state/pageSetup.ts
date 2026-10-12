// Sayfa yapısını uygulama: tek bölümlü belgede belge varsayılanı, çok bölümlüde imleçteki bölüm
// (Word: "Uygulama yeri: Bu bölüm / Bu noktadan sonra / Tüm belge").
import type { Editor } from "@tiptap/core";
import { samePage, type DocSettings, type PageSetup, type SectionAttrs } from "../model/document";
import { insertSectionBreak, sectionAt } from "../editor/sections";

export type PageScope = "auto" | "section" | "all" | "forward";

/** İmlecin (ya da verilen konumun) bulunduğu bölüm ve etkin sayfa yapısı. */
export function sectionInfo(editor: Editor, def: PageSetup, pos?: number) {
  const doc = editor.state.doc;
  const at = sectionAt(doc, pos ?? editor.state.selection.from);
  const page = (at?.node.attrs.page as PageSetup | null) ?? def;
  return { index: at?.index ?? 0, pos: at?.pos ?? 0, node: at?.node ?? null, page, count: doc.childCount };
}

/** Sıradaki bölümün etkin sayfa yapısı. */
export function sectionPageAt(editor: Editor, def: PageSetup, index: number): PageSetup {
  const doc = editor.state.doc;
  if (index < 0 || index >= doc.childCount) return def;
  return (doc.child(index).attrs.page as PageSetup | null) ?? def;
}

export function applyPageSetup(
  editor: Editor,
  settings: DocSettings,
  updateSettings: (fn: (s: DocSettings) => DocSettings) => void,
  fn: (p: PageSetup) => PageSetup,
  scope: PageScope = "auto",
  sectionIndex?: number,
) {
  const doc = editor.state.doc;
  if (scope === "forward") {
    // Bu noktadan sonra: imleçte yeni bölüm (sonraki sayfa) başlar ve yalnız ona uygulanır.
    if (!insertSectionBreak(editor.state, editor.view.dispatch, "nextPage")) return;
    applyPageSetup(editor, settings, updateSettings, fn, "section");
    return;
  }
  const multi = doc.childCount > 1;
  const idx = sectionIndex ?? sectionAt(doc, editor.state.selection.from)?.index ?? 0;
  const own = doc.child(Math.min(idx, doc.childCount - 1)).attrs.page as PageSetup | null;
  if (scope === "all" || (scope === "auto" && !multi && !own)) {
    updateSettings((s) => ({ ...s, page: fn(s.page) }));
    // Bölümlerin kendi sayfa yapıları kaldırılır: hepsi belge varsayılanını kullanır.
    const tr = editor.state.tr;
    let off = 0;
    let changed = false;
    doc.forEach((sec) => {
      if (sec.attrs.page) {
        tr.setNodeMarkup(off, undefined, { ...(sec.attrs as SectionAttrs), page: null });
        changed = true;
      }
      off += sec.nodeSize;
    });
    if (changed) editor.view.dispatch(tr);
    return;
  }
  let off = 0;
  for (let i = 0; i < idx; i++) off += doc.child(i).nodeSize;
  const sec = doc.child(idx);
  const next = fn(own ?? settings.page);
  const page = samePage(next, settings.page) ? null : next;
  editor.view.dispatch(editor.state.tr.setNodeMarkup(off, undefined, { ...(sec.attrs as SectionAttrs), page }));
}
