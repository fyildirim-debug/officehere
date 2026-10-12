// Odaktaki not (dipnot/son not) düzenleyicisi: şerit biçimlendirmesi bu düzenleyiciye uygulanır.
import { useSyncExternalStore } from "react";
import type { Editor } from "@tiptap/core";

let current: Editor | null = null;
const listeners = new Set<() => void>();

export function setActiveNoteEditor(ed: Editor | null) {
  if (current === ed) return;
  current = ed;
  listeners.forEach((l) => l());
}

export function getActiveNoteEditor(): Editor | null {
  return current && !current.isDestroyed ? current : null;
}

export function useActiveNoteEditor(): Editor | null {
  return useSyncExternalStore(
    (cb) => {
      listeners.add(cb);
      return () => listeners.delete(cb);
    },
    getActiveNoteEditor,
  );
}
