// Biçim Boyacısı: kopyalanan biçim bir sonraki seçime (kilitliyse Esc'e dek her seçime) uygulanır.
import { useCallback, useEffect, useState } from "react";
import type { Editor } from "@tiptap/core";
import { copyFormat, pasteFormat, type CopiedFormat } from "../editor/commands";

export type PainterState = { active: boolean; locked: boolean; fmt: CopiedFormat | null };

const OFF: PainterState = { active: false, locked: false, fmt: null };

export function useFormatPainter(editor: Editor) {
  const [painter, setPainter] = useState<PainterState>(OFF);

  const startFormatPainter = useCallback(
    (lock: boolean) => {
      if (painter.active && !lock) return setPainter(OFF);
      setPainter({ active: true, locked: lock, fmt: copyFormat(editor) });
    },
    [editor, painter.active],
  );

  useEffect(() => {
    if (!painter.active || !painter.fmt) return;
    const dom = editor.view.dom;
    // Hedef aralık fare konumlarından bulunur (seçimin güncellenmesini beklemeye gerek kalmaz):
    // tıklama sözcüğe, sürükleme basılan ve bırakılan noktalar arasına uygulanır.
    let down: number | null = null;
    const at = (e: MouseEvent) => editor.view.posAtCoords({ left: e.clientX, top: e.clientY })?.pos ?? null;
    const onDown = (e: MouseEvent) => {
      down = at(e);
    };
    const onUp = (e: MouseEvent) => {
      const up = at(e);
      const a = down;
      down = null;
      requestAnimationFrame(() => {
        const range = a !== null && up !== null ? { from: Math.min(a, up), to: Math.max(a, up) } : undefined;
        pasteFormat(editor, painter.fmt!, range);
        if (!painter.locked) setPainter(OFF);
      });
    };
    const onKey = (e: KeyboardEvent) => e.key === "Escape" && setPainter(OFF);
    dom.addEventListener("mousedown", onDown);
    dom.addEventListener("mouseup", onUp);
    window.addEventListener("keydown", onKey);
    return () => {
      dom.removeEventListener("mousedown", onDown);
      dom.removeEventListener("mouseup", onUp);
      window.removeEventListener("keydown", onKey);
    };
  }, [editor, painter]);

  return { painter, startFormatPainter };
}
