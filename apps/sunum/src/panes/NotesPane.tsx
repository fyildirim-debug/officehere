/** Konuşmacı notları bölmesi (yeniden boyutlanabilir ayraçla). */
import { useEffect, useState } from "react";
import type { Actions } from "../state/actions";
import type { AppState } from "../state/store";

export function NotesPane({ state, actions, height, spellcheck }: { state: AppState; actions: Actions; height: number; spellcheck: boolean }) {
  const slide = state.pres.slides[state.slideIdx];
  const [text, setText] = useState(slide?.notes ?? "");
  useEffect(() => setText(slide?.notes ?? ""), [slide?.id, slide?.notes]);
  return (
    <div className="notes-pane" style={{ height }}>
      <textarea
        className="notes-text oh-scroll"
        value={text}
        placeholder="Notlar eklemek için tıklayın"
        spellCheck={spellcheck}
        lang="tr"
        disabled={!slide}
        onFocus={() => actions.store.set({ focus: "notes", editing: null })}
        onChange={(e) => {
          setText(e.target.value);
          actions.setNotes(e.target.value);
        }}
        onKeyDown={(e) => {
          if (!e.ctrlKey && !e.metaKey && !e.key.startsWith("F")) e.stopPropagation();
        }}
      />
    </div>
  );
}

/** Sürüklenebilir yatay/dikey ayraç. */
export function Splitter({ dir, onDrag }: { dir: "h" | "v"; onDrag: (delta: number) => void }) {
  return (
    <div
      className={`splitter is-${dir}`}
      onPointerDown={(e) => {
        e.preventDefault();
        let last = dir === "h" ? e.clientY : e.clientX;
        const move = (ev: PointerEvent) => {
          const cur = dir === "h" ? ev.clientY : ev.clientX;
          onDrag(cur - last);
          last = cur;
        };
        const up = () => {
          window.removeEventListener("pointermove", move);
          window.removeEventListener("pointerup", up);
        };
        window.addEventListener("pointermove", move);
        window.addEventListener("pointerup", up);
      }}
    />
  );
}
