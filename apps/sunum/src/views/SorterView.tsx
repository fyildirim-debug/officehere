/** Slayt Sıralayıcısı ve Not Sayfası görünümleri. */
import { memo, useEffect, useRef, useState } from "react";
import type { Presentation, Slide } from "../model/types";
import { ScaledSlide } from "../render/SlideView";
import type { Actions } from "../state/actions";
import type { AppState } from "../state/store";

const Thumb = memo(function Thumb({ pres, slide, index, width }: { pres: Presentation; slide: Slide; index: number; width: number }) {
  return <ScaledSlide pres={pres} slide={slide} slideNumber={index + 1} width={width} />;
});

export function SorterView({ state, actions: a, onOpenSlide, onContextMenu }: { state: AppState; actions: Actions; onOpenSlide: (i: number) => void; onContextMenu: (e: React.MouseEvent) => void }) {
  const pres = state.pres;
  const w = Math.round(220 * (state.zoom / 66));
  const width = Math.max(120, Math.min(560, w));
  const ref = useRef<HTMLDivElement>(null);
  const [drop, setDrop] = useState<number | null>(null);
  const drag = useRef<{ id: string; x: number; y: number; active: boolean } | null>(null);

  useEffect(() => {
    ref.current?.focus();
  }, []);

  const indexAt = (x: number, y: number): number => {
    const items = Array.from(ref.current?.querySelectorAll<HTMLElement>(".sorter-item") ?? []);
    let best = items.length;
    for (let i = 0; i < items.length; i++) {
      const r = items[i].getBoundingClientRect();
      if (y < r.top - 10) {
        best = i;
        break;
      }
      if (y <= r.bottom + 10 && x < r.left + r.width / 2) {
        best = i;
        break;
      }
    }
    return best;
  };

  return (
    <div
      className="sorter oh-scroll"
      ref={ref}
      tabIndex={0}
      onFocus={() => a.store.set({ focus: "slides" })}
      onContextMenu={(e) => {
        e.preventDefault();
        onContextMenu(e);
      }}
      onPointerMove={(e) => {
        const d = drag.current;
        if (!d) return;
        if (!d.active && Math.hypot(e.clientX - d.x, e.clientY - d.y) > 6) d.active = true;
        if (d.active) setDrop(indexAt(e.clientX, e.clientY));
      }}
      onPointerUp={() => {
        const d = drag.current;
        drag.current = null;
        if (d?.active && drop !== null) {
          const ids = state.slideSel.includes(d.id) ? pres.slides.filter((s) => state.slideSel.includes(s.id)).map((s) => s.id) : [d.id];
          a.moveSlidesTo(ids, drop);
        }
        setDrop(null);
      }}
    >
      <div className="sorter-grid">
        {pres.slides.map((s, i) => {
          const sel = state.slideSel.includes(s.id);
          return (
            <div
              key={s.id}
              className={`sorter-item${sel ? " is-selected" : ""}${s.hidden ? " is-hidden" : ""}`}
              onPointerDown={(e) => {
                a.store.set({ focus: "slides" });
                if (e.button !== 0) {
                  if (!sel) a.goToSlide(i);
                  return;
                }
                if (e.ctrlKey || e.metaKey) a.goToSlide(i, "toggle");
                else if (e.shiftKey) a.goToSlide(i, "extend");
                else if (!sel) a.goToSlide(i);
                else a.store.set({ slideIdx: i });
                drag.current = { id: s.id, x: e.clientX, y: e.clientY, active: false };
              }}
              onDoubleClick={() => onOpenSlide(i)}
            >
              {drop === i && <div className="sorter-drop" />}
              <div className="thumb-box">
                <Thumb pres={pres} slide={s} index={i} width={width} />
              </div>
              <div className="sorter-meta">
                <span>
                  {(s.transition && s.transition.type !== "none") || s.animations.length ? "★ " : ""}
                  {s.transition?.advanceAfter !== undefined ? `${String(Math.floor(s.transition.advanceAfter / 60)).padStart(2, "0")}:${String(Math.round(s.transition.advanceAfter % 60)).padStart(2, "0")}` : ""}
                </span>
                <span style={{ textDecoration: s.hidden ? "line-through" : undefined }}>{i + 1}</span>
              </div>
            </div>
          );
        })}
        {drop === pres.slides.length && <div style={{ width: 3, background: "var(--oh-accent)" }} />}
      </div>
    </div>
  );
}

export function NotesPageView({ state, actions: a }: { state: AppState; actions: Actions }) {
  const pres = state.pres;
  const slide = pres.slides[state.slideIdx];
  const [text, setText] = useState(slide?.notes ?? "");
  useEffect(() => setText(slide?.notes ?? ""), [slide?.id, slide?.notes]);
  if (!slide) return <div className="notes-page-view" />;
  return (
    <div className="notes-page-view oh-scroll">
      <div className="notes-page">
        <div className="thumb-box">
          <ScaledSlide pres={pres} slide={slide} slideNumber={state.slideIdx + 1} width={560} />
        </div>
        <textarea
          value={text}
          placeholder="Notlar eklemek için tıklayın"
          onChange={(e) => {
            setText(e.target.value);
            a.setNotes(e.target.value);
          }}
          onFocus={() => a.store.set({ focus: "notes" })}
          onKeyDown={(e) => {
            if (!e.ctrlKey && !e.metaKey) e.stopPropagation();
          }}
        />
        <div style={{ alignSelf: "flex-end", color: "#666" }}>{state.slideIdx + 1}</div>
      </div>
    </div>
  );
}
