/** Soldaki slayt küçük resimleri bölmesi: seçim, sürükle-bırak sıralama, bağlam menüsü. */
import { memo, useEffect, useRef, useState } from "react";
import type { Presentation, Slide } from "../model/types";
import { ScaledSlide } from "../render/SlideView";
import type { Actions } from "../state/actions";
import type { AppState } from "../state/store";

export interface ThumbnailPaneProps {
  state: AppState;
  actions: Actions;
  width: number;
  onContextMenu: (e: React.MouseEvent, onSlide: boolean) => void;
}

const Thumb = memo(function Thumb({ pres, slide, index, width }: { pres: Presentation; slide: Slide; index: number; width: number }) {
  return <ScaledSlide pres={pres} slide={slide} slideNumber={index + 1} width={width} />;
});

export function ThumbnailPane({ state, actions, width, onContextMenu }: ThumbnailPaneProps) {
  const pres = state.pres;
  const listRef = useRef<HTMLDivElement>(null);
  const [dropAt, setDropAt] = useState<number | null>(null);
  const dragRef = useRef<{ id: string; y: number; active: boolean } | null>(null);
  const thumbW = Math.max(60, width - 46);

  useEffect(() => {
    const el = listRef.current?.querySelector(`[data-index="${state.slideIdx}"]`) as HTMLElement | null;
    el?.scrollIntoView({ block: "nearest" });
  }, [state.slideIdx]);

  const indexFromY = (clientY: number): number => {
    const items = Array.from(listRef.current?.querySelectorAll<HTMLElement>(".thumb-item") ?? []);
    for (let i = 0; i < items.length; i++) {
      const r = items[i].getBoundingClientRect();
      if (clientY < r.top + r.height / 2) return i;
    }
    return items.length;
  };

  const onPointerDown = (e: React.PointerEvent, i: number) => {
    actions.store.set({ focus: "slides" });
    const slide = pres.slides[i];
    if (e.button === 2) {
      if (!state.slideSel.includes(slide.id)) actions.goToSlide(i);
      return;
    }
    if (e.button !== 0) return;
    if (e.ctrlKey || e.metaKey) {
      actions.goToSlide(i, "toggle");
      return;
    }
    if (e.shiftKey) {
      actions.goToSlide(i, "extend");
      return;
    }
    if (!state.slideSel.includes(slide.id) || state.slideSel.length <= 1) actions.goToSlide(i);
    else actions.store.set({ slideIdx: i, sel: [], editing: null });
    dragRef.current = { id: slide.id, y: e.clientY, active: false };
    (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
  };

  const onPointerMove = (e: React.PointerEvent) => {
    const d = dragRef.current;
    if (!d) return;
    if (!d.active && Math.abs(e.clientY - d.y) > 5) d.active = true;
    if (d.active) setDropAt(indexFromY(e.clientY));
  };

  const onPointerUp = (e: React.PointerEvent, i: number) => {
    const d = dragRef.current;
    dragRef.current = null;
    if (!d) return;
    if (d.active && dropAt !== null) {
      const ids = state.slideSel.includes(d.id) ? pres.slides.filter((s) => state.slideSel.includes(s.id)).map((s) => s.id) : [d.id];
      actions.moveSlidesTo(ids, dropAt);
    } else if (!d.active && state.slideSel.length > 1 && !e.ctrlKey && !e.shiftKey) {
      actions.goToSlide(i);
    }
    setDropAt(null);
  };

  return (
    <div
      className="thumb-pane oh-scroll"
      ref={listRef}
      style={{ width }}
      tabIndex={0}
      onContextMenu={(e) => {
        e.preventDefault();
        onContextMenu(e, !!(e.target as Element).closest(".thumb-item"));
      }}
      onPointerDown={(e) => {
        if (!(e.target as Element).closest(".thumb-item")) actions.store.set({ focus: "slides" });
      }}
    >
      {pres.slides.map((slide, i) => {
        const selected = state.slideSel.includes(slide.id);
        const current = i === state.slideIdx;
        const hasFx = (slide.transition && slide.transition.type !== "none") || slide.animations.length > 0;
        return (
          <div key={slide.id}>
            {dropAt === i && <div className="thumb-drop" />}
            <div
              className={`thumb-item${selected ? " is-selected" : ""}${current ? " is-current" : ""}${slide.hidden ? " is-hidden" : ""}`}
              data-index={i}
              onPointerDown={(e) => onPointerDown(e, i)}
              onPointerMove={onPointerMove}
              onPointerUp={(e) => onPointerUp(e, i)}
              onDoubleClick={() => actions.store.set({ focus: "canvas" })}
            >
              <div className="thumb-num">
                <span className="thumb-num-text">{i + 1}</span>
                {hasFx && (
                  <svg className="thumb-star" viewBox="0 0 12 12" width="11" height="11" aria-label="Geçiş/animasyon var">
                    <path d="M6 0.8l1.5 3.3 3.6.4-2.7 2.4.8 3.6L6 8.7 2.8 10.5l.8-3.6L.9 4.5l3.6-.4z" fill="currentColor" />
                  </svg>
                )}
              </div>
              <div className="thumb-box">
                <Thumb pres={pres} slide={slide} index={i} width={thumbW} />
              </div>
            </div>
          </div>
        );
      })}
      {dropAt === pres.slides.length && <div className="thumb-drop" />}
      {pres.slides.length === 0 && (
        <button type="button" className="thumb-empty" onClick={() => actions.newSlide()}>
          Yeni slayt eklemek için tıklayın
        </button>
      )}
    </div>
  );
}
