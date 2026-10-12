/** Asıl Slayt Görünümünde soldaki bölme: asıl slayt ve ona bağlı düzenler. */
import { memo, useEffect, useRef } from "react";
import type { Presentation, Slide } from "../model/types";
import { MASTER_SLIDE_ID } from "../model/masterView";
import { ScaledSlide } from "../render/SlideView";
import type { Actions } from "../state/actions";
import type { AppState } from "../state/store";
import { layoutUsageText } from "../state/masterActions";
import { sameThumb } from "./ThumbnailPane";

const Thumb = memo(function Thumb({ pres, slide, index, width }: { pres: Presentation; slide: Slide; index: number; width: number }) {
  return <ScaledSlide pres={pres} slide={slide} slideNumber={index + 1} width={width} />;
}, sameThumb);

export function MasterThumbPane({ state, actions: a, width, onContextMenu }: { state: AppState; actions: Actions; width: number; onContextMenu: (e: React.MouseEvent) => void }) {
  const pres = state.pres;
  const ref = useRef<HTMLDivElement>(null);
  const masterW = Math.max(60, width - 40);
  const layoutW = Math.max(50, masterW - 34);

  useEffect(() => {
    const el = ref.current?.querySelector(`[data-index="${state.slideIdx}"]`) as HTMLElement | null;
    el?.scrollIntoView({ block: "nearest" });
  }, [state.slideIdx]);

  return (
    <div
      className="thumb-pane master-pane oh-scroll"
      ref={ref}
      style={{ width }}
      tabIndex={0}
      aria-label="Asıl slayt ve düzenler"
      onPointerDown={() => a.store.set({ focus: "slides" })}
      onContextMenu={(e) => {
        e.preventDefault();
        onContextMenu(e);
      }}
    >
      {pres.slides.map((s, i) => {
        const isMaster = s.id === MASTER_SLIDE_ID;
        return (
          <div
            key={s.id}
            className={`thumb-item master-item${isMaster ? " is-master" : " is-layout"}${i === state.slideIdx ? " is-current is-selected" : ""}`}
            data-index={i}
            title={layoutUsageText(a, s.id)}
            onPointerDown={(e) => {
              if (e.button === 0 || i !== state.slideIdx) a.goToSlide(i);
              a.store.set({ focus: "slides" });
            }}
          >
            <div className="thumb-num">
              <span className="thumb-num-text">{isMaster ? "1" : ""}</span>
            </div>
            <div className="thumb-box">
              <Thumb pres={pres} slide={s} index={i} width={isMaster ? masterW : layoutW} />
            </div>
          </div>
        );
      })}
    </div>
  );
}
