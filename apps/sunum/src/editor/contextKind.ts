/** Slayt alanında sağ tıklanan yere göre hangi bağlam menüsünün açılacağı. */
import { useMemo } from "react";
import { findShape } from "../model/ops";
import type { Shape, Slide } from "../model/types";
import type { Actions } from "../state/actions";
import { animationLabels } from "./canvasUtils";
import type { AppState } from "../state/store";

export type CanvasMenuKind = "shape" | "slide" | "text" | "cell";

export function contextKind(t: Element, st: AppState, slideEl: HTMLElement | null, actions: Actions): CanvasMenuKind {
  if (t.closest(".tb-edit")) return st.editing?.cell ? "cell" : "text";
  // Tablo hücresi: tıklanan hücre seçim dışındaysa o hücre seçilir.
  const td = t.closest("td[data-cell]") as HTMLElement | null;
  const tblEl = td?.closest("[data-shape-id]") as HTMLElement | null;
  if (td && tblEl && slideEl?.contains(td) && st.sel.includes(tblEl.dataset.shapeId!)) {
    const [r, c] = td.dataset.cell!.split(",").map(Number);
    const cs = st.cellSel;
    const inside = cs && r >= Math.min(cs.r1, cs.r2) && r <= Math.max(cs.r1, cs.r2) && c >= Math.min(cs.c1, cs.c2) && c <= Math.max(cs.c1, cs.c2);
    if (!inside) actions.store.set({ cellSel: { r1: r, c1: c, r2: r, c2: c }, editing: null });
    return "cell";
  }
  const onShape = !!t.closest("[data-shape-id]") && !!slideEl?.contains(t);
  return onShape ? "shape" : "slide";
}

/** Animasyon sekmesinde şekillerin yanında gösterilen sıra numaraları. */
export function useAnimBadges(show: boolean, slide: Slide | undefined) {
  return useMemo(() => {
    if (!show || !slide) return undefined;
    const labels = animationLabels(slide);
    const out: { shape: Shape; labels: string[] }[] = [];
    for (const [id, l] of labels) {
      const s = findShape(slide.shapes, id);
      if (s) out.push({ shape: s, labels: l });
    }
    return out;
  }, [show, slide]);
}
