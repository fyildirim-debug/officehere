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
  // Düzenlenen şeklin herhangi bir yerinde (metin bloğu dışında da) metin menüsü açılır.
  if (st.editing && !st.editing.cell && isInEditedShape(t, st, slideEl)) return "text";
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
  // Seçim çerçevesi ve tutamaçları seçili nesneye aittir.
  const onHandle = !!t.closest("[data-handle][data-shape]") && st.sel.length > 0;
  return onShape || onHandle ? "shape" : "slide";
}

/** Klavyeyle (Shift+F10, Menü tuşu) açılan menünün türü ve konumu: geçerli seçime göre. */
export function keyboardMenuTarget(st: AppState): { kind: CanvasMenuKind; x: number; y: number } {
  const slideEl = document.querySelector<HTMLElement>(".stage-slide");
  const clamp = (x: number, y: number) => ({ x: Math.max(4, Math.min(window.innerWidth - 4, x)), y: Math.max(4, Math.min(window.innerHeight - 4, y)) });
  const shapeEl = (id: string) => slideEl?.querySelector<HTMLElement>(`[data-shape-id="${CSS.escape(id)}"]`) ?? null;
  if (st.editing) {
    // Metin düzenlenirken menü imlecin/seçimin altında açılır.
    const sel = window.getSelection();
    const range = sel && sel.rangeCount ? sel.getRangeAt(0) : null;
    let r = range?.getBoundingClientRect();
    if (!r || (!r.width && !r.height)) r = range?.getClientRects()[0];
    if (!r || (!r.width && !r.height)) r = (document.querySelector(".tb-edit") ?? shapeEl(st.editing.shapeId))?.getBoundingClientRect();
    const p = r ? clamp(r.left, r.bottom) : { x: 0, y: 0 };
    return { kind: st.editing.cell ? "cell" : "text", ...p };
  }
  if (st.sel.length) {
    const r = shapeEl(st.sel[0])?.getBoundingClientRect();
    const p = r ? clamp(r.left + Math.min(r.width / 2, 40), r.top + Math.min(r.height / 2, 40)) : { x: 0, y: 0 };
    return { kind: st.cellSel ? "cell" : "shape", ...p };
  }
  const r = slideEl?.getBoundingClientRect();
  const p = r ? clamp(r.left + r.width / 2, r.top + r.height / 2) : { x: window.innerWidth / 2, y: window.innerHeight / 2 };
  return { kind: "slide", ...p };
}

/** Öğe, düzenlenmekte olan şeklin içinde mi? */
export function isInEditedShape(t: Element, st: AppState, slideEl: HTMLElement | null): boolean {
  const ed = st.editing;
  if (!ed || !slideEl) return false;
  const el = t.closest("[data-shape-id]") as HTMLElement | null;
  return !!el && slideEl.contains(el) && el.dataset.shapeId === ed.shapeId;
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
