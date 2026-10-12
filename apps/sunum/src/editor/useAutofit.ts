/**
 * Otomatik sığdırma ölçümü: "şekli metne uydur" ve "taşan metni küçült"
 * ayarlı metin kutularıyla tablo satır yükseklikleri her çizimden sonra
 * DOM'dan ölçülür; değişiklikler geçmişe eklenmeden (sessiz) uygulanır.
 */
import { useLayoutEffect, type RefObject } from "react";
import { effectiveBodyProps, type TextContext } from "../model/textStyle";
import type { Shape, Slide } from "../model/types";
import { isLineGeom } from "../render/geometry";
import type { Actions } from "../state/actions";

const FONT_STEPS = [1, 0.925, 0.85, 0.775, 0.7, 0.625, 0.55, 0.475, 0.4, 0.325, 0.25];

export function useAutofit({ slide, slideIdx, measuring, tctx, slideRef, actions }: { slide: Slide | undefined; slideIdx: number; measuring: boolean; tctx: TextContext; slideRef: RefObject<HTMLDivElement | null>; actions: Actions }) {
  useLayoutEffect(() => {
    if (!measuring || !slide) return;
    const root = slideRef.current;
    if (!root) return;
    const updates = new Map<string, (s: Shape) => Shape>();
    for (const s of slide.shapes) {
      if (s.kind === "sp" && s.text) {
        const bp = effectiveBodyProps(s, tctx);
        if (bp.autofit === "none") continue;
        const tb = root.querySelector(`[data-shape-id="${s.id}"] > .tb`) as HTMLElement | null;
        const inner = tb?.firstElementChild as HTMLElement | null;
        if (!tb || !inner) continue;
        const contentH = inner.scrollHeight;
        if (bp.autofit === "shape" && !isLineGeom(s.geom)) {
          let needW = s.w;
          if (!bp.wrap && !s.rot) {
            const prevW = inner.style.width;
            inner.style.width = "max-content";
            needW = Math.ceil(inner.offsetWidth) + bp.insets.l + bp.insets.r + 2;
            inner.style.width = prevW;
          }
          const needH = Math.ceil(contentH + bp.insets.t + bp.insets.b);
          if (Math.abs(needH - s.h) > 1 || Math.abs(needW - s.w) > 1) {
            updates.set(s.id, (sh) => ({ ...sh, h: needH, w: needW }));
          }
        } else if (bp.autofit === "shrink") {
          const boxH = tb.clientHeight;
          const scale = s.text.fontScale ?? 1;
          let next = scale;
          if (contentH > boxH + 1) {
            const target = (scale * boxH) / contentH;
            next = FONT_STEPS.find((f) => f <= target + 1e-6 && f < scale) ?? FONT_STEPS[FONT_STEPS.length - 1];
            if (next >= scale) next = FONT_STEPS.find((f) => f < scale) ?? scale;
          } else if (scale < 1) {
            const bigger = [...FONT_STEPS].reverse().find((f) => f > scale + 1e-6);
            if (bigger && (contentH * bigger) / scale <= boxH - 2) next = bigger;
          }
          if (next !== scale) {
            const red = next < 1 ? (next <= 0.7 ? 0.2 : 0.1) : 0;
            updates.set(s.id, (sh) => (sh.kind === "sp" && sh.text ? { ...sh, text: { ...sh.text, fontScale: next < 1 ? next : undefined, lnSpcReduction: red || undefined } } : sh));
          }
        }
      } else if (s.kind === "table") {
        const tbl = root.querySelector(`[data-shape-id="${s.id}"] table`) as HTMLTableElement | null;
        if (!tbl) continue;
        const rows = Array.from(tbl.rows);
        const hs = rows.map((r) => r.offsetHeight);
        const changed = hs.length === s.rows.length && hs.some((h, i) => Math.abs(h - s.rows[i].h) > 1);
        if (changed) {
          updates.set(s.id, (sh) => {
            if (sh.kind !== "table") return sh;
            const nrows = sh.rows.map((r, i) => ({ ...r, h: Math.max(r.h, hs[i]) }));
            return { ...sh, rows: nrows, h: nrows.reduce((t, r) => t + r.h, 0) };
          });
        } else {
          const total = tbl.offsetHeight;
          if (Math.abs(total - s.h) > 1) updates.set(s.id, (sh) => ({ ...sh, h: total }));
        }
      }
    }
    if (updates.size) {
      const i = slideIdx;
      actions.store.commit(
        (p) => {
          const sl = p.slides[i];
          if (!sl) return p;
          const slides = p.slides.slice();
          slides[i] = { ...sl, shapes: sl.shapes.map((sh) => (updates.has(sh.id) ? updates.get(sh.id)!(sh) : sh)) };
          return { ...p, slides };
        },
        { silent: true },
      );
    }
  });
}
