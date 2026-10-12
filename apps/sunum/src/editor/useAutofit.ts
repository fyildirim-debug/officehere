/**
 * Otomatik sığdırma ölçümü: "şekli metne uydur" ve "taşan metni küçült"
 * ayarlı metin kutularıyla tablo satır yükseklikleri her çizimden sonra
 * DOM'dan ölçülür; değişiklikler geçmişe eklenmeden (sessiz) uygulanır.
 *
 * Kararlılık: küçültme adımı görünmeyen bir ölçüm kutusunda, o anki ölçekten
 * bağımsız olarak ikili aramayla tek geçişte bulunur ve (metin + kutu)
 * anahtarıyla önbelleğe alınır. Aynı girdi her zaman aynı sonucu verdiği için
 * uygula → yeniden çiz → ölç döngüsü en fazla bir kez döner. Yine de bir
 * şekil kısa sürede tekrar tekrar güncellenirse koruma döngüyü keser.
 */
import { useLayoutEffect, useRef, type RefObject } from "react";
import { effectiveBodyProps, isTextEmpty, levelResolver, type TextContext } from "../model/textStyle";
import type { Shape, Slide, SpShape, TextBody, Theme } from "../model/types";
import { isLineGeom } from "../render/geometry";
import { paragraphsHtml, type TextRenderCtx } from "../render/textHtml";
import type { Actions } from "../state/actions";
import { CommitGuard, currentStep, objKey, pickShrinkStep, SHRINK_STEPS } from "./autofitFit";

/** Görünmeyen bir kopyada küçültme adımını ölçer. */
function measureShrink(tb: HTMLElement, body: TextBody, trc: TextRenderCtx): number {
  const host = tb.parentElement;
  if (!host) return 0;
  const cs = getComputedStyle(tb);
  const vertical = cs.writingMode.startsWith("vertical");
  const boxW = parseFloat(tb.style.width) || tb.clientWidth;
  const boxH = parseFloat(tb.style.height) || tb.clientHeight;
  const m = document.createElement("div");
  m.className = "tb tb-measure";
  m.setAttribute("aria-hidden", "true");
  m.style.cssText = `position:absolute;left:0;top:0;visibility:hidden;pointer-events:none;white-space:${cs.whiteSpace};writing-mode:${cs.writingMode};text-orientation:${cs.textOrientation};${vertical ? `height:${boxH}px;width:auto` : `width:${boxW}px;height:auto`}`;
  const inner = document.createElement("div");
  inner.className = "tb-c";
  m.appendChild(inner);
  host.appendChild(m);
  try {
    const sizeAt = (i: number) => {
      const st = SHRINK_STEPS[i];
      inner.innerHTML = paragraphsHtml(body, { ...trc, fontScale: st.scale < 1 ? st.scale : undefined, lnReduction: st.red || undefined, editing: true, prompt: undefined });
      return vertical ? inner.scrollWidth : inner.scrollHeight;
    };
    return pickShrinkStep(sizeAt, vertical ? boxW : boxH);
  } finally {
    m.remove();
  }
}

export function useAutofit({ slide, slideIdx, measuring, tctx, theme, slideRef, actions }: { slide: Slide | undefined; slideIdx: number; measuring: boolean; tctx: TextContext; theme: Theme; slideRef: RefObject<HTMLDivElement | null>; actions: Actions }) {
  const cache = useRef(new WeakMap<object, { key: string; step: number }>());
  const guard = useRef(new CommitGuard());

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
        if (bp.autofit === "shape" && !isLineGeom(s.geom)) {
          const contentH = inner.scrollHeight;
          let needW = s.w;
          if (!bp.wrap && !s.rot) {
            const prevW = inner.style.width;
            inner.style.width = "max-content";
            needW = Math.ceil(inner.offsetWidth) + bp.insets.l + bp.insets.r + 2;
            inner.style.width = prevW;
          }
          const needH = Math.ceil(contentH + bp.insets.t + bp.insets.b);
          if ((Math.abs(needH - s.h) > 1 || Math.abs(needW - s.w) > 1) && guard.current.allow(s.id)) {
            updates.set(s.id, (sh) => ({ ...sh, h: needH, w: needW }));
          }
        } else if (bp.autofit === "shrink") {
          const step = shrinkStep(s, s.text, tb, bp.vert);
          if (currentStep(s.text.fontScale, s.text.lnSpcReduction) !== step && guard.current.allow(s.id)) {
            const st = SHRINK_STEPS[step];
            updates.set(s.id, (sh) => (sh.kind === "sp" && sh.text ? { ...sh, text: { ...sh.text, fontScale: st.scale < 1 ? st.scale : undefined, lnSpcReduction: st.red || undefined } } : sh));
          }
        }
      } else if (s.kind === "table") {
        const tbl = root.querySelector(`[data-shape-id="${s.id}"] table`) as HTMLTableElement | null;
        if (!tbl) continue;
        const rows = Array.from(tbl.rows);
        const hs = rows.map((r) => r.offsetHeight);
        const changed = hs.length === s.rows.length && hs.some((h, i) => Math.abs(h - s.rows[i].h) > 1);
        if (changed) {
          if (guard.current.allow(s.id))
            updates.set(s.id, (sh) => {
              if (sh.kind !== "table") return sh;
              // Satır içerikle büyür; içerik azalınca en küçük yüksekliğe (minH) kadar geri küçülür.
              const nrows = sh.rows.map((r, i) => {
                const minH = r.minH ?? r.h;
                const h = Math.max(minH, hs[i]);
                return { ...r, h, minH: h > minH + 0.5 ? minH : undefined };
              });
              return { ...sh, rows: nrows, h: nrows.reduce((t, r) => t + r.h, 0) };
            });
        } else {
          const total = tbl.offsetHeight;
          if (Math.abs(total - s.h) > 1 && guard.current.allow(s.id)) updates.set(s.id, (sh) => ({ ...sh, h: total }));
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

    /** Önbellekli küçültme adımı: anahtar metin, kutu boyutu ve stil zinciridir (ölçek değil). */
    function shrinkStep(s: SpShape, body: TextBody, tb: HTMLElement, vert: string): number {
      if (isTextEmpty(body)) return 0;
      const key = [tb.style.width, tb.style.height, tb.style.whiteSpace, vert, objKey(tctx.master), objKey(tctx.layout), objKey(theme), objKey(body.levels), s.ph?.type ?? "", s.ph?.idx ?? "", slideIdx].join("|");
      const hit = cache.current.get(body.paragraphs);
      if (hit && hit.key === key) return hit.step;
      const trc: TextRenderCtx = { theme, lvl: levelResolver(s, tctx), slideNumber: slideIdx + 1 };
      const step = measureShrink(tb, body, trc);
      cache.current.set(body.paragraphs, { key, step });
      return step;
    }
  });
}
