/**
 * Normal görünümün slayt alanı: gri zemin, ölçeklenmiş slayt, seçim,
 * sürükle-taşı/boyutlandır/döndür, kement seçimi, akıllı kılavuzlar,
 * şekil çizme ve metin düzenleyicinin yerleştirilmesi.
 */
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import { aabb, findShape, unionBox, type Box } from "../model/ops";
import { effectiveBodyProps, layoutOf, type TextContext } from "../model/textStyle";
import type { GroupShape, Shape, Slide, SpShape, TableShape, TextBody } from "../model/types";
import { isLineGeom } from "../render/geometry";
import { SlideView, type GeomOverride } from "../render/SlideView";
import type { TextRenderCtx } from "../render/textHtml";
import type { Actions } from "../state/actions";
import type { AppState } from "../state/store";
import { applyFormatPainter } from "../state/painter";
import { ADJUST_HANDLES, effectiveAdj } from "./adjust";
import { angleFromCenter, constrainLine, geomFromEnds, geomOf, HANDLES, lineEnds, normAngle, resizeGeom, rotateVec, type Geom } from "./interact";
import { Rulers } from "./Rulers";
import { SelectionOverlay } from "./SelectionOverlay";
import { snapMove, snapValue, type GuideLine } from "./snap";
import { TextEditor } from "./TextEditor";

export interface CanvasProps {
  state: AppState;
  actions: Actions;
  onContextMenu: (e: React.MouseEvent, kind: "shape" | "slide") => void;
  onPlaceholderAction: (action: "table" | "picture", shapeId: string) => void;
  onShapeDoubleClick: (s: Shape) => void;
  showAnimBadges: boolean;
  guides: { h: number[]; v: number[] };
  onGuidesChange: (g: { h: number[]; v: number[] }) => void;
  onZoomChange: (zoom: number, fit: boolean) => void;
}

type Pt = [number, number];

type Drag =
  | { kind: "pending" | "move"; id: string; start: Pt; startClient: Pt; wasSelected: boolean; textHit: boolean; cell?: [number, number]; ids: string[]; geoms: Map<string, Geom>; ctrl: boolean; point: { x: number; y: number } }
  | { kind: "resize"; hx: number; hy: number; start: Pt; geoms: Map<string, Geom>; kinds: Map<string, string>; crop?: { id: string; c: { l: number; t: number; r: number; b: number } } }
  | { kind: "rotate"; start: Pt; startAngle: number; geoms: Map<string, Geom> }
  | { kind: "end"; id: string; which: "start" | "end"; geom: Geom }
  | { kind: "adj"; id: string; index: number; geom: Geom; adj: Record<string, number>; geomName: string }
  | { kind: "lasso"; start: Pt; additive: boolean }
  | { kind: "insert"; start: Pt }
  | { kind: "cells"; id: string; start: [number, number] }
  | { kind: "guide"; dir: "h" | "v"; index: number };

const MARGIN = 40;
const FONT_STEPS = [1, 0.925, 0.85, 0.775, 0.7, 0.625, 0.55, 0.475, 0.4, 0.325, 0.25];

function topLevelOf(shapes: Shape[], id: string): Shape | undefined {
  for (const s of shapes) {
    if (s.id === id) return s;
    if (s.kind === "group" && findShape(s.children, id)) return s;
  }
  return undefined;
}

function childOfGroupContaining(g: GroupShape, id: string): Shape | undefined {
  return g.children.find((c) => c.id === id || (c.kind === "group" && findShape(c.children, id)));
}

export function animationLabels(slide: Slide): Map<string, string[]> {
  const out = new Map<string, string[]>();
  let click = 0;
  for (const a of slide.animations) {
    if (a.trigger === "click") click += 1;
    const list = out.get(a.shapeId) ?? [];
    list.push(String(click));
    out.set(a.shapeId, list);
  }
  return out;
}

export function Canvas(props: CanvasProps) {
  const { state, actions } = props;
  const pres = state.pres;
  const slide = pres.slides[state.slideIdx];
  const wrapRef = useRef<HTMLDivElement>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const stageRef = useRef<HTMLDivElement>(null);
  const slideRef = useRef<HTMLDivElement>(null);
  const drag = useRef<Drag | null>(null);
  const [overrides, setOverrides] = useState<Map<string, GeomOverride> | undefined>(undefined);
  const [guideLines, setGuideLines] = useState<GuideLine[]>([]);
  const [lasso, setLasso] = useState<Box | null>(null);
  const [preview, setPreview] = useState<(Box & { line?: boolean; flipH?: boolean; flipV?: boolean }) | null>(null);
  const [rotInfo, setRotInfo] = useState<{ x: number; y: number; text: string } | null>(null);
  const [viewport, setViewport] = useState({ w: 800, h: 500 });
  const [scrollPos, setScrollPos] = useState({ x: 0, y: 0 });
  const propsRef = useRef(props);
  propsRef.current = props;

  const z = state.zoom / 100;
  const rulerSize = state.showRuler ? 18 : 0;

  // Pencere boyutu ve "Pencereye Sığdır".
  useLayoutEffect(() => {
    const el = scrollRef.current;
    if (!el) return;
    const ro = new ResizeObserver(() => setViewport({ w: el.clientWidth, h: el.clientHeight }));
    ro.observe(el);
    setViewport({ w: el.clientWidth, h: el.clientHeight });
    return () => ro.disconnect();
  }, []);

  const fitZoom = useMemo(() => {
    const zz = Math.min((viewport.w - MARGIN * 2) / pres.width, (viewport.h - MARGIN * 2) / pres.height);
    return Math.max(10, Math.min(400, Math.floor(zz * 100)));
  }, [viewport, pres.width, pres.height]);

  useEffect(() => {
    if (state.fit && Math.abs(fitZoom - state.zoom) > 0.5) propsRef.current.onZoomChange(fitZoom, true);
  }, [state.fit, fitZoom, state.zoom]);

  const contentW = Math.max(viewport.w, pres.width * z + MARGIN * 2);
  const contentH = Math.max(viewport.h, pres.height * z + MARGIN * 2);
  const stageLeft = (contentW - pres.width * z) / 2;
  const stageTop = (contentH - pres.height * z) / 2;

  // Ctrl + tekerlek yakınlaştırma.
  useEffect(() => {
    const el = scrollRef.current;
    if (!el) return;
    const onWheel = (e: WheelEvent) => {
      if (!e.ctrlKey && !e.metaKey) return;
      e.preventDefault();
      const cur = propsRef.current.state.zoom;
      const next = Math.max(10, Math.min(400, Math.round(cur * (e.deltaY < 0 ? 1.1 : 1 / 1.1))));
      propsRef.current.onZoomChange(next, false);
    };
    el.addEventListener("wheel", onWheel, { passive: false });
    return () => el.removeEventListener("wheel", onWheel);
  }, []);

  const toSlide = useCallback((cx: number, cy: number): Pt => {
    const r = stageRef.current!.getBoundingClientRect();
    const zz = propsRef.current.state.zoom / 100;
    return [(cx - r.left) / zz, (cy - r.top) / zz];
  }, []);

  const tctx: TextContext = useMemo(() => ({ master: pres.master, layout: layoutOf(pres, slide) }), [pres, slide]);

  /* ------------------------- Ölçüm: otomatik sığdırma ------------------------- */
  useLayoutEffect(() => {
    if (overrides || !slide) return;
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
      const i = state.slideIdx;
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

  /* ------------------------- Metin düzenleyici ------------------------- */
  const editing = state.editing;
  const renderEditor = useCallback(
    (shape: SpShape | TableShape, trc: TextRenderCtx, cell?: [number, number]) => {
      const st = propsRef.current.state;
      const ed = st.editing;
      const body: TextBody | undefined = shape.kind === "sp" ? shape.text : cell ? shape.rows[cell[0]]?.cells[cell[1]]?.text : undefined;
      if (!body || !ed) return null;
      const key = `${shape.id}:${cell?.join(",") ?? ""}`;
      return (
        <TextEditor
          key={key}
          body={body}
          ctx={trc}
          initial={{ selectAll: ed.selectAll, replaceWith: ed.replaceWith, point: ed.point, range: ed.range }}
          spellcheck={st.spellcheck}
          onChange={(nb, merge) => {
            const i = actions.s.slideIdx;
            actions.store.commit(
              (p) => {
                const sl = p.slides[i];
                const update = (s: Shape): Shape => {
                  if (s.id === shape.id) {
                    if (s.kind === "sp") return { ...s, text: nb };
                    if (s.kind === "table" && cell)
                      return { ...s, rows: s.rows.map((r, ri) => (ri !== cell[0] ? r : { ...r, cells: r.cells.map((c, ci) => (ci === cell[1] ? { ...c, text: nb } : c)) })) };
                  }
                  if (s.kind === "group") return { ...s, children: s.children.map(update) };
                  return s;
                };
                const slides = p.slides.slice();
                slides[i] = { ...sl, shapes: sl.shapes.map(update) };
                return { ...p, slides };
              },
              { merge: merge ? `${merge}:${key}` : undefined },
            );
            if (ed.replaceWith !== undefined || ed.selectAll || ed.point || ed.range) actions.store.set({ editing: { shapeId: ed.shapeId, cell: ed.cell } });
          }}
          onExit={() => actions.stopEditing()}
          register={(h) => {
            actions.textHandle = h;
          }}
          onSelectionChange={() => actions.store.set((s) => ({ textTick: s.textTick + 1 }))}
          onTab={
            cell && shape.kind === "table"
              ? (shift) => {
                  const nCols = shape.cols.length;
                  const nRows = shape.rows.length;
                  let idx = cell[0] * nCols + cell[1] + (shift ? -1 : 1);
                  if (idx < 0) idx = 0;
                  if (idx >= nRows * nCols) {
                    // Son hücrede Tab: yeni satır ekle.
                    actions.updateShapes((s) =>
                      s.kind === "table" ? { ...s, rows: [...s.rows, { h: s.rows[s.rows.length - 1].h, cells: s.cols.map(() => ({ text: { paragraphs: [{ runs: [], level: 0, props: {} }] } })) }], h: s.h + s.rows[s.rows.length - 1].h } : s,
                    [shape.id]);
                  }
                  actions.store.set({ editing: { shapeId: shape.id, cell: [Math.floor(idx / nCols), idx % nCols], selectAll: true } });
                  return true;
                }
              : undefined
          }
          onUndo={() => actions.store.undo()}
          onRedo={() => actions.store.redo()}
          onCopyRich={(paras) => {
            actions.store.clipboard = { kind: "text", paragraphs: paras };
          }}
          getRichClipboard={() => actions.store.clipboard?.paragraphs ?? null}
        />
      );
    },
    [actions],
  );

  /* ------------------------- İşaretçi olayları ------------------------- */

  const selectedTop = (): Shape[] => {
    const st = propsRef.current.state;
    const sl = st.pres.slides[st.slideIdx];
    if (!sl) return [];
    return st.sel.map((id) => findShape(sl.shapes, id)).filter((x): x is Shape => !!x);
  };

  const onPointerDown = (e: React.PointerEvent) => {
    const st = propsRef.current.state;
    if (!slide) return;
    const target = e.target as Element;
    if (!stageRef.current) return;
    const pt = toSlide(e.clientX, e.clientY);
    actions.store.set({ focus: "canvas" });
    if (e.button === 2) {
      const shapeEl = target.closest("[data-shape-id]") as HTMLElement | null;
      if (shapeEl) {
        const top = topLevelOf(slide.shapes, shapeEl.dataset.shapeId!);
        if (top && !st.sel.includes(top.id) && !st.sel.some((id) => top.kind === "group" && findShape(top.children, id))) actions.select([top.id]);
      }
      return;
    }
    if (e.button !== 0) return;
    // Yerel metin seçimini ve odak kaymasını engelle; klavye odağı slayt alanında kalsın.
    e.preventDefault();
    if (document.activeElement instanceof HTMLElement && document.activeElement !== scrollRef.current) document.activeElement.blur();
    scrollRef.current?.focus({ preventScroll: true });
    (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);

    if (st.insert) {
      drag.current = { kind: "insert", start: pt };
      setPreview({ x: pt[0], y: pt[1], w: 0, h: 0, line: st.insert.kind === "shape" && isLineGeom(st.insert.geom ?? "") });
      e.preventDefault();
      return;
    }

    if (target.closest("[data-comment]")) {
      actions.store.set({ pane: "comments" });
      return;
    }
    const guideEl = target.closest("[data-guide]") as HTMLElement | null;
    if (guideEl) {
      drag.current = { kind: "guide", dir: guideEl.dataset.guide as "h" | "v", index: Number(guideEl.dataset.index) };
      e.preventDefault();
      return;
    }

    const handleEl = target.closest("[data-handle]") as HTMLElement | null;
    if (handleEl) {
      e.preventDefault();
      const handle = handleEl.dataset.handle!;
      const sid = handleEl.dataset.shape!;
      const sel = selectedTop();
      const geoms = new Map(sel.map((s) => [s.id, geomOf(s)]));
      if (handle === "move") {
        drag.current = { kind: "pending", id: sid, start: pt, startClient: [e.clientX, e.clientY], wasSelected: true, textHit: false, ids: sel.map((s) => s.id), geoms, ctrl: e.ctrlKey, point: { x: e.clientX, y: e.clientY } };
        return;
      }
      if (handle === "rot") {
        const g = geoms.get(sid)!;
        drag.current = { kind: "rotate", start: pt, startAngle: angleFromCenter(g, pt[0], pt[1]), geoms };
        return;
      }
      if (handle === "start" || handle === "end") {
        drag.current = { kind: "end", id: sid, which: handle, geom: geoms.get(sid)! };
        return;
      }
      if (handle.startsWith("adj")) {
        const s = sel.find((x) => x.id === sid) as SpShape;
        drag.current = { kind: "adj", id: sid, index: Number(handle.slice(3)), geom: geomOf(s), adj: effectiveAdj(s.geom, s.adj), geomName: s.geom };
        return;
      }
      const h = HANDLES.find((x) => x.id === handle);
      if (h) {
        const cropShape = st.cropId === sid ? sel.find((s) => s.id === sid && s.kind === "pic") : undefined;
        drag.current = {
          kind: "resize",
          hx: h.hx,
          hy: h.hy,
          start: pt,
          geoms: cropShape ? new Map([[sid, geoms.get(sid)!]]) : geoms,
          kinds: new Map(sel.map((s) => [s.id, s.kind])),
          crop: cropShape && cropShape.kind === "pic" ? { id: sid, c: cropShape.crop ?? { l: 0, t: 0, r: 0, b: 0 } } : undefined,
        };
        if (st.editing) actions.stopEditing();
      }
      return;
    }

    const phAct = target.closest("[data-ph-action]") as HTMLElement | null;
    if (phAct) {
      const sid = (phAct.closest("[data-shape-id]") as HTMLElement | null)?.dataset.shapeId;
      if (sid) propsRef.current.onPlaceholderAction(phAct.dataset.phAction as "table" | "picture", sid);
      return;
    }

    const shapeEl = target.closest("[data-shape-id]") as HTMLElement | null;
    if (shapeEl && slideRef.current?.contains(shapeEl)) {
      const hitId = shapeEl.dataset.shapeId!;
      const top = topLevelOf(slide.shapes, hitId);
      if (!top) return;
      let id = top.id;
      if (top.kind === "group" && (st.sel.includes(top.id) || st.sel.some((x) => findShape(top.children, x)))) {
        id = childOfGroupContaining(top, hitId)?.id ?? top.id;
      }
      const wasSelected = st.sel.includes(id);
      let sel = st.sel;
      if (e.shiftKey || e.ctrlKey) {
        if (!(e.ctrlKey && wasSelected)) sel = wasSelected ? st.sel.filter((x) => x !== id) : [...st.sel.filter((x) => findShape(slide.shapes, x) && topLevelOf(slide.shapes, x)?.id !== id), id];
      } else if (!wasSelected) sel = [id];
      if (st.editing && st.editing.shapeId !== id) actions.stopEditing();
      if (sel !== st.sel) actions.store.set({ sel, cellSel: null, editing: st.editing && sel.includes(st.editing.shapeId) ? st.editing : null });
      const td = target.closest("td[data-cell]") as HTMLElement | null;
      const cell = td ? (td.dataset.cell!.split(",").map(Number) as [number, number]) : undefined;
      const textHit = !!target.closest(".tb.is-hit") || (!!cell && wasSelected);
      const shapes = sel.map((x) => findShape(slide.shapes, x)).filter((x): x is Shape => !!x);
      if (cell && wasSelected && top.kind === "table") {
        drag.current = { kind: "cells", id: top.id, start: cell };
        actions.store.set({ cellSel: null });
        // Tıklama ise hücre düzenlenir (pointerup'ta).
        (drag.current as Drag & { pendingClick?: boolean }).pendingClick = true;
        (drag.current as Drag & { point?: { x: number; y: number } }).point = { x: e.clientX, y: e.clientY };
        return;
      }
      drag.current = {
        kind: "pending",
        id,
        start: pt,
        startClient: [e.clientX, e.clientY],
        wasSelected,
        textHit,
        cell,
        ids: shapes.map((s) => s.id),
        geoms: new Map(shapes.map((s) => [s.id, geomOf(s)])),
        ctrl: e.ctrlKey,
        point: { x: e.clientX, y: e.clientY },
      };
      return;
    }

    // Boş alan: seçimi kaldır, kement başlat.
    if (st.editing) actions.stopEditing();
    if (!e.shiftKey && !e.ctrlKey) actions.store.set({ sel: [], cellSel: null });
    if (st.painter && !st.painter.sticky) actions.store.set({ painter: null });
    drag.current = { kind: "lasso", start: pt, additive: e.shiftKey || e.ctrlKey };
  };

  const otherBoxes = (exclude: Set<string>): Box[] => (slide ? slide.shapes.filter((s) => !exclude.has(s.id) && !s.hidden).map(aabb) : []);

  const onPointerMove = (e: React.PointerEvent) => {
    const d = drag.current;
    if (!d) return;
    const pt = toSlide(e.clientX, e.clientY);
    const st = propsRef.current.state;
    const thr = 6 / (st.zoom / 100);
    switch (d.kind) {
      case "pending":
      case "move": {
        if (d.kind === "pending") {
          if (Math.hypot(e.clientX - d.startClient[0], e.clientY - d.startClient[1]) < 4) return;
          if (st.editing) actions.stopEditing();
          (d as Drag).kind = "move";
        }
        let dx = pt[0] - d.start[0];
        let dy = pt[1] - d.start[1];
        if (e.shiftKey) {
          if (Math.abs(dx) > Math.abs(dy)) dy = 0;
          else dx = 0;
        }
        let lines: GuideLine[] = [];
        if (!e.altKey) {
          const u = unionBox([...d.geoms.values()].map((g) => aabb({ ...g, x: g.x + dx, y: g.y + dy })));
          const snap = snapMove(u, otherBoxes(new Set(d.ids)), pres.width, pres.height, thr, st.showGuides ? propsRef.current.guides : undefined);
          if (!e.shiftKey || dx) dx += snap.dx;
          if (!e.shiftKey || dy) dy += snap.dy;
          lines = snap.lines;
        }
        const ov = new Map<string, GeomOverride>();
        for (const [id, g] of d.geoms) ov.set(id, { ...g, x: g.x + dx, y: g.y + dy });
        setOverrides(ov);
        setGuideLines(lines);
        break;
      }
      case "resize": {
        const dxw = pt[0] - d.start[0];
        const dyw = pt[1] - d.start[1];
        const ov = new Map<string, GeomOverride>();
        let lines: GuideLine[] = [];
        for (const [id, g] of d.geoms) {
          const [lx, ly] = rotateVec(dxw, dyw, -g.rot);
          const keep = !d.crop && (e.shiftKey || (d.kinds.get(id) === "pic" && d.hx !== 0 && d.hy !== 0));
          let ng = resizeGeom(g, d.hx, d.hy, lx, ly, { keepAspect: keep, fromCenter: e.ctrlKey, min: 1 });
          if (!e.altKey && !g.rot && d.geoms.size === 1 && !keep) {
            const others = otherBoxes(new Set([id]));
            const xs = [0, pres.width / 2, pres.width, ...others.flatMap((o) => [o.x, o.x + o.w / 2, o.x + o.w])];
            const ys = [0, pres.height / 2, pres.height, ...others.flatMap((o) => [o.y, o.y + o.h / 2, o.y + o.h])];
            if (d.hx === 1) {
              const r = snapValue(ng.x + ng.w, xs, thr);
              if (r.hit !== undefined) {
                ng = { ...ng, w: r.v - ng.x };
                lines.push({ dir: "v", pos: r.v, from: Math.min(0, ng.y), to: Math.max(pres.height, ng.y + ng.h) });
              }
            } else if (d.hx === -1) {
              const r = snapValue(ng.x, xs, thr);
              if (r.hit !== undefined) {
                ng = { ...ng, w: ng.x + ng.w - r.v, x: r.v };
                lines.push({ dir: "v", pos: r.v, from: Math.min(0, ng.y), to: Math.max(pres.height, ng.y + ng.h) });
              }
            }
            if (d.hy === 1) {
              const r = snapValue(ng.y + ng.h, ys, thr);
              if (r.hit !== undefined) {
                ng = { ...ng, h: r.v - ng.y };
                lines.push({ dir: "h", pos: r.v, from: Math.min(0, ng.x), to: Math.max(pres.width, ng.x + ng.w) });
              }
            } else if (d.hy === -1) {
              const r = snapValue(ng.y, ys, thr);
              if (r.hit !== undefined) {
                ng = { ...ng, h: ng.y + ng.h - r.v, y: r.v };
                lines.push({ dir: "h", pos: r.v, from: Math.min(0, ng.x), to: Math.max(pres.width, ng.x + ng.w) });
              }
            }
            lines = lines.filter((l) => l.to > l.from);
          }
          if (d.crop && d.crop.id === id) {
            const c = d.crop.c;
            const fw = g.w / Math.max(0.01, 1 - c.l - c.r);
            const fh = g.h / Math.max(0.01, 1 - c.t - c.b);
            const left0 = g.x - c.l * fw;
            const top0 = g.y - c.t * fh;
            const cl = (v: number) => Math.max(0, Math.min(0.98, v));
            const nc = { l: cl((ng.x - left0) / fw), t: cl((ng.y - top0) / fh), r: cl((left0 + fw - (ng.x + ng.w)) / fw), b: cl((top0 + fh - (ng.y + ng.h)) / fh) };
            ng = { ...ng, x: left0 + nc.l * fw, y: top0 + nc.t * fh, w: fw * (1 - nc.l - nc.r), h: fh * (1 - nc.t - nc.b), flipH: g.flipH, flipV: g.flipV };
            ov.set(id, { ...ng, crop: nc });
            continue;
          }
          ov.set(id, ng);
        }
        setOverrides(ov);
        setGuideLines(lines);
        break;
      }
      case "rotate": {
        const first = [...d.geoms.values()][0];
        const ang = angleFromCenter(first, pt[0], pt[1]);
        let delta = ang - d.startAngle;
        const ov = new Map<string, GeomOverride>();
        let shown = 0;
        for (const [id, g] of d.geoms) {
          let r = normAngle(g.rot + delta);
          if (e.shiftKey) r = normAngle(Math.round(r / 15) * 15);
          if (id === [...d.geoms.keys()][0]) shown = r;
          ov.set(id, { ...g, rot: r });
        }
        void delta;
        delta = 0;
        setOverrides(ov);
        const rr = stageRef.current!.getBoundingClientRect();
        setRotInfo({ x: e.clientX - rr.left + 14, y: e.clientY - rr.top + 14, text: `${Math.round(shown)}°` });
        break;
      }
      case "end": {
        const [a, b] = lineEnds(d.geom);
        let p: Pt = pt;
        const fixed = d.which === "start" ? b : a;
        if (e.shiftKey) p = constrainLine(fixed, p);
        const ng = d.which === "start" ? geomFromEnds(p, b) : geomFromEnds(a, p);
        setOverrides(new Map([[d.id, ng]]));
        break;
      }
      case "adj": {
        const g = d.geom;
        const cx = g.x + g.w / 2;
        const cy = g.y + g.h / 2;
        let [lx, ly] = rotateVec(pt[0] - cx, pt[1] - cy, -g.rot);
        lx += g.w / 2;
        ly += g.h / 2;
        if (g.flipH) lx = g.w - lx;
        if (g.flipV) ly = g.h - ly;
        const defs = ADJUST_HANDLES[d.geomName];
        if (!defs) break;
        const patch = defs[d.index].set(lx, ly, g.w, g.h, d.adj);
        const i = st.slideIdx;
        actions.store.commit(
          (p) => {
            const sl = p.slides[i];
            const slides = p.slides.slice();
            slides[i] = { ...sl, shapes: sl.shapes.map((s) => (s.id === d.id && s.kind === "sp" ? { ...s, adj: { ...d.adj, ...patch } } : s)) };
            return { ...p, slides };
          },
          { merge: `adj:${d.id}` },
        );
        break;
      }
      case "lasso": {
        const x = Math.min(d.start[0], pt[0]);
        const y = Math.min(d.start[1], pt[1]);
        setLasso({ x, y, w: Math.abs(pt[0] - d.start[0]), h: Math.abs(pt[1] - d.start[1]) });
        break;
      }
      case "insert": {
        let w = pt[0] - d.start[0];
        let h = pt[1] - d.start[1];
        const line = st.insert?.kind === "shape" && isLineGeom(st.insert.geom ?? "");
        if (line) {
          let end: Pt = pt;
          if (e.shiftKey) end = constrainLine(d.start, pt);
          const g = geomFromEnds(d.start, end);
          setPreview({ ...g, line: true });
          break;
        }
        if (e.shiftKey) {
          const m = Math.max(Math.abs(w), Math.abs(h));
          w = Math.sign(w || 1) * m;
          h = Math.sign(h || 1) * m;
        }
        setPreview({ x: Math.min(d.start[0], d.start[0] + w), y: Math.min(d.start[1], d.start[1] + h), w: Math.abs(w), h: Math.abs(h) });
        break;
      }
      case "cells": {
        const el = document.elementFromPoint(e.clientX, e.clientY)?.closest("td[data-cell]") as HTMLElement | null;
        if (!el || !el.closest(`[data-shape-id="${d.id}"]`)) break;
        const [r, c] = el.dataset.cell!.split(",").map(Number);
        if (r !== d.start[0] || c !== d.start[1]) {
          (d as Drag & { pendingClick?: boolean }).pendingClick = false;
          if (st.editing) actions.stopEditing();
          actions.store.set({ cellSel: { r1: d.start[0], c1: d.start[1], r2: r, c2: c }, sel: [d.id] });
        }
        break;
      }
      case "guide": {
        const g = { h: [...propsRef.current.guides.h], v: [...propsRef.current.guides.v] };
        const v = d.dir === "h" ? Math.max(0, Math.min(pres.height, pt[1])) : Math.max(0, Math.min(pres.width, pt[0]));
        g[d.dir][d.index] = v;
        propsRef.current.onGuidesChange(g);
        const rr = stageRef.current!.getBoundingClientRect();
        const cm = ((v - (d.dir === "h" ? pres.height : pres.width) / 2) / (96 / 2.54)).toLocaleString("tr-TR", { maximumFractionDigits: 2 });
        setRotInfo({ x: e.clientX - rr.left + 12, y: e.clientY - rr.top + 12, text: cm });
        break;
      }
    }
  };

  const onPointerUp = (e: React.PointerEvent) => {
    const d = drag.current;
    drag.current = null;
    const st = propsRef.current.state;
    setGuideLines([]);
    setRotInfo(null);
    if (!d) return;
    const ov = overrides;
    setOverrides(undefined);
    switch (d.kind) {
      case "pending": {
        if (st.painter) {
          applyFormatPainter(actions, d.id);
          return;
        }
        const sh = slide ? findShape(slide.shapes, d.id) : undefined;
        if (!sh || e.shiftKey || e.ctrlKey) return;
        if (sh.kind === "table" && d.cell) {
          actions.startEditing({ shapeId: sh.id, cell: d.cell, point: d.point });
          return;
        }
        const top = slide ? topLevelOf(slide.shapes, d.id) : undefined;
        if (sh.kind === "sp" && d.textHit && !(top?.kind === "group" && top.id === sh.id)) {
          if (st.editing?.shapeId === sh.id) return;
          actions.startEditing({ shapeId: sh.id, point: d.point });
        }
        return;
      }
      case "move": {
        if (!ov) return;
        const map = new Map<string, Box & { rot?: number }>();
        for (const [id, g] of ov) map.set(id, g);
        if (d.ctrl) {
          // Ctrl+sürükle: kopya bırak.
          const first = ov.get(d.ids[0])!;
          const g0 = d.geoms.get(d.ids[0])!;
          const dx = first.x - g0.x;
          const dy = first.y - g0.y;
          actions.duplicateSelection();
          actions.moveSelection(dx - 24, dy - 24);
          return;
        }
        actions.setGeometries(map);
        return;
      }
      case "resize":
      case "rotate":
      case "end": {
        if (d.kind === "resize" && d.crop && ov) {
          const o = ov.get(d.crop.id);
          if (o?.crop) {
            const nc = o.crop;
            actions.updateShapes((s) => (s.kind === "pic" ? { ...s, x: o.x, y: o.y, w: o.w, h: o.h, crop: nc.l || nc.t || nc.r || nc.b ? nc : undefined } : s), [d.crop.id]);
          }
          return;
        }
        if (!ov) return;
        const map = new Map<string, Box & { rot?: number; flipH?: boolean; flipV?: boolean }>();
        for (const [id, g] of ov) map.set(id, g);
        actions.setGeometries(map);
        return;
      }
      case "lasso": {
        const box = lasso;
        setLasso(null);
        if (!box || !slide || (box.w < 2 && box.h < 2)) return;
        const ids = slide.shapes
          .filter((s) => !s.hidden)
          .filter((s) => {
            const b = aabb(s);
            return b.x >= box.x && b.y >= box.y && b.x + b.w <= box.x + box.w && b.y + b.h <= box.y + box.h;
          })
          .map((s) => s.id);
        actions.store.set({ sel: d.additive ? Array.from(new Set([...st.sel, ...ids])) : ids, editing: null });
        return;
      }
      case "insert": {
        const pv = preview;
        setPreview(null);
        const ins = st.insert;
        actions.store.set({ insert: null });
        if (!ins || !pv) return;
        const tiny = pv.w * z < 4 && pv.h * z < 4;
        if (ins.kind === "textbox") {
          if (tiny) actions.insertTextBox({ x: d.start[0], y: d.start[1], w: 24, h: 40 }, true, ins.vertical);
          else actions.insertTextBox({ x: pv.x, y: pv.y, w: Math.max(24, pv.w), h: Math.max(30, pv.h) }, false, ins.vertical);
          return;
        }
        if (ins.kind === "shape" && ins.geom) {
          if (pv.line) {
            const box = tiny ? { x: d.start[0], y: d.start[1], w: 96, h: 0 } : pv;
            actions.insertShape(ins.geom, box, { head: ins.head, tail: ins.tail, flipH: tiny ? false : pv.flipH, flipV: tiny ? false : pv.flipV });
          } else {
            const box = tiny ? { x: d.start[0], y: d.start[1], w: 96, h: 96 } : pv;
            actions.insertShape(ins.geom, box);
          }
        }
        return;
      }
      case "cells": {
        const dd = d as Drag & { pendingClick?: boolean; point?: { x: number; y: number } };
        if (dd.pendingClick) actions.startEditing({ shapeId: d.id, cell: d.start, point: dd.point });
        return;
      }
      default:
        return;
    }
  };

  const onDoubleClick = (e: React.MouseEvent) => {
    const target = e.target as Element;
    const shapeEl = target.closest("[data-shape-id]") as HTMLElement | null;
    if (!shapeEl || !slide) return;
    const sh = findShape(slide.shapes, shapeEl.dataset.shapeId!);
    if (!sh) return;
    const top = topLevelOf(slide.shapes, sh.id);
    const target2 = top?.kind === "group" && !state.sel.includes(sh.id) ? top : sh;
    if (target2.kind === "sp" && !isLineGeom(target2.geom)) {
      if (state.editing?.shapeId === target2.id) return;
      actions.startEditing({ shapeId: target2.id, point: { x: e.clientX, y: e.clientY } });
    }
    props.onShapeDoubleClick(target2);
  };

  const onContextMenu = (e: React.MouseEvent) => {
    const t = e.target as Element;
    if (t.closest(".tb-edit")) return;
    e.preventDefault();
    const onShape = !!t.closest("[data-shape-id]") && !!slideRef.current?.contains(t);
    props.onContextMenu(e, onShape ? "shape" : "slide");
  };

  /* ------------------------- Çizim ------------------------- */

  const selShapes = slide ? state.sel.map((id) => findShape(slide.shapes, id)).filter((x): x is Shape => !!x) : [];
  const parentGroup = slide && selShapes.length ? slide.shapes.find((s) => s.kind === "group" && selShapes.some((x) => x.id !== s.id && findShape(s.children, x.id))) : null;
  const badges = useMemo(() => {
    if (!props.showAnimBadges || !slide) return undefined;
    const labels = animationLabels(slide);
    const out: { shape: Shape; labels: string[] }[] = [];
    for (const [id, l] of labels) {
      const s = findShape(slide.shapes, id);
      if (s) out.push({ shape: s, labels: l });
    }
    return out;
  }, [props.showAnimBadges, slide]);

  const cellRects = useMemo(() => {
    const cs = state.cellSel;
    if (!cs || !slide || selShapes.length !== 1 || selShapes[0].kind !== "table") return null;
    const t = selShapes[0] as TableShape;
    const r1 = Math.min(cs.r1, cs.r2);
    const r2 = Math.max(cs.r1, cs.r2);
    const c1 = Math.min(cs.c1, cs.c2);
    const c2 = Math.max(cs.c1, cs.c2);
    let y = t.y;
    for (let r = 0; r < r1; r++) y += t.rows[r].h;
    let x = t.x;
    for (let c = 0; c < c1; c++) x += t.cols[c];
    let h = 0;
    for (let r = r1; r <= r2; r++) h += t.rows[r]?.h ?? 0;
    let w = 0;
    for (let c = c1; c <= c2; c++) w += t.cols[c] ?? 0;
    return { shape: t, rects: [{ x, y, w, h }] };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [state.cellSel, slide, state.sel]);

  const cursor = state.insert ? "crosshair" : state.painter ? "copy" : undefined;
  const cropGhost = (() => {
    if (!state.cropId || !slide) return null;
    const pic = findShape(slide.shapes, state.cropId);
    if (!pic || pic.kind !== "pic") return null;
    const g = overrides?.get(pic.id) ?? pic;
    const c = (overrides?.get(pic.id)?.crop ?? pic.crop) ?? { l: 0, t: 0, r: 0, b: 0 };
    const fw = g.w / Math.max(0.01, 1 - c.l - c.r);
    const fh = g.h / Math.max(0.01, 1 - c.t - c.b);
    return { src: pic.src, x: g.x - c.l * fw, y: g.y - c.t * fh, w: fw, h: fh };
  })();

  return (
    <div className="canvas-wrap" ref={wrapRef}>
      {state.showRuler && <Rulers width={viewport.w} height={viewport.h} stageLeft={stageLeft - scrollPos.x} stageTop={stageTop - scrollPos.y} zoom={z} slideW={pres.width} slideH={pres.height} size={rulerSize} />}
      <div
        className="canvas-scroll oh-scroll"
        ref={scrollRef}
        style={{ left: rulerSize, top: rulerSize, cursor }}
        onScroll={(e) => setScrollPos({ x: e.currentTarget.scrollLeft, y: e.currentTarget.scrollTop })}
        onPointerDown={onPointerDown}
        onPointerMove={onPointerMove}
        onPointerUp={onPointerUp}
        onPointerCancel={onPointerUp}
        onDoubleClick={onDoubleClick}
        onContextMenu={onContextMenu}
        tabIndex={-1}
      >
        <div className="canvas-content" style={{ width: contentW, height: contentH }}>
          <div className="stage" ref={stageRef} style={{ left: stageLeft, top: stageTop, width: pres.width * z, height: pres.height * z }}>
            {slide ? (
              <div className="stage-slide" ref={slideRef} style={{ width: pres.width, height: pres.height, transform: `scale(${z})` }}>
                <SlideView pres={pres} slide={slide} slideNumber={state.slideIdx + 1} mode="edit" overrides={overrides} editing={editing} renderEditor={renderEditor} />
              </div>
            ) : (
              <div className="stage-empty" style={{ width: pres.width * z, height: pres.height * z }}>
                Slayt eklemek için tıklayın
              </div>
            )}
            {state.showGrid && <GridOverlay w={pres.width} h={pres.height} z={z} />}
            {state.showGuides && (
              <>
                {props.guides.v.map((x, i) => (
                  <div key={`gv${i}`} className="drawing-guide is-v" data-guide="v" data-index={i} style={{ left: x * z - 3, height: pres.height * z }} />
                ))}
                {props.guides.h.map((y, i) => (
                  <div key={`gh${i}`} className="drawing-guide is-h" data-guide="h" data-index={i} style={{ top: y * z - 3, width: pres.width * z }} />
                ))}
              </>
            )}
            <SelectionOverlay
              shapes={selShapes}
              overrides={overrides}
              zoom={z}
              editingId={editing?.shapeId}
              parentGroup={parentGroup ?? null}
              guides={guideLines}
              lasso={lasso}
              preview={preview}
              animBadges={badges}
              cellSel={cellRects}
              rotInfo={rotInfo}
              cropGhost={cropGhost}
              comments={slide?.comments}
            />
          </div>
        </div>
      </div>
    </div>
  );
}

function GridOverlay({ w, h, z }: { w: number; h: number; z: number }) {
  const step = (96 / 2.54) * z;
  return (
    <svg className="grid-overlay" width={w * z} height={h * z}>
      <defs>
        <pattern id="oh-grid" width={step} height={step} patternUnits="userSpaceOnUse" x={((w * z) / 2) % step} y={((h * z) / 2) % step}>
          <circle cx={0.5} cy={0.5} r={0.7} fill="rgba(0,0,0,0.35)" />
        </pattern>
      </defs>
      <rect width="100%" height="100%" fill="url(#oh-grid)" />
    </svg>
  );
}
