/**
 * Normal görünümün slayt alanı: gri zemin, ölçeklenmiş slayt, seçim,
 * sürükle-taşı/boyutlandır/döndür, kement seçimi, akıllı kılavuzlar,
 * şekil çizme ve metin düzenleyicinin yerleştirilmesi.
 */
import { useCallback, useMemo, useRef, useState } from "react";
import { aabb, findShape, unionBox, type Box } from "../model/ops";
import { layoutOf, type TextContext } from "../model/textStyle";
import type { Shape, SpShape } from "../model/types";
import { isLineGeom } from "../render/geometry";
import { SlideView, type GeomOverride, type PlaceholderAction } from "../render/SlideView";
import type { Actions } from "../state/actions";
import type { AppState } from "../state/store";
import { applyFormatPainter } from "../state/painter";
import { ADJUST_HANDLES, effectiveAdj } from "./adjust";
import { angleFromCenter, constrainLine, geomFromEnds, geomOf, HANDLES, lineEnds, normAngle, resizeGeom, rotateVec } from "./interact";
import { Rulers } from "./Rulers";
import { SelectionOverlay } from "./SelectionOverlay";
import { snapMove, snapValue, type GuideLine } from "./snap";
import { cellSelRects, childOfGroupContaining, cropGhostOf, InkPreview, inkIdsAt, MARGIN, topLevelOf, type Drag, type Pt } from "./canvasUtils";
import { useAutofit } from "./useAutofit";
import { contextKind, isInEditedShape, useAnimBadges, type CanvasMenuKind } from "./contextKind";
import { StageGuides } from "./MotionPathOverlay";
import { PointEditor } from "./PointEditor";
import { SunumMiniToolbar } from "./MiniToolbar";
import { useViewportZoom } from "./useViewportZoom";
import { useEditorRenderer } from "./useEditorRenderer";
import { EyedropperOverlay } from "./Eyedropper";
import { slideToCanvas } from "../app/exporting";

export interface CanvasProps {
  state: AppState;
  actions: Actions;
  onContextMenu: (e: React.MouseEvent, kind: CanvasMenuKind) => void;
  onPlaceholderAction: (action: PlaceholderAction, shapeId: string) => void;
  onShapeDoubleClick: (s: Shape) => void;
  showAnimBadges: boolean;
  guides: { h: number[]; v: number[] };
  onGuidesChange: (g: { h: number[]; v: number[] }) => void;
  onZoomChange: (zoom: number, fit: boolean) => void;
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
  const [preview, setPreview] = useState<(Box & { line?: boolean; flipH?: boolean; flipV?: boolean; geom?: string }) | null>(null);
  const [inkPts, setInkPts] = useState<Pt[] | null>(null);
  const [rotInfo, setRotInfo] = useState<{ x: number; y: number; text: string } | null>(null);
  const [scrollPos, setScrollPos] = useState({ x: 0, y: 0 });
  const propsRef = useRef(props);
  propsRef.current = props;

  const z = state.zoom / 100;
  const rulerSize = state.showRuler ? 18 : 0;

  const viewport = useViewportZoom(scrollRef, pres.width, pres.height, () => propsRef.current);

  const contentW = Math.max(viewport.w, pres.width * z + MARGIN * 2);
  const contentH = Math.max(viewport.h, pres.height * z + MARGIN * 2);
  const stageLeft = (contentW - pres.width * z) / 2;
  const stageTop = (contentH - pres.height * z) / 2;

  const toSlide = useCallback((cx: number, cy: number): Pt => {
    const r = stageRef.current!.getBoundingClientRect();
    const zz = propsRef.current.state.zoom / 100;
    return [(cx - r.left) / zz, (cy - r.top) / zz];
  }, []);

  const tctx: TextContext = useMemo(() => ({ master: pres.master, layout: layoutOf(pres, slide) }), [pres, slide]);

  useAutofit({ slide, slideIdx: state.slideIdx, measuring: !overrides, tctx, theme: tctx.layout?.theme ?? pres.theme, slideRef, actions });

  /* ------------------------- Metin düzenleyici ------------------------- */
  const editing = state.editing;
  const renderEditor = useEditorRenderer(actions, () => propsRef.current.state);

  /* ------------------------- İşaretçi olayları ------------------------- */

  const selectedTop = (): Shape[] => {
    const st = propsRef.current.state;
    const sl = st.pres.slides[st.slideIdx];
    if (!sl) return [];
    return st.sel.map((id) => findShape(sl.shapes, id)).filter((x): x is Shape => !!x);
  };

  const eraseAt = (cx: number, cy: number, key: string) => {
    const ids = slide ? inkIdsAt(slide, cx, cy) : [];
    if (ids.length) actions.eraseInk(ids, key);
  };

  const onPointerDown = (e: React.PointerEvent) => {
    const st = propsRef.current.state;
    if (!slide) return;
    const target = e.target as Element;
    if (!stageRef.current) return;
    const pt = toSlide(e.clientX, e.clientY);
    actions.store.set({ focus: "canvas" });
    if (e.button === 2) {
      // Düzenlenen şeklin içinde sağ tık: odak metin düzenleyicide kalır.
      if (isInEditedShape(target, st, slideRef.current)) return e.preventDefault();
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

    if (st.ink) {
      if (st.ink.tool === "eraser") {
        const key = `erase:${Date.now()}`;
        drag.current = { kind: "erase", key };
        eraseAt(e.clientX, e.clientY, key);
      } else {
        drag.current = { kind: "ink", pts: [pt] };
        setInkPts([pt]);
      }
      return;
    }

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

    // Kırpma modunda resmin içinde sürükleme: resmi kırpma çerçevesi içinde kaydırır.
    if (st.cropId && (target.closest(".crop-ghost") || (target.closest("[data-shape-id]") as HTMLElement | null)?.dataset.shapeId === st.cropId)) {
      const pic = findShape(slide.shapes, st.cropId);
      if (pic && pic.kind === "pic") {
        e.preventDefault();
        drag.current = { kind: "cropPan", id: pic.id, start: pt, geom: geomOf(pic), c: pic.crop ?? { l: 0, t: 0, r: 0, b: 0 } };
        return;
      }
    }

    const phAct = target.closest("[data-ph-action]") as HTMLElement | null;
    if (phAct) {
      const sid = (phAct.closest("[data-shape-id]") as HTMLElement | null)?.dataset.shapeId;
      if (sid) propsRef.current.onPlaceholderAction(phAct.dataset.phAction as PlaceholderAction, sid);
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
      // Tablo hücresinde (seçili olmasa da) sürükleme hücre aralığı seçer, tıklama hücreyi düzenler; tablo kenarlıktan taşınır.
      if (cell && top.kind === "table" && !e.shiftKey && !e.ctrlKey) {
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
      case "ink": {
        const last = d.pts[d.pts.length - 1];
        if (Math.hypot(pt[0] - last[0], pt[1] - last[1]) * (st.zoom / 100) >= 1.5) {
          d.pts.push(pt);
          setInkPts(d.pts.slice());
        }
        break;
      }
      case "erase":
        eraseAt(e.clientX, e.clientY, d.key);
        break;
      case "cropPan": {
        const g = d.geom;
        const c = d.c;
        const fw = g.w / Math.max(0.01, 1 - c.l - c.r);
        const fh = g.h / Math.max(0.01, 1 - c.t - c.b);
        // Resmin sol/üst kenarı: çerçeveyi her zaman örtecek biçimde sınırlanır.
        const left = Math.min(g.x, Math.max(g.x + g.w - fw, g.x - c.l * fw + (pt[0] - d.start[0])));
        const top = Math.min(g.y, Math.max(g.y + g.h - fh, g.y - c.t * fh + (pt[1] - d.start[1])));
        const l = (g.x - left) / fw;
        const t = (g.y - top) / fh;
        const nc = { l, t, r: Math.max(0, 1 - l - g.w / fw), b: Math.max(0, 1 - t - g.h / fh) };
        setOverrides(new Map([[d.id, { x: g.x, y: g.y, w: g.w, h: g.h, rot: g.rot, flipH: g.flipH, flipV: g.flipV, crop: nc }]]));
        break;
      }
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
        // Çizilirken şeklin kendisi gösterilir (üçgen, yıldız, ok…); metin kutusunda çerçeve.
        const geom = st.insert?.kind === "shape" ? st.insert.geom : undefined;
        setPreview({ x: Math.min(d.start[0], d.start[0] + w), y: Math.min(d.start[1], d.start[1] + h), w: Math.abs(w), h: Math.abs(h), geom });
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
      case "ink":
        setInkPts(null);
        if (st.ink && st.ink.tool !== "eraser") actions.addInk(d.pts, st.ink);
        return;
      case "erase":
        actions.store.breakMerge();
        return;
      case "cropPan": {
        const nc = ov?.get(d.id)?.crop;
        if (nc) actions.updateShapes((sh) => (sh.kind === "pic" ? { ...sh, crop: nc.l || nc.t || nc.r || nc.b ? nc : undefined } : sh), [d.id]);
        return;
      }
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
          if (st.editing?.shapeId === sh.id) {
            // Düzenlenen şeklin metin dışı alanına tıklama: imleç en yakın konuma, odak metne.
            if (d.point) actions.textHandle?.placeCaret(d.point.x, d.point.y);
            return;
          }
          actions.startEditing({ shapeId: sh.id, point: d.point });
        } else if (sh.kind === "sp" && st.editing?.shapeId === sh.id && d.point) actions.textHandle?.placeCaret(d.point.x, d.point.y);
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
          // Kopya bırakılan yere tek adımda konur (geri al kopyayı tümden kaldırır).
          actions.duplicateSelection(dx, dy);
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
    // Shift + sağ tık: düzenleyicide tarayıcının menüsü (yazım önerileri).
    if (e.shiftKey && (e.target as Element).closest(".tb-edit")) return;
    e.preventDefault();
    props.onContextMenu(e, contextKind(e.target as Element, propsRef.current.state, slideRef.current, actions));
  };

  /* ------------------------- Çizim ------------------------- */

  const selShapes = slide ? state.sel.map((id) => findShape(slide.shapes, id)).filter((x): x is Shape => !!x) : [];
  const parentGroup = slide && selShapes.length ? slide.shapes.find((s) => s.kind === "group" && selShapes.some((x) => x.id !== s.id && findShape(s.children, x.id))) : null;
  const badges = useAnimBadges(props.showAnimBadges, slide);
  const cellRects = useMemo(
    () => (state.cellSel && selShapes.length === 1 && selShapes[0].kind === "table" ? cellSelRects(selShapes[0], state.cellSel) : null),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [state.cellSel, slide, state.sel],
  );

  const cursor = state.insert ? "crosshair" : state.ink ? (state.ink.tool === "eraser" ? "cell" : "crosshair") : state.painter ? "copy" : undefined;
  const cropGhost = slide && state.cropId ? cropGhostOf(slide, state.cropId, overrides) : null;

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
        onMouseDown={(e) => e.button === 2 && isInEditedShape(e.target as Element, propsRef.current.state, slideRef.current) && e.preventDefault()}
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
            <StageGuides state={state} slide={slide} z={z} W={pres.width} H={pres.height} guides={props.guides} motion={props.showAnimBadges} />
            {inkPts && state.ink && <InkPreview pts={inkPts} ink={state.ink} z={z} w={pres.width} h={pres.height} />}
            {state.pointEdit && selShapes.length === 1 && selShapes[0].id === state.pointEdit && selShapes[0].kind === "sp" && <PointEditor shape={selShapes[0]} zoom={z} actions={actions} />}
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
              cropId={state.cropId}
              comments={slide?.comments}
            />
          </div>
        </div>
      </div>
      <SunumMiniToolbar state={state} actions={actions} slideRef={slideRef} />
      <EyedropperOverlay pres={pres} slideIdx={state.slideIdx} toSlide={toSlide} rasterize={slideToCanvas} />
    </div>
  );
}

