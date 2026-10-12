/** Şekil ağacı üzerinde değişmez (immutable) işlemler. */
import { cloneShape, uid } from "./factory";
import type { GroupShape, Presentation, Shape, Slide } from "./types";

export interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

export function findShape(shapes: Shape[], id: string): Shape | undefined {
  for (const s of shapes) {
    if (s.id === id) return s;
    if (s.kind === "group") {
      const f = findShape(s.children, id);
      if (f) return f;
    }
  }
  return undefined;
}

/** Şeklin bulunduğu üst grup (yoksa undefined: slayt düzeyi). */
export function parentOf(shapes: Shape[], id: string, parent?: GroupShape): GroupShape | undefined | null {
  for (const s of shapes) {
    if (s.id === id) return parent;
    if (s.kind === "group") {
      const f = parentOf(s.children, id, s);
      if (f !== null) return f;
    }
  }
  return null;
}

export function flatShapes(shapes: Shape[]): Shape[] {
  const out: Shape[] = [];
  const walk = (list: Shape[]) => {
    for (const s of list) {
      out.push(s);
      if (s.kind === "group") walk(s.children);
    }
  };
  walk(shapes);
  return out;
}

/** Döndürülmüş kutunun eksen hizalı sınırları. */
export function aabb(s: Box & { rot?: number }): Box {
  const r = ((s.rot ?? 0) * Math.PI) / 180;
  if (!r) return { x: s.x, y: s.y, w: s.w, h: s.h };
  const c = Math.abs(Math.cos(r));
  const n = Math.abs(Math.sin(r));
  const w = s.w * c + s.h * n;
  const h = s.w * n + s.h * c;
  const cx = s.x + s.w / 2;
  const cy = s.y + s.h / 2;
  return { x: cx - w / 2, y: cy - h / 2, w, h };
}

export function unionBox(boxes: Box[]): Box {
  if (!boxes.length) return { x: 0, y: 0, w: 0, h: 0 };
  let x1 = Infinity;
  let y1 = Infinity;
  let x2 = -Infinity;
  let y2 = -Infinity;
  for (const b of boxes) {
    x1 = Math.min(x1, b.x);
    y1 = Math.min(y1, b.y);
    x2 = Math.max(x2, b.x + b.w);
    y2 = Math.max(y2, b.y + b.h);
  }
  return { x: x1, y: y1, w: x2 - x1, h: y2 - y1 };
}

function fitGroup(g: GroupShape): GroupShape {
  const b = unionBox(g.children.map(aabb));
  return { ...g, x: b.x, y: b.y, w: b.w, h: b.h };
}

/** Şekli (gruplar için alt şekilleri ölçekleyerek) yeni kutuya yerleştirir. */
export function setBox<T extends Shape>(s: T, b: Box): T {
  if (s.kind === "group") {
    const sx = s.w ? b.w / s.w : 1;
    const sy = s.h ? b.h / s.h : 1;
    const children = s.children.map((c) => setBox(c, { x: b.x + (c.x - s.x) * sx, y: b.y + (c.y - s.y) * sy, w: c.w * sx, h: c.h * sy }));
    return { ...s, ...b, children } as T;
  }
  if (s.kind === "table") {
    const sx = s.w ? b.w / s.w : 1;
    const sy = s.h ? b.h / s.h : 1;
    return { ...s, ...b, cols: s.cols.map((c) => c * sx), rows: s.rows.map((r) => ({ ...r, h: r.h * sy })) } as T;
  }
  return { ...s, ...b };
}

export function moveShape<T extends Shape>(s: T, dx: number, dy: number): T {
  if (s.kind === "group") return { ...s, x: s.x + dx, y: s.y + dy, children: s.children.map((c) => moveShape(c, dx, dy)) } as T;
  return { ...s, x: s.x + dx, y: s.y + dy };
}

/** Kimliği verilen şekilleri (iç içe dahil) dönüştürür; grup sınırlarını günceller. */
export function mapShapes(shapes: Shape[], ids: Set<string> | ((s: Shape) => boolean), fn: (s: Shape) => Shape): Shape[] {
  const test = typeof ids === "function" ? ids : (s: Shape) => ids.has(s.id);
  let changed = false;
  const out = shapes.map((s) => {
    if (test(s)) {
      changed = true;
      return fn(s);
    }
    if (s.kind === "group") {
      const kids = mapShapes(s.children, test, fn);
      if (kids !== s.children) {
        changed = true;
        return fitGroup({ ...s, children: kids });
      }
    }
    return s;
  });
  return changed ? out : shapes;
}

export function removeShapes(shapes: Shape[], ids: Set<string>): Shape[] {
  const out: Shape[] = [];
  for (const s of shapes) {
    if (ids.has(s.id)) continue;
    if (s.kind === "group") {
      const kids = removeShapes(s.children, ids);
      if (kids.length === 0) continue;
      if (kids.length === 1 && kids.length !== s.children.length) {
        out.push(kids[0]);
        continue;
      }
      out.push(kids === s.children ? s : fitGroup({ ...s, children: kids }));
    } else out.push(s);
  }
  return out;
}

export function updateSlide(pres: Presentation, index: number, fn: (s: Slide) => Slide): Presentation {
  const slides = pres.slides.slice();
  slides[index] = fn(slides[index]);
  return { ...pres, slides };
}

export function updateSlideShapes(pres: Presentation, index: number, fn: (shapes: Shape[]) => Shape[]): Presentation {
  return updateSlide(pres, index, (s) => ({ ...s, shapes: fn(s.shapes) }));
}

/* ------------------------------------------------------------------ */
/* Sıralama                                                            */
/* ------------------------------------------------------------------ */

export type ZOrder = "front" | "back" | "forward" | "backward";

function reorderList(list: Shape[], ids: Set<string>, op: ZOrder): Shape[] {
  if (!list.some((s) => ids.has(s.id))) {
    let changed = false;
    const out = list.map((s) => {
      if (s.kind !== "group") return s;
      const kids = reorderList(s.children, ids, op);
      if (kids === s.children) return s;
      changed = true;
      return { ...s, children: kids };
    });
    return changed ? out : list;
  }
  const arr = list.slice();
  if (op === "front") return [...arr.filter((s) => !ids.has(s.id)), ...arr.filter((s) => ids.has(s.id))];
  if (op === "back") return [...arr.filter((s) => ids.has(s.id)), ...arr.filter((s) => !ids.has(s.id))];
  if (op === "forward") {
    for (let i = arr.length - 2; i >= 0; i--) {
      if (ids.has(arr[i].id) && !ids.has(arr[i + 1].id)) [arr[i], arr[i + 1]] = [arr[i + 1], arr[i]];
    }
  } else {
    for (let i = 1; i < arr.length; i++) {
      if (ids.has(arr[i].id) && !ids.has(arr[i - 1].id)) [arr[i], arr[i - 1]] = [arr[i - 1], arr[i]];
    }
  }
  return arr;
}

export function reorder(shapes: Shape[], ids: string[], op: ZOrder): Shape[] {
  return reorderList(shapes, new Set(ids), op);
}

/* ------------------------------------------------------------------ */
/* Gruplama                                                            */
/* ------------------------------------------------------------------ */

export function groupShapes(shapes: Shape[], ids: string[]): { shapes: Shape[]; groupId?: string } {
  const set = new Set(ids);
  const members = shapes.filter((s) => set.has(s.id));
  if (members.length < 2) return { shapes };
  const b = unionBox(members.map(aabb));
  const g: GroupShape = { kind: "group", id: uid(), name: `Grup ${Math.floor(Math.random() * 90) + 10}`, ...b, rot: 0, children: members };
  const firstIdx = shapes.findIndex((s) => set.has(s.id));
  const lastIdx = shapes.reduce((m, s, i) => (set.has(s.id) ? i : m), -1);
  const out: Shape[] = [];
  shapes.forEach((s, i) => {
    if (set.has(s.id)) {
      if (i === lastIdx) out.push(g);
      return;
    }
    out.push(s);
  });
  void firstIdx;
  return { shapes: out, groupId: g.id };
}

/** Grubu çözer; döndürülmüş grupta alt şekillerin konum/döndürmesini dünyaya taşır. */
export function ungroupShapes(shapes: Shape[], ids: string[]): { shapes: Shape[]; childIds: string[] } {
  const set = new Set(ids);
  const childIds: string[] = [];
  const out: Shape[] = [];
  for (const s of shapes) {
    if (s.kind === "group" && set.has(s.id)) {
      const cx = s.x + s.w / 2;
      const cy = s.y + s.h / 2;
      const r = (s.rot * Math.PI) / 180;
      for (const c of s.children) {
        let child: Shape = c;
        if (s.rot) {
          const ccx = c.x + c.w / 2 - cx;
          const ccy = c.y + c.h / 2 - cy;
          const nx = cx + ccx * Math.cos(r) - ccy * Math.sin(r);
          const ny = cy + ccx * Math.sin(r) + ccy * Math.cos(r);
          child = moveShape(c, nx - (c.x + c.w / 2), ny - (c.y + c.h / 2));
          child = { ...child, rot: (((c.rot + s.rot) % 360) + 360) % 360 };
        }
        childIds.push(child.id);
        out.push(child);
      }
    } else out.push(s);
  }
  return { shapes: out, childIds };
}

/* ------------------------------------------------------------------ */
/* Hizalama ve dağıtma                                                  */
/* ------------------------------------------------------------------ */

export type AlignOp = "left" | "center" | "right" | "top" | "middle" | "bottom" | "distH" | "distV";

export function alignShapes(shapes: Shape[], ids: string[], op: AlignOp, toSlide: boolean, slideW: number, slideH: number): Shape[] {
  const set = new Set(ids);
  const targets = shapes.filter((s) => set.has(s.id));
  if (!targets.length) return shapes;
  const boxes = new Map(targets.map((s) => [s.id, aabb(s)]));
  const ref = toSlide || targets.length === 1 ? { x: 0, y: 0, w: slideW, h: slideH } : unionBox([...boxes.values()]);
  const delta = new Map<string, [number, number]>();
  if (op === "distH" || op === "distV") {
    const horiz = op === "distH";
    const sorted = targets.slice().sort((a, b) => (horiz ? boxes.get(a.id)!.x - boxes.get(b.id)!.x : boxes.get(a.id)!.y - boxes.get(b.id)!.y));
    const total = sorted.reduce((t, s) => t + (horiz ? boxes.get(s.id)!.w : boxes.get(s.id)!.h), 0);
    let start: number;
    let span: number;
    if (toSlide || targets.length < 3) {
      start = horiz ? ref.x : ref.y;
      span = horiz ? ref.w : ref.h;
    } else {
      const u = unionBox([...boxes.values()]);
      start = horiz ? u.x : u.y;
      span = horiz ? u.w : u.h;
    }
    const gap = sorted.length > 1 ? (span - total) / (sorted.length - 1) : 0;
    let pos = start;
    for (const s of sorted) {
      const b = boxes.get(s.id)!;
      if (sorted.length === 1) pos = start + (span - (horiz ? b.w : b.h)) / 2;
      delta.set(s.id, horiz ? [pos - b.x, 0] : [0, pos - b.y]);
      pos += (horiz ? b.w : b.h) + gap;
    }
  } else {
    for (const s of targets) {
      const b = boxes.get(s.id)!;
      let dx = 0;
      let dy = 0;
      if (op === "left") dx = ref.x - b.x;
      if (op === "center") dx = ref.x + ref.w / 2 - (b.x + b.w / 2);
      if (op === "right") dx = ref.x + ref.w - (b.x + b.w);
      if (op === "top") dy = ref.y - b.y;
      if (op === "middle") dy = ref.y + ref.h / 2 - (b.y + b.h / 2);
      if (op === "bottom") dy = ref.y + ref.h - (b.y + b.h);
      delta.set(s.id, [dx, dy]);
    }
  }
  return shapes.map((s) => {
    const d = delta.get(s.id);
    return d ? moveShape(s, d[0], d[1]) : s;
  });
}

/* ------------------------------------------------------------------ */
/* Çoğaltma, slayt işlemleri                                           */
/* ------------------------------------------------------------------ */

export function duplicateShapes(shapes: Shape[], ids: string[], offset = 16): { shapes: Shape[]; newIds: string[] } {
  const set = new Set(ids);
  const copies = shapes.filter((s) => set.has(s.id)).map((s) => moveShape(cloneShape(s), offset, offset));
  return { shapes: [...shapes, ...copies], newIds: copies.map((c) => c.id) };
}

export function cloneSlide(slide: Slide): Slide {
  const idMap = new Map<string, string>();
  const shapes = slide.shapes.map((s) => cloneShape(s, idMap));
  return {
    ...structuredClone(slide),
    id: uid("sl"),
    shapes,
    animations: slide.animations.filter((a) => idMap.has(a.shapeId)).map((a) => ({ ...a, id: uid("a"), shapeId: idMap.get(a.shapeId)! })),
  };
}

export function moveSlides(slides: Slide[], ids: string[], toIndex: number): Slide[] {
  const set = new Set(ids);
  const moving = slides.filter((s) => set.has(s.id));
  const before = slides.slice(0, toIndex).filter((s) => !set.has(s.id));
  const after = slides.slice(toIndex).filter((s) => !set.has(s.id));
  return [...before, ...moving, ...after];
}

/** Şekillerin silinmesiyle sahipsiz kalan animasyonları temizler. */
export function pruneAnimations(slide: Slide): Slide {
  const ids = new Set(flatShapes(slide.shapes).map((s) => s.id));
  const anims = slide.animations.filter((a) => ids.has(a.shapeId));
  return anims.length === slide.animations.length ? slide : { ...slide, animations: anims };
}
