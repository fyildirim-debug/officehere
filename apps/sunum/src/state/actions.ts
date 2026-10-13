/** Slayt ve şekil komutları (şerit, kısayollar ve bağlam menüleri bunları kullanır). */
import { defaultActionFor } from "../render/actionButtons";
import { cloneShape, DEFAULT_SHAPE_FILL, DEFAULT_SHAPE_LINE, makePicture, makeSp, makeTable, makeTextBox, NO_FILL, NO_LINE, uid } from "../model/factory";
import {
  aabb,
  alignShapes,
  cloneSlide,
  duplicateShapes,
  findShape,
  flatShapes,
  groupShapes,
  mapShapes,
  moveShape,
  pruneAnimations,
  removeShapes,
  reorder,
  setBox,
  ungroupShapes,
  unionBox,
  updateSlide,
  updateSlideShapes,
  type AlignOp,
  type Box,
  type ZOrder,
} from "../model/ops";
import { footerShapes, layoutByType, relayoutSlide, newSlide as makeSlide } from "../model/presentation";
import { layoutOf } from "../model/textStyle";
import type { Animation, Fill, Layout, Presentation, Shape, Slide, SpShape, TableShape, Transition } from "../model/types";
import { isLineGeom } from "../render/geometry";
import type { TextEditorHandle } from "../editor/TextEditor";
import type { AppState, EditTarget, InkTool, Store } from "./store";
import { inkShape } from "../model/ink";
import { moveSlidesKeepSections, removeSlidesKeepSections } from "../model/sections";
import { inMasterView, masterDeleteLayout, masterDuplicateLayout, masterInsertLayout } from "./masterActions";
import type { ShowStroke } from "../show/ShowInk";

export class Actions {
  textHandle: TextEditorHandle | null = null;
  constructor(readonly store: Store) {}

  get s(): AppState {
    return this.store.get();
  }

  get pres(): Presentation {
    return this.s.pres;
  }

  slide(): Slide {
    const s = this.s;
    return s.pres.slides[s.slideIdx];
  }

  layout(): Layout | undefined {
    return layoutOf(this.pres, this.slide());
  }

  selectedShapes(): Shape[] {
    const sl = this.slide();
    if (!sl) return [];
    return this.s.sel.map((id) => findShape(sl.shapes, id)).filter((x): x is Shape => !!x);
  }

  message(text: string, kind: "info" | "warning" | "error" | "success" = "info", action?: "startShow") {
    this.store.set({ message: action ? { text, kind, action } : { text, kind } });
  }

  /* ------------------------------ Slaytlar ------------------------------ */

  goToSlide(i: number, mode: "replace" | "extend" | "toggle" = "replace") {
    const s = this.s;
    const idx = Math.max(0, Math.min(s.pres.slides.length - 1, i));
    const id = s.pres.slides[idx]?.id;
    if (!id) return;
    let slideSel = [id];
    if (mode === "toggle") slideSel = s.slideSel.includes(id) ? s.slideSel.filter((x) => x !== id) : [...s.slideSel, id];
    if (mode === "extend") {
      const cur = s.slideIdx;
      const [a, b] = cur < idx ? [cur, idx] : [idx, cur];
      slideSel = s.pres.slides.slice(a, b + 1).map((x) => x.id);
    }
    if (!slideSel.length) slideSel = [id];
    this.store.set({ slideIdx: idx, slideSel, sel: idx === s.slideIdx ? s.sel : [], editing: idx === s.slideIdx ? s.editing : null, cellSel: null });
  }

  selectedSlideIds(): string[] {
    const s = this.s;
    const cur = s.pres.slides[s.slideIdx]?.id;
    const ids = s.slideSel.filter((id) => s.pres.slides.some((x) => x.id === id));
    return ids.length ? ids : cur ? [cur] : [];
  }

  newSlide(layoutId?: string) {
    if (inMasterView(this)) return masterInsertLayout(this);
    const s = this.s;
    const cur = this.slide();
    let lid = layoutId;
    if (!lid) {
      const curLayout = cur ? layoutOf(s.pres, cur) : undefined;
      // Boş sunuda ilk slayt Başlık Slaydı'dır; sonrakiler geçerli düzeni (başlıktan sonra Başlık ve İçerik) alır.
      if (!s.pres.slides.length) lid = layoutByType(s.pres, "title").id;
      else if (curLayout?.type === "title") {
        // Başlık slaydından sonra aynı temanın (asıl slaydın) Başlık ve İçerik düzeni.
        lid = (s.pres.layouts.find((l) => l.type === "obj" && l.theme?.name === curLayout.theme?.name) ?? layoutByType(s.pres, "obj")).id;
      } else lid = curLayout ? curLayout.id : layoutByType(s.pres, "obj").id;
    }
    const slide = makeSlide(s.pres, lid);
    const at = s.slideIdx + 1;
    this.store.commit((p) => ({ ...p, slides: [...p.slides.slice(0, at), slide, ...p.slides.slice(at)] }), { slideIdx: at, sel: [], editing: null });
  }

  /** Sununun sonuna yeni slayt ekler (küçük resim bölmesindeki + düğmesi); düzen son slayttan alınır. */
  appendSlide() {
    const last = this.s.pres.slides.length - 1;
    if (last >= 0 && this.s.slideIdx !== last) this.goToSlide(last);
    this.newSlide();
  }

  duplicateSlides(ids = this.selectedSlideIds()) {
    if (inMasterView(this)) return masterDuplicateLayout(this);
    const s = this.s;
    const set = new Set(ids);
    const lastIdx = s.pres.slides.reduce((m, x, i) => (set.has(x.id) ? i : m), -1);
    if (lastIdx < 0) return;
    const copies = s.pres.slides.filter((x) => set.has(x.id)).map(cloneSlide);
    this.store.commit((p) => ({ ...p, slides: [...p.slides.slice(0, lastIdx + 1), ...copies, ...p.slides.slice(lastIdx + 1)] }), { slideIdx: lastIdx + copies.length, sel: [] });
    this.store.set({ slideSel: copies.map((c) => c.id) });
  }

  deleteSlides(ids = this.selectedSlideIds()) {
    if (inMasterView(this)) return masterDeleteLayout(this);
    const s = this.s;
    const set = new Set(ids);
    if (set.size >= s.pres.slides.length) {
      // Tüm slaytlar silinirse sunu boş kalır (PowerPoint gibi).
      this.store.commit((p) => ({ ...p, slides: [] }), { sel: [], editing: null });
      this.store.set({ slideIdx: 0, slideSel: [] });
      return;
    }
    const firstIdx = s.pres.slides.findIndex((x) => set.has(x.id));
    this.store.commit((p) => ({ ...p, slides: removeSlidesKeepSections(p.slides, set) }), { slideIdx: Math.max(0, firstIdx - 1), sel: [], editing: null });
  }

  toggleHideSlides(ids = this.selectedSlideIds()) {
    if (inMasterView(this)) return;
    const set = new Set(ids);
    const s = this.s;
    const allHidden = s.pres.slides.filter((x) => set.has(x.id)).every((x) => x.hidden);
    this.store.commit((p) => ({ ...p, slides: p.slides.map((x) => (set.has(x.id) ? { ...x, hidden: !allHidden } : x)) }));
  }

  moveSlidesTo(ids: string[], toIndex: number) {
    if (inMasterView(this)) return;
    const s = this.s;
    const moved = moveSlidesKeepSections(s.pres.slides, ids, toIndex);
    const cur = s.pres.slides[s.slideIdx]?.id;
    this.store.commit((p) => ({ ...p, slides: moved }), { slideIdx: Math.max(0, moved.findIndex((x) => x.id === cur)) });
    this.store.set({ slideSel: ids });
  }

  setLayout(layoutId: string, ids = this.selectedSlideIds()) {
    if (inMasterView(this)) return;
    const set = new Set(ids);
    this.store.commit((p) => {
      const nl = p.layouts.find((l) => l.id === layoutId);
      if (!nl) return p;
      return { ...p, slides: p.slides.map((x) => (set.has(x.id) ? relayoutSlide(p, x, nl, layoutOf(p, x)) : x)) };
    }, { sel: [], editing: null });
  }

  resetSlides(ids = this.selectedSlideIds()) {
    if (inMasterView(this)) return;
    const set = new Set(ids);
    this.store.commit((p) => ({
      ...p,
      slides: p.slides.map((x) => {
        if (!set.has(x.id)) return x;
        const l = layoutOf(p, x)!;
        return { ...relayoutSlide(p, x, l, l, true), background: undefined, hideMasterShapes: undefined };
      }),
    }), { sel: [], editing: null });
  }

  setBackground(fill: Fill | undefined, all = false) {
    const cur = this.selectedSlideIds();
    const set = new Set(cur);
    this.store.commit((p) => ({ ...p, slides: p.slides.map((x) => (all || set.has(x.id) ? { ...x, background: fill } : x)) }), { merge: "bg" });
  }

  setHideBackgroundGraphics(v: boolean, all = false) {
    const set = new Set(this.selectedSlideIds());
    this.store.commit((p) => ({ ...p, slides: p.slides.map((x) => (all || set.has(x.id) ? { ...x, hideMasterShapes: v || undefined } : x)) }));
  }

  setTransition(patch: Partial<Transition> | null, all = false) {
    const set = new Set(this.selectedSlideIds());
    this.store.commit((p) => ({
      ...p,
      slides: p.slides.map((x) => {
        if (!all && !set.has(x.id)) return x;
        if (patch === null) return { ...x, transition: undefined };
        const base: Transition = x.transition ?? { type: "none", duration: 1, advanceOnClick: true };
        return { ...x, transition: { ...base, ...patch } };
      }),
    }), { merge: patch && ("duration" in patch || "advanceAfter" in patch) ? "trans-time" : undefined });
  }

  applyTransitionToAll() {
    const t = this.slide()?.transition;
    this.store.commit((p) => ({ ...p, slides: p.slides.map((x) => ({ ...x, transition: t ? { ...t } : undefined })) }));
  }

  setNotes(text: string) {
    const i = this.s.slideIdx;
    this.store.commit((p) => updateSlide(p, i, (sl) => ({ ...sl, notes: text })), { merge: `notes:${this.slide()?.id}` });
  }

  /** Geçerli slayda boş bir açıklama ekler (seçimin sağ üstüne; üst üste binmez). */
  addComment(author: string) {
    if (inMasterView(this)) {
      this.message("Açıklamalar slaytlara eklenir; Asıl Slayt Görünümünü kapatın.", "info");
      return;
    }
    const i = this.s.slideIdx;
    const sl = this.slide();
    if (!sl) return;
    const box = this.selectionBox();
    const x = box ? Math.min(this.pres.width - 24, box.x + box.w) : 20;
    let y = box ? Math.max(0, box.y) : 20;
    const list = sl.comments ?? [];
    while (list.some((c) => Math.abs(c.x - x) < 6 && Math.abs(c.y - y) < 6)) y += 24;
    const c = { id: uid("c"), author, text: "", x, y, date: Date.now() };
    this.store.commit((p) => updateSlide(p, i, (s2) => ({ ...s2, comments: [...(s2.comments ?? []), c] })));
    this.store.set({ pane: "comments" });
  }

  /* ------------------------------ Seçim ------------------------------ */

  select(ids: string[]) {
    const pe = this.s.pointEdit;
    this.store.set({ sel: ids, editing: null, cellSel: null, focus: "canvas", pointEdit: pe && ids.includes(pe) ? pe : null });
  }

  selectAll() {
    const sl = this.slide();
    if (!sl) return;
    this.select(sl.shapes.filter((x) => !x.hidden).map((x) => x.id));
  }

  cycleSelection(back: boolean) {
    const sl = this.slide();
    if (!sl?.shapes.length) return;
    const list = sl.shapes;
    const cur = this.s.sel[this.s.sel.length - 1];
    const i = list.findIndex((x) => x.id === cur);
    const n = list.length;
    const next = i < 0 ? (back ? n - 1 : 0) : (i + (back ? -1 : 1) + n) % n;
    this.select([list[next].id]);
  }

  startEditing(target: EditTarget) {
    this.store.breakMerge();
    this.store.set({ editing: target, sel: [target.shapeId], focus: "canvas" });
  }

  stopEditing() {
    if (!this.s.editing) return;
    this.store.breakMerge();
    this.store.set({ editing: null });
  }

  /* ------------------------------ Şekil ekleme ------------------------------ */

  private addShapes(shapes: Shape[], opts: { edit?: boolean; selectAll?: boolean } = {}) {
    const i = this.s.slideIdx;
    if (!this.slide()) {
      this.newSlide();
    }
    this.store.commit((p) => updateSlideShapes(p, i, (list) => [...list, ...shapes]), {
      sel: shapes.map((x) => x.id),
      editing: opts.edit ? { shapeId: shapes[0].id, selectAll: opts.selectAll } : null,
    });
  }

  insertShape(geom: string, box: Box, opts: { head?: "triangle"; tail?: "triangle"; flipH?: boolean; flipV?: boolean } = {}) {
    const line = isLineGeom(geom);
    const sh = makeSp(geom, box.x, box.y, box.w, box.h, {
      name: `${geomName(geom)} ${this.nextShapeNumber()}`,
      fill: line ? NO_FILL : DEFAULT_SHAPE_FILL,
      line: line ? { fill: { type: "solid", color: { scheme: "accent1" } }, width: 1, dash: "solid", head: opts.head, tail: opts.tail } : DEFAULT_SHAPE_LINE,
      flipH: opts.flipH,
      flipV: opts.flipV,
      connector: line && geom !== "line" && geom !== "arc" ? true : undefined,
      text: line ? undefined : { paragraphs: [{ runs: [], level: 0, props: { align: "ctr" } }], anchor: "ctr", levels: [{ run: { color: { scheme: "lt1" } }, para: { align: "ctr" } }] },
    });
    // Eylem düğmesi: türüne uygun varsayılan eylem atanır ve Eylem Ayarları açılır.
    const act = geom.startsWith("actionButton") ? defaultActionFor(geom) : undefined;
    if (act) sh.link = act;
    this.addShapes([sh]);
    if (geom.startsWith("actionButton")) this.store.set({ dialogRequest: "action" });
  }

  insertTextBox(box: Box, autoWidth: boolean, vertical = false) {
    const tb = makeTextBox(box.x, box.y, box.w, box.h, "", autoWidth);
    tb.name = `Metin Kutusu ${this.nextShapeNumber()}`;
    if (vertical) tb.text = { ...tb.text!, vert: "vert" };
    this.addShapes([tb], { edit: true });
  }

  nextShapeNumber(): number {
    const sl = this.slide();
    return (sl ? flatShapes(sl.shapes).length : 0) + 1;
  }

  /** Resim ekler; boş resim/içerik yer tutucusu seçiliyse onun yerine. */
  insertPicture(src: string, natW: number, natH: number, name = "Resim") {
    const p = this.pres;
    const target = this.placeholderTarget(["pic", "obj"]);
    let box: Box;
    if (target) {
      const s = Math.min(target.w / natW, target.h / natH);
      box = { x: target.x + (target.w - natW * s) / 2, y: target.y + (target.h - natH * s) / 2, w: natW * s, h: natH * s };
    } else {
      const s = Math.min(1, (p.width * 0.8) / natW, (p.height * 0.8) / natH);
      box = { x: (p.width - natW * s) / 2, y: (p.height - natH * s) / 2, w: natW * s, h: natH * s };
    }
    const pic = makePicture(src, box.x, box.y, box.w, box.h, `${name} ${this.nextShapeNumber()}`);
    if (target) {
      pic.ph = target.ph;
      const i = this.s.slideIdx;
      this.store.commit((pp) => updateSlideShapes(pp, i, (list) => list.map((x) => (x.id === target.id ? pic : x))), { sel: [pic.id], editing: null });
    } else this.addShapes([pic]);
  }

  insertTable(rows: number, cols: number) {
    const p = this.pres;
    const target = this.placeholderTarget(["tbl", "obj"]);
    const w = target ? target.w : Math.min(p.width * 0.8, cols * 160);
    const x = target ? target.x : (p.width - w) / 2;
    const y = target ? target.y : (p.height - rows * 40) / 2;
    const t = makeTable(rows, cols, x, y, w, 40);
    t.name = `Tablo ${this.nextShapeNumber()}`;
    if (target) {
      const i = this.s.slideIdx;
      this.store.commit((pp) => updateSlideShapes(pp, i, (list) => list.map((s) => (s.id === target.id ? t : s))), { sel: [t.id], editing: { shapeId: t.id, cell: [0, 0] } });
    } else this.addShapes([t], {});
    this.store.set({ editing: { shapeId: t.id, cell: [0, 0] } });
  }

  /** Seçili ya da ilk boş uygun yer tutucu. */
  placeholderTarget(types: string[]): SpShape | undefined {
    const sl = this.slide();
    if (!sl) return undefined;
    const empty = (x: Shape): x is SpShape => x.kind === "sp" && !!x.ph && types.includes(x.ph.type) && (!x.text || x.text.paragraphs.every((pp) => pp.runs.every((r) => !r.text)));
    const selected = this.selectedShapes().find(empty);
    if (selected) return selected;
    if (this.pendingPlaceholder) {
      const pending = sl.shapes.find((x) => x.id === this.pendingPlaceholder && empty(x)) as SpShape | undefined;
      if (pending) return pending;
    }
    // Hiçbir şey seçili değilse ilk boş uygun yer tutucu kullanılır (ör. "Başlık ve İçerik" slaytında içerik alanı).
    if (this.s.sel.length) return undefined;
    for (const t of types) {
      const hit = sl.shapes.find((x): x is SpShape => empty(x) && (x as SpShape).ph!.type === t);
      if (hit) return hit;
    }
    return undefined;
  }

  /** Yer tutucu simgesinden ekleme yapılırken hedef. */
  pendingPlaceholder: string | null = null;

  /* ------------------------------ Şekil düzenleme ------------------------------ */

  updateShapes(fn: (s: Shape) => Shape, ids = this.s.sel, merge?: string) {
    if (!ids.length) return;
    const i = this.s.slideIdx;
    const set = new Set(ids);
    this.store.commit((p) => updateSlideShapes(p, i, (list) => mapShapes(list, set, fn)), { merge });
  }

  /** Grup içindeki şekiller dahil; gruplarda alt şekillere uygular. */
  updateSelectedDeep(fn: (s: Shape) => Shape, merge?: string) {
    const deep = (s: Shape): Shape => (s.kind === "group" ? { ...s, children: s.children.map(deep) } : fn(s));
    this.updateShapes(deep, this.s.sel, merge);
  }

  /* ------------------------------ Mürekkep (Çizim sekmesi) ------------------------------ */

  /** Slayt koordinatlarındaki noktalardan mürekkep şekli oluşturur (seçimi değiştirmez). */
  addInk(points: [number, number][], tool: InkTool) {
    const i = this.s.slideIdx;
    if (!this.slide() || !points.length) return;
    const count = (this.slide()?.shapes.filter((x) => x.kind === "sp" && x.ink).length ?? 0) + 1;
    const sh = inkShape(points, { color: tool.color, width: tool.width, alpha: tool.tool === "highlighter" ? 0.5 : tool.tool === "pencil" ? 0.85 : undefined }, `Mürekkep ${count}`);
    if (!sh) return;
    this.store.commit((p) => updateSlideShapes(p, i, (list) => [...list, sh]));
  }

  /** Gösteride çizilen mürekkebi slaytlara şekil olarak ekler (tek geri alma adımı). */
  keepShowInk(ink: Record<string, ShowStroke[]>) {
    this.store.commit((p) => ({
      ...p,
      slides: p.slides.map((sl) => {
        const strokes = ink[sl.id];
        if (!strokes?.length) return sl;
        let n = sl.shapes.filter((x) => x.kind === "sp" && x.ink).length;
        const added = strokes.map((st) => inkShape(st.pts, { color: st.color, width: st.width, alpha: st.tool === "highlighter" ? 0.5 : undefined }, `Mürekkep ${++n}`)).filter((x): x is SpShape => !!x);
        return { ...sl, shapes: [...sl.shapes, ...added] };
      }),
    }));
  }

  /** Mürekkep şekillerini siler; aynı sürüklemedeki silmeler tek geri alma adımıdır. */
  eraseInk(ids: string[], mergeKey?: string) {
    const i = this.s.slideIdx;
    const set = new Set(ids);
    if (!set.size) return;
    this.store.commit((p) => updateSlide(p, i, (sl) => pruneAnimations({ ...sl, shapes: sl.shapes.filter((x) => !set.has(x.id)) })), { merge: mergeKey, sel: this.s.sel.filter((x) => !set.has(x)) });
  }

  clearInk() {
    const ids = (this.slide()?.shapes ?? []).filter((x) => x.kind === "sp" && x.ink).map((x) => x.id);
    if (ids.length) this.eraseInk(ids);
  }

  deleteSelection() {
    const sel = this.s.sel;
    if (!sel.length) return;
    const i = this.s.slideIdx;
    const set = new Set(sel);
    this.store.commit((p) => updateSlide(p, i, (sl) => pruneAnimations({ ...sl, shapes: removeShapes(sl.shapes, set) })), { sel: [], editing: null });
  }

  duplicateSelection() {
    const sel = this.s.sel;
    if (!sel.length) return;
    const i = this.s.slideIdx;
    let newIds: string[] = [];
    this.store.commit((p) =>
      updateSlideShapes(p, i, (list) => {
        const r = duplicateShapes(list, sel, 24);
        newIds = r.newIds;
        return r.shapes;
      }),
    );
    this.store.set({ sel: newIds, editing: null });
  }

  moveSelection(dx: number, dy: number, merge?: string) {
    this.updateShapes((s) => moveShape(s, dx, dy), this.s.sel, merge);
  }

  setGeometry(id: string, box: Box & { rot?: number; flipH?: boolean; flipV?: boolean }) {
    this.updateShapes((s) => {
      const r = setBox(s, box);
      return { ...r, rot: box.rot ?? r.rot, flipH: box.flipH ?? r.flipH, flipV: box.flipV ?? r.flipV };
    }, [id]);
  }

  setGeometries(map: Map<string, Box & { rot?: number; flipH?: boolean; flipV?: boolean }>) {
    const i = this.s.slideIdx;
    const set = new Set(map.keys());
    this.store.commit((p) =>
      updateSlideShapes(p, i, (list) =>
        mapShapes(list, set, (s) => {
          const b = map.get(s.id)!;
          const r = setBox(s, b);
          return { ...r, rot: b.rot ?? r.rot, flipH: b.flipH ?? r.flipH, flipV: b.flipV ?? r.flipV };
        }),
      ),
    );
  }

  setSize(w?: number, h?: number, lock?: boolean) {
    this.updateShapes((s) => {
      let nw = w ?? s.w;
      let nh = h ?? s.h;
      if (lock && s.w && s.h) {
        if (w !== undefined && h === undefined) nh = (s.h * nw) / s.w;
        if (h !== undefined && w === undefined) nw = (s.w * nh) / s.h;
      }
      const cx = s.x;
      const cy = s.y;
      return setBox(s, { x: cx, y: cy, w: Math.max(1, nw), h: Math.max(0, nh) });
    });
  }

  rotateSelection(delta: number | null, flip?: "h" | "v") {
    this.updateShapes((s) => {
      if (flip === "h") return { ...s, flipH: !s.flipH };
      if (flip === "v") return { ...s, flipV: !s.flipV };
      return { ...s, rot: ((((s.rot + (delta ?? 0)) % 360) + 360) % 360) };
    });
  }

  setRotation(deg: number) {
    this.updateShapes((s) => ({ ...s, rot: ((deg % 360) + 360) % 360 }));
  }

  reorderSelection(op: ZOrder) {
    const sel = this.s.sel;
    if (!sel.length) return;
    const i = this.s.slideIdx;
    this.store.commit((p) => updateSlideShapes(p, i, (list) => reorder(list, sel, op)));
  }

  alignSelection(op: AlignOp, toSlide: boolean) {
    const sel = this.s.sel;
    if (!sel.length) return;
    const i = this.s.slideIdx;
    const p0 = this.pres;
    this.store.commit((p) => updateSlideShapes(p, i, (list) => alignShapes(list, sel, op, toSlide || sel.length === 1, p0.width, p0.height)));
  }

  /** Seçim gruplanabilir mi (en az iki şekil, yer tutucu yok — PowerPoint'te yer tutucular gruplanamaz). */
  canGroup(): boolean {
    const shapes = this.selectedShapes();
    return shapes.length >= 2 && !shapes.some((x) => (x.kind === "sp" || x.kind === "pic") && x.ph);
  }

  groupSelection() {
    const sel = this.s.sel;
    if (sel.length < 2) return;
    if (!this.canGroup()) {
      this.message("Yer tutucular gruplanamaz. Gruplamak için yalnızca şekilleri seçin.", "info");
      return;
    }
    const i = this.s.slideIdx;
    let gid: string | undefined;
    this.store.commit((p) =>
      updateSlideShapes(p, i, (list) => {
        const r = groupShapes(list, sel);
        gid = r.groupId;
        return r.shapes;
      }),
    );
    if (gid) this.store.set({ sel: [gid], editing: null });
  }

  ungroupSelection() {
    const groups = this.selectedShapes().filter((x) => x.kind === "group");
    if (!groups.length) return;
    const i = this.s.slideIdx;
    let ids: string[] = [];
    this.store.commit((p) =>
      updateSlideShapes(p, i, (list) => {
        const r = ungroupShapes(list, groups.map((g) => g.id));
        ids = r.childIds;
        return r.shapes;
      }),
    );
    this.store.set({ sel: ids, editing: null });
  }

  /* ------------------------------ Pano ------------------------------ */

  copySelection(cut = false): boolean {
    const shapes = this.selectedShapes();
    if (!shapes.length) return false;
    this.store.clipboard = { kind: "shapes", shapes: structuredClone(shapes), sourceTheme: this.pres.theme.name };
    if (cut) this.deleteSelection();
    return true;
  }

  copySlides(cut = false) {
    if (inMasterView(this)) return;
    const ids = new Set(this.selectedSlideIds());
    const slides = this.pres.slides.filter((x) => ids.has(x.id));
    this.store.clipboard = { kind: "slides", slides: structuredClone(slides) };
    if (cut) this.deleteSlides([...ids]);
  }

  pasteInternal(): boolean {
    const cb = this.store.clipboard;
    if (!cb) return false;
    if (cb.kind === "slides" && cb.slides?.length) {
      if (inMasterView(this)) return false;
      const at = this.s.slideIdx + 1;
      const p = this.pres;
      const copies = cb.slides.map((sl) => {
        const c = cloneSlide(sl);
        const lt = p.layouts.find((l) => l.id === c.layoutId) ? c.layoutId : layoutByType(p, "obj").id;
        return { ...c, layoutId: lt };
      });
      this.store.commit((pp) => ({ ...pp, slides: [...pp.slides.slice(0, at), ...copies, ...pp.slides.slice(at)] }), { slideIdx: at + copies.length - 1, sel: [] });
      return true;
    }
    if (cb.kind === "shapes" && cb.shapes?.length) {
      const sl = this.slide();
      if (!sl) return false;
      // Aynı slayta yapıştırılıyorsa kaydır.
      const existing = new Set(flatShapes(sl.shapes).map((x) => `${x.x.toFixed(1)},${x.y.toFixed(1)},${x.w.toFixed(1)}`));
      let off = 0;
      while (cb.shapes.some((x) => existing.has(`${(x.x + off).toFixed(1)},${(x.y + off).toFixed(1)},${x.w.toFixed(1)}`)) && off < 400) off += 16;
      const copies = cb.shapes.map((x) => moveShape(cloneShape(x), off, off));
      this.addShapes(copies.map((c) => (c.kind === "sp" && c.ph ? { ...c, ph: undefined, txBox: c.text ? true : undefined } as Shape : c)));
      return true;
    }
    return false;
  }

  pasteText(text: string) {
    const p = this.pres;
    const lines = text.replace(/\r\n?/g, "\n");
    const w = Math.min(p.width * 0.6, 640);
    const tb = makeTextBox((p.width - w) / 2, p.height / 2 - 30, w, 60, lines);
    tb.name = `Metin Kutusu ${this.nextShapeNumber()}`;
    this.addShapes([tb]);
  }

  /* ------------------------------ Animasyonlar ------------------------------ */

  setAnimations(fn: (anims: Animation[], slide: Slide) => Animation[], merge?: string) {
    const i = this.s.slideIdx;
    this.store.commit((p) => updateSlide(p, i, (sl) => ({ ...sl, animations: fn(sl.animations, sl) })), { merge });
  }

  /** Seçili şekillere animasyon uygular (galeri: var olanı değiştirir; ekle: yenisini ekler). */
  applyAnimation(effect: string, cls: Animation["cls"], add: boolean, option?: string) {
    const sel = this.s.sel;
    if (!sel.length) return;
    this.setAnimations((anims) => {
      let list = anims.slice();
      for (const id of sel) {
        if (effect === "none") {
          list = list.filter((a) => a.shapeId !== id);
          continue;
        }
        const existing = list.findIndex((a) => a.shapeId === id);
        const anim: Animation = { id: uid("a"), shapeId: id, cls, effect, option, trigger: "click", duration: defaultDuration(effect), delay: 0 };
        if (!add && existing >= 0) {
          const ex = list[existing];
          list[existing] = { ...anim, id: ex.id, trigger: ex.trigger, delay: ex.delay };
          list = list.filter((a, k) => k === existing || a.shapeId !== id);
        } else list.push(anim);
      }
      return list;
    });
  }

  /* ------------------------------ Üst bilgi / alt bilgi ------------------------------ */

  applyHeaderFooter(hf: Presentation["headerFooter"], all: boolean) {
    if (inMasterView(this)) {
      this.message("Üst bilgi ve alt bilgi, Asıl Slayt Görünümü kapatıldıktan sonra slaytlara uygulanabilir.", "info");
      return;
    }
    const set = new Set(this.selectedSlideIds());
    this.store.commit((p) => {
      const np = { ...p, headerFooter: hf };
      return {
        ...np,
        slides: p.slides.map((sl) => {
          if (!all && !set.has(sl.id)) return sl;
          const l = layoutOf(np, sl)!;
          const kept = sl.shapes.filter((x) => !(x.kind === "sp" && x.ph && (x.ph.type === "dt" || x.ph.type === "ftr" || x.ph.type === "sldNum")));
          return { ...sl, shapes: [...kept, ...footerShapes(np, l, l.type === "title")] };
        }),
      };
    });
  }

  /* ------------------------------ Yardımcılar ------------------------------ */

  selectionBox(): Box | null {
    const shapes = this.selectedShapes();
    if (!shapes.length) return null;
    return unionBox(shapes.map(aabb));
  }

  tableSelected(): TableShape | undefined {
    const s = this.selectedShapes();
    return s.length === 1 && s[0].kind === "table" ? s[0] : undefined;
  }
}

export function defaultDuration(effect: string): number {
  if (effect === "appear" || effect === "disappear") return 0;
  if (effect === "spin" || effect === "pulse" || effect === "teeter") return 1;
  return 0.5;
}

function geomName(geom: string): string {
  const map: Record<string, string> = {
    rect: "Dikdörtgen",
    roundRect: "Dikdörtgen: Köşeleri Yuvarlatılmış",
    ellipse: "Oval",
    triangle: "İkizkenar Üçgen",
    line: "Düz Bağlayıcı",
    straightConnector1: "Düz Ok Bağlayıcısı",
  };
  return map[geom] ?? "Şekil";
}

export { NO_LINE };
