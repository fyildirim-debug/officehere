/** Asıl Slayt Görünümü komutları (Görünüm > Asıl Slayt, Asıl Slayt sekmesi). */
import {
  addLayoutPlaceholder,
  deleteLayout,
  duplicateLayout,
  fromMasterView,
  insertLayout,
  layoutUsage,
  MASTER_SLIDE_ID,
  pseudoLayoutOf,
  renameLayout,
  syncMasterView,
  toggleLayoutPlaceholders,
  toMasterView,
} from "../model/masterView";
import type { Presentation } from "../model/types";
import type { Actions } from "./actions";
import type { ViewMode } from "./store";

export function inMasterView(a: Actions): boolean {
  return a.s.view === "master" && !!a.s.masterBase;
}

/** Kaydetme, yazdırma ve gösteri için belgenin gerçek hali (asıl görünüm açıksa değişiklikler uygulanmış). */
export function documentPres(a: Actions): Presentation {
  const mb = a.s.masterBase;
  return mb && a.s.view === "master" ? fromMasterView(mb.pres, a.s.pres) : a.s.pres;
}

export function openMasterView(a: Actions) {
  const s = a.s;
  if (s.masterBase) return;
  a.stopEditing();
  const base = s.pres;
  const mv = toMasterView(base);
  const cur = base.slides[s.slideIdx];
  const li = cur ? mv.slides.findIndex((x) => x.id === cur.layoutId) : 0;
  const idx = li >= 0 ? li : 0;
  a.store.beginSession();
  a.store.normalize = syncMasterView;
  a.store.set({
    masterBase: { pres: base, slideIdx: s.slideIdx, dirty: s.dirty, tab: s.ribbonTab },
    pres: mv,
    view: "master",
    slideIdx: idx,
    slideSel: [mv.slides[idx].id],
    sel: [],
    editing: null,
    cellSel: null,
    insert: null,
    cropId: null,
    ink: null,
    pane: s.pane === "background" || s.pane === "shape" || s.pane === "selection" ? s.pane : null,
    ribbonTab: "master",
  });
}

/** Asıl görünümü kapatır ve değişiklikleri sunuya tek geri alma adımı olarak uygular. */
export function closeMasterView(a: Actions, view: ViewMode = "normal") {
  const s = a.s;
  const mb = s.masterBase;
  if (!mb) {
    a.store.set({ view });
    return;
  }
  a.stopEditing();
  const real = fromMasterView(mb.pres, s.pres);
  const changed = a.store.endSession();
  const idx = Math.min(mb.slideIdx, Math.max(0, mb.pres.slides.length - 1));
  a.store.set({
    pres: mb.pres,
    masterBase: null,
    view,
    slideIdx: idx,
    slideSel: mb.pres.slides[idx] ? [mb.pres.slides[idx].id] : [],
    sel: [],
    editing: null,
    cellSel: null,
    dirty: mb.dirty,
    ribbonTab: mb.tab === "master" ? "home" : mb.tab,
  });
  if (changed) {
    a.store.commit(() => real, { slideIdx: idx });
    // Asıl görünümdeyken kaydedilip sonra değiştirilmediyse belge kaydedilmiş sayılır.
    if (!s.dirty) a.store.set({ dirty: false });
  }
}

/** Asıl görünüm açıkken oturumu değişiklikleri atarak sonlandırır (yeni dosya açılırken). */
export function abortMasterView(a: Actions) {
  if (!a.s.masterBase) return;
  a.store.endSession();
  a.store.set({ masterBase: null });
}

function current(a: Actions) {
  return a.s.pres.slides[a.s.slideIdx];
}

export function masterInsertLayout(a: Actions) {
  const s = a.s;
  const after = s.slideIdx === 0 ? s.pres.slides.length - 1 : s.slideIdx;
  const r = insertLayout(s.pres, after);
  a.store.commit(() => r.pres, { slideIdx: r.index, sel: [], editing: null });
}

export function masterDuplicateLayout(a: Actions) {
  const r = duplicateLayout(a.s.pres, a.s.slideIdx);
  if (!r) return;
  a.store.commit(() => r.pres, { slideIdx: r.index, sel: [], editing: null });
}

export function masterDeleteLayout(a: Actions) {
  const cur = current(a);
  const mb = a.s.masterBase;
  if (!cur || !mb) return;
  if (cur.id === MASTER_SLIDE_ID) {
    a.message("Asıl slayt silinemez.", "warning");
    return;
  }
  const n = layoutUsage(mb.pres, cur.id);
  if (n > 0) {
    a.message(`Bu düzen sunudaki ${n} slayt tarafından kullanıldığı için silinemez.`, "warning");
    return;
  }
  const i = a.s.slideIdx;
  a.store.commit((p) => deleteLayout(p, cur.id), { slideIdx: Math.max(0, i - 1), sel: [], editing: null });
}

export function masterRenameLayout(a: Actions, name: string) {
  const cur = current(a);
  if (!cur || cur.id === MASTER_SLIDE_ID) return;
  a.store.commit((p) => renameLayout(p, cur.id, name));
}

export function masterLayoutName(a: Actions): string {
  const cur = current(a);
  if (!cur) return "";
  return pseudoLayoutOf(a.s.pres, cur.id)?.name ?? "";
}

export function masterToggle(a: Actions, kind: "title" | "footers", on: boolean) {
  const cur = current(a);
  if (!cur || cur.id === MASTER_SLIDE_ID) return;
  a.store.commit((p) => toggleLayoutPlaceholders(p, cur.id, kind, on), { sel: [], editing: null });
}

export function masterHas(a: Actions, kind: "title" | "footers"): boolean {
  const cur = current(a);
  if (!cur) return false;
  const types = kind === "title" ? ["title", "ctrTitle"] : ["dt", "ftr", "sldNum"];
  return cur.shapes.some((x) => x.kind === "sp" && !!x.ph && types.includes(x.ph.type));
}

export function masterAddPlaceholder(a: Actions, type: "obj" | "body" | "pic" | "tbl") {
  const cur = current(a);
  if (!cur) return;
  if (cur.id === MASTER_SLIDE_ID) {
    a.message("Yer tutucular düzenlere eklenir; soldan bir düzen seçin.", "info");
    return;
  }
  const r = addLayoutPlaceholder(a.s.pres, cur.id, type);
  if (!r) return;
  a.store.commit(() => r.pres, { sel: [r.shapeId], editing: null });
}

/** Düzeni kullanan slaytların açıklaması (küçük resim ipucu). */
export function layoutUsageText(a: Actions, slideId: string): string {
  const mb = a.s.masterBase;
  if (!mb) return "";
  if (slideId === MASTER_SLIDE_ID) return "Asıl Slayt: Tüm slaytlar tarafından kullanılıyor";
  const name = pseudoLayoutOf(a.s.pres, slideId)?.name ?? "Düzen";
  const n = layoutUsage(mb.pres, slideId);
  return `${name} Düzeni: ${n ? `${n} slayt tarafından kullanılıyor` : "Hiçbir slayt tarafından kullanılmıyor"}`;
}
