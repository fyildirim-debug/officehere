/** Yüzen nesne (grafik, resim, şekil) klavye ve pano işlemleri: Sil, Esc, ok tuşlarıyla kaydırma, kopyala/kes/yapıştır. */
import type { Controller } from "../controller";
import type { ChartDef, ImageDef, ShapeDef } from "../model/types";
import { deleteObject } from "../charts/FloatingLayer";
import { anchorAt } from "../charts/FloatingLayer";

type ObjKind = "chart" | "image" | "shape";
type AnyObj = ChartDef | ImageDef | ShapeDef;

/** Nesne panosu (uygulama içi). Panoya yazılan HTML'deki işaret bu kaydı gösterir. */
export interface ObjClip {
  kind: ObjKind;
  obj: AnyObj;
  stamp: string;
  /** Kaç kez yapıştırıldı (her yapıştırmada biraz kaydırılır). */
  pasted: number;
}

let clip: ObjClip | null = null;

export function objectClipMarker(): string | null {
  return clip ? `<!--tb-obj:${clip.stamp}-->` : null;
}

function listOf(ctl: Controller, kind: ObjKind): AnyObj[] {
  const s = ctl.sheet;
  return kind === "chart" ? s.charts : kind === "image" ? s.images : s.shapes;
}

export function selectedObject(ctl: Controller): { kind: ObjKind; obj: AnyObj } | null {
  const sel = ctl.selObj;
  if (!sel) return null;
  const obj = listOf(ctl, sel.kind).find((o) => o.id === sel.id);
  return obj ? { kind: sel.kind, obj } : null;
}

/** Seçili nesneyi panoya kopyalar (cut: sonra siler). */
export function copyObject(ctl: Controller, cut: boolean): boolean {
  const s = selectedObject(ctl);
  if (!s) return false;
  clip = { kind: s.kind, obj: structuredClone(s.obj), stamp: `${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`, pasted: 0 };
  if (cut) {
    clip.pasted = -1;
    deleteObject(ctl, s.kind, s.obj.id);
  }
  return true;
}

export function hasObjectClip(): boolean {
  return !!clip;
}

/** Panodaki nesneyi etkin sayfaya yapıştırır (etkin hücreye ya da kopyanın yanına). */
export function pasteObject(ctl: Controller): boolean {
  if (!clip) return false;
  const sheet = ctl.sheet;
  const wb = ctl.wb;
  clip.pasted++;
  const src = structuredClone(clip.obj);
  const prefix = clip.kind === "chart" ? "chart" : clip.kind === "image" ? "img" : "shape";
  const id = wb.newId(prefix);
  // Nesne hâlâ seçiliyken aynı sayfaya kopya: biraz sağa/aşağı kaydır; bir hücre seçildiyse ya da
  // başka sayfadaysa etkin hücreye yerleştir
  const sameSheet = listOf(ctl, clip.kind).some((o) => o.id === clip!.obj.id);
  let anchor = src.anchor;
  if (!ctl.selObj) anchor = { row: ctl.sel.active.row, col: ctl.sel.active.col, dx: 0, dy: 0 };
  else if (sameSheet && clip.pasted > 0) {
    const x = sheet.cols.offset(src.anchor.col) + src.anchor.dx + 16 * clip.pasted;
    const y = sheet.rows.offset(src.anchor.row) + src.anchor.dy + 16 * clip.pasted;
    anchor = anchorAt(sheet, x, y);
  } else if (!sameSheet) anchor = { row: ctl.sel.active.row, col: ctl.sel.active.col, dx: 0, dy: 0 };
  const copy = { ...src, id, anchor } as AnyObj;
  const kind = clip.kind;
  ctl.ops.sheetOp(sheet, "Yapıştır", () => {
    if (kind === "chart") sheet.charts.push(copy as ChartDef);
    else if (kind === "image") sheet.images.push(copy as ImageDef);
    else sheet.shapes.push(copy as ShapeDef);
  });
  ctl.selObj = { kind, id };
  ctl.emit();
  return true;
}

/** Ok tuşlarıyla seçili nesneyi kaydırır (Ctrl ile 1 piksel). */
export function nudgeObject(ctl: Controller, dx: number, dy: number): void {
  const s = selectedObject(ctl);
  if (!s) return;
  const sheet = ctl.sheet;
  ctl.ops.sheetOp(sheet, "Taşı", () => {
    const o = listOf(ctl, s.kind).find((x) => x.id === s.obj.id);
    if (!o) return;
    const x = sheet.cols.offset(o.anchor.col) + o.anchor.dx + dx;
    const y = sheet.rows.offset(o.anchor.row) + o.anchor.dy + dy;
    o.anchor = anchorAt(sheet, Math.max(0, x), Math.max(0, y));
  });
}

/** Nesne seçiliyken ızgara klavye olayları. İşlendiyse true. */
export function handleObjectKey(ctl: Controller, ev: KeyboardEvent | React.KeyboardEvent): boolean {
  const s = selectedObject(ctl);
  if (!s) return false;
  const ctrl = ev.ctrlKey || ev.metaKey;
  switch (ev.key) {
    case "Delete":
    case "Backspace":
      deleteObject(ctl, s.kind, s.obj.id);
      return true;
    case "Escape":
      ctl.selObj = null;
      ctl.emit();
      return true;
    case "ArrowUp":
    case "ArrowDown":
    case "ArrowLeft":
    case "ArrowRight": {
      const step = ctrl ? 1 : 8;
      nudgeObject(ctl, ev.key === "ArrowLeft" ? -step : ev.key === "ArrowRight" ? step : 0, ev.key === "ArrowUp" ? -step : ev.key === "ArrowDown" ? step : 0);
      return true;
    }
    case "Enter":
      if (s.kind === "shape") {
        ctl.openDialog("shapeText", { id: s.obj.id });
        return true;
      }
      if (s.kind === "chart") {
        ctl.openDialog("chartFormat");
        return true;
      }
      return false;
  }
  // Şekil seçiliyken yazılan karakter şeklin metnini başlatır
  if (s.kind === "shape" && !ctrl && !ev.altKey && ev.key.length === 1) {
    ctl.openDialog("shapeText", { id: s.obj.id, text: ev.key });
    return true;
  }
  return false;
}

/** Bul ve Seç > Nesneleri Seç: sayfadaki ilk görünür nesne seçilir (Sekme ile sonrakine geçilir). */
export function selectFirstObject(ctl: Controller): void {
  const all = objectList(ctl).filter((o) => !o.hidden);
  if (!all.length) return;
  ctl.selObj = { kind: all[0].kind, id: all[0].id };
  ctl.emit();
}

export interface ObjectEntry {
  kind: ObjKind;
  id: string;
  name: string;
  hidden: boolean;
}

/** Sayfadaki nesneler (çizim sırasıyla, en üstteki en sonda). */
export function objectList(ctl: Controller): ObjectEntry[] {
  const s = ctl.sheet;
  return [
    ...s.images.map((o, i) => ({ kind: "image" as const, id: o.id, name: o.name || `Resim ${i + 1}`, hidden: !!o.hidden })),
    ...s.shapes.map((o, i) => ({ kind: "shape" as const, id: o.id, name: o.text ? `${o.text.slice(0, 20)}` : `Şekil ${i + 1}`, hidden: !!o.hidden })),
    ...s.charts.map((o, i) => ({ kind: "chart" as const, id: o.id, name: o.title ? `Grafik ${i + 1} (${o.title})` : `Grafik ${i + 1}`, hidden: !!o.hidden })),
  ];
}

/** Nesneyi göster/gizle (Seçim Bölmesi). */
export function setObjectHidden(ctl: Controller, kind: ObjKind, id: string, hidden: boolean): void {
  const sheet = ctl.sheet;
  ctl.ops.sheetOp(sheet, hidden ? "Gizle" : "Göster", () => {
    const o = listOf(ctl, kind).find((x) => x.id === id);
    if (o) o.hidden = hidden;
  });
  if (hidden && ctl.selObj?.id === id) ctl.selObj = null;
  ctl.emit();
}
