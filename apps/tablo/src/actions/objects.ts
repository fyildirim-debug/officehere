/** Yüzen nesneler: grafik, resim, şekil ekleme ve güncelleme. */
import { colToName, quoteSheetName } from "@officehere/formula";
import type { Controller } from "../controller";
import { currentRegion } from "../model/selection";
import type { ChartDef, ChartType, ImageDef, Rect, ShapeDef } from "../model/types";
import { defaultSeriesInRows } from "../charts/chartData";
import { STRUCTURE_MESSAGE, structureLocked } from "../model/protection";
import { pickImage } from "./imageFile";

function absRange(r: Rect): string {
  return `$${colToName(r.c1)}$${r.r1 + 1}:$${colToName(r.c2)}$${r.r2 + 1}`;
}

/** Grafik için veri aralığı: çok hücreli seçim ya da geçerli bölge. */
export function chartSourceRange(ctl: Controller): Rect | null {
  const r = ctl.range;
  if (r.r1 !== r.r2 || r.c1 !== r.c2) {
    const ext = ctl.wb.dataExtent(ctl.sheet);
    return { ...r, r2: Math.min(r.r2, Math.max(r.r1, ext.rows - 1)), c2: Math.min(r.c2, Math.max(r.c1, ext.cols - 1)) };
  }
  if (ctl.wb.isEmpty(ctl.sheet, r.r1, r.c1)) return null;
  return currentRegion(ctl.wb, ctl.sheet, ctl.sel.active);
}

const TYPE_TITLES: Partial<Record<ChartType, string>> = { pie: "Grafik Başlığı", doughnut: "Grafik Başlığı" };

export function insertChart(ctl: Controller, type: ChartType): void {
  if (ctl.edit) ctl.commitEdit(null);
  const sheet = ctl.sheet;
  const r = chartSourceRange(ctl);
  if (!r) {
    ctl.alert("Grafik oluşturmak için önce veri içeren bir hücre aralığı seçin.");
    return;
  }
  // Tek seri ise başlık olarak seri adını kullan
  const header = ctl.wb.getValue(sheet, r.r1, r.c2);
  const title = r.c2 - r.c1 <= 1 && typeof header === "string" ? header : (TYPE_TITLES[type] ?? "Grafik Başlığı");
  // Görünür alanın içine, verinin sağına yerleştir
  const col = Math.min(r.c2 + 2, sheet.scroll.col + 8);
  const def: ChartDef = {
    id: ctl.wb.newId("chart"),
    type,
    range: `${quoteSheetName(sheet.name)}!${absRange(r)}`,
    seriesInRows: defaultSeriesInRows(r),
    title,
    showTitle: true,
    showLegend: true,
    legendPos: type === "pie" || type === "doughnut" ? "right" : "bottom",
    showDataLabels: false,
    showGridlines: type !== "pie" && type !== "doughnut",
    palette: 0,
    style: 1,
    anchor: { row: Math.max(sheet.scroll.row + 1, r.r1), col: Math.max(col, sheet.scroll.col + 1), dx: 0, dy: 0 },
    width: 480,
    height: 288,
  };
  ctl.ops.sheetOp(sheet, "Grafik Ekle", () => sheet.charts.push(def));
  ctl.selObj = { kind: "chart", id: def.id };
  ctl.emit();
}

export function updateChart(ctl: Controller, id: string, patch: Partial<ChartDef>, label = "Grafik"): void {
  const sheet = ctl.sheet;
  if (!sheet.charts.some((c) => c.id === id)) return;
  ctl.ops.sheetOp(sheet, label, () => {
    // Veri aralığı ya da seri yönü değişince açık seri tanımları bırakılır (aralıktan yeniden kurulur)
    const reset = patch.range !== undefined || patch.seriesInRows !== undefined;
    sheet.charts = sheet.charts.map((c) => (c.id === id ? { ...c, ...patch, ...(reset && !("series" in patch) ? { series: undefined } : {}) } : c));
  });
}

/** Grafiği yeni sayfaya taşır (Grafiği Taşı). */
export function moveChartToSheet(ctl: Controller, id: string, toIndex: number): void {
  const from = ctl.sheet;
  const to = ctl.wb.sheets[toIndex];
  const c = from.charts.find((x) => x.id === id);
  if (!c || !to || to === from) return;
  ctl.ops.snapshotOp("Grafiği Taşı", () => {
    from.charts = from.charts.filter((x) => x.id !== id);
    to.charts.push({ ...c, anchor: { row: 1, col: 1, dx: 0, dy: 0 } });
  });
  ctl.selObj = null;
  ctl.activateSheet(toIndex);
}

export function insertImage(ctl: Controller, name: string, src: string, natW: number, natH: number): void {
  const sheet = ctl.sheet;
  const scale = Math.min(1, 480 / natW, 360 / natH);
  const { row, col } = ctl.sel.active;
  const img: ImageDef = { id: ctl.wb.newId("img"), name, src, anchor: { row, col, dx: 0, dy: 0 }, width: Math.round(natW * scale), height: Math.round(natH * scale) };
  ctl.ops.sheetOp(sheet, "Resim Ekle", () => sheet.images.push(img));
  ctl.selObj = { kind: "image", id: img.id };
  ctl.emit();
}

/** Dosya seçtirip resmi ekler (PNG, JPEG, GIF, SVG, WEBP). */
export function pickAndInsertImage(ctl: Controller): void {
  pickImage(ctl, (name, src, w, h) => insertImage(ctl, name, src, w, h));
}

export const SHAPE_KINDS: { kind: ShapeDef["kind"]; label: string }[] = [
  { kind: "rect", label: "Dikdörtgen" },
  { kind: "roundRect", label: "Yuvarlatılmış Dikdörtgen" },
  { kind: "ellipse", label: "Oval" },
  { kind: "triangle", label: "İkizkenar Üçgen" },
  { kind: "arrow", label: "Sağ Ok" },
  { kind: "line", label: "Çizgi" },
  { kind: "textbox", label: "Metin Kutusu" },
];

export function insertShape(ctl: Controller, kind: ShapeDef["kind"]): void {
  // Metin kutusu sürükleyerek çizilir (Excel gibi); diğer şekiller etkin hücrenin yanına eklenir
  if (kind === "textbox") {
    if (ctl.edit) ctl.commitEdit(null);
    ctl.drawShape = "textbox";
    ctl.status = "Metin kutusunu çizmek için sayfada sürükleyin (iptal için Esc)";
    ctl.emit();
    return;
  }
  const { row, col } = ctl.sel.active;
  insertShapeAt(ctl, kind, { row, col: col + 1, dx: 0, dy: 0 }, kind === "line" ? 160 : 140, kind === "line" ? 2 : 90);
}

/** Şekli verilen konuma ekler ve seçer; metin kutusunda yerinde yazma başlar. */
export function insertShapeAt(ctl: Controller, kind: ShapeDef["kind"], anchor: ShapeDef["anchor"], width: number, height: number): void {
  const sheet = ctl.sheet;
  const s: ShapeDef = {
    id: ctl.wb.newId("shape"),
    kind,
    text: "",
    fill: kind === "textbox" ? "#FFFFFF" : kind === "line" ? "transparent" : "#156082",
    stroke: kind === "textbox" ? "#7F7F7F" : kind === "line" ? "#156082" : "#0F4761",
    anchor,
    width,
    height,
  };
  ctl.ops.sheetOp(sheet, "Şekil Ekle", () => sheet.shapes.push(s));
  ctl.selObj = { kind: "shape", id: s.id };
  if (kind === "textbox") ctl.shapeEdit = { id: s.id, text: "" };
  ctl.emit();
}

export function updateShape(ctl: Controller, id: string, patch: Partial<ShapeDef>, label = "Şekil Biçimi"): void {
  const sheet = ctl.sheet;
  ctl.ops.sheetOp(sheet, label, () => {
    sheet.shapes = sheet.shapes.map((s) => (s.id === id ? { ...s, ...patch } : s));
  });
}

/** Nesneyi öne/arkaya getirir (dizideki sıra çizim sırasıdır). */
/** Yığın sırası: z verilmemişse kimlikteki oluşturulma numarası. */
export function objZ(o: { id: string; z?: number }): number {
  return o.z ?? Number(/(\d+)$/.exec(o.id)?.[1] ?? 0);
}

/** Sayfadaki tüm yüzen nesneler (arkadan öne). */
export function stackOf(sheet: Controller["sheet"]): { kind: "chart" | "image" | "shape"; o: { id: string; z?: number } }[] {
  return [...sheet.images.map((o) => ({ kind: "image" as const, o })), ...sheet.shapes.map((o) => ({ kind: "shape" as const, o })), ...sheet.charts.map((o) => ({ kind: "chart" as const, o }))].sort((a, b) => objZ(a.o) - objZ(b.o));
}

export function reorderObject(ctl: Controller, kind: "chart" | "image" | "shape", id: string, front: boolean): void {
  const sheet = ctl.sheet;
  ctl.ops.sheetOp(sheet, front ? "Öne Getir" : "Arkaya Gönder", () => {
    const all = stackOf(sheet);
    const zs = all.map((x) => objZ(x.o));
    const target = all.find((x) => x.kind === kind && x.o.id === id);
    if (target) target.o.z = front ? Math.max(...zs) + 1 : Math.min(...zs) - 1;
  });
}

/**
 * F11: seçili veriden grafiği kendi sayfasına ("Grafik1") yerleştirir. Sayfa
 * yalnız grafiği taşır (kılavuz çizgileri kapalı, grafik pencereyi doldurur).
 */
export function chartOnNewSheet(ctl: Controller): void {
  if (ctl.edit) ctl.commitEdit(null);
  if (structureLocked(ctl.wb)) return ctl.alert(STRUCTURE_MESSAGE);
  const src = ctl.sheet;
  const r = chartSourceRange(ctl);
  if (!r) {
    ctl.alert("Grafik oluşturmak için önce veri içeren bir hücre aralığı seçin.");
    return;
  }
  const header = ctl.wb.getValue(src, r.r1, r.c2);
  ctl.ops.snapshotOp("Grafik Sayfası", () => {
    const wb = ctl.wb;
    let n = 1;
    while (wb.sheetByName(`Grafik${n}`)) n++;
    const sheet = wb.addSheet(`Grafik${n}`, wb.activeIndex);
    sheet.showGridlines = false;
    sheet.showHeadings = false;
    sheet.charts.push({
      id: wb.newId("chart"),
      type: "column",
      range: `${quoteSheetName(src.name)}!${absRange(r)}`,
      seriesInRows: defaultSeriesInRows(r),
      title: r.c2 - r.c1 <= 1 && typeof header === "string" ? header : "Grafik Başlığı",
      showTitle: true,
      showLegend: true,
      legendPos: "bottom",
      showDataLabels: false,
      showGridlines: true,
      palette: 0,
      style: 1,
      anchor: { row: 0, col: 0, dx: 8, dy: 8 },
      width: 900,
      height: 560,
    });
    wb.activeIndex = wb.sheets.indexOf(sheet);
  });
  const sheet = ctl.sheet;
  ctl.selObj = sheet.charts[0] ? { kind: "chart", id: sheet.charts[0].id } : null;
  ctl.emit();
}

/** Alt+F1: seçili veriden varsayılan türde (kümelenmiş sütun) hızlı grafik. */
export function quickChart(ctl: Controller): void {
  insertChart(ctl, "column");
}
