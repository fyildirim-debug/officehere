/** Yüzen nesneler: grafik, resim, şekil ekleme ve güncelleme. */
import { colToName, quoteSheetName } from "@officehere/formula";
import type { Controller } from "../controller";
import { currentRegion } from "../model/selection";
import type { ChartDef, ChartType, ImageDef, Rect, ShapeDef } from "../model/types";
import { defaultSeriesInRows } from "../charts/chartData";

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
    sheet.charts = sheet.charts.map((c) => (c.id === id ? { ...c, ...patch } : c));
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
  const input = document.createElement("input");
  input.type = "file";
  input.accept = "image/png,image/jpeg,image/gif,image/svg+xml,image/webp,image/bmp";
  input.onchange = () => {
    const f = input.files?.[0];
    if (!f) return;
    const reader = new FileReader();
    reader.onload = () => {
      const src = String(reader.result);
      const im = new Image();
      im.onload = () => insertImage(ctl, f.name, src, im.naturalWidth || 200, im.naturalHeight || 150);
      im.onerror = () => ctl.alert("Resim dosyası okunamadı.");
      im.src = src;
    };
    reader.readAsDataURL(f);
  };
  input.click();
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
  const sheet = ctl.sheet;
  const { row, col } = ctl.sel.active;
  const s: ShapeDef = {
    id: ctl.wb.newId("shape"),
    kind,
    text: kind === "textbox" ? "Metin" : "",
    fill: kind === "textbox" || kind === "line" ? "transparent" : "#156082",
    stroke: kind === "textbox" ? "#7F7F7F" : kind === "line" ? "#156082" : "#0F4761",
    anchor: { row, col: col + 1, dx: 0, dy: 0 },
    width: kind === "line" ? 160 : 140,
    height: kind === "line" ? 2 : 90,
  };
  ctl.ops.sheetOp(sheet, "Şekil Ekle", () => sheet.shapes.push(s));
  ctl.selObj = { kind: "shape", id: s.id };
  ctl.emit();
  if (kind === "textbox") ctl.openDialog("shapeText", { id: s.id });
}

export function updateShape(ctl: Controller, id: string, patch: Partial<ShapeDef>, label = "Şekil Biçimi"): void {
  const sheet = ctl.sheet;
  ctl.ops.sheetOp(sheet, label, () => {
    sheet.shapes = sheet.shapes.map((s) => (s.id === id ? { ...s, ...patch } : s));
  });
}

/** Nesneyi öne/arkaya getirir (dizideki sıra çizim sırasıdır). */
export function reorderObject(ctl: Controller, kind: "chart" | "image" | "shape", id: string, front: boolean): void {
  const sheet = ctl.sheet;
  ctl.ops.sheetOp(sheet, front ? "Öne Getir" : "Arkaya Gönder", () => {
    const move = <T extends { id: string }>(list: T[]): T[] => {
      const o = list.find((x) => x.id === id);
      if (!o) return list;
      const rest = list.filter((x) => x.id !== id);
      return front ? [...rest, o] : [o, ...rest];
    };
    if (kind === "chart") sheet.charts = move(sheet.charts);
    else if (kind === "image") sheet.images = move(sheet.images);
    else sheet.shapes = move(sheet.shapes);
  });
}
