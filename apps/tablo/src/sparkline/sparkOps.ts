/** Mini grafik işlemleri: oluşturma, temizleme, grup ayarları, gruplama ve veri düzenleme (geri alınabilir). */
import { themeShades } from "@officehere/core";
import { quoteSheetName } from "@officehere/formula";
import type { Controller } from "../controller";
import { parseRangeLoose } from "../model/inputRules";
import { contains, rectToA1 } from "../model/rect";
import type { Rect, SparklineGroup } from "../model/types";
import type { SheetModel, WorkbookModel } from "../model/workbook";
import { POINT_RED, sparkAt, type SparkItem } from "./geometry";

export const SPARK_INVALID = "Konum başvurusu veya veri aralığı geçerli değil.";

/** "Sayfa1!$A$1:$D$20" ya da "A1:D20" → sayfa + aralık. */
export function parseRangeText(wb: WorkbookModel, current: SheetModel, text: string): { sheet: SheetModel; rect: Rect } | null {
  const t = text.trim().replace(/^=/, "");
  if (!t) return null;
  const bang = t.lastIndexOf("!");
  let sheet: SheetModel | undefined = current;
  let body = t;
  if (bang > 0) {
    let name = t.slice(0, bang);
    if (name.startsWith("'")) name = name.slice(1, -1).replace(/''/g, "'");
    sheet = wb.sheetByName(name);
    body = t.slice(bang + 1);
  }
  const r = parseRangeLoose(body);
  return sheet && r ? { sheet, rect: r } : null;
}

/** Varsayılan mini grafik rengi: tema Vurgu 1, %25 koyu. */
export function defaultSparkColor(wb: WorkbookModel): string {
  return themeShades(wb.theme.colors[4])[4].toUpperCase();
}

/**
 * Veri aralığını konum hücrelerine paylaştırır. Konum tek satır ya da tek
 * sütun olmalı; birden çok hücrede veri satır ya da sütun sayısı eşleşmeli.
 */
export function planItems(data: { sheet: SheetModel; rect: Rect }, loc: Rect): SparkItem[] | null {
  const vertical = loc.c1 === loc.c2;
  const horizontal = loc.r1 === loc.r2;
  if (!vertical && !horizontal) return null;
  const n = vertical ? loc.r2 - loc.r1 + 1 : loc.c2 - loc.c1 + 1;
  const R = data.rect.r2 - data.rect.r1 + 1;
  const C = data.rect.c2 - data.rect.c1 + 1;
  const name = quoteSheetName(data.sheet.name);
  const ref = (r: Rect) => `${name}!${rectToA1(r)}`;
  const cell = (i: number) => (vertical ? { row: loc.r1 + i, col: loc.c1 } : { row: loc.r1, col: loc.c1 + i });
  if (n === 1) {
    if (R > 1 && C > 1) return null;
    return [{ ...cell(0), range: ref(data.rect) }];
  }
  // Dikey konumda önce satırlar, yatay konumda önce sütunlar eşleştirilir
  const byRows = vertical ? R === n : R === n && C !== n;
  const byCols = !byRows && C === n;
  if (!byRows && !byCols) return null;
  const items: SparkItem[] = [];
  for (let i = 0; i < n; i++) {
    const r = byRows ? { r1: data.rect.r1 + i, r2: data.rect.r1 + i, c1: data.rect.c1, c2: data.rect.c2 } : { r1: data.rect.r1, r2: data.rect.r2, c1: data.rect.c1 + i, c2: data.rect.c1 + i };
    items.push({ ...cell(i), range: ref(r) });
  }
  return items;
}

/** Hücrelerdeki mini grafikleri gruplarından çıkarır (boş kalan gruplar silinir). */
function removeAt(sheet: SheetModel, cells: { row: number; col: number }[]): void {
  const keys = new Set(cells.map((c) => `${c.row},${c.col}`));
  for (const g of sheet.sparklines) g.items = g.items.filter((it) => !keys.has(`${it.row},${it.col}`));
  sheet.sparklines = sheet.sparklines.filter((g) => g.items.length);
}

/** Ekle > Mini Grafikler: yeni grup oluşturur. Hata metni ya da null döner. */
export function createSparklines(ctl: Controller, type: SparklineGroup["type"], dataText: string, locText: string): string | null {
  const wb = ctl.wb;
  const data = parseRangeText(wb, ctl.sheet, dataText);
  const loc = parseRangeText(wb, ctl.sheet, locText);
  if (!data || !loc) return SPARK_INVALID;
  if (loc.sheet !== ctl.sheet) return "Konum başvurusu etkin sayfada olmalıdır.";
  const items = planItems(data, loc.rect);
  if (!items) return SPARK_INVALID;
  const sheet = ctl.sheet;
  const color = defaultSparkColor(wb);
  const group: SparklineGroup = { id: wb.newId("spark"), type, items, color, markerColor: color, negativeColor: POINT_RED, negative: type === "winLoss" ? true : undefined };
  ctl.ops.sheetOp(sheet, "Mini Grafik Ekle", () => {
    removeAt(sheet, items);
    sheet.sparklines.push(group);
  });
  ctl.selectRange(loc.rect, { row: loc.rect.r1, col: loc.rect.c1 });
  ctl.requestTab = "sparkline";
  ctl.emit();
  return null;
}

/** Grup Konumunu ve Verilerini Düzenle: grubun öğelerini yeniden kurar (ayarlar korunur). */
export function editGroupData(ctl: Controller, groupId: string, dataText: string, locText: string): string | null {
  const sheet = ctl.sheet;
  const g = sheet.sparklines.find((x) => x.id === groupId);
  if (!g) return null;
  const data = parseRangeText(ctl.wb, sheet, dataText);
  const loc = parseRangeText(ctl.wb, sheet, locText);
  if (!data || !loc) return SPARK_INVALID;
  if (loc.sheet !== sheet) return "Konum başvurusu etkin sayfada olmalıdır.";
  const items = planItems(data, loc.rect);
  if (!items) return SPARK_INVALID;
  ctl.ops.sheetOp(sheet, "Mini Grafik Verilerini Düzenle", () => {
    const live = sheet.sparklines.find((x) => x.id === groupId)!;
    live.items = [];
    removeAt(sheet, items);
    live.items = items;
    if (!sheet.sparklines.includes(live)) sheet.sparklines.push(live);
  });
  ctl.emit();
  return null;
}

/** Tek mini grafiğin veri aralığını değiştirir (Tek Mini Grafiğin Verilerini Düzenle). */
export function editItemData(ctl: Controller, row: number, col: number, dataText: string): string | null {
  const sheet = ctl.sheet;
  const data = parseRangeText(ctl.wb, sheet, dataText);
  if (!data || (data.rect.r1 !== data.rect.r2 && data.rect.c1 !== data.rect.c2)) return "Veri başvurusu geçerli değil.";
  const ref = `${quoteSheetName(data.sheet.name)}!${rectToA1(data.rect)}`;
  ctl.ops.sheetOp(sheet, "Mini Grafik Verilerini Düzenle", () => {
    const s = sparkAt(sheet, row, col);
    if (s) s.item.range = ref;
  });
  ctl.emit();
  return null;
}

/** Grup ayarlarını değiştirir. */
export function updateGroup(ctl: Controller, groupId: string, patch: Partial<SparklineGroup>, label: string): void {
  const sheet = ctl.sheet;
  ctl.ops.sheetOp(sheet, label, () => {
    const g = sheet.sparklines.find((x) => x.id === groupId);
    if (!g) return;
    Object.assign(g, patch);
    for (const k of Object.keys(g) as (keyof SparklineGroup)[]) if (g[k] === undefined) delete g[k];
  });
  ctl.emit();
}

/** Seçimdeki hücreler. */
function selectedCells(ctl: Controller): (r: number, c: number) => boolean {
  const ranges = ctl.sel.ranges;
  return (r, c) => ranges.some((x) => contains(x, r, c));
}

/** Temizle: seçili mini grafikler ya da seçimle kesişen grupların tamamı. */
export function clearSparklines(ctl: Controller, whole: boolean): void {
  const sheet = ctl.sheet;
  const inSel = selectedCells(ctl);
  ctl.ops.sheetOp(sheet, whole ? "Mini Grafik Gruplarını Temizle" : "Mini Grafikleri Temizle", () => {
    if (whole) sheet.sparklines = sheet.sparklines.filter((g) => !g.items.some((it) => inSel(it.row, it.col)));
    else {
      for (const g of sheet.sparklines) g.items = g.items.filter((it) => !inSel(it.row, it.col));
      sheet.sparklines = sheet.sparklines.filter((g) => g.items.length);
    }
  });
  ctl.emit();
}

/** Grubu Çöz: seçimdeki grupların her mini grafiği kendi grubuna ayrılır. */
export function ungroupSparklines(ctl: Controller): void {
  const sheet = ctl.sheet;
  const inSel = selectedCells(ctl);
  ctl.ops.sheetOp(sheet, "Mini Grafik Grubunu Çöz", () => {
    const next: SparklineGroup[] = [];
    for (const g of sheet.sparklines) {
      if (g.items.length < 2 || !g.items.some((it) => inSel(it.row, it.col))) {
        next.push(g);
        continue;
      }
      g.items.forEach((it, i) => next.push({ ...structuredClone(g), id: i === 0 ? g.id : ctl.wb.newId("spark"), items: [{ ...it }] }));
    }
    sheet.sparklines = next;
  });
  ctl.emit();
}

/** Gruplandır: seçimle kesişen gruplar etkin hücredeki grubun ayarlarıyla birleşir. */
export function groupSparklines(ctl: Controller): void {
  const sheet = ctl.sheet;
  const inSel = selectedCells(ctl);
  const { row, col } = ctl.sel.active;
  const main = sparkAt(sheet, row, col)?.group ?? sheet.sparklines.find((g) => g.items.some((it) => inSel(it.row, it.col)));
  if (!main) return;
  ctl.ops.sheetOp(sheet, "Mini Grafikleri Gruplandır", () => {
    const live = sheet.sparklines.find((g) => g.id === main.id)!;
    for (const g of sheet.sparklines) {
      if (g === live) continue;
      const take = g.items.filter((it) => inSel(it.row, it.col));
      if (!take.length) continue;
      live.items.push(...take);
      g.items = g.items.filter((it) => !inSel(it.row, it.col));
    }
    live.items.sort((a, b) => a.row - b.row || a.col - b.col);
    sheet.sparklines = sheet.sparklines.filter((g) => g.items.length);
  });
  ctl.emit();
}

/** Grubun konum aralığı metni (tek dikdörtgense "G2:G5", değilse virgülle). */
export function groupLocationText(g: SparklineGroup): string {
  const rows = g.items.map((i) => i.row);
  const cols = g.items.map((i) => i.col);
  const r: Rect = { r1: Math.min(...rows), r2: Math.max(...rows), c1: Math.min(...cols), c2: Math.max(...cols) };
  if ((r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1) === g.items.length) return rectToA1(r);
  return g.items.map((i) => rectToA1({ r1: i.row, r2: i.row, c1: i.col, c2: i.col })).join(";");
}

/** Grubun veri aralığı metni: öğe aralıklarının kapsayanı (aynı sayfadaysa). */
export function groupDataText(wb: WorkbookModel, owner: SheetModel, g: SparklineGroup): string {
  if (g.items.length === 1) return g.items[0].range;
  let sheetName: string | null = null;
  let box: Rect | null = null;
  for (const it of g.items) {
    const p = parseRangeText(wb, owner, it.range);
    if (!p) return g.items[0].range;
    if (sheetName !== null && sheetName !== p.sheet.name) return g.items[0].range;
    sheetName = p.sheet.name;
    box = box ? { r1: Math.min(box.r1, p.rect.r1), c1: Math.min(box.c1, p.rect.c1), r2: Math.max(box.r2, p.rect.r2), c2: Math.max(box.c2, p.rect.c2) } : p.rect;
  }
  return `${quoteSheetName(sheetName!)}!${rectToA1(box!)}`;
}

/** Bir hücreden diğerine kopyalanan mini grafik: veri aralığı göreli kaydırılır (doldurma, yapıştırma). */
export function copySparkline(sheet: SheetModel, from: { row: number; col: number }, to: { row: number; col: number }, wb: WorkbookModel): void {
  const s = sparkAt(sheet, from.row, from.col);
  if (!s) return;
  const p = parseRangeText(wb, sheet, s.item.range);
  if (!p) return;
  const dr = to.row - from.row;
  const dc = to.col - from.col;
  const r = { r1: p.rect.r1 + dr, r2: p.rect.r2 + dr, c1: p.rect.c1 + dc, c2: p.rect.c2 + dc };
  if (r.r1 < 0 || r.c1 < 0) return;
  removeAt(sheet, [to]);
  s.group.items.push({ row: to.row, col: to.col, range: `${quoteSheetName(p.sheet.name)}!${rectToA1(r)}` });
}

/** Kes-yapıştır: aralıktaki mini grafikleri taşır; aralık içindeki veri kaynakları da kayar. */
export function moveSparklines(sheet: SheetModel, src: Rect, dr: number, dc: number, wb: WorkbookModel): void {
  const inside = (r: number, c: number) => contains(src, r, c);
  const moved: SparkItem[] = [];
  for (const g of sheet.sparklines)
    for (const it of g.items) {
      if (!inside(it.row, it.col)) continue;
      const p = parseRangeText(wb, sheet, it.range);
      if (p && p.sheet === sheet && inside(p.rect.r1, p.rect.c1) && inside(p.rect.r2, p.rect.c2)) it.range = `${quoteSheetName(sheet.name)}!${rectToA1({ r1: p.rect.r1 + dr, r2: p.rect.r2 + dr, c1: p.rect.c1 + dc, c2: p.rect.c2 + dc })}`;
      it.row += dr;
      it.col += dc;
      moved.push(it);
    }
  // Hedefte önceden olan mini grafikler silinir
  const keys = new Set(moved.map((m) => `${m.row},${m.col}`));
  for (const g of sheet.sparklines) g.items = g.items.filter((it) => moved.includes(it) || !keys.has(`${it.row},${it.col}`));
  sheet.sparklines = sheet.sparklines.filter((g) => g.items.length);
}
