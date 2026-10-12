/**
 * Anahat simgeleri: satır başlıklarının solunda ve sütun başlıklarının
 * üstünde grup köşeli çizgileri, +/- düğmeleri ve düzey düğmeleri (1, 2, 3…).
 */
import { outlineGroups, type OutlineAxis, type OutlineGroup } from "../model/outline";
import type { SheetModel } from "../model/workbook";
import type { GridLayout } from "./layout";

export type OutlineHit = { axis: OutlineAxis; kind: "level"; level: number } | { axis: OutlineAxis; kind: "group"; group: OutlineGroup };

function step(L: GridLayout): number {
  return 14 * Math.max(0.75, Math.min(L.z, 1.5));
}

function button(ctx: CanvasRenderingContext2D, cx: number, cy: number, s: number, text: string, dark: boolean): void {
  ctx.fillStyle = dark ? "#3a3a3a" : "#FFFFFF";
  ctx.strokeStyle = dark ? "#9a9a9a" : "#8A8A8A";
  ctx.lineWidth = 1;
  ctx.fillRect(Math.round(cx - s / 2) + 0.5, Math.round(cy - s / 2) + 0.5, s, s);
  ctx.strokeRect(Math.round(cx - s / 2) + 0.5, Math.round(cy - s / 2) + 0.5, s, s);
  ctx.fillStyle = dark ? "#e6e6e6" : "#333333";
  if (text === "+" || text === "-") {
    const a = s * 0.28;
    ctx.fillRect(Math.round(cx - a), Math.round(cy), Math.round(2 * a) + 1, 1);
    if (text === "+") ctx.fillRect(Math.round(cx), Math.round(cy - a), 1, Math.round(2 * a) + 1);
  } else {
    ctx.font = `${Math.max(8, s * 0.72)}px "Segoe UI", system-ui, sans-serif`;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText(text, cx, cy + 0.5);
  }
}

/** Anahat şeritlerini çizer. */
export function drawOutline(ctx: CanvasRenderingContext2D, sheet: SheetModel, L: GridLayout, dark: boolean): void {
  if (!L.outlineW && !L.outlineH) return;
  const st = step(L);
  const bs = Math.round(st * 0.78);
  const bg = dark ? "#2b2b2b" : "#F3F3F3";
  const line = dark ? "#8a8a8a" : "#6E6E6E";
  ctx.save();
  if (L.outlineW) {
    ctx.fillStyle = bg;
    ctx.fillRect(0, 0, L.outlineW, L.height);
    const max = sheet.rows.maxLevel();
    // Düzey düğmeleri (sütun başlığı hizasında)
    const by = L.headerH - (L.headerH - L.outlineH) / 2;
    for (let lv = 1; lv <= max + 1; lv++) button(ctx, 2 + (lv - 0.5) * st, by, bs, String(lv), dark);
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, L.headerH, L.outlineW, L.height - L.headerH);
    ctx.clip();
    for (const g of outlineGroups(sheet.rows)) {
      const x = 2 + (g.level - 0.5) * st;
      const yS = g.summary < sheet.rows.count ? L.rowY(g.summary) + L.rowH(g.summary) / 2 : L.rowY(g.end) + L.rowH(g.end);
      if (!g.collapsed) {
        const y1 = L.rowY(g.start) + 2;
        ctx.fillStyle = line;
        ctx.fillRect(Math.round(x), y1, 1, Math.max(0, yS - y1 - bs / 2));
        ctx.fillRect(Math.round(x) - 3, y1, 4, 1);
      }
      if (yS > L.headerH - 20 && yS < L.height + 20) button(ctx, x, yS, bs, g.collapsed ? "+" : "-", dark);
    }
    // Ayrıntı satırlarındaki noktalar (en derin düzey)
    ctx.restore();
  }
  if (L.outlineH) {
    ctx.fillStyle = bg;
    ctx.fillRect(0, 0, L.width, L.outlineH);
    const max = sheet.cols.maxLevel();
    const bx = L.headerW - (L.headerW - L.outlineW) / 2;
    for (let lv = 1; lv <= max + 1; lv++) button(ctx, bx, 2 + (lv - 0.5) * st, bs, String(lv), dark);
    ctx.save();
    ctx.beginPath();
    ctx.rect(L.headerW, 0, L.width - L.headerW, L.outlineH);
    ctx.clip();
    for (const g of outlineGroups(sheet.cols)) {
      const y = 2 + (g.level - 0.5) * st;
      const xS = g.summary < sheet.cols.count ? L.colX(g.summary) + L.colW(g.summary) / 2 : L.colX(g.end) + L.colW(g.end);
      if (!g.collapsed) {
        const x1 = L.colX(g.start) + 2;
        ctx.fillStyle = line;
        ctx.fillRect(x1, Math.round(y), Math.max(0, xS - x1 - bs / 2), 1);
        ctx.fillRect(x1, Math.round(y) - 3, 1, 4);
      }
      if (xS > L.headerW - 20 && xS < L.width + 20) button(ctx, xS, y, bs, g.collapsed ? "+" : "-", dark);
    }
    ctx.restore();
  }
  ctx.restore();
}

/** Anahat şeritlerinde tıklanan düğme. */
export function hitTestOutline(L: GridLayout, sheet: SheetModel, x: number, y: number): OutlineHit | null {
  const st = step(L);
  const half = st * 0.5;
  if (L.outlineW && x < L.outlineW) {
    if (y < L.headerH && y >= L.outlineH) {
      const lv = Math.floor((x - 2) / st) + 1;
      return lv >= 1 && lv <= sheet.rows.maxLevel() + 1 ? { axis: "row", kind: "level", level: lv } : null;
    }
    for (const g of outlineGroups(sheet.rows)) {
      const cx = 2 + (g.level - 0.5) * st;
      const r = g.summary < sheet.rows.count ? g.summary : g.end;
      const cy = L.rowY(r) + L.rowH(r) / 2;
      if (Math.abs(x - cx) <= half && Math.abs(y - cy) <= half) return { axis: "row", kind: "group", group: g };
    }
    return null;
  }
  if (L.outlineH && y < L.outlineH) {
    if (x < L.headerW && x >= L.outlineW) {
      const lv = Math.floor((y - 2) / st) + 1;
      return lv >= 1 && lv <= sheet.cols.maxLevel() + 1 ? { axis: "col", kind: "level", level: lv } : null;
    }
    for (const g of outlineGroups(sheet.cols)) {
      const cy = 2 + (g.level - 0.5) * st;
      const c = g.summary < sheet.cols.count ? g.summary : g.end;
      const cx = L.colX(c) + L.colW(c) / 2;
      if (Math.abs(x - cx) <= half && Math.abs(y - cy) <= half) return { axis: "col", kind: "group", group: g };
    }
  }
  return null;
}

/** Nokta anahat şeritlerinden birinde mi (başlık/hücre tıklaması sayılmaz)? */
export function inOutlineGutter(L: GridLayout, x: number, y: number): boolean {
  return (L.outlineW > 0 && x < L.outlineW) || (L.outlineH > 0 && y < L.outlineH);
}
