/** Yazdırma/PDF/HTML için yüzen nesnelerin (grafik, resim, şekil) durağan işaretlemesi. */
import { renderToStaticMarkup } from "react-dom/server";
import { ChartView } from "../charts/ChartView";
import { extractChartData } from "../charts/chartData";
import type { Anchor, Rect, ShapeDef } from "../model/types";
import type { SheetModel, WorkbookModel } from "../model/workbook";

function esc(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function pos(sheet: SheetModel, a: Anchor): { x: number; y: number } {
  return { x: sheet.cols.offset(a.col) + a.dx, y: sheet.rows.offset(a.row) + a.dy };
}

/** Nesnelerin kapladığı en sağ-alt hücre (yazdırma alanını genişletmek için). */
export function objectsExtent(sheet: SheetModel): { r2: number; c2: number } | null {
  let r2 = -1;
  let c2 = -1;
  for (const o of [...sheet.charts, ...sheet.images, ...sheet.shapes]) {
    const p = pos(sheet, o.anchor);
    r2 = Math.max(r2, sheet.rows.indexAt(p.y + o.height - 1));
    c2 = Math.max(c2, sheet.cols.indexAt(p.x + o.width - 1));
  }
  return r2 < 0 ? null : { r2, c2 };
}

function shapeSvg(s: ShapeDef): string {
  const w = s.width;
  const h = s.height;
  const sw = 1.5;
  let el: string;
  switch (s.kind) {
    case "ellipse":
      el = `<ellipse cx="${w / 2}" cy="${h / 2}" rx="${w / 2 - sw}" ry="${h / 2 - sw}"/>`;
      break;
    case "roundRect":
      el = `<rect x="${sw}" y="${sw}" width="${w - 2 * sw}" height="${h - 2 * sw}" rx="${Math.min(w, h) * 0.15}"/>`;
      break;
    case "triangle":
      el = `<path d="M${w / 2},${sw}L${w - sw},${h - sw}L${sw},${h - sw}Z"/>`;
      break;
    case "arrow":
      el = `<path d="M${sw},${h * 0.3}H${w * 0.65}V${sw}L${w - sw},${h / 2}L${w * 0.65},${h - sw}V${h * 0.7}H${sw}Z"/>`;
      break;
    case "line":
      el = `<line x1="${sw}" y1="${h - sw}" x2="${w - sw}" y2="${sw}"/>`;
      break;
    default:
      el = `<rect x="${sw}" y="${sw}" width="${w - 2 * sw}" height="${h - 2 * sw}"/>`;
  }
  const text = s.text ? `<div class="tb-print-shape-text">${esc(s.text).replace(/\n/g, "<br>")}</div>` : "";
  return `<svg width="${w}" height="${h}" fill="${s.kind === "line" ? "none" : esc(s.fill)}" stroke="${esc(s.stroke)}" stroke-width="${sw}">${el}</svg>${text}`;
}

/** Aralığın sol üstüne göre konumlanmış nesne katmanı. headings: satır/sütun başlıkları basılıyor mu. */
export function objectsHtml(wb: WorkbookModel, sheet: SheetModel, r: Rect, headings: boolean): string {
  const ox = sheet.cols.offset(r.c1) - (headings ? 32 : 0);
  const oy = sheet.rows.offset(r.r1) - (headings ? 20 : 0);
  const box = (a: Anchor, w: number, h: number, inner: string, cls: string) => {
    const p = pos(sheet, a);
    return `<div class="tb-print-obj ${cls}" style="left:${p.x - ox}px;top:${p.y - oy}px;width:${w}px;height:${h}px">${inner}</div>`;
  };
  const parts: string[] = [];
  for (const img of sheet.images) parts.push(box(img.anchor, img.width, img.height, `<img src="${esc(img.src)}" alt="${esc(img.name)}" style="width:100%;height:100%">`, "is-image"));
  for (const s of sheet.shapes) parts.push(box(s.anchor, s.width, s.height, shapeSvg(s), "is-shape"));
  for (const c of sheet.charts) {
    const svg = renderToStaticMarkup(<ChartView def={c} data={extractChartData(wb, c)} width={c.width} height={c.height} />);
    parts.push(box(c.anchor, c.width, c.height, svg, "is-chart"));
  }
  return parts.join("");
}
