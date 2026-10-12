/** Hücrelerin üzerinde yüzen nesneler: grafikler, resimler, şekiller (taşıma ve boyutlandırma). */
import { useMemo, useRef, useState } from "react";
import type { Controller } from "../controller";
import type { GridLayout } from "../grid/layout";
import type { Anchor, ChartDef, ImageDef, ShapeDef } from "../model/types";
import type { SheetModel } from "../model/workbook";
import { ChartView } from "./ChartView";
import { extractChartData } from "./chartData";
import { handleObjectKey } from "../actions/objectClip";

type Obj = { kind: "chart"; o: ChartDef } | { kind: "image"; o: ImageDef } | { kind: "shape"; o: ShapeDef };

function absPos(sheet: SheetModel, a: Anchor): { x: number; y: number } {
  return { x: sheet.cols.offset(a.col) + a.dx, y: sheet.rows.offset(a.row) + a.dy };
}

export function anchorAt(sheet: SheetModel, x: number, y: number): Anchor {
  const col = sheet.cols.indexAt(Math.max(0, x));
  const row = sheet.rows.indexAt(Math.max(0, y));
  return { col, row, dx: Math.max(0, x - sheet.cols.offset(col)), dy: Math.max(0, y - sheet.rows.offset(row)) };
}

const HANDLES = ["nw", "n", "ne", "e", "se", "s", "sw", "w"] as const;

export function FloatingLayer({ ctl, layout }: { ctl: Controller; layout: GridLayout | null }) {
  const sheet = ctl.sheet;
  const [drag, setDrag] = useState<{ id: string; dx: number; dy: number; dw: number; dh: number } | null>(null);
  const start = useRef<{ x: number; y: number; handle: string | null; id: string; kind: Obj["kind"] } | null>(null);
  const version = ctl.wb.engine.version;
  // Seçim Bölmesinde gizlenen nesneler çizilmez
  const objs: Obj[] = [...sheet.images.map((o) => ({ kind: "image" as const, o })), ...sheet.shapes.map((o) => ({ kind: "shape" as const, o })), ...sheet.charts.map((o) => ({ kind: "chart" as const, o }))].filter((x) => !x.o.hidden);
  const chartData = useMemo(() => new Map(sheet.charts.map((c) => [c.id, extractChartData(ctl.wb, c)])), [sheet.charts, version, ctl.version]); // eslint-disable-line react-hooks/exhaustive-deps
  if (!layout || !objs.length) return null;
  const z = layout.z;
  const fr0 = sheet.rows.offset(layout.fr);
  const fc0 = sheet.cols.offset(layout.fc);

  const onDown = (e: React.PointerEvent, obj: Obj, handle: string | null) => {
    if (e.button !== 0) return;
    e.stopPropagation();
    e.preventDefault();
    if (ctl.edit) ctl.commitEdit(null);
    ctl.selObj = { kind: obj.kind, id: obj.o.id };
    ctl.emit();
    start.current = { x: e.clientX, y: e.clientY, handle, id: obj.o.id, kind: obj.kind };
    (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
    setDrag({ id: obj.o.id, dx: 0, dy: 0, dw: 0, dh: 0 });
  };
  const onMove = (e: React.PointerEvent) => {
    const s = start.current;
    if (!s) return;
    const mx = (e.clientX - s.x) / z;
    const my = (e.clientY - s.y) / z;
    const h = s.handle;
    if (!h) setDrag({ id: s.id, dx: mx, dy: my, dw: 0, dh: 0 });
    else {
      let dx = 0;
      let dy = 0;
      let dw = 0;
      let dh = 0;
      if (h.includes("e")) dw = mx;
      if (h.includes("s")) dh = my;
      if (h.includes("w")) {
        dx = mx;
        dw = -mx;
      }
      if (h.includes("n")) {
        dy = my;
        dh = -my;
      }
      if (e.shiftKey && (h.length === 2)) {
        // Oranı koru
        const obj = objs.find((o) => o.o.id === s.id)!;
        const ratio = obj.o.height / obj.o.width;
        dh = dw * ratio;
        if (h.includes("n")) dy = -dh;
      }
      setDrag({ id: s.id, dx, dy, dw, dh });
    }
  };
  const onUp = () => {
    const s = start.current;
    start.current = null;
    if (!s || !drag || (drag.dx === 0 && drag.dy === 0 && drag.dw === 0 && drag.dh === 0)) {
      setDrag(null);
      return;
    }
    const d = drag;
    ctl.ops.sheetOp(sheet, s.handle ? "Boyutlandır" : "Taşı", () => {
      const list = s.kind === "chart" ? sheet.charts : s.kind === "image" ? sheet.images : sheet.shapes;
      const o = (list as { id: string; anchor: Anchor; width: number; height: number }[]).find((x) => x.id === s.id);
      if (!o) return;
      const p = absPos(sheet, o.anchor);
      o.anchor = anchorAt(sheet, p.x + d.dx, p.y + d.dy);
      o.width = Math.max(20, o.width + d.dw);
      o.height = Math.max(20, o.height + d.dh);
    });
    setDrag(null);
  };

  return (
    <div className="tb-floats" style={{ left: layout.cellsX - layout.frozenW, top: layout.cellsY - layout.frozenH, width: Math.max(0, layout.width - layout.cellsX + layout.frozenW), height: Math.max(0, layout.height - layout.cellsY + layout.frozenH) }}>
      {objs.map((obj) => {
        const o = obj.o;
        const p = absPos(sheet, o.anchor);
        const d = drag && drag.id === o.id ? drag : null;
        // Dondurulmuş bölmede başlayan nesne o eksende kaymaz; kayan bölmedekiler donmuş alanın altında kırpılır
        const px = p.x + (d?.dx ?? 0);
        const py = p.y + (d?.dy ?? 0);
        const fixedX = px < fc0;
        const fixedY = py < fr0;
        const x = fixedX ? px * z : layout.frozenW + (px - fc0) * z - layout.scrollX;
        const y = fixedY ? py * z : layout.frozenH + (py - fr0) * z - layout.scrollY;
        const clipL = fixedX ? 0 : Math.max(0, layout.frozenW - x);
        const clipT = fixedY ? 0 : Math.max(0, layout.frozenH - y);
        const w = Math.max(20, o.width + (d?.dw ?? 0)) * z;
        const h = Math.max(20, o.height + (d?.dh ?? 0)) * z;
        if (x > layout.width || y > layout.height || x + w < -50 || y + h < -50) return null;
        const selected = ctl.selObj?.id === o.id;
        return (
          <div
            key={o.id}
            className={`tb-float${selected ? " is-selected" : ""} is-${obj.kind}`}
            style={{ left: x, top: y, width: w, height: h, clipPath: clipL || clipT ? `inset(${clipT}px 0 0 ${clipL}px)` : undefined }}
            tabIndex={-1}
            onPointerDown={(e) => onDown(e, obj, null)}
            onPointerMove={onMove}
            onPointerUp={onUp}
            onDoubleClick={() => {
              // Çift tık: grafikte biçimlendirme bölmesi, şekilde metin düzenleme
              if (obj.kind === "chart") ctl.openDialog("chartFormat");
              else if (obj.kind === "shape") ctl.openDialog("shapeText", { id: o.id });
            }}
            onKeyDown={(e) => {
              if (e.key === "Escape") {
                ctl.selObj = null;
                ctl.emit();
                ctl.focusGrid();
                return;
              }
              if (handleObjectKey(ctl, e)) e.preventDefault();
            }}
            onContextMenu={(e) => {
              e.preventDefault();
              e.stopPropagation();
              ctl.selObj = { kind: obj.kind, id: o.id };
              ctl.openDialog("objectMenu", { x: e.clientX, y: e.clientY, kind: obj.kind, id: o.id });
            }}
            ref={(el) => {
              if (el && selected && document.activeElement?.closest(".tb-float") !== el && !ctl.edit && !document.activeElement?.closest(".oh-ribbon, .oh-popover, .oh-dialog, input, textarea, select")) el.focus({ preventScroll: true });
            }}
          >
            {obj.kind === "chart" && (
              <div style={{ transform: `scale(${z})`, transformOrigin: "0 0", width: w / z, height: h / z }}>
                <ChartView def={obj.o} data={chartData.get(o.id) ?? { categories: [], series: [] }} width={Math.max(20, o.width + (d?.dw ?? 0))} height={Math.max(20, o.height + (d?.dh ?? 0))} />
              </div>
            )}
            {obj.kind === "image" && <img src={obj.o.src} alt={obj.o.name} draggable={false} />}
            {obj.kind === "shape" && <ShapeSvg s={obj.o} w={w} h={h} />}
            {selected &&
              HANDLES.map((hd) => (
                <span key={hd} className={`tb-handle is-${hd}`} onPointerDown={(e) => onDown(e, obj, hd)} onPointerMove={onMove} onPointerUp={onUp} />
              ))}
          </div>
        );
      })}
    </div>
  );
}

function ShapeSvg({ s, w, h }: { s: ShapeDef; w: number; h: number }) {
  const sw = 1.5;
  let shape: React.ReactNode;
  switch (s.kind) {
    case "ellipse":
      shape = <ellipse cx={w / 2} cy={h / 2} rx={w / 2 - sw} ry={h / 2 - sw} />;
      break;
    case "roundRect":
      shape = <rect x={sw} y={sw} width={w - 2 * sw} height={h - 2 * sw} rx={Math.min(w, h) * 0.15} />;
      break;
    case "triangle":
      shape = <path d={`M${w / 2},${sw}L${w - sw},${h - sw}L${sw},${h - sw}Z`} />;
      break;
    case "arrow":
      shape = <path d={`M${sw},${h * 0.3}H${w * 0.65}V${sw}L${w - sw},${h / 2}L${w * 0.65},${h - sw}V${h * 0.7}H${sw}Z`} />;
      break;
    case "line":
      shape = <line x1={sw} y1={h - sw} x2={w - sw} y2={sw} />;
      break;
    default:
      shape = <rect x={sw} y={sw} width={w - 2 * sw} height={h - 2 * sw} />;
  }
  return (
    <div className="tb-shape">
      <svg width={w} height={h} fill={s.kind === "line" ? "none" : s.fill} stroke={s.stroke} strokeWidth={sw}>
        {shape}
      </svg>
      {s.text && <div className="tb-shape-text">{s.text}</div>}
    </div>
  );
}

export function deleteObject(ctl: Controller, kind: "chart" | "image" | "shape", id: string): void {
  const sheet = ctl.sheet;
  ctl.ops.sheetOp(sheet, "Sil", () => {
    if (kind === "chart") sheet.charts = sheet.charts.filter((c) => c.id !== id);
    else if (kind === "image") sheet.images = sheet.images.filter((c) => c.id !== id);
    else sheet.shapes = sheet.shapes.filter((c) => c.id !== id);
  });
  ctl.selObj = null;
  ctl.focusGrid();
  ctl.emit();
}
