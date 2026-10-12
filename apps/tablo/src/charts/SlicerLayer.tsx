/** Sayfadaki dilimleyiciler: başlıktan sürüklenir, köşeden boyutlandırılır; öğeler filtreyi değiştirir. */
import { useRef, useState } from "react";
import { MultiselectLtr16Regular, FilterDismiss16Regular } from "@fluentui/react-icons";
import { tipProps } from "@officehere/ui";
import type { Controller } from "../controller";
import type { GridLayout } from "../grid/layout";
import type { SlicerDef } from "../model/types";
import { anchorAt } from "./FloatingLayer";
import { applySlicer, clickSlicerItem, deleteSlicer, slicerItems, updateSlicer } from "../actions/slicers";

export function SlicerLayer({ ctl, layout }: { ctl: Controller; layout: GridLayout }) {
  const sheet = ctl.sheet;
  const [drag, setDrag] = useState<{ id: string; dx: number; dy: number; dw: number; dh: number } | null>(null);
  const start = useRef<{ x: number; y: number; id: string; resize: boolean } | null>(null);
  const [multi, setMulti] = useState<Record<string, boolean>>({});
  if (!sheet.slicers.length) return null;
  const z = layout.z;
  const fr0 = sheet.rows.offset(layout.fr);
  const fc0 = sheet.cols.offset(layout.fc);
  const down = (e: React.PointerEvent, s: SlicerDef, resize: boolean) => {
    if (e.button !== 0) return;
    e.stopPropagation();
    e.preventDefault();
    if (ctl.edit) ctl.commitEdit(null);
    ctl.selSlicer = s.id;
    ctl.selObj = null;
    ctl.emit();
    start.current = { x: e.clientX, y: e.clientY, id: s.id, resize };
    (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
    ((e.currentTarget as HTMLElement).closest(".tb-slicer") as HTMLElement | null)?.focus({ preventScroll: true });
    setDrag({ id: s.id, dx: 0, dy: 0, dw: 0, dh: 0 });
  };
  const move = (e: React.PointerEvent) => {
    const st = start.current;
    if (!st) return;
    const mx = (e.clientX - st.x) / z;
    const my = (e.clientY - st.y) / z;
    setDrag(st.resize ? { id: st.id, dx: 0, dy: 0, dw: mx, dh: my } : { id: st.id, dx: mx, dy: my, dw: 0, dh: 0 });
  };
  const up = () => {
    const st = start.current;
    start.current = null;
    const d = drag;
    setDrag(null);
    if (!st || !d || (!d.dx && !d.dy && !d.dw && !d.dh)) return;
    const s = sheet.slicers.find((x) => x.id === st.id);
    if (!s) return;
    if (st.resize) updateSlicer(ctl, s.id, { width: Math.max(90, s.width + d.dw), height: Math.max(70, s.height + d.dh) }, "Boyutlandır");
    else {
      const x = sheet.cols.offset(s.anchor.col) + s.anchor.dx + d.dx;
      const y = sheet.rows.offset(s.anchor.row) + s.anchor.dy + d.dy;
      updateSlicer(ctl, s.id, { anchor: anchorAt(sheet, x, y) }, "Taşı");
    }
  };
  return (
    <div className="tb-floats tb-slicers" style={{ left: layout.cellsX - layout.frozenW, top: layout.cellsY - layout.frozenH, width: Math.max(0, layout.width - layout.cellsX + layout.frozenW), height: Math.max(0, layout.height - layout.cellsY + layout.frozenH) }}>
      {sheet.slicers.map((s) => {
        const d = drag?.id === s.id ? drag : null;
        const px = sheet.cols.offset(s.anchor.col) + s.anchor.dx + (d?.dx ?? 0);
        const py = sheet.rows.offset(s.anchor.row) + s.anchor.dy + (d?.dy ?? 0);
        const x = px < fc0 ? px * z : layout.frozenW + (px - fc0) * z - layout.scrollX;
        const y = py < fr0 ? py * z : layout.frozenH + (py - fr0) * z - layout.scrollY;
        const w = Math.max(90, s.width + (d?.dw ?? 0)) * z;
        const h = Math.max(70, s.height + (d?.dh ?? 0)) * z;
        if (x > layout.width || y > layout.height || x + w < -50 || y + h < -50) return null;
        const items = slicerItems(ctl, s);
        const filtered = items.some((i) => !i.selected);
        const sel = ctl.selSlicer === s.id;
        const isMulti = !!multi[s.id];
        return (
          <div
            key={s.id}
            className={`tb-slicer${sel ? " is-selected" : ""}`}
            role="group"
            aria-label={`Dilimleyici: ${s.caption}`}
            tabIndex={-1}
            style={{ left: x, top: y, width: w, height: h, fontSize: 12 * z }}
            onPointerDown={(e) => {
              e.stopPropagation();
              if (ctl.selSlicer !== s.id) {
                ctl.selSlicer = s.id;
                ctl.selObj = null;
                ctl.emit();
              }
            }}
            onKeyDown={(e) => {
              if (e.key === "Delete" || e.key === "Backspace") {
                e.preventDefault();
                deleteSlicer(ctl, s.id);
                ctl.focusGrid();
              } else if (e.key === "Escape") {
                ctl.selSlicer = null;
                ctl.emit();
                ctl.focusGrid();
              }
            }}
          >
            <div className="tb-slicer-head" onPointerDown={(e) => down(e, s, false)} onPointerMove={move} onPointerUp={up}>
              <span className="tb-slicer-title">{s.caption}</span>
              <button
                type="button"
                className={`tb-slicer-btn${isMulti ? " is-on" : ""}`}
                aria-label="Çoklu Seçim"
                aria-pressed={isMulti}
                onPointerDown={(e) => e.stopPropagation()}
                onClick={() => setMulti({ ...multi, [s.id]: !isMulti })}
                {...tipProps("Çoklu Seçim (Alt+S)")}
              >
                <MultiselectLtr16Regular />
              </button>
              <button type="button" className="tb-slicer-btn" aria-label="Filtreyi Temizle" disabled={!filtered} onPointerDown={(e) => e.stopPropagation()} onClick={() => applySlicer(ctl, s, [])} {...tipProps("Filtreyi Temizle (Alt+C)")}>
                <FilterDismiss16Regular />
              </button>
            </div>
            <div className="tb-slicer-items" style={{ gridTemplateColumns: `repeat(${Math.max(1, s.columns)}, 1fr)` }}>
              {items.map((it) => (
                <button
                  key={it.key}
                  type="button"
                  className={`tb-slicer-item${it.selected ? " is-selected" : ""}`}
                  aria-pressed={it.selected}
                  onPointerDown={(e) => e.stopPropagation()}
                  onClick={(e) => clickSlicerItem(ctl, s, it.key, isMulti || e.ctrlKey || e.metaKey)}
                >
                  {it.label}
                </button>
              ))}
            </div>
            {sel && <span className="tb-handle is-se tb-slicer-resize" onPointerDown={(e) => down(e, s, true)} onPointerMove={move} onPointerUp={up} />}
          </div>
        );
      })}
    </div>
  );
}
