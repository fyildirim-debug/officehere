/** İzleme Penceresi: seçili hücrelerin değer/formüllerini başka sayfadayken de gösteren yüzen pencere. */
import { useRef, useState } from "react";
import { baseName } from "@officehere/core";
import { addr, toLocalFormula } from "@officehere/formula";
import { DialogButton } from "@officehere/ui";
import type { Controller } from "../controller";
import { addWatches } from "../actions/audit";
import { displayText } from "../model/display";

export function WatchWindow({ ctl }: { ctl: Controller }) {
  const [pos, setPos] = useState({ x: 0, y: 0 });
  const [sel, setSel] = useState<number | null>(null);
  const drag = useRef<{ sx: number; sy: number } | null>(null);
  const names = ctl.wb.engine.getNames();
  const rows = ctl.watches
    .map((w, i) => {
      const sh = ctl.wb.sheetById(w.sheetId);
      if (!sh) return null;
      const f = ctl.wb.getFormula(sh, w.row, w.col);
      const ref = `${sh.name}!$${addr(0, w.col).replace(/\d+$/, "")}$${w.row + 1}`;
      const name = names.find((n) => n.formula.replace(/'/g, "").toUpperCase() === `=${ref}`.toUpperCase())?.name ?? "";
      return { i, w, sheet: sh, name, value: displayText(ctl.wb, sh, w.row, w.col), formula: f ? toLocalFormula(f, "tr") : "" };
    })
    .filter((x): x is NonNullable<typeof x> => !!x);
  const close = () => {
    ctl.showWatch = false;
    ctl.emit();
    ctl.focusGrid();
  };
  return (
    <div className="tb-watch" style={{ transform: `translate(${pos.x}px, ${pos.y}px)` }} role="dialog" aria-label="İzleme Penceresi">
      <div
        className="tb-watch-title"
        onPointerDown={(e) => {
          if ((e.target as HTMLElement).closest("button")) return;
          drag.current = { sx: e.clientX - pos.x, sy: e.clientY - pos.y };
          (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
        }}
        onPointerMove={(e) => drag.current && setPos({ x: e.clientX - drag.current.sx, y: e.clientY - drag.current.sy })}
        onPointerUp={() => (drag.current = null)}
      >
        <span>İzleme Penceresi</span>
        <button type="button" className="tb-pane-close" aria-label="Kapat" onClick={close}>
          <svg width="10" height="10" viewBox="0 0 10 10" aria-hidden="true">
            <path d="M1 1l8 8M9 1L1 9" stroke="currentColor" strokeWidth="1.2" />
          </svg>
        </button>
      </div>
      <div className="tb-watch-bar">
        <DialogButton onClick={() => addWatches(ctl)}>İzleme Ekle...</DialogButton>
        <DialogButton
          disabled={sel === null}
          onClick={() => {
            if (sel === null) return;
            ctl.watches.splice(sel, 1);
            setSel(null);
            ctl.emit();
          }}
        >
          İzleme Sil
        </DialogButton>
      </div>
      <div className="tb-watch-table" role="table">
        <div className="tb-watch-row is-head" role="row">
          <span>Kitap</span>
          <span>Sayfa</span>
          <span>Ad</span>
          <span>Hücre</span>
          <span>Değer</span>
          <span>Formül</span>
        </div>
        {rows.map((r) => (
          <div
            key={`${r.w.sheetId}:${r.w.row}:${r.w.col}`}
            role="row"
            className={`tb-watch-row${sel === r.i ? " is-active" : ""}`}
            onClick={() => setSel(r.i)}
            onDoubleClick={() => {
              const idx = ctl.wb.sheets.indexOf(r.sheet);
              if (idx !== ctl.wb.activeIndex) ctl.activateSheet(idx);
              ctl.selectCell(r.w.row, r.w.col);
              ctl.focusGrid();
            }}
          >
            <span>{baseName(ctl.fileName)}</span>
            <span>{r.sheet.name}</span>
            <span>{r.name}</span>
            <span>{addr(r.w.row, r.w.col)}</span>
            <span>{r.value}</span>
            <span>{r.formula}</span>
          </div>
        ))}
      </div>
    </div>
  );
}
