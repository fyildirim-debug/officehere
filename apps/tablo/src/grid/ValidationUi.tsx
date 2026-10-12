/** Veri doğrulama arayüzü: giriş iletisi ve hücre içi açılan liste. */
import { useEffect, useRef, useState } from "react";
import type { Controller } from "../controller";
import { validationList } from "../controller";
import { contains } from "../model/rect";
import type { GridLayout } from "./layout";

export function ValidationUi({ ctl, layout }: { ctl: Controller; layout: GridLayout }) {
  const sheet = ctl.sheet;
  const { row, col } = ctl.sel.active;
  const v = sheet.validations.find((x) => x.ranges.some((r) => contains(r, row, col)));
  const [sel, setSel] = useState(0);
  const listRef = useRef<HTMLDivElement>(null);
  const open = ctl.listDropdown;
  useEffect(() => {
    if (open) {
      setSel(0);
      requestAnimationFrame(() => listRef.current?.focus());
    }
  }, [open]);
  useEffect(() => {
    // Seçim değişince açılır liste kapanır
    if (ctl.listDropdown) ctl.listDropdown = false;
  }, [row, col, sheet.id]); // eslint-disable-line react-hooks/exhaustive-deps
  if (!v || ctl.edit || ctl.selObj) return null;
  const [r1, r2] = layout.visibleRows();
  const [c1, c2] = layout.visibleCols();
  const frozen = row < sheet.frozen.rows || col < sheet.frozen.cols;
  if (!frozen && (row < r1 || row > r2 || col < c1 || col > c2)) return null;
  const m = sheet.mergeAt(row, col);
  const rc = m ? layout.cellRect(m.r1, m.c1, m.r2, m.c2) : layout.cellRect(row, col);
  const rule = v.rule;
  const isList = rule.type === "list" && rule.showDropdown !== false;
  const items = isList && open ? validationList(ctl.wb, sheet, rule.formula1 ?? "", row, col) : [];
  const close = () => {
    ctl.listDropdown = false;
    ctl.emit();
    ctl.focusGrid();
  };
  const pick = (s: string) => {
    ctl.listDropdown = false;
    ctl.ops.enterText(sheet, [{ row, col }], s, "Listeden Seç");
    ctl.focusGrid();
  };
  return (
    <>
      {(rule.prompt || rule.promptTitle) && !open && (
        <div className="tb-tip tb-dv-prompt" style={{ left: rc.x + Math.min(rc.w, 40), top: rc.y + rc.h + 6 }}>
          {rule.promptTitle && <strong>{rule.promptTitle}</strong>}
          {(rule.prompt ?? "").split("\n").map((l, i) => (
            <div key={i}>{l}</div>
          ))}
        </div>
      )}
      {isList && (
        <button
          type="button"
          className="tb-dv-arrow"
          style={{ left: rc.x + rc.w + 1, top: rc.y + rc.h - 18 }}
          aria-label="Açılan liste"
          onMouseDown={(e) => e.preventDefault()}
          onPointerDown={(e) => e.stopPropagation()}
          onClick={() => {
            ctl.listDropdown = !ctl.listDropdown;
            ctl.emit();
          }}
        >
          <svg width="8" height="8" viewBox="0 0 8 8" aria-hidden="true">
            <path d="M1 2.5l3 3 3-3" fill="none" stroke="currentColor" strokeWidth="1.2" />
          </svg>
        </button>
      )}
      {isList && open && (
        <div
          ref={listRef}
          className="tb-dv-list"
          role="listbox"
          tabIndex={0}
          style={{ left: rc.x, top: rc.y + rc.h + 1, minWidth: rc.w + 18 }}
          onPointerDown={(e) => e.stopPropagation()}
          onKeyDown={(e) => {
            e.stopPropagation();
            if (e.key === "ArrowDown") setSel((s) => Math.min(items.length - 1, s + 1));
            else if (e.key === "ArrowUp") setSel((s) => Math.max(0, s - 1));
            else if (e.key === "Enter" && items[sel] !== undefined) pick(items[sel]);
            else if (e.key === "Escape") close();
            else return;
            e.preventDefault();
          }}
          onBlur={() => ctl.listDropdown && close()}
        >
          {items.length === 0 && <div className="tb-dv-item is-empty">(Liste boş)</div>}
          {items.map((s, i) => (
            <div key={s + i} role="option" aria-selected={i === sel} className={`tb-dv-item${i === sel ? " is-active" : ""}`} onMouseEnter={() => setSel(i)} onMouseDown={(e) => (e.preventDefault(), pick(s))}>
              {s}
            </div>
          ))}
        </div>
      )}
    </>
  );
}
