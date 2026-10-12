/** Tablonun Toplam Satırı hücresinde işlev açılır listesi (Yok, Ortalama, Say, …, Diğer İşlevler). */
import { useEffect, useRef, useState } from "react";
import type { Controller } from "../controller";
import type { TableDef } from "../model/types";
import type { GridLayout } from "./layout";

/** ALTTOPLAM işlev numaraları (gizli satırları yok sayan 10x biçimi). */
export const TOTAL_FUNCTIONS: { label: string; code: number | null }[] = [
  { label: "Yok", code: null },
  { label: "Ortalama", code: 101 },
  { label: "Say", code: 103 },
  { label: "Sayıları Say", code: 102 },
  { label: "En Büyük", code: 104 },
  { label: "En Küçük", code: 105 },
  { label: "Topla", code: 109 },
  { label: "StdSapma", code: 107 },
  { label: "Var", code: 110 },
];

/** Tablo sütununa yapısal başvuru: Tablo1[Tutar] ya da Tablo1[[Birim Fiyat]]. */
export function columnRef(t: TableDef, col: number): string {
  const name = t.columns[col - t.ref.c1] ?? "";
  const simple = /^[\p{L}\p{N}_.][\p{L}\p{N}_. ]*$/u.test(name) && !name.endsWith(" ");
  return `${t.name}[${simple ? name : `[${name.replace(/(['[\]#])/g, "'$1")}]`}]`;
}

/** Toplam satırı hücresine işlev yazar (null: hücreyi temizler). */
export function setTotalFunction(ctl: Controller, t: TableDef, col: number, code: number | null): void {
  const row = t.ref.r2;
  if (code === null) ctl.ops.setInputs(ctl.sheet, [{ row, col, input: null }], "Toplam Satırı");
  else ctl.ops.setInputs(ctl.sheet, [{ row, col, input: { formula: `=SUBTOTAL(${code},${columnRef(t, col)})` } }], "Toplam Satırı");
}

export function TotalRowUi({ ctl, layout }: { ctl: Controller; layout: GridLayout }) {
  const sheet = ctl.sheet;
  const { row, col } = ctl.sel.active;
  const t = sheet.tableAt(row, col);
  const [open, setOpen] = useState(false);
  const [sel, setSel] = useState(0);
  const listRef = useRef<HTMLDivElement>(null);
  useEffect(() => setOpen(false), [row, col, sheet.id]);
  useEffect(() => {
    if (open) requestAnimationFrame(() => listRef.current?.focus());
  }, [open]);
  if (!t || !t.totalRow || row !== t.ref.r2 || ctl.edit || ctl.selObj) return null;
  const [r1, r2] = layout.visibleRows();
  const [c1, c2] = layout.visibleCols();
  const frozen = row < sheet.frozen.rows || col < sheet.frozen.cols;
  if (!frozen && (row < r1 || row > r2 || col < c1 || col > c2)) return null;
  const rc = layout.cellRect(row, col);
  const f = ctl.wb.getFormula(sheet, row, col) ?? "";
  const cur = /^=SUBTOTAL\((\d+),/i.exec(f)?.[1];
  const items = [...TOTAL_FUNCTIONS.map((x) => x.label), "Diğer İşlevler..."];
  const pick = (i: number) => {
    setOpen(false);
    if (i < TOTAL_FUNCTIONS.length) setTotalFunction(ctl, t, col, TOTAL_FUNCTIONS[i].code);
    else {
      ctl.startEdit("=");
      ctl.openDialog("insertFunction");
      return;
    }
    ctl.focusGrid();
  };
  return (
    <>
      <button
        type="button"
        className="tb-dv-arrow"
        style={{ left: rc.x + rc.w + 1, top: rc.y + rc.h - 18 }}
        aria-label="Toplam işlevi"
        onMouseDown={(e) => e.preventDefault()}
        onPointerDown={(e) => e.stopPropagation()}
        onClick={() => {
          const i = TOTAL_FUNCTIONS.findIndex((x) => String(x.code) === cur);
          setSel(i < 0 ? 0 : i);
          setOpen(!open);
        }}
      >
        <svg width="8" height="8" viewBox="0 0 8 8" aria-hidden="true">
          <path d="M1 2.5l3 3 3-3" fill="none" stroke="currentColor" strokeWidth="1.2" />
        </svg>
      </button>
      {open && (
        <div
          ref={listRef}
          className="tb-dv-list"
          role="listbox"
          aria-label="Toplam işlevi"
          tabIndex={0}
          style={{ left: rc.x, top: rc.y + rc.h + 1, minWidth: Math.max(rc.w + 18, 140) }}
          onPointerDown={(e) => e.stopPropagation()}
          onKeyDown={(e) => {
            e.stopPropagation();
            if (e.key === "ArrowDown") setSel((s) => Math.min(items.length - 1, s + 1));
            else if (e.key === "ArrowUp") setSel((s) => Math.max(0, s - 1));
            else if (e.key === "Enter") pick(sel);
            else if (e.key === "Escape") (setOpen(false), ctl.focusGrid());
            else return;
            e.preventDefault();
          }}
          onBlur={() => setOpen(false)}
        >
          {items.map((s, i) => (
            <div key={s} role="option" aria-selected={i === sel} className={`tb-dv-item${i === sel ? " is-active" : ""}`} onMouseEnter={() => setSel(i)} onMouseDown={(e) => (e.preventDefault(), pick(i))}>
              {s}
            </div>
          ))}
        </div>
      )}
    </>
  );
}
