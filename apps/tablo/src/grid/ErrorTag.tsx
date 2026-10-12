/** Hata göstergeli etkin hücrenin yanındaki uyarı düğmesi ve menüsü (Excel akıllı etiketi). */
import { useRef, useState } from "react";
import { Menu, MenuHeader, MenuItem, MenuProvider, MenuSeparator, Popover } from "@officehere/ui";
import type { Controller } from "../controller";
import { cellError, convertToNumber, ignoreError } from "../actions/audit";
import type { GridLayout } from "./layout";

export function ErrorTag({ ctl, layout }: { ctl: Controller; layout: GridLayout }) {
  const [open, setOpen] = useState(false);
  const btn = useRef<HTMLButtonElement>(null);
  const sheet = ctl.sheet;
  const { row, col } = ctl.sel.active;
  if (!ctl.errorChecking || ctl.edit || ctl.selObj || sheet.showFormulas) return null;
  const err = cellError(ctl.wb, sheet, row, col);
  if (!err) return null;
  const [r1, r2] = layout.visibleRows();
  const [c1, c2] = layout.visibleCols();
  const frozen = row < sheet.frozen.rows || col < sheet.frozen.cols;
  if (!frozen && (row < r1 || row > r2 || col < c1 || col > c2)) return null;
  const rc = layout.cellRect(row, col);
  const left = rc.x - 22 < layout.headerW ? rc.x + rc.w + 2 : rc.x - 22;
  return (
    <>
      <button ref={btn} type="button" className="tb-errtag" style={{ left, top: rc.y + Math.max(0, (rc.h - 18) / 2) }} aria-label="Hata denetimi seçenekleri" title={err.help} onMouseDown={(e) => e.preventDefault()} onClick={() => setOpen(!open)}>
        <svg width="14" height="14" viewBox="0 0 14 14" aria-hidden="true">
          <path d="M7 1.2L13 12.3H1z" fill="#FFC83D" stroke="#9C6A00" strokeWidth="0.8" />
          <path d="M7 5v3.6M7 10v.9" stroke="#3B2A00" strokeWidth="1.3" />
        </svg>
      </button>
      <Popover open={open} onClose={() => setOpen(false)} anchor={btn} ignore={[btn]}>
        <MenuProvider onClose={() => setOpen(false)}>
          <Menu width={260}>
            <MenuHeader>{err.title}</MenuHeader>
            {err.kind === "numberAsText" && <MenuItem label="Sayıya Dönüştür" onClick={() => convertToNumber(ctl)} />}
            {err.kind === "formulaError" && <MenuItem label="Hesaplama Adımlarını Göster..." onClick={() => ctl.openDialog("evaluateFormula")} />}
            {err.kind === "formulaError" && <MenuItem label="Hatayı İzle" onClick={() => ctl.openDialog("traceError")} />}
            <MenuItem label="Hatayı Yoksay" onClick={() => ignoreError(ctl, row, col)} />
            <MenuItem label="Formül Çubuğunda Düzenle" onClick={() => ctl.startEdit(undefined, true)} />
            <MenuSeparator />
            <MenuItem label="Hata Denetimi Seçenekleri..." onClick={() => ctl.openDialog("errorOptions")} />
          </Menu>
        </MenuProvider>
      </Popover>
    </>
  );
}
