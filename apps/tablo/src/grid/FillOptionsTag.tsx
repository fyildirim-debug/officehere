/**
 * Otomatik Doldurma Seçenekleri: doldurma tutamacıyla sürüklemeden sonra hedefin sağ altında
 * beliren düğme. Seçilen seçenek son doldurmayı geri alıp aynı aralığı o kiple yeniden doldurur
 * (Hücreleri Kopyala, Seri Doldur, yalnız biçim, biçimsiz, gün/hafta içi/ay/yıl, Akıllı Doldurma).
 */
import { Fragment, useRef, useState } from "react";
import { Menu, MenuItem, MenuProvider, MenuSeparator, Popover } from "@officehere/ui";
import { isDateFormat } from "@officehere/formula";
import type { Controller } from "../controller";
import type { FillMode } from "../model/ops";
import type { Rect } from "../model/types";
import { flashFill } from "../actions/whatIf";
import type { GridLayout } from "./layout";

type Option = { mode: FillMode | "flash"; label: string };

/** Kaynağın içeriğine göre gösterilecek seçenekler ve doldurmanın şu anki karşılığı. */
export function fillOptionsFor(ctl: Controller, src: Rect, mode: FillMode): { options: Option[]; current: FillMode } {
  const { wb, sheet } = ctl;
  let hasDate = false;
  let hasText = false;
  let cells = 0;
  let plainNumbers = 0;
  for (let r = src.r1; r <= src.r2; r++)
    for (let c = src.c1; c <= src.c2; c++) {
      cells++;
      const v = wb.getInput(sheet, r, c);
      const fmt = wb.styleAt(sheet, r, c).numFmt;
      if (typeof v === "number") {
        if (fmt && isDateFormat(fmt)) hasDate = true;
        else plainNumbers++;
      } else if (typeof v === "string") hasText = true;
    }
  const singleNumber = cells === 1 && plainNumbers === 1;
  const options: Option[] = [
    { mode: "copy", label: "Hücreleri Kopyala" },
    { mode: "seriesStep", label: "Seri Doldur" },
    { mode: "formats", label: "Yalnızca Biçimlendirmeyi Doldur" },
    { mode: "noFormats", label: "Biçimlendirmeden Doldur" },
  ];
  if (hasDate)
    options.push({ mode: "days", label: "Günleri Doldur" }, { mode: "weekdays", label: "Hafta İçi Günleri Doldur" }, { mode: "months", label: "Ayları Doldur" }, { mode: "years", label: "Yılları Doldur" });
  if (hasText && src.c1 === src.c2) options.push({ mode: "flash", label: "Akıllı Doldurma" });
  // Sürüklemenin yaptığı: tek sayıda varsayılan kopya, Ctrl ile seri; diğerlerinde tersi
  const current: FillMode = mode === "series" ? (singleNumber ? "copy" : "seriesStep") : mode === "alt" ? (singleNumber ? "seriesStep" : "copy") : mode;
  return { options, current };
}

/** Son doldurmayı seçilen kiple yineler (tek geri alma adımı olarak). */
export function refill(ctl: Controller, mode: FillMode | "flash"): void {
  const fo = ctl.fillOptions;
  if (!fo || ctl.ops.history.undoCount !== fo.undo) return;
  const sheet = ctl.wb.sheetById(fo.sheetId);
  if (!sheet) return;
  ctl.ops.undo();
  if (mode === "flash") {
    // Akıllı Doldurma örnekleri kaynaktan öğrenir; hedefin boş hücrelerini doldurur
    ctl.selectRange(fo.target);
    flashFill(ctl);
    ctl.fillOptions = null;
    ctl.emit();
    return;
  }
  ctl.ops.fill(sheet, fo.src, fo.target, mode);
  ctl.fillOptions = { ...fo, mode, undo: ctl.ops.history.undoCount };
  ctl.selectRange(fo.target, ctl.sel.active);
}

export function FillOptionsTag({ ctl, layout }: { ctl: Controller; layout: GridLayout }) {
  const [open, setOpen] = useState(false);
  const btn = useRef<HTMLButtonElement>(null);
  const fo = ctl.fillOptions;
  if (!fo || fo.sheetId !== ctl.sheet.id || ctl.edit || ctl.selObj || ctl.ops.history.undoCount !== fo.undo) return null;
  const t = fo.target;
  const [r1, r2] = layout.visibleRows();
  const [c1, c2] = layout.visibleCols();
  if (t.r2 < r1 || t.r2 > r2 || t.c2 < c1 || t.c2 > c2) return null;
  const rc = layout.cellRect(t.r2, t.c2);
  const { options, current } = fillOptionsFor(ctl, fo.src, fo.mode);
  const pick = (m: FillMode | "flash") => {
    setOpen(false);
    refill(ctl, m);
    ctl.focusGrid();
  };
  return (
    <>
      <button
        ref={btn}
        type="button"
        className="tb-filltag"
        style={{ left: rc.x + rc.w + 3, top: rc.y + rc.h + 3 }}
        aria-label="Otomatik Doldurma Seçenekleri"
        title="Otomatik Doldurma Seçenekleri"
        aria-expanded={open}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => setOpen(!open)}
      >
        <svg width="16" height="14" viewBox="0 0 16 14" aria-hidden="true">
          <rect x="1.5" y="1.5" width="8" height="9" rx="1" fill="#fff" stroke="currentColor" />
          <path d="M3.5 4h4M3.5 6h4M3.5 8h2.5" stroke="currentColor" />
          <path d="M11 6l2 2.5 2-2.5" fill="none" stroke="currentColor" />
        </svg>
      </button>
      <Popover open={open} onClose={() => setOpen(false)} anchor={btn} ignore={[btn]}>
        <MenuProvider onClose={() => setOpen(false)}>
          <Menu width={250}>
            {options.map((o, i) => (
              <Fragment key={o.mode}>
                {(i === 4 || o.mode === "flash") && <MenuSeparator />}
                <MenuItem label={o.label} checked={o.mode === current} onClick={() => pick(o.mode)} />
              </Fragment>
            ))}
          </Menu>
        </MenuProvider>
      </Popover>
    </>
  );
}
