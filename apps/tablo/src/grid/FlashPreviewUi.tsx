/** Akıllı Doldurma önizlemesi: önerilen değerler hücrelerde gri yazıyla gösterilir (Enter kabul eder). */
import type { Controller } from "../controller";
import { flashPreviewFor } from "../actions/flashPreview";
import type { GridLayout } from "./layout";

export function FlashPreviewUi({ ctl, layout }: { ctl: Controller; layout: GridLayout }) {
  const p = flashPreviewFor(ctl);
  if (!p || p.sheetId !== ctl.sheet.id) return null;
  const [r1, r2] = layout.visibleRows();
  const items = p.fills.filter((f) => f.row >= r1 && f.row <= r2);
  return (
    <>
      {items.map((f) => {
        const rc = layout.cellRect(f.row, p.col);
        return (
          <div key={f.row} className="tb-flash-preview" style={{ left: rc.x, top: rc.y, width: rc.w, height: rc.h, lineHeight: `${rc.h}px`, fontSize: `${((11 * 4) / 3) * layout.z}px` }}>
            {f.text}
          </div>
        );
      })}
    </>
  );
}
