/** Alan seçilmemiş PivotTable için ızgara üzerindeki yer tutucu kutu. */
import type { Controller } from "../controller";
import type { GridLayout } from "../grid/layout";
import { placeholderRect } from "./pivotOps";

export function PivotPlaceholders({ ctl, layout }: { ctl: Controller; layout: GridLayout | null }) {
  if (!layout) return null;
  const empty = ctl.sheet.pivots.filter((p) => p.rows.length + p.cols.length + p.values.length === 0);
  if (!empty.length) return null;
  return (
    <>
      {empty.map((p) => {
        const r = placeholderRect(p);
        const rc = layout.cellRect(r.r1, r.c1, r.r2, r.c2);
        if (rc.x + rc.w < layout.headerW || rc.y + rc.h < layout.headerH) return null;
        return (
          <div key={p.id} className="tb-pv-placeholder" style={{ left: rc.x, top: rc.y, width: rc.w, height: rc.h }}>
            <strong>{p.name}</strong>
            <span>Rapor oluşturmak için Özet Tablo Alanları bölmesinden alanları seçin</span>
          </div>
        );
      })}
    </>
  );
}
