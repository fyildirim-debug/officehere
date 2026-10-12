/** Sayfa Düzeni görünümünün parçaları: yatay/dikey cetvel (cm) ve sayfa kenarındaki sütun harfleri / satır numaraları. */
import { colToName } from "@officehere/formula";
import type { Rect } from "../model/types";
import type { SheetModel } from "../model/workbook";

const CM = 96 / 2.54;
export const RULER = 16;
export const HEAD_W = 30;
export const HEAD_H = 16;

/** Cetvel: kenar boşlukları gri, her santimetrede numara, yarım santimetrede kısa çizgi. */
export function Ruler({ length, z, start, end, vertical }: { length: number; z: number; start: number; end: number; vertical?: boolean }) {
  const n = Math.floor(length / CM);
  const marks: React.ReactNode[] = [];
  // Numaralar içerik alanının başından (sol/üst kenar boşluğundan) sayılır
  for (let i = 0; i <= n * 2; i++) {
    const pos = (i * CM) / 2;
    if (pos > length) break;
    const rel = Math.round((pos - start) / CM);
    const full = i % 2 === 0;
    const st = vertical ? { top: pos * z } : { left: pos * z };
    marks.push(full && rel !== 0 ? <span key={i} className="tb-pl-rnum" style={st}>{Math.abs(rel)}</span> : <span key={i} className={`tb-pl-tick${full ? " is-full" : ""}`} style={st} />);
  }
  const box = vertical ? { width: RULER, height: length * z } : { width: length * z, height: RULER };
  return (
    <div className={`tb-pl-ruler${vertical ? " is-v" : ""}`} style={box} aria-hidden="true">
      <div className="tb-pl-rmargin" style={vertical ? { top: 0, height: start * z } : { left: 0, width: start * z }} />
      <div className="tb-pl-rmargin" style={vertical ? { bottom: 0, height: end * z } : { right: 0, width: end * z }} />
      {marks}
    </div>
  );
}

/** Sayfanın üstünde sütun harfleri, solunda satır numaraları (içerikle aynı ölçekte). */
export function Headings({ sheet, r, k, axis, active }: { sheet: SheetModel; r: Rect; k: number; axis: "col" | "row"; active: { r1: number; c1: number; r2: number; c2: number }[] }) {
  const cells: React.ReactNode[] = [];
  if (axis === "col") {
    for (let c = r.c1; c <= r.c2; c++) {
      const w = sheet.cols.size(c);
      if (!w) continue;
      const on = active.some((s) => c >= s.c1 && c <= s.c2);
      cells.push(
        <div key={c} className={`tb-pl-head${on ? " is-on" : ""}`} style={{ width: w * k }}>
          {colToName(c)}
        </div>,
      );
    }
    return <div className="tb-pl-heads is-col">{cells}</div>;
  }
  for (let row = r.r1; row <= r.r2; row++) {
    const h = sheet.rows.size(row);
    if (!h) continue;
    const on = active.some((s) => row >= s.r1 && row <= s.r2);
    cells.push(
      <div key={row} className={`tb-pl-head${on ? " is-on" : ""}`} style={{ height: h * k, lineHeight: `${h * k}px` }}>
        {row + 1}
      </div>,
    );
  }
  return <div className="tb-pl-heads is-row">{cells}</div>;
}
