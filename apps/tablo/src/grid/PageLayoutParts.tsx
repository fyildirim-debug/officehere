/** Sayfa Düzeni görünümünün parçaları: yatay/dikey cetvel (cm) ve sayfa kenarındaki sütun harfleri / satır numaraları. */
import { useRef, useState } from "react";
import { colToName } from "@officehere/formula";
import type { Rect } from "../model/types";
import type { SheetModel } from "../model/workbook";

const CM = 96 / 2.54;
export const RULER = 16;
export const HEAD_W = 30;
export const HEAD_H = 16;

/** Üst/alt bilgi yerinde düzenleme kutusu (&[Sayfa], &[Sayfalar], &[Tarih] kodları kullanılabilir). */
export function HfEditor({ value, onDone }: { value: string; onDone: (v: string | null) => void }) {
  const [text, setText] = useState(value);
  const done = useRef(false);
  const finish = (v: string | null) => {
    if (done.current) return;
    done.current = true;
    onDone(v);
  };
  return (
    <input
      className="tb-pl-hf-input"
      aria-label="Üst bilgi / alt bilgi"
      autoFocus
      value={text}
      onChange={(e) => setText(e.target.value)}
      onClick={(e) => e.stopPropagation()}
      onBlur={() => finish(text)}
      onKeyDown={(e) => {
        e.stopPropagation();
        if (e.key === "Enter") {
          e.preventDefault();
          finish(text);
        } else if (e.key === "Escape") finish(null);
      }}
    />
  );
}

/** Cetvel: kenar boşlukları gri, her santimetrede numara, yarım santimetrede kısa çizgi; kenar boşluğu sınırları sürüklenebilir. */
export function Ruler({ length, z, start, end, vertical, onChange }: { length: number; z: number; start: number; end: number; vertical?: boolean; onChange?: (start: number, end: number) => void }) {
  const [live, setLive] = useState<{ start: number; end: number } | null>(null);
  const cur = live ?? { start, end };
  const handle = (which: "start" | "end") => (ev: React.PointerEvent<HTMLDivElement>) => {
    if (!onChange) return;
    ev.preventDefault();
    ev.stopPropagation();
    const el = ev.currentTarget;
    el.setPointerCapture(ev.pointerId);
    const origin = vertical ? ev.clientY : ev.clientX;
    const base = which === "start" ? start : end;
    let next = { start, end };
    const move = (e: PointerEvent) => {
      const d = ((vertical ? e.clientY : e.clientX) - origin) / z;
      const v = Math.max(0, Math.min(length / 2 - 20, which === "start" ? base + d : base - d));
      next = which === "start" ? { start: v, end } : { start, end: v };
      setLive(next);
    };
    const up = () => {
      el.removeEventListener("pointermove", move);
      el.removeEventListener("pointerup", up);
      setLive(null);
      if (next.start !== start || next.end !== end) onChange(next.start, next.end);
    };
    el.addEventListener("pointermove", move);
    el.addEventListener("pointerup", up);
  };
  const cs = cur.start;
  const ce = cur.end;
  const n = Math.floor(length / CM);
  const marks: React.ReactNode[] = [];
  // Numaralar içerik alanının başından (sol/üst kenar boşluğundan) sayılır
  for (let i = 0; i <= n * 2; i++) {
    const pos = (i * CM) / 2;
    if (pos > length) break;
    const rel = Math.round((pos - cs) / CM);
    const full = i % 2 === 0;
    const st = vertical ? { top: pos * z } : { left: pos * z };
    marks.push(full && rel !== 0 ? <span key={i} className="tb-pl-rnum" style={st}>{Math.abs(rel)}</span> : <span key={i} className={`tb-pl-tick${full ? " is-full" : ""}`} style={st} />);
  }
  const box = vertical ? { width: RULER, height: length * z } : { width: length * z, height: RULER };
  return (
    <div className={`tb-pl-ruler${vertical ? " is-v" : ""}`} style={box}>
      <div className="tb-pl-rmargin" style={vertical ? { top: 0, height: cs * z } : { left: 0, width: cs * z }} />
      <div className="tb-pl-rmargin" style={vertical ? { bottom: 0, height: ce * z } : { right: 0, width: ce * z }} />
      {marks}
      {onChange && (
        <>
          <div className={`tb-pl-rhandle${vertical ? " is-v" : ""}`} title={vertical ? "Üst Kenar Boşluğu" : "Sol Kenar Boşluğu"} style={vertical ? { top: cs * z - 3 } : { left: cs * z - 3 }} onPointerDown={handle("start")} />
          <div className={`tb-pl-rhandle${vertical ? " is-v" : ""}`} title={vertical ? "Alt Kenar Boşluğu" : "Sağ Kenar Boşluğu"} style={vertical ? { top: (length - ce) * z - 3 } : { left: (length - ce) * z - 3 }} onPointerDown={handle("end")} />
        </>
      )}
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
