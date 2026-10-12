// Yatay cetvel: kenar boşlukları ve girinti işaretleri sürüklenebilir.
import { useRef, useState } from "react";
import { tipProps } from "@officehere/ui";
import { TAB_ALIGN_LABEL, normalizeTabs, type TabStop } from "../editor/tabStops";

const PX_PER_CM = 96 / 2.54;
const PT_TO_PX = 4 / 3;

interface Props {
  /** Cetvel alanında sayfanın sol kenarının x konumu (px, ekran). */
  offset: number;
  scale: number;
  pageW: number;
  ml: number;
  mr: number;
  indentLeft: number;
  indentRight: number;
  indentFirst: number;
  /** Çok sütunlu bölümde sütunlar (sayfa koordinatı) ve imlecin sütunu. */
  columns?: { left: number; width: number }[];
  column?: number;
  onIndent: (v: { indentLeft?: number; indentRight?: number; indentFirst?: number }) => void;
  onMargins: (v: { left?: number; right?: number }) => void;
  onDoubleClick: () => void;
  /** İmleçteki paragrafın sekme durakları; tıklayınca sol durak eklenir, sürükleyince taşınır/kaldırılır. */
  tabs?: TabStop[];
  onTabs?: (tabs: TabStop[] | null) => void;
}

type DragKind = "first" | "hanging" | "left" | "right" | "marginL" | "marginR";

export function Ruler({ offset, scale, pageW, ml, mr, indentLeft, indentRight, indentFirst, columns, column = 0, onIndent, onMargins, onDoubleClick, tabs = [], onTabs }: Props) {
  const ref = useRef<HTMLDivElement>(null);
  const [drag, setDrag] = useState<{ kind: DragKind; x: number } | null>(null);
  const X = (px: number) => offset + px * scale;

  // Girintiler imlecin bulunduğu sütuna göredir.
  const col = columns?.[Math.min(column, columns.length - 1)];
  const cl = col ? col.left : ml;
  const cr = col ? pageW - col.left - col.width : mr;
  const firstX = cl + (indentLeft + indentFirst) * PT_TO_PX;
  const leftX = cl + indentLeft * PT_TO_PX;
  const rightX = pageW - cr - indentRight * PT_TO_PX;

  const start = (kind: DragKind) => (e: React.PointerEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const host = ref.current;
    if (!host) return;
    const rect = host.getBoundingClientRect();
    const snap = (px: number, fine: boolean) => (fine ? px : Math.round(px / (PX_PER_CM / 4)) * (PX_PER_CM / 4));
    const toPage = (cx: number) => (cx - rect.left - offset) / scale;
    const move = (ev: PointerEvent) => {
      const p = toPage(ev.clientX);
      setDrag({ kind, x: ev.clientX - rect.left });
      const fine = ev.altKey;
      const rel = (v: number) => +(snap(v, fine) / PT_TO_PX).toFixed(2);
      switch (kind) {
        case "first":
          onIndent({ indentFirst: rel(Math.max(0, Math.min(pageW - cr, p)) - cl) - indentLeft });
          break;
        case "hanging": {
          const nl = rel(Math.max(0, Math.min(pageW - cr, p)) - cl);
          onIndent({ indentLeft: nl, indentFirst: +(indentLeft + indentFirst - nl).toFixed(2) });
          break;
        }
        case "left":
          onIndent({ indentLeft: rel(Math.max(0, Math.min(pageW - cr, p)) - cl) });
          break;
        case "right":
          onIndent({ indentRight: rel(pageW - cr - Math.max(cl, Math.min(pageW, p))) });
          break;
        case "marginL":
          onMargins({ left: Math.max(0, Math.min(pageW - mr - PX_PER_CM, snap(p, fine))) });
          break;
        case "marginR":
          onMargins({ right: Math.max(0, Math.min(pageW - ml - PX_PER_CM, pageW - snap(p, fine))) });
          break;
      }
    };
    const up = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      setDrag(null);
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  };

  /** Cetvelin alt kısmına tıklama sol sekme durağı ekler; durak sürüklenince taşınır, cetvelden çekilince kaldırılır. */
  const tabDown = (index: number | null) => (e: React.PointerEvent) => {
    if (!onTabs || e.button !== 0) return;
    const host = ref.current;
    if (!host) return;
    if (index === null && (e.target as HTMLElement).closest(".bl-ind, .bl-ruler-mhandle, .bl-ruler-tab")) return;
    e.preventDefault();
    e.stopPropagation();
    const rect = host.getBoundingClientRect();
    const toPt = (cx: number) => +((Math.round((cx - rect.left - offset) / scale / (PX_PER_CM / 4)) * (PX_PER_CM / 4) - cl) / PT_TO_PX).toFixed(1);
    if (index === null) {
      const pt = toPt(e.clientX);
      if (pt <= 0 || cl + pt * PT_TO_PX >= pageW - cr) return;
      onTabs(normalizeTabs([...tabs, { pos: pt, align: "left", leader: "none" }]));
      return;
    }
    const orig = tabs[index];
    let last = tabs;
    const move = (ev: PointerEvent) => {
      const off = Math.abs(ev.clientY - rect.top - rect.height / 2) > 24;
      const rest = tabs.filter((_, i) => i !== index);
      last = off ? rest : normalizeTabs([...rest, { ...orig, pos: Math.max(0.5, toPt(ev.clientX)) }]);
      setDrag({ kind: "first", x: ev.clientX - rect.left });
    };
    const up = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      setDrag(null);
      onTabs(last.length ? last : null);
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  };

  // Ölçek çizgileri: her 0,25 cm; sol kenar boşluğu 0 noktasıdır.
  const ticks = [];
  const step = PX_PER_CM / 4;
  const from = -Math.floor(ml / step);
  const to = Math.floor((pageW - ml) / step);
  const cmPx = PX_PER_CM * scale;
  // Düşük yakınlaştırmada sayılar ve çizgiler seyreltilir.
  const numEvery = cmPx < 14 ? 4 : cmPx < 24 ? 2 : 1;
  for (let i = from; i <= to; i++) {
    const x = X(ml + i * step);
    if (i % 4 !== 0 && cmPx < 20 && i % 2 !== 0) continue;
    if (i % 4 === 0) {
      const n = Math.abs(i / 4);
      if (n !== 0 && n % numEvery === 0)
        ticks.push(
          <span key={i} className="bl-ruler-num" style={{ left: x }}>
            {n}
          </span>,
        );
      else if (n !== 0) ticks.push(<span key={i} className="bl-ruler-tick is-half" style={{ left: x }} />);
    } else {
      ticks.push(<span key={i} className={`bl-ruler-tick${i % 2 === 0 ? " is-half" : ""}`} style={{ left: x }} />);
    }
  }

  return (
    <div className="bl-ruler" ref={ref} onDoubleClick={onDoubleClick} onPointerDown={tabDown(null)}>
      <div className="bl-ruler-track" style={{ left: X(0), width: pageW * scale }}>
        <div className="bl-ruler-margin" style={{ left: 0, width: ml * scale }} />
        <div className="bl-ruler-margin" style={{ right: 0, width: mr * scale }} />
        {columns?.slice(0, -1).map((c, i) => (
          <div key={i} className="bl-ruler-margin is-colgap" style={{ left: (c.left + c.width) * scale, width: (columns[i + 1].left - c.left - c.width) * scale }} />
        ))}
      </div>
      {ticks}
      <div className="bl-ruler-mhandle" style={{ left: X(ml) - 3 }} onPointerDown={start("marginL")} {...tipProps("Sol Kenar Boşluğu")} />
      <div className="bl-ruler-mhandle" style={{ left: X(pageW - mr) - 3 }} onPointerDown={start("marginR")} {...tipProps("Sağ Kenar Boşluğu")} />
      <div className="bl-ind bl-ind-first" style={{ left: X(firstX) }} onPointerDown={start("first")} {...tipProps("İlk Satır Girintisi")} />
      <div className="bl-ind bl-ind-hanging" style={{ left: X(leftX) }} onPointerDown={start("hanging")} {...tipProps("Asılı Girinti")} />
      <div className="bl-ind bl-ind-left" style={{ left: X(leftX) }} onPointerDown={start("left")} {...tipProps("Sol Girinti")} />
      <div className="bl-ind bl-ind-right" style={{ left: X(rightX) }} onPointerDown={start("right")} {...tipProps("Sağ Girinti")} />
      {tabs.map((t, i) => (
        <div key={`${t.pos}-${t.align}`} className={`bl-ruler-tab is-${t.align}`} style={{ left: X(cl + t.pos * PT_TO_PX) }} onPointerDown={tabDown(i)} {...tipProps(`${TAB_ALIGN_LABEL[t.align]} Sekme`)} />
      ))}
      {drag && <div className="bl-ruler-guide" style={{ left: drag.x }} />}
    </div>
  );
}
