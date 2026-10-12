// Yatay cetvel: kenar boşlukları ve girinti işaretleri sürüklenebilir.
import { useRef, useState } from "react";
import { tipProps } from "@officehere/ui";

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
  onIndent: (v: { indentLeft?: number; indentRight?: number; indentFirst?: number }) => void;
  onMargins: (v: { left?: number; right?: number }) => void;
  onDoubleClick: () => void;
}

type DragKind = "first" | "hanging" | "left" | "right" | "marginL" | "marginR";

export function Ruler({ offset, scale, pageW, ml, mr, indentLeft, indentRight, indentFirst, onIndent, onMargins, onDoubleClick }: Props) {
  const ref = useRef<HTMLDivElement>(null);
  const [drag, setDrag] = useState<{ kind: DragKind; x: number } | null>(null);
  const X = (px: number) => offset + px * scale;

  const firstX = ml + (indentLeft + indentFirst) * PT_TO_PX;
  const leftX = ml + indentLeft * PT_TO_PX;
  const rightX = pageW - mr - indentRight * PT_TO_PX;

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
          onIndent({ indentFirst: rel(Math.max(0, Math.min(pageW - mr, p)) - ml) - indentLeft });
          break;
        case "hanging": {
          const nl = rel(Math.max(0, Math.min(pageW - mr, p)) - ml);
          onIndent({ indentLeft: nl, indentFirst: +(indentLeft + indentFirst - nl).toFixed(2) });
          break;
        }
        case "left":
          onIndent({ indentLeft: rel(Math.max(0, Math.min(pageW - mr, p)) - ml) });
          break;
        case "right":
          onIndent({ indentRight: rel(pageW - mr - Math.max(ml, Math.min(pageW, p))) });
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
    <div className="bl-ruler" ref={ref} onDoubleClick={onDoubleClick}>
      <div className="bl-ruler-track" style={{ left: X(0), width: pageW * scale }}>
        <div className="bl-ruler-margin" style={{ left: 0, width: ml * scale }} />
        <div className="bl-ruler-margin" style={{ right: 0, width: mr * scale }} />
      </div>
      {ticks}
      <div className="bl-ruler-mhandle" style={{ left: X(ml) - 3 }} onPointerDown={start("marginL")} {...tipProps("Sol Kenar Boşluğu")} />
      <div className="bl-ruler-mhandle" style={{ left: X(pageW - mr) - 3 }} onPointerDown={start("marginR")} {...tipProps("Sağ Kenar Boşluğu")} />
      <div className="bl-ind bl-ind-first" style={{ left: X(firstX) }} onPointerDown={start("first")} {...tipProps("İlk Satır Girintisi")} />
      <div className="bl-ind bl-ind-hanging" style={{ left: X(leftX) }} onPointerDown={start("hanging")} {...tipProps("Asılı Girinti")} />
      <div className="bl-ind bl-ind-left" style={{ left: X(leftX) }} onPointerDown={start("left")} {...tipProps("Sol Girinti")} />
      <div className="bl-ind bl-ind-right" style={{ left: X(rightX) }} onPointerDown={start("right")} {...tipProps("Sağ Girinti")} />
      {drag && <div className="bl-ruler-guide" style={{ left: drag.x }} />}
    </div>
  );
}
