// Dikey cetvel: imleçteki sayfanın üst/alt kenar boşlukları; kenar boşlukları sürüklenebilir.
import { useRef, useState } from "react";
import { tipProps } from "@officehere/ui";

const PX_PER_CM = 96 / 2.54;

interface Props {
  /** Cetvel alanında sayfanın üst kenarının y konumu (px, ekran). */
  offset: number;
  scale: number;
  pageH: number;
  mt: number;
  mb: number;
  onMargins: (v: { top?: number; bottom?: number }) => void;
  onDoubleClick: () => void;
}

export function VRuler({ offset, scale, pageH, mt, mb, onMargins, onDoubleClick }: Props) {
  const ref = useRef<HTMLDivElement>(null);
  const [drag, setDrag] = useState<number | null>(null);
  const Yp = (px: number) => offset + px * scale;

  const start = (kind: "top" | "bottom") => (e: React.PointerEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const host = ref.current;
    if (!host) return;
    const rect = host.getBoundingClientRect();
    const snap = (px: number, fine: boolean) => (fine ? px : Math.round(px / (PX_PER_CM / 4)) * (PX_PER_CM / 4));
    const move = (ev: PointerEvent) => {
      const p = (ev.clientY - rect.top - offset) / scale;
      setDrag(ev.clientY - rect.top);
      if (kind === "top") onMargins({ top: Math.max(0, Math.min(pageH - mb - PX_PER_CM, snap(p, ev.altKey))) });
      else onMargins({ bottom: Math.max(0, Math.min(pageH - mt - PX_PER_CM, pageH - snap(p, ev.altKey))) });
    };
    const up = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      setDrag(null);
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  };

  // Ölçek: üst kenar boşluğu 0 noktası, her 0,25 cm çizgi, her cm sayı.
  const ticks = [];
  const step = PX_PER_CM / 4;
  const from = -Math.floor(mt / step);
  const to = Math.floor((pageH - mt) / step);
  const cmPx = PX_PER_CM * scale;
  const numEvery = cmPx < 14 ? 4 : cmPx < 24 ? 2 : 1;
  for (let i = from; i <= to; i++) {
    const y = Yp(mt + i * step);
    if (i % 4 !== 0 && cmPx < 20 && i % 2 !== 0) continue;
    if (i % 4 === 0) {
      const n = Math.abs(i / 4);
      if (n !== 0 && n % numEvery === 0)
        ticks.push(
          <span key={i} className="bl-vruler-num" style={{ top: y }}>
            {n}
          </span>,
        );
      else if (n !== 0) ticks.push(<span key={i} className="bl-vruler-tick is-half" style={{ top: y }} />);
    } else ticks.push(<span key={i} className={`bl-vruler-tick${i % 2 === 0 ? " is-half" : ""}`} style={{ top: y }} />);
  }

  return (
    <div className="bl-vruler" ref={ref} onDoubleClick={onDoubleClick} aria-hidden="true">
      <div className="bl-vruler-track" style={{ top: Yp(0), height: pageH * scale }}>
        <div className="bl-vruler-margin" style={{ top: 0, height: mt * scale }} />
        <div className="bl-vruler-margin" style={{ bottom: 0, height: mb * scale }} />
      </div>
      {ticks}
      <div className="bl-vruler-mhandle" style={{ top: Yp(mt) - 3 }} onPointerDown={start("top")} {...tipProps("Üst Kenar Boşluğu")} />
      <div className="bl-vruler-mhandle" style={{ top: Yp(pageH - mb) - 3 }} onPointerDown={start("bottom")} {...tipProps("Alt Kenar Boşluğu")} />
      {drag !== null && <div className="bl-vruler-guide" style={{ top: drag }} />}
    </div>
  );
}
