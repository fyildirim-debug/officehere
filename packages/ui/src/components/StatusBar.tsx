import { useRef, type ReactElement, type ReactNode } from "react";
import { tipProps } from "./Tooltip";

export function StatusBar({ left, right }: { left?: ReactNode; right?: ReactNode }) {
  return (
    <footer className="oh-statusbar">
      <div className="oh-statusbar-left">{left}</div>
      <div className="oh-statusbar-right">{right}</div>
    </footer>
  );
}

export function StatusItem({ children, onClick, tip }: { children: ReactNode; onClick?: () => void; tip?: string }) {
  if (onClick) {
    return (
      <button type="button" className="oh-status-item is-button" onClick={onClick} {...tipProps(tip)}>
        {children}
      </button>
    );
  }
  return (
    <span className="oh-status-item" {...tipProps(tip)}>
      {children}
    </span>
  );
}

export interface ViewOption {
  id: string;
  icon: ReactElement;
  tip: string;
}

export function ViewSwitcher({ views, active, onChange }: { views: ViewOption[]; active: string; onChange: (id: string) => void }) {
  return (
    <div className="oh-viewswitch">
      {views.map((v) => (
        <button key={v.id} type="button" className={`oh-status-icon${v.id === active ? " is-checked" : ""}`} onClick={() => onChange(v.id)} aria-label={v.tip} {...tipProps(v.tip)}>
          {v.icon}
        </button>
      ))}
    </div>
  );
}

export interface ZoomSliderProps {
  /** Yüzde (100 = %100). */
  value: number;
  onChange: (value: number) => void;
  min?: number;
  max?: number;
  /** Yüzdeye tıklanınca (Yakınlaştır iletişim kutusu). */
  onPercentClick?: () => void;
}

/**
 * Ofis yakınlaştırma kaydırıcısı: orta nokta %100'dür; sol yarı min..100,
 * sağ yarı 100..max aralığını doğrusal gösterir.
 */
export function ZoomSlider({ value, onChange, min = 10, max = 500, onPercentClick }: ZoomSliderProps) {
  const track = useRef<HTMLDivElement>(null);
  const toPos = (v: number) => (v <= 100 ? ((v - min) / (100 - min)) * 0.5 : 0.5 + ((v - 100) / (max - 100)) * 0.5);
  const fromPos = (p: number) => {
    const c = Math.min(1, Math.max(0, p));
    const v = c <= 0.5 ? min + (c / 0.5) * (100 - min) : 100 + ((c - 0.5) / 0.5) * (max - 100);
    return Math.abs(v - 100) < 4 ? 100 : Math.round(v);
  };
  const step = (dir: 1 | -1) => {
    const v = value;
    const next = dir > 0 ? Math.floor(v / 10) * 10 + 10 : Math.ceil(v / 10) * 10 - 10;
    onChange(Math.min(max, Math.max(min, next)));
  };
  const drag = (e: React.PointerEvent) => {
    const el = track.current;
    if (!el) return;
    const r = el.getBoundingClientRect();
    const move = (x: number) => onChange(fromPos((x - r.left) / r.width));
    move(e.clientX);
    const onMove = (ev: PointerEvent) => move(ev.clientX);
    const onUp = () => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
  };
  return (
    <div className="oh-zoom">
      <button type="button" className="oh-zoom-btn" onClick={() => step(-1)} aria-label="Uzaklaştır" {...tipProps("Uzaklaştır")}>
        −
      </button>
      <div className="oh-zoom-track" ref={track} onPointerDown={drag} {...tipProps("Yakınlaştırma")}>
        <div className="oh-zoom-mid" />
        <div className="oh-zoom-thumb" style={{ left: `${toPos(value) * 100}%` }} />
      </div>
      <button type="button" className="oh-zoom-btn" onClick={() => step(1)} aria-label="Yakınlaştır" {...tipProps("Yakınlaştır")}>
        +
      </button>
      <button type="button" className="oh-zoom-pct" onClick={onPercentClick} {...tipProps("Yakınlaştırma düzeyi")}>
        %{Math.round(value)}
      </button>
    </div>
  );
}
