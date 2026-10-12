/** Sunucu Görünümü yerleşimi: geçerli slayt, sonraki slayt, notlar, süre ve denetimler. */
import { useEffect, useState, type ReactElement, type ReactNode, type RefObject } from "react";
import type { Presentation, Slide } from "../model/types";
import { ScaledSlide } from "../render/SlideView";
import type { ShowTools } from "./ShowControls";
import { formatTime } from "./showKeys";

export interface PresenterLayoutProps {
  pres: Presentation;
  index: number;
  slide: Slide | undefined;
  /** Sonraki tıklamada gösterilecek slayt (adım kaldıysa aynı slayt). */
  nextSlide: Slide | undefined;
  nextNumber: number;
  stage: ReactNode;
  areaRef: RefObject<HTMLDivElement | null>;
  rootRef: RefObject<HTMLDivElement | null>;
  onKeyDown: (e: React.KeyboardEvent) => void;
  onClick: (e: React.MouseEvent) => void;
  onMenu: (e: React.MouseEvent) => void;
  onOpenMenu: (e: React.MouseEvent, m: ReactElement) => void;
  onExit: () => void;
  onBack: () => void;
  onNext: () => void;
  onGrid: () => void;
  onBlack: () => void;
  tools: ShowTools;
  /** İzleyici penceresi durumu (yoksa tek pencere). */
  audience: "none" | "waiting" | "connected";
  /** Gösteri ızgarası ve menüler. */
  overlay: ReactNode;
}

export function PresenterLayout(p: PresenterLayoutProps) {
  const [startTime, setStartTime] = useState(() => Date.now());
  const [now, setNow] = useState(() => Date.now());
  const [paused, setPaused] = useState<number | null>(null);
  const [notesZoom, setNotesZoom] = useState(1);
  useEffect(() => {
    const t = window.setInterval(() => setNow(Date.now()), 1000);
    return () => window.clearInterval(t);
  }, []);
  const elapsed = (paused ?? now) - startTime;
  return (
    <div ref={p.rootRef} className="show-root presenter" tabIndex={0} onKeyDown={p.onKeyDown} onContextMenu={(e) => e.preventDefault()}>
      <div className="presenter-top">
        <span className="presenter-timer" aria-label="Geçen süre">
          {formatTime(elapsed, true)}
        </span>
        <button type="button" className="presenter-btn" title={paused ? "Süreölçeri sürdür" : "Süreölçeri duraklat"} onClick={() => (paused ? (setStartTime((t) => t + (Date.now() - paused)), setPaused(null)) : setPaused(Date.now()))}>
          {paused ? "▶" : "❚❚"}
        </button>
        <button type="button" className="presenter-btn" title="Süreölçeri yeniden başlat" onClick={() => (setStartTime(Date.now()), setPaused(null))}>
          ↺
        </button>
        <span>{new Intl.DateTimeFormat("tr-TR", { hour: "2-digit", minute: "2-digit" }).format(now)}</span>
        {p.audience !== "none" && <span className={`presenter-audience is-${p.audience}`}>{p.audience === "connected" ? "İzleyici penceresi bağlı" : "İzleyici penceresi açılıyor…"}</span>}
        <button type="button" className="presenter-btn" onClick={p.onExit}>
          Gösteriyi Bitir
        </button>
      </div>
      <div className="presenter-main">
        <div className="presenter-current" ref={p.areaRef} onClick={p.onClick} onContextMenu={p.onMenu}>
          {p.stage}
        </div>
        <div className="presenter-side">
          <div className="presenter-label">Sonraki slayt</div>
          {p.nextSlide ? <ScaledSlide pres={p.pres} slide={p.nextSlide} slideNumber={p.nextNumber} width={340} mode="show" /> : <div className="presenter-endbox">Slayt gösterisinin sonu</div>}
          <div className="presenter-label">
            Notlar
            <span>
              <button type="button" className="presenter-btn" onClick={() => setNotesZoom((z) => Math.min(2.5, z + 0.2))} title="Metni büyüt">
                A+
              </button>
              <button type="button" className="presenter-btn" onClick={() => setNotesZoom((z) => Math.max(0.6, z - 0.2))} title="Metni küçült">
                A−
              </button>
            </span>
          </div>
          <div className="presenter-notes" style={{ fontSize: 16 * notesZoom }}>
            {p.slide?.notes || <em>Bu slayt için not yok.</em>}
          </div>
        </div>
      </div>
      <div className="presenter-bottom">
        <button type="button" className="presenter-btn" onClick={p.onBack} title="Önceki">
          ◀
        </button>
        <span>
          Slayt {p.index + 1} / {p.pres.slides.length}
        </span>
        <button type="button" className="presenter-btn" onClick={p.onNext} title="Sonraki">
          ▶
        </button>
        <button type="button" className={`presenter-btn${p.tools.tool !== "pointer" ? " is-on" : ""}`} onClick={(e) => p.onOpenMenu(e, p.tools.toolMenu())}>
          Kalem ve Lazer İşaretçisi
        </button>
        <button type="button" className="presenter-btn" onClick={p.onGrid}>
          Tüm Slaytları Gör
        </button>
        <button type="button" className="presenter-btn" onClick={p.onBlack}>
          Siyah Ekran
        </button>
      </div>
      {p.overlay}
    </div>
  );
}
