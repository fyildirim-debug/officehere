/**
 * Damlalık: renk menülerinden açılır; slayttaki herhangi bir noktanın rengini
 * (resimler ve gradyanlar dahil) alır. Slayt bir kez tuvale çizilir, imleç
 * altındaki piksel okunur; tıklama rengi uygular, Esc vazgeçer.
 */
import { useEffect, useRef, useState, useSyncExternalStore } from "react";
import type { Presentation } from "../model/types";

type Pick = (hex: string) => void;

let active: Pick | null = null;
const listeners = new Set<() => void>();
const emit = () => listeners.forEach((l) => l());

/** Damlalık kipini başlatır; seçilen renk "#RRGGBB" olarak verilir. */
export function startEyedropper(onPick: Pick) {
  active = onPick;
  emit();
}

export function stopEyedropper() {
  if (!active) return;
  active = null;
  emit();
}

export function useEyedropper(): Pick | null {
  return useSyncExternalStore(
    (l) => {
      listeners.add(l);
      return () => listeners.delete(l);
    },
    () => active,
  );
}

/** Damlalık kullanılabilir mi (Normal görünümde slayt alanı var mı)? */
export function eyedropperAvailable(): boolean {
  // Kalıcı (modal) iletişim kutusu açıkken slayta tıklanamaz.
  return typeof document !== "undefined" && !!document.querySelector(".stage .stage-slide") && !document.querySelector('[role="dialog"][aria-modal="true"]');
}

const toHex = (r: number, g: number, b: number) => `#${[r, g, b].map((v) => v.toString(16).padStart(2, "0")).join("").toUpperCase()}`;

/** Çizilmiş slaydın (ölçekli) piksel verisinde slayt noktasının rengi; slayt dışındaysa null. */
export function samplePixels(img: { width: number; height: number; data: ArrayLike<number> }, scale: number, sx: number, sy: number): string | null {
  const x = Math.floor(sx * scale);
  const y = Math.floor(sy * scale);
  if (x < 0 || y < 0 || x >= img.width || y >= img.height) return null;
  const i = (y * img.width + x) * 4;
  return toHex(img.data[i], img.data[i + 1], img.data[i + 2]);
}

interface Props {
  pres: Presentation;
  slideIdx: number;
  /** Ekran noktasını slayt koordinatına çevirir. */
  toSlide: (cx: number, cy: number) => [number, number];
  /** Slaydı tuvale çizer (dışa aktarma ile aynı çizim). */
  rasterize: (pres: Presentation, index: number, scale: number) => Promise<HTMLCanvasElement>;
}

const SCALE = 1;

export function EyedropperOverlay({ pres, slideIdx, toSlide, rasterize }: Props) {
  const pick = useEyedropper();
  // Piksel verisi bir kez okunur (her fare hareketinde tuvalden okumak yavaştır).
  const pixelsRef = useRef<ImageData | null>(null);
  const [tip, setTip] = useState<{ x: number; y: number; hex: string } | null>(null);
  const [ready, setReady] = useState(false);
  const rootRef = useRef<HTMLDivElement>(null);

  // Kip açılınca slayt bir kez çizilir.
  useEffect(() => {
    if (!pick) return;
    let alive = true;
    setReady(false);
    pixelsRef.current = null;
    void rasterize(pres, slideIdx, SCALE)
      .then((c) => {
        if (!alive) return;
        pixelsRef.current = c.getContext("2d")?.getImageData(0, 0, c.width, c.height) ?? null;
        setReady(true);
      })
      .catch(() => alive && stopEyedropper());
    return () => {
      alive = false;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [pick, slideIdx]);

  // Esc vazgeçer (uygulama kısayollarına ulaşmadan).
  useEffect(() => {
    if (!pick) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== "Escape") return;
      e.preventDefault();
      e.stopPropagation();
      stopEyedropper();
    };
    window.addEventListener("keydown", onKey, true);
    return () => window.removeEventListener("keydown", onKey, true);
  }, [pick]);

  // Görünüm değişince (slayt alanı kalkınca) kip kapanır.
  useEffect(() => () => stopEyedropper(), []);

  if (!pick) return null;

  const colorAt = (cx: number, cy: number): string | null => {
    const c = pixelsRef.current;
    if (!c) return null;
    const [sx, sy] = toSlide(cx, cy);
    return samplePixels(c, SCALE, sx, sy);
  };

  const local = (cx: number, cy: number) => {
    const r = rootRef.current!.getBoundingClientRect();
    return { x: cx - r.left, y: cy - r.top };
  };

  return (
    <div
      ref={rootRef}
      className={`eyedropper-layer${ready ? " is-ready" : ""}`}
      role="application"
      aria-label="Damlalık: renk almak için slayta tıklayın"
      onPointerDown={(e) => {
        e.preventDefault();
        e.stopPropagation();
        if (e.button !== 0) {
          stopEyedropper();
          return;
        }
        const hex = colorAt(e.clientX, e.clientY);
        if (!hex) return;
        stopEyedropper();
        pick(hex);
      }}
      onPointerMove={(e) => {
        const hex = colorAt(e.clientX, e.clientY);
        setTip(hex ? { ...local(e.clientX, e.clientY), hex } : null);
      }}
      onPointerLeave={() => setTip(null)}
      onContextMenu={(e) => e.preventDefault()}
    >
      {tip && (
        <div className="eyedropper-tip" style={{ left: tip.x + 18, top: tip.y + 18 }} data-color={tip.hex}>
          <span className="eyedropper-swatch" style={{ background: tip.hex }} />
          <span>
            RGB({parseInt(tip.hex.slice(1, 3), 16)}, {parseInt(tip.hex.slice(3, 5), 16)}, {parseInt(tip.hex.slice(5, 7), 16)})
          </span>
        </div>
      )}
    </div>
  );
}
