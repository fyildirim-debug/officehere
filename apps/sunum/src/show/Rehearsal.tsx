/**
 * Zamanlama Provası: gösteri sırasında her slaytta geçen süre ölçülür
 * (duraklatılabilir, slayt süresi yinelenebilir). Gösteri bitince süreler
 * slaytların "Sonra" (otomatik ilerleme) zamanlaması olarak kaydedilebilir.
 */
import { useCallback, useEffect, useRef, useState } from "react";
import { formatTime } from "./showKeys";

export interface RehearsalResult {
  /** Slayt kimliği → saniye. */
  times: Record<string, number>;
  /** Toplam saniye. */
  total: number;
}

export interface Rehearsal {
  slideMs: number;
  totalMs: number;
  paused: boolean;
  togglePause: () => void;
  repeat: () => void;
  result: () => RehearsalResult;
}

export function useRehearsal(active: boolean, slideId: string | undefined): Rehearsal {
  const acc = useRef<Record<string, number>>({});
  const cur = useRef<{ id: string | undefined; start: number }>({ id: slideId, start: Date.now() });
  const [pausedAt, setPausedAt] = useState<number | null>(null);
  const pausedRef = useRef<number | null>(null);
  pausedRef.current = pausedAt;
  const [now, setNow] = useState(() => Date.now());

  useEffect(() => {
    if (!active) return;
    const t = window.setInterval(() => setNow(Date.now()), 250);
    return () => window.clearInterval(t);
  }, [active]);

  const commit = useCallback(() => {
    const c = cur.current;
    if (!c.id) return;
    const end = pausedRef.current ?? Date.now();
    acc.current[c.id] = (acc.current[c.id] ?? 0) + Math.max(0, end - c.start);
    c.start = end;
  }, []);

  // Slayt değişince önceki slaydın süresi eklenir.
  useEffect(() => {
    if (!active || cur.current.id === slideId) return;
    commit();
    cur.current = { id: slideId, start: pausedRef.current ?? Date.now() };
  }, [active, slideId, commit]);

  const t = pausedAt ?? now;
  const slideMs = Math.max(0, t - cur.current.start);
  const committed = Object.values(acc.current).reduce((s, v) => s + v, 0);
  return {
    slideMs,
    totalMs: committed + slideMs,
    paused: pausedAt !== null,
    togglePause: () => {
      if (pausedAt !== null) {
        // Duraklatılan süre sayılmaz.
        cur.current.start += Date.now() - pausedAt;
        setPausedAt(null);
      } else setPausedAt(Date.now());
    },
    repeat: () => {
      cur.current.start = pausedAt ?? Date.now();
      setNow(Date.now());
    },
    result: () => {
      commit();
      const times: Record<string, number> = {};
      let total = 0;
      for (const [id, ms] of Object.entries(acc.current)) {
        times[id] = Math.max(0, Math.round(ms / 1000));
        total += ms;
      }
      return { times, total: Math.round(total / 1000) };
    },
  };
}

/** Sol üstteki "Kayıt" araç çubuğu. */
export function RehearsalBar({ r, onNext }: { r: Rehearsal; onNext: () => void }) {
  return (
    <div className="rehearse-bar" role="toolbar" aria-label="Kayıt" onClick={(e) => e.stopPropagation()} onPointerDown={(e) => e.stopPropagation()}>
      <span className="rehearse-title">Kayıt</span>
      <button type="button" className="rehearse-btn" title="Sonraki" aria-label="Sonraki" onClick={onNext}>
        ➜
      </button>
      <button type="button" className="rehearse-btn" title={r.paused ? "Kayda Devam Et" : "Kaydı Duraklat"} aria-label={r.paused ? "Kayda Devam Et" : "Kaydı Duraklat"} onClick={r.togglePause}>
        {r.paused ? "▶" : "❚❚"}
      </button>
      <span className="rehearse-time" aria-label="Slayt süresi">
        {formatTime(r.slideMs, true)}
      </span>
      <button type="button" className="rehearse-btn" title="Yinele" aria-label="Yinele" onClick={r.repeat}>
        ↺
      </button>
      <span className="rehearse-total" aria-label="Toplam süre">
        {formatTime(r.totalMs, true)}
      </span>
    </div>
  );
}
