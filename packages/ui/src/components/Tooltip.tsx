import { useEffect, useRef, useState } from "react";
import { createPortal } from "react-dom";

/**
 * Ofis tarzı "süper ipucu" katmanı. Herhangi bir öğeye
 *   data-tip="Kalın" data-tip-key="Ctrl+B" data-tip-desc="Metni kalın yapar."
 * vermek yeterlidir; katman AppShell içinde bir kez çizilir.
 */
export function tipProps(title?: string, shortcut?: string, description?: string): Record<string, string> {
  const p: Record<string, string> = {};
  if (title) p["data-tip"] = title;
  if (shortcut) p["data-tip-key"] = shortcut;
  if (description) p["data-tip-desc"] = description;
  return p;
}

interface TipState {
  title: string;
  shortcut?: string;
  desc?: string;
  x: number;
  y: number;
  below: boolean;
}

export function TooltipLayer() {
  const [tip, setTip] = useState<TipState | null>(null);
  const timer = useRef<number | undefined>(undefined);
  const current = useRef<Element | null>(null);

  useEffect(() => {
    const hide = () => {
      window.clearTimeout(timer.current);
      current.current = null;
      setTip(null);
    };
    const onOver = (e: MouseEvent) => {
      const el = (e.target as Element | null)?.closest?.("[data-tip]");
      if (el === current.current) return;
      hide();
      if (!el) return;
      current.current = el;
      timer.current = window.setTimeout(() => {
        if (!el.isConnected) return;
        const r = el.getBoundingClientRect();
        const below = r.bottom + 90 < window.innerHeight;
        setTip({
          title: el.getAttribute("data-tip") ?? "",
          shortcut: el.getAttribute("data-tip-key") ?? undefined,
          desc: el.getAttribute("data-tip-desc") ?? undefined,
          x: Math.min(r.left, window.innerWidth - 280),
          y: below ? r.bottom + 4 : r.top - 4,
          below,
        });
      }, 550);
    };
    document.addEventListener("mouseover", onOver);
    document.addEventListener("mousedown", hide, true);
    document.addEventListener("keydown", hide, true);
    window.addEventListener("blur", hide);
    return () => {
      document.removeEventListener("mouseover", onOver);
      document.removeEventListener("mousedown", hide, true);
      document.removeEventListener("keydown", hide, true);
      window.removeEventListener("blur", hide);
      window.clearTimeout(timer.current);
    };
  }, []);

  if (!tip) return null;
  return createPortal(
    <div
      className="oh-tooltip"
      role="tooltip"
      style={{ left: Math.max(4, tip.x), top: tip.y, transform: tip.below ? undefined : "translateY(-100%)" }}
    >
      <div className="oh-tooltip-title">
        {tip.title}
        {tip.shortcut && <span className="oh-tooltip-key"> ({tip.shortcut})</span>}
      </div>
      {tip.desc && <div className="oh-tooltip-desc">{tip.desc}</div>}
    </div>,
    document.body,
  );
}
