import { useEffect, useLayoutEffect, useRef, useState, type CSSProperties, type ReactNode, type RefObject } from "react";
import { createPortal } from "react-dom";

export type Placement = "bottom-start" | "bottom-end" | "right-start" | "left-start" | "top-start";

export interface PopoverProps {
  open: boolean;
  onClose: () => void;
  /** Konumlandırmanın bağlanacağı öğe. */
  anchor?: RefObject<HTMLElement | null>;
  /** Öğe yerine sabit nokta (sağ tık menüsü). */
  point?: { x: number; y: number };
  placement?: Placement;
  className?: string;
  children: ReactNode;
  /** Dışarı tıklamada kapanmasın diye hariç tutulacak öğeler. */
  ignore?: RefObject<HTMLElement | null>[];
  style?: CSSProperties;
  role?: string;
}

/**
 * Bir öğeye ya da noktaya bağlı, ekrandan taşmayan açılır katman.
 * Dışarı tıklama ve Escape ile kapanır.
 */
export function Popover({ open, onClose, anchor, point, placement = "bottom-start", className, children, ignore, style, role }: PopoverProps) {
  const ref = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState<{ left: number; top: number } | null>(null);

  useLayoutEffect(() => {
    if (!open) {
      setPos(null);
      return;
    }
    const el = ref.current;
    if (!el) return;
    const w = el.offsetWidth;
    const h = el.offsetHeight;
    const vw = window.innerWidth;
    const vh = window.innerHeight;
    let left = 0;
    let top = 0;
    if (point) {
      left = point.x;
      top = point.y;
      if (left + w > vw - 4) left = Math.max(4, point.x - w);
      if (top + h > vh - 4) top = Math.max(4, point.y - h);
    } else if (anchor?.current) {
      const r = anchor.current.getBoundingClientRect();
      switch (placement) {
        case "bottom-end":
          left = r.right - w;
          top = r.bottom + 1;
          break;
        case "right-start":
          left = r.right - 2;
          top = r.top - 4;
          if (left + w > vw - 4) left = r.left - w + 2;
          break;
        case "left-start":
          left = r.left - w + 2;
          top = r.top - 4;
          break;
        case "top-start":
          left = r.left;
          top = r.top - h - 1;
          break;
        default:
          left = r.left;
          top = r.bottom + 1;
      }
      if (top + h > vh - 4) top = placement === "right-start" ? Math.max(4, vh - h - 4) : Math.max(4, r.top - h - 1);
      if (left + w > vw - 4) left = Math.max(4, vw - w - 4);
      if (left < 4) left = 4;
      if (top < 4) top = 4;
    }
    setPos({ left, top });
  }, [open, point?.x, point?.y, anchor, placement]);

  useEffect(() => {
    if (!open) return;
    const onDown = (e: MouseEvent) => {
      const t = e.target as Node;
      if (ref.current?.contains(t)) return;
      if (anchor?.current?.contains(t)) return;
      if (ignore?.some((r) => r.current?.contains(t))) return;
      // İç içe açılan (portal) alt menüler daha üst katmandadır; onlara tıklamak üst menüyü kapatmaz.
      const el = t instanceof Element ? t : t.parentElement;
      const host = el?.closest("[data-oh-popover]");
      if (host) {
        const myLevel = Number(ref.current?.getAttribute("data-oh-level") ?? 1);
        if (Number(host.getAttribute("data-oh-level") ?? 1) > myLevel) return;
      }
      onClose();
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        e.stopPropagation();
        onClose();
      }
    };
    const onBlur = () => onClose();
    document.addEventListener("mousedown", onDown, true);
    document.addEventListener("keydown", onKey, true);
    window.addEventListener("blur", onBlur);
    return () => {
      document.removeEventListener("mousedown", onDown, true);
      document.removeEventListener("keydown", onKey, true);
      window.removeEventListener("blur", onBlur);
    };
  }, [open, onClose, anchor, ignore]);

  if (!open) return null;
  const level = anchor?.current?.closest("[data-oh-popover]")?.getAttribute("data-oh-level");
  return createPortal(
    <div
      ref={ref}
      data-oh-popover=""
      data-oh-level={level ? Number(level) + 1 : 1}
      role={role}
      className={`oh-popover ${className ?? ""}`}
      style={{ ...style, left: pos?.left ?? -9999, top: pos?.top ?? -9999, visibility: pos ? "visible" : "hidden" }}
      onContextMenu={(e) => e.preventDefault()}
    >
      {children}
    </div>,
    document.body,
  );
}
