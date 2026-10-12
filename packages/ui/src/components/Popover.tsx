import { useEffect, useLayoutEffect, useRef, useState, type CSSProperties, type ReactNode, type RefObject } from "react";
import { createPortal } from "react-dom";
import { createMenuNav, focusBeforeLastPointer, focusOnOpen as focusLayer, lastInputWasKeyboard, layerInfo, prepareLayer } from "./menuKeyboard";

export type Placement = "bottom-start" | "bottom-end" | "right-start" | "left-start" | "top-start";

/** Kapanma nedeni: Esc, dışarı tıklama, pencere odağı kaybı, alt menüde ←, saf menüde Tab. */
export type PopoverCloseReason = "escape" | "outside" | "blur" | "arrow" | "tab";

export interface PopoverProps {
  open: boolean;
  onClose: (reason?: PopoverCloseReason) => void;
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
  /**
   * Açılınca odak katmana taşınsın mı. "auto" (varsayılan): içerik menüyse; "keyboard": yalnız
   * klavyeyle açıldıysa; true/false: her zaman/hiçbir zaman. Klavyeyle açılınca ilk öğeye,
   * fareyle açılınca vurgusuz olarak kaba odaklanılır. Kapanınca odak eski yerine döner.
   */
  focusOnOpen?: "auto" | "keyboard" | boolean;
}

/* ------------------------------------------------------------------ */
/* Açık katman yığını: Esc yalnız en üsttekini kapatır.                  */
/* ------------------------------------------------------------------ */

interface Layer {
  el: HTMLElement | null;
  close: (reason: PopoverCloseReason) => void;
  anchor: HTMLElement | null;
  /** Alt menüyse onu açan üst menü öğesi. */
  parentItem: HTMLElement | null;
  /** Kapanınca odağın döneceği öğe (alt menüde üst öğe, değilse açılıştaki odak). */
  restoreTo: HTMLElement | null;
}

const layers: Layer[] = [];
let escInstalled = false;

function installEscape() {
  if (escInstalled) return;
  escInstalled = true;
  document.addEventListener(
    "keydown",
    (e) => {
      if (e.key !== "Escape" || !layers.length) return;
      e.stopPropagation();
      layers[layers.length - 1].close("escape");
    },
    true,
  );
}

/** Alt menü zinciri (en üstten başlayarak, alt menü olmayan ilk katmana kadar). */
function menuChain(layer: Layer): Layer[] {
  const chain = [layer];
  for (let i = layers.indexOf(layer) - 1, cur = layer; i >= 0 && cur.parentItem; i--) {
    if (!layers[i].el?.contains(cur.parentItem)) break;
    cur = layers[i];
    chain.push(cur);
  }
  return chain;
}

/**
 * Bir öğeye ya da noktaya bağlı, ekrandan taşmayan açılır katman.
 * Dışarı tıklama ve Escape ile kapanır; menü içeriğinde klavye gezintisi sağlar.
 */
export function Popover(props: PopoverProps) {
  if (!props.open) return null;
  return createPortal(<PopoverLayer {...props} />, document.body);
}

function PopoverLayer({ onClose, anchor, point, placement = "bottom-start", className, children, ignore, style, role, focusOnOpen = "auto" }: PopoverProps) {
  const ref = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState<{ left: number; top: number } | null>(null);
  const closeRef = useRef(onClose);
  closeRef.current = onClose;
  const layerRef = useRef<Layer | null>(null);
  const [nav] = useState(createMenuNav);

  // Katmanı yığına ekle; kapanırken (DOM henüz yerindeyken) odak içerideyse eski yerine döndür.
  useLayoutEffect(() => {
    installEscape();
    const a = anchor?.current ?? null;
    const ae = document.activeElement;
    const parentItem = a?.closest("[data-oh-popover]") && a.matches('[role^="menuitem"]') ? a : null;
    const layer: Layer = {
      el: ref.current,
      close: (r) => closeRef.current(r),
      anchor: a,
      parentItem,
      restoreTo: parentItem ?? (ae instanceof HTMLElement && ae !== document.body ? ae : focusBeforeLastPointer()),
    };
    layers.push(layer);
    layerRef.current = layer;
    return () => {
      const i = layers.indexOf(layer);
      if (i >= 0) layers.splice(i, 1);
      const now = document.activeElement;
      const inside = !!now && [layer, ...layers.slice(Math.max(0, i))].some((l) => l.el?.contains(now));
      const to = layer.restoreTo;
      if (inside && to?.isConnected && !to.matches(":disabled")) to.focus({ preventScroll: true });
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useLayoutEffect(() => {
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
  }, [point?.x, point?.y, anchor, placement]);

  // Görünür olunca (konum hesaplandıktan sonra) bir kez: menü rolü ve açılış odağı.
  const shown = pos !== null;
  useLayoutEffect(() => {
    const el = ref.current;
    if (!shown || !el) return;
    const info = layerInfo(el);
    prepareLayer(info);
    const kbd = lastInputWasKeyboard();
    const want = focusOnOpen === "auto" ? info.kind !== null : focusOnOpen === "keyboard" ? kbd : focusOnOpen;
    if (want) focusLayer(el, info, kbd ? "first" : "container");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [shown]);

  useEffect(() => {
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
      closeRef.current("outside");
    };
    const onBlur = () => closeRef.current("blur");
    document.addEventListener("mousedown", onDown, true);
    window.addEventListener("blur", onBlur);
    return () => {
      document.removeEventListener("mousedown", onDown, true);
      window.removeEventListener("blur", onBlur);
    };
  }, [anchor, ignore]);

  const level = anchor?.current?.closest("[data-oh-popover]")?.getAttribute("data-oh-level");
  return (
    <div
      ref={ref}
      data-oh-popover=""
      data-oh-level={level ? Number(level) + 1 : 1}
      role={role}
      className={`oh-popover ${className ?? ""}`}
      style={{ ...style, left: pos?.left ?? -9999, top: pos?.top ?? -9999, visibility: pos ? "visible" : "hidden" }}
      onContextMenu={(e) => e.preventDefault()}
      onMouseMove={(e) => ref.current && nav.onMouseMove(e, ref.current)}
      onKeyDown={(e) => {
        const el = ref.current;
        const layer = layerRef.current;
        if (!el || !layer) return;
        nav.onKeyDown(e, el, {
          submenu: !!layer.parentItem,
          tabLoop: focusOnOpen === "keyboard" || focusOnOpen === true,
          close: (r) => {
            if (r !== "tab") return layer.close(r);
            for (const l of menuChain(layer)) l.close(r);
          },
          tabTarget: () => {
            const chain = menuChain(layer);
            const root = chain[chain.length - 1];
            const to = root.restoreTo;
            return { el: to, opener: !!to && !!root.anchor?.contains(to) };
          },
        });
      }}
    >
      {children}
    </div>
  );
}
