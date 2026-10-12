// Mini araç çubuğu: fareyle metin seçilince seçimin üstünde, sağ tıkta bağlam menüsünün üstünde açılan
// küçük biçimlendirme yüzeyi. Konumlandırma, fareyle uzaklaşınca saydamlaşıp kapanma ve açılıp kapanma
// yaşam döngüsü buradadır; komutlar (içerik) uygulamadadır.
import { useEffect, useLayoutEffect, useRef, useState, type ReactNode } from "react";
import { createPortal } from "react-dom";
import { miniBarPosition, nearestRect, type MiniBarAnchor, type MiniBarRect } from "./MiniToolbarGeometry";

export { miniBarPosition, type MiniBarAnchor, type MiniBarRect } from "./MiniToolbarGeometry";

export interface MiniToolbarTracking {
  /** Fare seçiminin izlendiği düzenleyici kökü (her olayda yeniden okunur). */
  root: () => HTMLElement | null | undefined;
  /** Sol tuş bırakıldıktan sonra seçimin ekran dikdörtgeni; yoksa çubuk açılmaz. */
  selectionRect?: () => DOMRect | null | undefined;
  /** Sağ tıkta (uygulama kendi bağlam menüsünü açtıysa) çubuk açılsın mı? */
  contextTarget?: (target: Element) => boolean;
}

const BAR = ".oh-minibar";
const POPOVER = "[data-oh-popover]";

/** Tıklama noktasında açılmış bağlam menüsü (ilk düzey açılır katman). */
function contextMenuRect(a: MiniBarAnchor, self: Element | null): MiniBarRect | null {
  const rects = Array.from(document.querySelectorAll<HTMLElement>('[data-oh-popover][data-oh-level="1"]'))
    .filter((el) => el !== self)
    .map((el) => el.getBoundingClientRect());
  return nearestRect(a.x, a.y, rects);
}

/** Çubuğun açılıp kapanmasını düzenleyicinin fare ve klavye olaylarına bağlar. */
export function useMiniToolbar(tracking: MiniToolbarTracking) {
  const [anchor, setAnchor] = useState<MiniBarAnchor | null>(null);
  const opts = useRef(tracking);
  opts.current = tracking;

  useEffect(() => {
    let down = false;
    const inRoot = (t: EventTarget | null) => t instanceof Node && !!opts.current.root()?.contains(t);
    const onDown = (e: MouseEvent) => {
      const t = e.target as Element;
      if (t.closest?.(BAR) || t.closest?.(POPOVER)) return;
      if (inRoot(t)) {
        down = e.button === 0;
        setAnchor((a) => (a?.context ? a : null));
      } else {
        down = false;
        setAnchor(null);
      }
    };
    const onUp = (e: MouseEvent) => {
      if (!down || e.button !== 0) return;
      down = false;
      const rectOf = opts.current.selectionRect;
      if (!rectOf) return;
      window.setTimeout(() => {
        const r = rectOf();
        if (!r || (r.width === 0 && r.height === 0)) return;
        setAnchor({ x: Math.max(8, r.left), y: r.top, bottom: r.bottom, context: false });
      }, 0);
    };
    const onKey = (e: KeyboardEvent) => {
      if (inRoot(e.target)) setAnchor(null);
    };
    // Yalnız belge kaydırılınca kapan; çubuğun kendi açılır listeleri (yazı tipi, boyut) kaydırılınca açık kalsın.
    const onScroll = (e: Event) => {
      const t = e.target;
      if (t instanceof Element && (t.closest(POPOVER) || t.closest(BAR))) return;
      setAnchor((a) => (a && !a.context ? null : a));
    };
    // Bağlam menüsü kapanınca (Esc ya da komut) çubuk da kapanır.
    const afterInput = (e: Event) => {
      if (e instanceof KeyboardEvent && e.key === "Escape") return setAnchor(null);
      window.setTimeout(() => setAnchor((a) => (a?.context && !document.querySelector(POPOVER) ? null : a)), 60);
    };
    // Uygulama sağ tıkta kendi menüsünü açtıysa (olayı engellediyse) çubuk menünün üstünde açılır.
    const onContext = (e: MouseEvent) => {
      const pick = opts.current.contextTarget;
      if (!pick || !e.defaultPrevented || !(e.target instanceof Element) || !pick(e.target)) return;
      setAnchor({ x: e.clientX, y: e.clientY, context: true });
    };
    document.addEventListener("mousedown", onDown, true);
    window.addEventListener("mouseup", onUp);
    document.addEventListener("keydown", onKey, true);
    document.addEventListener("scroll", onScroll, true);
    document.addEventListener("keyup", afterInput, true);
    document.addEventListener("mouseup", afterInput, true);
    document.addEventListener("contextmenu", onContext);
    return () => {
      document.removeEventListener("mousedown", onDown, true);
      window.removeEventListener("mouseup", onUp);
      document.removeEventListener("keydown", onKey, true);
      document.removeEventListener("scroll", onScroll, true);
      document.removeEventListener("keyup", afterInput, true);
      document.removeEventListener("mouseup", afterInput, true);
      document.removeEventListener("contextmenu", onContext);
    };
  }, []);

  return { anchor, setAnchor, close: () => setAnchor(null) };
}

export interface MiniToolbarProps {
  anchor: MiniBarAnchor | null;
  onClose: () => void;
  children: ReactNode;
  /** Ek sınıf (uygulamaya özgü seçiciler için). */
  className?: string;
  label?: string;
  /**
   * Çubuk bir açılır katmanın (bağlam menüsünün) parçasıysa katman düzeyi: menünün dışarı tıklama
   * denetimi çubuğu kendi içi sayar, çubuğun açılır listeleri bir üst düzeyde açılır.
   */
  layer?: number;
  /** Fareyle uzaklaşınca saydamlaşıp kapanma (yalnız seçim kipinde). */
  fade?: boolean;
}

/** Fare çubuktan bu kadar (px) daha uzaklaşınca çubuk kapanır; yol boyunca saydamlaşır. */
const FADE_PX = 160;
const NEAR_PX = 24;

/** Konumlanan cam yüzey (document.body'ye çizilir). */
export function MiniToolbar({ anchor, onClose, children, className, label = "Mini araç çubuğu", layer, fade = true }: MiniToolbarProps) {
  const ref = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState<{ left: number; top: number } | null>(null);
  const closeRef = useRef(onClose);
  closeRef.current = onClose;

  useLayoutEffect(() => {
    if (!anchor) return setPos(null);
    const place = () => {
      const el = ref.current;
      const size = { w: el?.offsetWidth ?? 360, h: el?.offsetHeight ?? 70 };
      const menu = anchor.context ? contextMenuRect(anchor, el) : null;
      setPos(miniBarPosition(anchor, size, { w: window.innerWidth, h: window.innerHeight }, menu));
    };
    place();
    // Bağlam menüsü kendi konumunu (taşmaya göre) bir sonraki çizimde bulur: çubuk ona göre yeniden yerleşir.
    if (!anchor.context) return;
    const id = requestAnimationFrame(place);
    return () => cancelAnimationFrame(id);
  }, [anchor]);

  // Seçim kipinde fare çubuktan uzaklaştıkça saydamlaşır, yeterince uzaklaşınca kapanır.
  useEffect(() => {
    const el = ref.current;
    if (!anchor || anchor.context || !fade || !el) return;
    let base = -1;
    const onMove = (e: MouseEvent) => {
      // Açılır liste açıkken ya da çubukta yazılırken (yazı tipi kutusu) saydamlaşmaz.
      if (document.querySelector(POPOVER) || el.contains(document.activeElement)) {
        el.style.opacity = "";
        return;
      }
      const r = el.getBoundingClientRect();
      const d = Math.max(NEAR_PX, Math.hypot(Math.max(r.left - e.clientX, 0, e.clientX - r.right), Math.max(r.top - e.clientY, 0, e.clientY - r.bottom)));
      // Eşik farenin çubuğa en çok yaklaştığı uzaklıktır (çubuk açıldığında fare seçimin ucunda olabilir).
      base = base < 0 ? d : Math.min(base, d);
      const t = (d - base) / FADE_PX;
      if (t >= 1) return closeRef.current();
      el.style.opacity = t <= 0 ? "" : String(1 - t * 0.85);
    };
    document.addEventListener("mousemove", onMove, { passive: true });
    return () => {
      document.removeEventListener("mousemove", onMove);
      el.style.opacity = "";
    };
  }, [anchor, fade]);

  if (!anchor) return null;
  return createPortal(
    <div
      ref={ref}
      className={`oh-minibar${className ? ` ${className}` : ""}`}
      role="toolbar"
      aria-label={label}
      {...(layer ? { "data-oh-popover": "", "data-oh-level": layer } : {})}
      style={pos ? { left: pos.left, top: pos.top } : { left: -9999, top: 0 }}
      onMouseDown={(e) => e.target instanceof HTMLInputElement || e.preventDefault()}
      onContextMenu={(e) => e.preventDefault()}
    >
      {children}
    </div>,
    document.body,
  );
}

export function MiniToolbarRow({ children }: { children: ReactNode }) {
  return <div className="oh-minibar-row">{children}</div>;
}

export interface MiniToolbarHostProps extends MiniToolbarTracking {
  /** Çubuğun içeriği; `close` bağlam kipindeki komutlardan sonra çağrılır. */
  children: (ctx: { anchor: MiniBarAnchor; close: () => void }) => ReactNode;
  /** Dışarıdan açmak için (ör. uygulamanın bağlam menüsü işleyicisi). */
  controlRef?: { current: ((a: MiniBarAnchor | null) => void) | null };
  className?: string;
  label?: string;
}

/** Yaşam döngüsü ve yüzey bir arada: düzenleyicinin yanında bir kez çizilir. */
export function MiniToolbarHost({ children, controlRef, className, label, ...tracking }: MiniToolbarHostProps) {
  const { anchor, setAnchor, close } = useMiniToolbar(tracking);
  if (controlRef) controlRef.current = setAnchor;
  return (
    <MiniToolbar anchor={anchor} onClose={close} className={className} label={label}>
      {anchor && children({ anchor, close })}
    </MiniToolbar>
  );
}
