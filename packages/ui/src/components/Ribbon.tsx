import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useId,
  useLayoutEffect,
  useRef,
  useState,
  type KeyboardEvent as ReactKeyboardEvent,
  type ReactElement,
  type ReactNode,
} from "react";
import { Popover } from "./Popover";
import { MenuProvider } from "./Menu";
import { tipProps } from "./Tooltip";
import { focusEl, tabbables } from "./menuKeyboard";
import { nextEnabled } from "./menuNav";

/* ------------------------------------------------------------------ */
/* Şerit (Ribbon)                                                      */
/* ------------------------------------------------------------------ */

export interface RibbonTabDef {
  id: string;
  label: string;
  content: ReactNode;
  /** Bağlamsal sekme (ör. "Tablo Tasarımı"); vurgu rengiyle gösterilir. */
  contextual?: boolean;
  /** Bağlamsal sekme grubunun başlığı (ör. "Tablo Araçları"). */
  contextualGroup?: string;
  hidden?: boolean;
}

export interface RibbonProps {
  tabs: RibbonTabDef[];
  activeTab: string;
  onTabChange: (id: string) => void;
  /** "Dosya" sekmesi tıklanınca (Backstage açılır). */
  onFileClick: () => void;
  fileLabel?: string;
  /** Sekme satırının sağına (Paylaş, Açıklamalar vb.). */
  tabsRight?: ReactNode;
  collapsed?: boolean;
  onCollapsedChange?: (collapsed: boolean) => void;
}

interface LayoutCtx {
  groupCount: number;
  collapseCount: number;
  version: number;
}

const RibbonLayoutContext = createContext<LayoutCtx>({ groupCount: 0, collapseCount: 0, version: 0 });

export function Ribbon({ tabs, activeTab, onTabChange, onFileClick, fileLabel = "Dosya", tabsRight, collapsed = false, onCollapsedChange }: RibbonProps) {
  const visible = tabs.filter((t) => !t.hidden);
  const active = visible.find((t) => t.id === activeTab) ?? visible[0];
  const bodyRef = useRef<HTMLDivElement>(null);
  const [layout, setLayout] = useState<LayoutCtx>({ groupCount: 0, collapseCount: 0, version: 0 });
  const [floating, setFloating] = useState(false);
  const [width, setWidth] = useState(0);
  const ribbonRef = useRef<HTMLDivElement>(null);
  const tabRefs = useRef(new Map<string, HTMLButtonElement>());
  const pendingBodyFocus = useRef(false);
  const uid = useId();

  // Genişlik değişince küçültmeyi sıfırla ve yeniden ölç.
  useEffect(() => {
    const el = bodyRef.current;
    if (!el) return;
    const ro = new ResizeObserver(() => setWidth(el.clientWidth));
    ro.observe(el);
    return () => ro.disconnect();
  }, [collapsed, floating]);

  useLayoutEffect(() => {
    setLayout((l) => ({ ...l, collapseCount: 0, version: l.version + 1 }));
  }, [width, active?.id]);

  useLayoutEffect(() => {
    const el = bodyRef.current;
    if (!el) return;
    const inner = el.firstElementChild as HTMLElement | null;
    const groups = el.querySelectorAll(":scope > .oh-ribbon-content > .oh-ribbon-group").length;
    if (groups !== layout.groupCount) {
      setLayout((l) => ({ ...l, groupCount: groups }));
      return;
    }
    if (inner && inner.scrollWidth > el.clientWidth + 1 && layout.collapseCount < groups) {
      setLayout((l) => ({ ...l, collapseCount: l.collapseCount + 1 }));
    }
  });

  const showBody = !collapsed || floating;

  useEffect(() => {
    if (!floating) return;
    const onDown = (e: MouseEvent) => {
      const t = e.target as Element;
      if (ribbonRef.current?.contains(t) || t.closest?.("[data-oh-popover]")) return;
      setFloating(false);
    };
    document.addEventListener("mousedown", onDown, true);
    return () => document.removeEventListener("mousedown", onDown, true);
  }, [floating]);

  /** Şerit gövdesindeki ilk denetime iner (↓ ya da Tab). */
  const focusBody = () => {
    const first = bodyRef.current && tabbables(bodyRef.current)[0];
    if (!first) return false;
    focusEl(first, bodyRef.current ?? undefined);
    return true;
  };

  useEffect(() => {
    if (floating && pendingBodyFocus.current) focusBody();
    pendingBodyFocus.current = false;
  }, [floating]);

  // Sekmeler: ←/→ odak ve seçim birlikte (uçlarda döner), Home/End; ↓ ya da Tab gövdeye iner.
  const onTabKey = (e: ReactKeyboardEvent, i: number) => {
    let next = -1;
    switch (e.key) {
      case "ArrowRight":
      case "ArrowLeft":
        next = nextEnabled(visible.length, i, e.key === "ArrowRight" ? 1 : -1);
        break;
      case "Home":
        next = 0;
        break;
      case "End":
        next = visible.length - 1;
        break;
      case "ArrowDown":
      case "Tab":
        if (e.key === "Tab" && e.shiftKey) return;
        if (collapsed && !floating) {
          e.preventDefault();
          pendingBodyFocus.current = true;
          setFloating(true);
        } else if (focusBody()) e.preventDefault();
        return;
      default:
        return;
    }
    const t = visible[next];
    if (!t) return;
    e.preventDefault();
    if (t.id !== active?.id) onTabChange(t.id);
    tabRefs.current.get(t.id)?.focus();
  };

  let lastGroup: string | undefined;
  return (
    <div ref={ribbonRef} className={`oh-ribbon${collapsed ? " is-collapsed" : ""}${floating ? " is-floating" : ""}`}>
      <div className="oh-ribbon-tabs">
        <button type="button" className="oh-ribbon-tab oh-ribbon-file" onClick={onFileClick}>
          {fileLabel}
        </button>
        <div className="oh-ribbon-tablist" role="tablist" aria-label="Şerit">
          {visible.map((t, i) => {
            const showGroup = t.contextual && t.contextualGroup && t.contextualGroup !== lastGroup;
            lastGroup = t.contextual ? t.contextualGroup : undefined;
            const selected = t.id === active?.id;
            return (
              <button
                key={t.id}
                ref={(el) => {
                  if (el) tabRefs.current.set(t.id, el);
                  else tabRefs.current.delete(t.id);
                }}
                type="button"
                role="tab"
                id={`${uid}-tab-${t.id}`}
                aria-selected={selected}
                aria-controls={selected && showBody ? `${uid}-panel` : undefined}
                tabIndex={selected ? 0 : -1}
                className={`oh-ribbon-tab${selected && showBody ? " is-active" : ""}${t.contextual ? " is-contextual" : ""}`}
                data-group={showGroup ? t.contextualGroup : undefined}
                onKeyDown={(e) => onTabKey(e, i)}
                // Fareyle sekme değiştirmek klavye odağını belgeden/ızgaradan almaz (yazma ve
                // ok tuşları belgeye gitmeye devam eder); odak zaten şeritteyse olağan davranış.
                onMouseDown={(e) => {
                  const a = document.activeElement;
                  if (e.button === 0 && a && a !== document.body && !e.currentTarget.closest(".oh-ribbon")?.contains(a)) e.preventDefault();
                }}
                onClick={() => {
                  onTabChange(t.id);
                  if (collapsed) setFloating(!(floating && t.id === active?.id));
                }}
                onDoubleClick={() => {
                  setFloating(false);
                  onCollapsedChange?.(!collapsed);
                }}
              >
                {t.label}
              </button>
            );
          })}
        </div>
        <div className="oh-ribbon-tabs-spacer" />
        {tabsRight}
      </div>
      {showBody && active && (
        <RibbonLayoutContext.Provider value={layout}>
          <div
            className="oh-ribbon-body"
            ref={bodyRef}
            role="tabpanel"
            id={`${uid}-panel`}
            aria-labelledby={`${uid}-tab-${active.id}`}
            onKeyDown={(e) => {
              // Daraltılmış şeridin açılan gövdesinde Esc gövdeyi kapatıp odağı sekmeye verir.
              if (e.key !== "Escape" || !floating || e.defaultPrevented) return;
              setFloating(false);
              tabRefs.current.get(active.id)?.focus();
            }}
          >
            <div className="oh-ribbon-content" key={active.id}>
              {active.content}
            </div>
            {onCollapsedChange && (
              <button
                type="button"
                className="oh-ribbon-pin"
                {...tipProps(collapsed ? "Şeridi Sabitle" : "Şeridi Daralt", "Ctrl+F1")}
                onClick={() => {
                  setFloating(false);
                  onCollapsedChange(!collapsed);
                }}
              >
                <svg width="12" height="12" viewBox="0 0 12 12" aria-hidden="true">
                  {collapsed ? (
                    <path d="M2 4.5l4 4 4-4" fill="none" stroke="currentColor" strokeWidth="1.2" />
                  ) : (
                    <path d="M2 7.5l4-4 4 4" fill="none" stroke="currentColor" strokeWidth="1.2" />
                  )}
                </svg>
              </button>
            )}
          </div>
        </RibbonLayoutContext.Provider>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Grup                                                                */
/* ------------------------------------------------------------------ */

export interface RibbonGroupProps {
  label: string;
  children: ReactNode;
  /** Grup sağ altındaki iletişim kutusu başlatıcısı. */
  onLauncher?: () => void;
  launcherTip?: string;
  /** Pencere daralınca grubun tek düğmeye dönüştüğündeki simgesi. */
  icon?: ReactElement;
}

export function RibbonGroup({ label, children, onLauncher, launcherTip, icon }: RibbonGroupProps) {
  const layout = useContext(RibbonLayoutContext);
  const ref = useRef<HTMLDivElement>(null);
  const btnRef = useRef<HTMLButtonElement>(null);
  const [index, setIndex] = useState(-1);
  const [open, setOpen] = useState(false);

  useLayoutEffect(() => {
    const el = ref.current;
    if (!el?.parentElement) return;
    const siblings = Array.from(el.parentElement.children).filter((c) => c.classList.contains("oh-ribbon-group"));
    const i = siblings.indexOf(el);
    if (i !== index) setIndex(i);
  }, [layout.version, layout.groupCount, index]);

  const collapsed = index >= 0 && layout.collapseCount > 0 && index >= layout.groupCount - layout.collapseCount;
  const close = useCallback(() => setOpen(false), []);

  const launcher = onLauncher && (
    <button type="button" className="oh-ribbon-launcher" onClick={onLauncher} {...tipProps(launcherTip ?? `${label} Ayarları`)}>
      <svg width="10" height="10" viewBox="0 0 10 10" aria-hidden="true">
        <path d="M1 1h4M1 1v4M9 9L3.5 3.5M9 9V5.5M9 9H5.5" fill="none" stroke="currentColor" strokeWidth="1" />
      </svg>
    </button>
  );

  if (collapsed) {
    return (
      <div ref={ref} className="oh-ribbon-group is-collapsed">
        <button
          ref={btnRef}
          type="button"
          className={`oh-rbtn oh-rbtn-large${open ? " is-open" : ""}`}
          aria-haspopup="true"
          aria-expanded={open}
          onClick={() => setOpen((o) => !o)}
          onKeyDown={(e) => openOnArrow(e, open, () => setOpen(true))}
        >
          <span className="oh-rbtn-icon">{icon ?? <span className="oh-ribbon-group-dot" />}</span>
          <span className="oh-rbtn-label">
            {label}
            <Chevron />
          </span>
        </button>
        {/* Klavyeyle açılınca odak ilk denetime geçer, Tab içeride döner; Esc odağı düğmeye verir. */}
        <Popover open={open} onClose={close} anchor={btnRef} focusOnOpen="keyboard">
          <MenuProvider onClose={close}>
            <div className="oh-ribbon-group oh-ribbon-group-popup">
              <div className="oh-ribbon-group-items">{children}</div>
              <div className="oh-ribbon-group-label">
                {label}
                {launcher}
              </div>
            </div>
          </MenuProvider>
        </Popover>
      </div>
    );
  }

  return (
    <div ref={ref} className="oh-ribbon-group" role="group" aria-label={label}>
      <div className="oh-ribbon-group-items">{children}</div>
      <div className="oh-ribbon-group-label">
        {label}
        {launcher}
      </div>
    </div>
  );
}

/** Küçük denetimleri alt alta (en çok 3 satır) dizer. */
export function RibbonStack({ children, align = "start" }: { children: ReactNode; align?: "start" | "center" }) {
  return <div className={`oh-ribbon-stack is-${align}`}>{children}</div>;
}

/** Yığın içinde yan yana denetim satırı (ör. Kalın/İtalik/Altı Çizili). */
export function RibbonRow({ children, joined }: { children: ReactNode; joined?: boolean }) {
  return <div className={`oh-ribbon-row${joined ? " is-joined" : ""}`}>{children}</div>;
}

export function RibbonSeparator() {
  return <div className="oh-ribbon-sep" />;
}

export function RibbonLabel({ children, width }: { children: ReactNode; width?: number }) {
  return (
    <span className="oh-ribbon-text" style={width ? { width } : undefined}>
      {children}
    </span>
  );
}

/* ------------------------------------------------------------------ */
/* Düğmeler                                                            */
/* ------------------------------------------------------------------ */

export type RibbonButtonSize = "large" | "medium" | "small";

interface BaseButtonProps {
  icon?: ReactElement;
  label?: string;
  size?: RibbonButtonSize;
  disabled?: boolean;
  /** Basılı (aç/kapa) durum. */
  checked?: boolean;
  /** İpucu başlığı (yoksa etiket). */
  tip?: string;
  shortcut?: string;
  description?: string;
  className?: string;
}

export interface RibbonButtonProps extends BaseButtonProps {
  onClick?: () => void;
  /** Fare basılıyken odak editörden çıkmasın (varsayılan: true). */
  preventFocus?: boolean;
}

function Chevron() {
  return (
    <svg className="oh-chevron" width="8" height="8" viewBox="0 0 8 8" aria-hidden="true">
      <path d="M1 2.5l3 3 3-3" fill="none" stroke="currentColor" strokeWidth="1.1" />
    </svg>
  );
}

function sizeClass(size: RibbonButtonSize) {
  return `oh-rbtn oh-rbtn-${size}`;
}

function preventMouseFocus(e: React.MouseEvent) {
  e.preventDefault();
}

/** Menü düğmesinde ↓ (ya da Alt+↓) menüyü açar; Enter/Space doğal tıklamayla açar. */
function openOnArrow(e: ReactKeyboardEvent, open: boolean, openMenu: () => void) {
  if (e.key !== "ArrowDown" || open || e.ctrlKey || e.metaKey) return;
  e.preventDefault();
  openMenu();
}

export function RibbonButton({
  icon,
  label,
  size = "medium",
  disabled,
  checked,
  tip,
  shortcut,
  description,
  onClick,
  className,
  preventFocus = true,
}: RibbonButtonProps) {
  return (
    <button
      type="button"
      className={`${sizeClass(size)}${checked ? " is-checked" : ""} ${className ?? ""}`}
      disabled={disabled}
      aria-pressed={checked}
      aria-label={label ?? tip}
      onMouseDown={preventFocus ? preventMouseFocus : undefined}
      onClick={onClick}
      {...tipProps(tip ?? label, shortcut, description)}
    >
      {icon && <span className="oh-rbtn-icon">{icon}</span>}
      {label && size !== "small" && <span className="oh-rbtn-label">{label}</span>}
    </button>
  );
}

type MenuContent = ReactNode | ((close: () => void) => ReactNode);

function renderMenu(menu: MenuContent, close: () => void) {
  return <MenuProvider onClose={close}>{typeof menu === "function" ? menu(close) : menu}</MenuProvider>;
}

export interface RibbonMenuButtonProps extends BaseButtonProps {
  menu: MenuContent;
  /** Menü açılınca (ör. içeriği tazelemek için). */
  onOpen?: () => void;
}

/** Tıklanınca menü açan düğme. */
export function RibbonMenuButton({ icon, label, size = "medium", disabled, checked, tip, shortcut, description, menu, onOpen, className }: RibbonMenuButtonProps) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLButtonElement>(null);
  const close = useCallback(() => setOpen(false), []);
  return (
    <>
      <button
        ref={ref}
        type="button"
        className={`${sizeClass(size)} has-menu${open ? " is-open" : ""}${checked ? " is-checked" : ""} ${className ?? ""}`}
        disabled={disabled}
        aria-haspopup="menu"
        aria-expanded={open}
        aria-label={label ?? tip}
        onMouseDown={preventMouseFocus}
        onClick={() => {
          if (!open) onOpen?.();
          setOpen(!open);
        }}
        onKeyDown={(e) =>
          openOnArrow(e, open, () => {
            onOpen?.();
            setOpen(true);
          })
        }
        {...tipProps(tip ?? label, shortcut, description)}
      >
        {icon && <span className="oh-rbtn-icon">{icon}</span>}
        {size === "large" ? (
          <span className="oh-rbtn-label">
            {label}
            <Chevron />
          </span>
        ) : (
          <>
            {label && size !== "small" && <span className="oh-rbtn-label">{label}</span>}
            <Chevron />
          </>
        )}
      </button>
      <Popover open={open} onClose={close} anchor={ref}>
        {renderMenu(menu, close)}
      </Popover>
    </>
  );
}

export interface RibbonSplitButtonProps extends BaseButtonProps {
  onClick: () => void;
  menu: MenuContent;
  /** Simgenin altına çizilen renk şeridi (yazı tipi rengi vb.). */
  colorBar?: string | null;
}

/** Ana eylem + açılır menüden oluşan bölünmüş düğme. */
export function RibbonSplitButton({ icon, label, size = "medium", disabled, checked, tip, shortcut, description, onClick, menu, colorBar, className }: RibbonSplitButtonProps) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);
  const close = useCallback(() => setOpen(false), []);
  const iconEl = icon && (
    <span className="oh-rbtn-icon">
      {icon}
      {colorBar !== undefined && <span className="oh-color-bar" style={{ background: colorBar ?? "transparent" }} />}
    </span>
  );
  return (
    <>
      <div
        ref={ref}
        className={`oh-split oh-split-${size}${open ? " is-open" : ""}${checked ? " is-checked" : ""}${disabled ? " is-disabled" : ""} ${className ?? ""}`}
        {...tipProps(tip ?? label, shortcut, description)}
      >
        <button
          type="button"
          className="oh-split-main"
          disabled={disabled}
          aria-label={label ?? tip}
          onMouseDown={preventMouseFocus}
          onClick={onClick}
          onKeyDown={(e) => openOnArrow(e, open, () => setOpen(true))}
        >
          {iconEl}
          {label && size === "medium" && <span className="oh-rbtn-label">{label}</span>}
        </button>
        <button
          type="button"
          className="oh-split-arrow"
          disabled={disabled}
          aria-haspopup="menu"
          aria-expanded={open}
          aria-label={`${label ?? tip ?? ""} seçenekleri`}
          onMouseDown={preventMouseFocus}
          onClick={() => setOpen(!open)}
          onKeyDown={(e) => openOnArrow(e, open, () => setOpen(true))}
        >
          {size === "large" && label && <span className="oh-rbtn-label">{label}</span>}
          <Chevron />
        </button>
      </div>
      <Popover open={open} onClose={close} anchor={ref}>
        {renderMenu(menu, close)}
      </Popover>
    </>
  );
}

/* ------------------------------------------------------------------ */
/* Açılan kutu (yazı tipi, boyut): RibbonCombo.tsx                      */
/* ------------------------------------------------------------------ */

export { RibbonCombo, type ComboOption, type RibbonComboProps } from "./RibbonCombo";

/* ------------------------------------------------------------------ */
/* Onay kutusu ve sayı kutusu                                           */
/* ------------------------------------------------------------------ */

export function RibbonCheckbox({ label, checked, onChange, disabled }: { label: string; checked: boolean; onChange: (v: boolean) => void; disabled?: boolean }) {
  return (
    <label className={`oh-rcheck${disabled ? " is-disabled" : ""}`}>
      <input type="checkbox" checked={checked} disabled={disabled} onChange={(e) => onChange(e.target.checked)} />
      <span>{label}</span>
    </label>
  );
}

export interface RibbonSpinnerProps {
  label?: string;
  icon?: ReactElement;
  value: number;
  onChange: (v: number) => void;
  step?: number;
  min?: number;
  max?: number;
  unit?: string;
  width?: number;
  tip?: string;
  decimals?: number;
}

export function RibbonSpinner({ label, icon, value, onChange, step = 1, min = -Infinity, max = Infinity, unit = "", width = 64, tip, decimals = 1 }: RibbonSpinnerProps) {
  const fmt = (v: number) => `${Number(v.toFixed(decimals)).toLocaleString("tr-TR")}${unit ? ` ${unit}` : ""}`;
  const [text, setText] = useState(fmt(value));
  // Kullanıcı kutuya yazdı mı: yazmadıysa odak kaybında (başka bir komut değeri değiştirip odağı aldığında) eski metin geri yazılmaz.
  const edited = useRef(false);
  useEffect(() => {
    setText(fmt(value));
    edited.current = false;
  }, [value]); // eslint-disable-line react-hooks/exhaustive-deps
  const clamp = (v: number) => Math.min(max, Math.max(min, v));
  const commit = () => {
    if (!edited.current) return;
    edited.current = false;
    const n = parseFloat(text.replace(",", ".").replace(/[^\d.-]/g, ""));
    if (Number.isFinite(n)) onChange(clamp(n));
    else setText(fmt(value));
  };
  return (
    <div className="oh-spinner-wrap" {...tipProps(tip ?? label)}>
      {icon && <span className="oh-rbtn-icon oh-spinner-icon">{icon}</span>}
      {label && <span className="oh-spinner-label">{label}</span>}
      <div className="oh-spinner" style={{ width }}>
        <input
          value={text}
          onChange={(e) => {
            edited.current = true;
            setText(e.target.value);
          }}
          onBlur={commit}
          onKeyDown={(e) => {
            if (e.key === "Enter") commit();
            if (e.key === "ArrowUp") (e.preventDefault(), onChange(clamp(value + step)));
            if (e.key === "ArrowDown") (e.preventDefault(), onChange(clamp(value - step)));
          }}
          aria-label={tip ?? label}
        />
        <div className="oh-spinner-btns">
          <button type="button" tabIndex={-1} onMouseDown={(e) => e.preventDefault()} onClick={() => onChange(clamp(value + step))} aria-label="Artır">
            <svg width="7" height="4" viewBox="0 0 7 4" aria-hidden="true">
              <path d="M0 4l3.5-4L7 4z" fill="currentColor" />
            </svg>
          </button>
          <button type="button" tabIndex={-1} onMouseDown={(e) => e.preventDefault()} onClick={() => onChange(clamp(value - step))} aria-label="Azalt">
            <svg width="7" height="4" viewBox="0 0 7 4" aria-hidden="true">
              <path d="M0 0l3.5 4L7 0z" fill="currentColor" />
            </svg>
          </button>
        </div>
      </div>
    </div>
  );
}
