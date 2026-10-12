import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  type ReactElement,
  type ReactNode,
} from "react";
import { Popover } from "./Popover";
import { MenuProvider } from "./Menu";
import { tipProps } from "./Tooltip";

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

  let lastGroup: string | undefined;
  return (
    <div ref={ribbonRef} className={`oh-ribbon${collapsed ? " is-collapsed" : ""}${floating ? " is-floating" : ""}`}>
      <div className="oh-ribbon-tabs" role="tablist">
        <button type="button" className="oh-ribbon-tab oh-ribbon-file" onClick={onFileClick}>
          {fileLabel}
        </button>
        {visible.map((t) => {
          const showGroup = t.contextual && t.contextualGroup && t.contextualGroup !== lastGroup;
          lastGroup = t.contextual ? t.contextualGroup : undefined;
          return (
            <button
              key={t.id}
              type="button"
              role="tab"
              aria-selected={t.id === active?.id}
              className={`oh-ribbon-tab${t.id === active?.id && showBody ? " is-active" : ""}${t.contextual ? " is-contextual" : ""}`}
              data-group={showGroup ? t.contextualGroup : undefined}
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
        <div className="oh-ribbon-tabs-spacer" />
        {tabsRight}
      </div>
      {showBody && active && (
        <RibbonLayoutContext.Provider value={layout}>
          <div className="oh-ribbon-body" ref={bodyRef} role="tabpanel">
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
        <button ref={btnRef} type="button" className={`oh-rbtn oh-rbtn-large${open ? " is-open" : ""}`} onClick={() => setOpen((o) => !o)}>
          <span className="oh-rbtn-icon">{icon ?? <span className="oh-ribbon-group-dot" />}</span>
          <span className="oh-rbtn-label">
            {label}
            <Chevron />
          </span>
        </button>
        <Popover open={open} onClose={close} anchor={btnRef}>
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
        <button type="button" className="oh-split-main" disabled={disabled} aria-label={label ?? tip} onMouseDown={preventMouseFocus} onClick={onClick}>
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
/* Açılan kutu (yazı tipi, boyut)                                       */
/* ------------------------------------------------------------------ */

export interface ComboOption {
  value: string;
  label?: string;
  /** Seçenek satırına uygulanacak stil (yazı tipi önizlemesi). */
  style?: React.CSSProperties;
  group?: string;
}

export interface RibbonComboProps {
  value: string;
  options: (string | ComboOption)[];
  onCommit: (value: string) => void;
  width?: number;
  editable?: boolean;
  tip?: string;
  shortcut?: string;
  disabled?: boolean;
  placeholder?: string;
  /** Liste yüksekliği (px). */
  listHeight?: number;
}

export function RibbonCombo({ value, options, onCommit, width = 120, editable = true, tip, shortcut, disabled, placeholder, listHeight = 320 }: RibbonComboProps) {
  const [text, setText] = useState(value);
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);
  const listRef = useRef<HTMLDivElement>(null);
  const opts = options.map((o) => (typeof o === "string" ? { value: o } : o));

  useEffect(() => setText(value), [value]);

  useEffect(() => {
    if (!open) return;
    // Seçili öğeyi ortala; scrollIntoView yerine yalnız listeyi kaydır ki üst öğeler (sayfa) kaymasın.
    const list = listRef.current;
    const sel = list?.querySelector(".is-selected") as HTMLElement | null;
    if (list && sel) list.scrollTop = sel.offsetTop - list.clientHeight / 2 + sel.offsetHeight / 2;
  }, [open]);

  const commit = (v: string) => {
    setOpen(false);
    if (v.trim() === "") {
      setText(value);
      return;
    }
    onCommit(v);
  };

  let lastGroup: string | undefined;
  return (
    <>
      <div ref={ref} className={`oh-combo${open ? " is-open" : ""}${disabled ? " is-disabled" : ""}`} style={{ width }} {...tipProps(tip, shortcut)}>
        <input
          className="oh-combo-input"
          value={text}
          disabled={disabled}
          readOnly={!editable}
          placeholder={placeholder}
          aria-label={tip}
          onChange={(e) => setText(e.target.value)}
          onFocus={(e) => e.currentTarget.select()}
          onMouseDown={!editable ? (e) => (e.preventDefault(), setOpen(!open)) : undefined}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              e.preventDefault();
              commit(text);
            } else if (e.key === "Escape") {
              setText(value);
              setOpen(false);
            } else if (e.key === "ArrowDown" && e.altKey) {
              setOpen(true);
            }
          }}
          onBlur={() => {
            if (!open && text !== value) commit(text);
          }}
        />
        <button type="button" className="oh-combo-arrow" tabIndex={-1} disabled={disabled} onMouseDown={(e) => e.preventDefault()} onClick={() => setOpen(!open)} aria-label="Listeyi aç">
          <Chevron />
        </button>
      </div>
      <Popover open={open} onClose={() => setOpen(false)} anchor={ref}>
        <div ref={listRef} className="oh-combo-list" style={{ minWidth: width, maxHeight: listHeight }} role="listbox">
          {opts.map((o) => {
            const header = o.group && o.group !== lastGroup ? o.group : null;
            lastGroup = o.group;
            return (
              <div key={`${o.group ?? ""}:${o.value}`}>
                {header && <div className="oh-menu-header">{header}</div>}
                <button
                  type="button"
                  role="option"
                  aria-selected={o.value === value}
                  className={`oh-combo-option${o.value === value ? " is-selected" : ""}`}
                  style={o.style}
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={() => {
                    setText(o.value);
                    commit(o.value);
                  }}
                >
                  {o.label ?? o.value}
                </button>
              </div>
            );
          })}
        </div>
      </Popover>
    </>
  );
}

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
  useEffect(() => setText(fmt(value)), [value]); // eslint-disable-line react-hooks/exhaustive-deps
  const clamp = (v: number) => Math.min(max, Math.max(min, v));
  const commit = () => {
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
          onChange={(e) => setText(e.target.value)}
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
