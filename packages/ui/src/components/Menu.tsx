import { createContext, useCallback, useContext, useRef, useState, type ReactElement, type ReactNode } from "react";
import { Popover } from "./Popover";
import { navItems } from "./menuKeyboard";

interface MenuCtx {
  /** Tüm menü zincirini kapatır. */
  closeAll: () => void;
}

const MenuContext = createContext<MenuCtx>({ closeAll: () => {} });

export function useMenuClose() {
  return useContext(MenuContext).closeAll;
}

export function MenuProvider({ onClose, children }: { onClose: () => void; children: ReactNode }) {
  return <MenuContext.Provider value={{ closeAll: onClose }}>{children}</MenuContext.Provider>;
}

/** Menü içindeki öğeler Tab durağı değildir (odak ok tuşlarıyla taşınır); menü dışındaki öğe olağan kalır. */
const MenuListContext = createContext(false);

/**
 * Menü kabı. Klavye gezintisi (↑/↓, Home/End, harfle atlama, →/← alt menü, Esc, Tab) içinde
 * bulunduğu Popover'dan gelir; uygulamaların düz .oh-menu kapları da aynı davranışı alır.
 */
export function Menu({ children, className, width, label }: { children: ReactNode; className?: string; width?: number; label?: string }) {
  return (
    <div className={`oh-menu ${className ?? ""}`} role="menu" aria-label={label} tabIndex={-1} style={width ? { minWidth: width } : undefined}>
      <MenuListContext.Provider value={true}>{children}</MenuListContext.Provider>
    </div>
  );
}

export interface MenuItemProps {
  icon?: ReactElement;
  label: ReactNode;
  shortcut?: string;
  checked?: boolean;
  disabled?: boolean;
  /** Seçiliyse soldaki onay işareti yerine vurgulu arka plan. */
  danger?: boolean;
  description?: string;
  onClick?: () => void;
  /** Alt menü içeriği. */
  submenu?: ReactNode;
  /** Tıklanınca menü kapanmasın. */
  keepOpen?: boolean;
}

export function MenuItem({ icon, label, shortcut, checked, disabled, danger, description, onClick, submenu, keepOpen }: MenuItemProps) {
  const { closeAll } = useContext(MenuContext);
  const inList = useContext(MenuListContext);
  const ref = useRef<HTMLButtonElement>(null);
  const subRef = useRef<HTMLDivElement>(null);
  const [subOpen, setSubOpen] = useState(false);
  // Alt menü klavyeyle açıldıysa odak ilk öğesine geçer; fareyle açılınca odak yerinde kalır.
  const [subKbd, setSubKbd] = useState(false);
  const timer = useRef<number | undefined>(undefined);

  const openSub = () => {
    window.clearTimeout(timer.current);
    timer.current = window.setTimeout(() => (setSubKbd(false), setSubOpen(true)), 120);
  };
  const closeSub = () => {
    window.clearTimeout(timer.current);
    timer.current = window.setTimeout(() => setSubOpen(false), 200);
  };
  const openSubByKey = () => {
    window.clearTimeout(timer.current);
    if (subOpen) {
      const first = subRef.current && navItems(subRef.current)[0];
      first?.focus({ preventScroll: true });
    } else {
      setSubKbd(true);
      setSubOpen(true);
    }
  };

  return (
    <>
      <button
        ref={ref}
        type="button"
        role={checked !== undefined ? "menuitemcheckbox" : "menuitem"}
        tabIndex={inList ? -1 : undefined}
        className={`oh-menu-item${checked ? " is-checked" : ""}${danger ? " is-danger" : ""}${submenu ? " has-sub" : ""}${subOpen ? " is-open" : ""}`}
        disabled={disabled}
        aria-checked={checked}
        aria-haspopup={submenu ? "menu" : undefined}
        aria-expanded={submenu ? subOpen : undefined}
        onMouseEnter={submenu ? openSub : undefined}
        onMouseLeave={submenu ? closeSub : undefined}
        onKeyDown={(e) => {
          if (!submenu || e.key !== "ArrowRight") return;
          e.preventDefault();
          e.stopPropagation();
          openSubByKey();
        }}
        onClick={(e) => {
          if (submenu) {
            // Klavyeyle (Enter/Space) tetiklenen tıklamanın ayrıntısı 0'dır.
            if (e.detail === 0) openSubByKey();
            else setSubOpen(true);
            return;
          }
          onClick?.();
          if (!keepOpen) closeAll();
        }}
      >
        <span className="oh-menu-icon">{checked && !icon ? <CheckGlyph /> : icon}</span>
        <span className="oh-menu-label">
          {label}
          {description && <span className="oh-menu-desc">{description}</span>}
        </span>
        {shortcut && <span className="oh-menu-shortcut">{shortcut}</span>}
        {submenu && <span className="oh-menu-chevron">›</span>}
      </button>
      {submenu && (
        <Popover
          open={subOpen}
          onClose={() => {
            window.clearTimeout(timer.current);
            setSubOpen(false);
          }}
          anchor={ref}
          placement="right-start"
          focusOnOpen={subKbd}
        >
          <div ref={subRef} onMouseEnter={() => window.clearTimeout(timer.current)} onMouseLeave={closeSub}>
            {submenu}
          </div>
        </Popover>
      )}
    </>
  );
}

function CheckGlyph() {
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
      <path d="M3.5 8.5l3 3 6-7" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

export function MenuSeparator() {
  return <div className="oh-menu-sep" role="separator" />;
}

export function MenuHeader({ children }: { children: ReactNode }) {
  return <div className="oh-menu-header">{children}</div>;
}

/** Menü içinde serbest içerik (renk paleti, ızgara vb.) için dolgulu alan. */
export function MenuSection({ children }: { children: ReactNode }) {
  return <div className="oh-menu-section">{children}</div>;
}

export interface ContextMenuState {
  x: number;
  y: number;
  content: ReactNode;
}

/**
 * Sağ tık menüsü. Kullanım:
 *   const ctx = useContextMenu();
 *   <div onContextMenu={(e) => ctx.open(e, <Menu>...</Menu>)} />
 *   {ctx.element}
 */
export function useContextMenu() {
  const [state, setState] = useState<ContextMenuState | null>(null);
  const close = useCallback(() => setState(null), []);
  const open = useCallback((e: { clientX: number; clientY: number; preventDefault: () => void }, content: ReactNode) => {
    e.preventDefault();
    setState({ x: e.clientX, y: e.clientY, content });
  }, []);
  const element = (
    <Popover open={state !== null} onClose={close} point={state ? { x: state.x, y: state.y } : undefined}>
      <MenuProvider onClose={close}>{state?.content}</MenuProvider>
    </Popover>
  );
  return { open, close, element, isOpen: state !== null };
}
