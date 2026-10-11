import { createContext, useCallback, useContext, useRef, useState, type ReactElement, type ReactNode } from "react";
import { Popover } from "./Popover";

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

export function Menu({ children, className, width }: { children: ReactNode; className?: string; width?: number }) {
  return (
    <div className={`oh-menu ${className ?? ""}`} role="menu" style={width ? { minWidth: width } : undefined}>
      {children}
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
  const ref = useRef<HTMLButtonElement>(null);
  const [subOpen, setSubOpen] = useState(false);
  const timer = useRef<number | undefined>(undefined);

  const openSub = () => {
    window.clearTimeout(timer.current);
    timer.current = window.setTimeout(() => setSubOpen(true), 120);
  };
  const closeSub = () => {
    window.clearTimeout(timer.current);
    timer.current = window.setTimeout(() => setSubOpen(false), 200);
  };

  return (
    <>
      <button
        ref={ref}
        type="button"
        role="menuitem"
        className={`oh-menu-item${checked ? " is-checked" : ""}${danger ? " is-danger" : ""}${submenu ? " has-sub" : ""}${subOpen ? " is-open" : ""}`}
        disabled={disabled}
        aria-checked={checked}
        onMouseEnter={submenu ? openSub : undefined}
        onMouseLeave={submenu ? closeSub : undefined}
        onClick={() => {
          if (submenu) {
            setSubOpen(true);
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
        <Popover open={subOpen} onClose={() => setSubOpen(false)} anchor={ref} placement="right-start">
          <div onMouseEnter={() => window.clearTimeout(timer.current)} onMouseLeave={closeSub}>
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
