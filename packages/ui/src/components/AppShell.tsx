import { useEffect, useRef, useState, type ReactNode } from "react";
import type { AppKind } from "@officehere/core";
import { TooltipLayer, tipProps } from "./Tooltip";

export const APP_NAMES: Record<AppKind | "baslangic", string> = {
  belge: "OfficeHere Belge",
  tablo: "OfficeHere Tablo",
  sunum: "OfficeHere Sunum",
  baslangic: "OfficeHere",
};

const APP_LETTER: Record<AppKind | "baslangic", string> = {
  belge: "B",
  tablo: "T",
  sunum: "S",
  baslangic: "O",
};

/** Uygulama simgesi: köşesi kıvrık sayfa üzerinde harf. Microsoft logolarıyla ilgisi yoktur. */
export function AppLogo({ app, size = 20 }: { app: AppKind | "baslangic"; size?: number }) {
  return (
    <svg className="oh-applogo" width={size} height={size} viewBox="0 0 20 20" aria-hidden="true">
      <path d="M4 1.5h8.5L17 6v12.5H4z" fill="#fff" stroke="rgba(0,0,0,0.3)" strokeWidth="1" />
      <path d="M12.5 1.5V6H17" fill="none" stroke="rgba(0,0,0,0.3)" strokeWidth="1" />
      <rect x="0.5" y="6.5" width="12" height="11" rx="2" fill="var(--oh-logo, var(--oh-accent))" stroke="#fff" strokeWidth="1" />
      <text x="6.5" y="15.2" textAnchor="middle" fontSize="8.5" fontWeight="700" fill="#fff" fontFamily="Segoe UI, Arial, sans-serif">
        {APP_LETTER[app]}
      </text>
    </svg>
  );
}

export interface TitleBarProps {
  app: AppKind | "baslangic";
  /** Belge adı (ör. "Belge1"). */
  title: string;
  dirty?: boolean;
  /** Hızlı Erişim Araç Çubuğu düğmeleri. */
  quickAccess?: ReactNode;
  onSearch?: (query: string) => void;
  searchPlaceholder?: string;
  /** Sağ köşe (hesap, paylaş vb.). */
  right?: ReactNode;
}

export function TitleBar({ app, title, dirty, quickAccess, onSearch, searchPlaceholder = "Ara (Alt+Q)", right }: TitleBarProps) {
  const [q, setQ] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.altKey && (e.key === "q" || e.key === "Q")) {
        e.preventDefault();
        inputRef.current?.focus();
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);

  useEffect(() => {
    document.title = `${dirty ? "● " : ""}${title} - ${APP_NAMES[app]}`;
  }, [title, dirty, app]);

  return (
    <header className="oh-titlebar">
      <div className="oh-titlebar-left">
        <span className="oh-titlebar-logo">
          <AppLogo app={app} />
        </span>
        {quickAccess && <div className="oh-qat">{quickAccess}</div>}
      </div>
      <div className="oh-titlebar-center">
        <span className="oh-titlebar-title" title={title}>
          {title}
          <span className="oh-titlebar-app"> - {APP_NAMES[app]}</span>
          {dirty && <span className="oh-titlebar-dirty" {...tipProps("Kaydedilmemiş değişiklikler var")}> • Kaydedilmedi</span>}
        </span>
      </div>
      <div className="oh-titlebar-search">
        {onSearch && (
          <label className="oh-search">
            <svg width="14" height="14" viewBox="0 0 16 16" aria-hidden="true">
              <circle cx="6.5" cy="6.5" r="4.5" fill="none" stroke="currentColor" strokeWidth="1.3" />
              <path d="M10 10l4.5 4.5" stroke="currentColor" strokeWidth="1.3" />
            </svg>
            <input
              ref={inputRef}
              value={q}
              placeholder={searchPlaceholder}
              onChange={(e) => setQ(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter") onSearch(q);
                if (e.key === "Escape") {
                  setQ("");
                  e.currentTarget.blur();
                }
              }}
            />
          </label>
        )}
      </div>
      <div className="oh-titlebar-right">{right}</div>
      <div className="oh-titlebar-controls-space" />
    </header>
  );
}

/** Hızlı Erişim Araç Çubuğu düğmesi. */
export function QatButton({ icon, tip, shortcut, onClick, disabled }: { icon: ReactNode; tip: string; shortcut?: string; onClick: () => void; disabled?: boolean }) {
  return (
    <button type="button" className="oh-qat-btn" disabled={disabled} onMouseDown={(e) => e.preventDefault()} onClick={onClick} aria-label={tip} {...tipProps(tip, shortcut)}>
      {icon}
    </button>
  );
}

export interface AppShellProps {
  app: AppKind | "baslangic";
  titleBar: ReactNode;
  ribbon?: ReactNode;
  /** Şeridin altındaki bilgi çubukları, formül çubuğu vb. */
  belowRibbon?: ReactNode;
  statusBar?: ReactNode;
  /** Dosya menüsü (açıksa tüm çalışma alanını kaplar). */
  backstage?: ReactNode;
  children: ReactNode;
}

/** Başlık çubuğu + şerit + çalışma alanı + durum çubuğu iskeleti. */
export function AppShell({ app, titleBar, ribbon, belowRibbon, statusBar, backstage, children }: AppShellProps) {
  useEffect(() => {
    document.documentElement.setAttribute("data-app", app);
  }, [app]);
  return (
    <div className="oh-shell">
      {titleBar}
      {backstage ? (
        <div className="oh-shell-backstage">{backstage}</div>
      ) : (
        <>
          {ribbon}
          {belowRibbon}
          <main className="oh-shell-main">{children}</main>
          {statusBar}
        </>
      )}
      <TooltipLayer />
    </div>
  );
}
