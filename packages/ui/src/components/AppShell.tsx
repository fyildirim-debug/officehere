import { useEffect, useRef, useState, type ReactNode } from "react";
import { getNative, type AppKind } from "@officehere/core";
import { TooltipLayer, tipProps } from "./Tooltip";
import { UpdatePill } from "./Updates";
import { SaveAlertDialog, SaveBadge } from "./SaveBadge";
import { useUiTheme } from "../hooks/useUiTheme";
import { resolveSaveState, useSaveStatus, type SaveState } from "../hooks/saveStatus";

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

/** Uygulama simgesi: köşesi kıvrık sayfanın alt bandında harf. */
export function AppLogo({ app, size = 20 }: { app: AppKind | "baslangic"; size?: number }) {
  if (app === "baslangic") {
    return (
      <svg className="oh-applogo" width={size} height={size} viewBox="0 0 20 20" aria-hidden="true">
        <rect x="1" y="1" width="18" height="18" rx="4.5" fill="#3d3d8f" stroke="rgba(255,255,255,0.6)" strokeWidth="0.8" />
        <rect x="4.5" y="5" width="11" height="2.4" rx="0.8" fill="#6d9bff" />
        <rect x="4.5" y="8.8" width="11" height="2.4" rx="0.8" fill="#48c98a" />
        <rect x="4.5" y="12.6" width="11" height="2.4" rx="0.8" fill="#ff8a66" />
      </svg>
    );
  }
  return (
    <svg className="oh-applogo" data-app={app} width={size} height={size} viewBox="0 0 20 20" aria-hidden="true">
      <path d="M3.5 1h9L17 5.5V19H3.5z" fill="#fff" stroke="rgba(0,0,0,0.35)" strokeWidth="1" />
      <path d="M12.5 1v4.5H17" fill="none" stroke="rgba(0,0,0,0.35)" strokeWidth="1" />
      <path d="M3.5 11H17v8H3.5z" fill="var(--oh-logo, var(--oh-accent))" />
      <text x="10.25" y="17.6" textAnchor="middle" fontSize="7.5" fontWeight="700" fill="#fff" fontFamily="Segoe UI, Arial, sans-serif">
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
  /**
   * Kayıt durumu rozeti. Verilmezse `dirty` ve kaydetme akışlarının bildirdiği
   * durumdan (saveStatus: otomatik kaydetme, elle Kaydet) çıkarılır.
   */
  saveState?: SaveState;
  /** Son başarılı kaydetme zamanı (ms; "Kaydedildi 14:02"). */
  savedAt?: number;
  /** Hızlı Erişim Araç Çubuğu düğmeleri. */
  quickAccess?: ReactNode;
  onSearch?: (query: string) => void;
  searchPlaceholder?: string;
  /** Sağ köşe (hesap, paylaş vb.). */
  right?: ReactNode;
}

export function TitleBar({ app, title, dirty, saveState, savedAt, quickAccess, onSearch, searchPlaceholder = "Ara (Alt+Q)", right }: TitleBarProps) {
  const [q, setQ] = useState("");
  const status = useSaveStatus();
  const badge = saveState ?? resolveSaveState(!!dirty, status);
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

  // Masaüstünde yerel pencere düğmelerini başlık çubuğu rengine uydur.
  const barRef = useRef<HTMLElement>(null);
  useEffect(() => {
    const native = getNative();
    if (!native?.setTitleBarOverlay) return;
    const sync = () => {
      const el = barRef.current;
      if (!el) return;
      const cs = getComputedStyle(el);
      native.setTitleBarOverlay?.(cs.backgroundColor, cs.color);
    };
    sync();
    const mo = new MutationObserver(sync);
    mo.observe(document.documentElement, { attributes: true, attributeFilter: ["data-oh-theme", "data-app"] });
    const mq = window.matchMedia("(prefers-color-scheme: dark)");
    mq.addEventListener("change", sync);
    return () => {
      mo.disconnect();
      mq.removeEventListener("change", sync);
    };
  }, []);

  return (
    <header className="oh-titlebar" ref={barRef}>
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
        </span>
        {badge && <SaveBadge state={badge} savedAt={savedAt ?? status.savedAt} error={status.error} errorDetail={status.errorDetail} pausedReason={status.paused} pausedKind={status.pausedKind} />}
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
      <div className="oh-titlebar-right">
        <UpdatePill />
        {right}
      </div>
      <div className="oh-titlebar-controls-space" />
      <SaveAlertDialog />
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
  // Kayıtlı arayüz temasını (Renkli/Beyaz/Koyu Gri/Siyah/Sistem) açılışta uygula ve pencereler arası eşitle.
  useUiTheme();
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
