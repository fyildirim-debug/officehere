import type { ReactElement, ReactNode } from "react";

export { RecentFileList, type RecentFileListProps } from "./RecentFileList";

export interface BackstageNavItem {
  id: string;
  label: string;
  icon?: ReactElement;
  /** Panel açmadan doğrudan eylem (Kaydet, Kapat). */
  onClick?: () => void;
  disabled?: boolean;
  /** Üstüne ayırıcı çizgi. */
  separator?: boolean;
  /** Alt bölüme (Hesap, Seçenekler) yerleşir. */
  bottom?: boolean;
}

export interface BackstageProps {
  items: BackstageNavItem[];
  activeId: string;
  onSelect: (id: string) => void;
  onBack: () => void;
  children: ReactNode;
}

/** "Dosya" sekmesi görünümü: solda vurgulu gezinme, sağda içerik. */
export function Backstage({ items, activeId, onSelect, onBack, children }: BackstageProps) {
  const nav = (it: BackstageNavItem) => (
    <div key={it.id}>
      {it.separator && <div className="oh-bs-sep" />}
      <button
        type="button"
        className={`oh-bs-item${it.id === activeId && !it.onClick ? " is-active" : ""}`}
        disabled={it.disabled}
        onClick={() => (it.onClick ? it.onClick() : onSelect(it.id))}
      >
        {it.icon && <span className="oh-bs-icon">{it.icon}</span>}
        <span>{it.label}</span>
      </button>
    </div>
  );
  return (
    <div className="oh-backstage" onKeyDown={(e) => e.key === "Escape" && onBack()}>
      <nav className="oh-bs-nav">
        <button type="button" className="oh-bs-back" onClick={onBack} aria-label="Geri" autoFocus>
          <svg width="20" height="20" viewBox="0 0 20 20" aria-hidden="true">
            <circle cx="10" cy="10" r="9" fill="none" stroke="currentColor" strokeWidth="1.2" />
            <path d="M11.5 6L7.5 10l4 4M7.8 10H14" fill="none" stroke="currentColor" strokeWidth="1.3" />
          </svg>
        </button>
        <div className="oh-bs-top">{items.filter((i) => !i.bottom).map(nav)}</div>
        <div className="oh-bs-bottom">{items.filter((i) => i.bottom).map(nav)}</div>
      </nav>
      <section className="oh-bs-content">{children}</section>
    </div>
  );
}

export function BackstageHeading({ children }: { children: ReactNode }) {
  return <h1 className="oh-bs-h1">{children}</h1>;
}

export function BackstageSubheading({ children }: { children: ReactNode }) {
  return <h2 className="oh-bs-h2">{children}</h2>;
}

export interface TemplateTileProps {
  label: string;
  preview: ReactNode;
  onClick: () => void;
  /** Kutucuk en-boy oranı: belge dikey, sunum yatay. */
  shape?: "portrait" | "landscape" | "square";
}

export function TemplateTile({ label, preview, onClick, shape = "portrait" }: TemplateTileProps) {
  return (
    <button type="button" className={`oh-template is-${shape}`} onClick={onClick}>
      <span className="oh-template-preview">{preview}</span>
      <span className="oh-template-label">{label}</span>
    </button>
  );
}

export function TemplateRow({ children }: { children: ReactNode }) {
  return <div className="oh-template-row">{children}</div>;
}

/** Backstage içindeki büyük eylem kartı (Aç → "Gözat", Dışarı Aktar → "PDF Oluştur"). */
export function BackstageAction({ icon, title, description, onClick }: { icon?: ReactElement; title: string; description?: string; onClick: () => void }) {
  return (
    <button type="button" className="oh-bs-action" onClick={onClick}>
      {icon && <span className="oh-bs-action-icon">{icon}</span>}
      <span className="oh-bs-action-text">
        <strong>{title}</strong>
        {description && <span>{description}</span>}
      </span>
    </button>
  );
}

/** Bilgi paneli için özellik satırları. */
export function PropertyList({ rows }: { rows: [string, ReactNode][] }) {
  return (
    <dl className="oh-props">
      {rows.map(([k, v]) => (
        <div key={k}>
          <dt>{k}</dt>
          <dd>{v}</dd>
        </div>
      ))}
    </dl>
  );
}
