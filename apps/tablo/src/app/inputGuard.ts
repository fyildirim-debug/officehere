/**
 * Şerit ve bölmelerdeki girdi kutularında Enter/Esc: Enter'ın varsayılan
 * eylemi (satır sonu ekleme) odağın geçtiği ızgara düzenleyicisine
 * "yazarak düzenleme" olarak ulaşmaz; şeritteki kutudan sonra odak
 * çalışma sayfasına döner (hücre düzenlenmez).
 */
import type { Controller } from "../controller";

const NON_TEXT = new Set(["checkbox", "radio", "button", "submit", "reset", "range", "color", "file"]);

export function installInputGuard(ctl: Controller): () => void {
  const onKeyDown = (e: KeyboardEvent) => {
    if ((e.key !== "Enter" && e.key !== "Escape") || e.isComposing || e.altKey) return;
    const t = e.target;
    if (!(t instanceof HTMLInputElement) || NON_TEXT.has(t.type)) return;
    if (e.key === "Enter") e.preventDefault();
    if (!t.closest(".oh-ribbon")) return;
    // Kutunun kendi işleyicisi çalıştıktan sonra: odak hâlâ kutudaysa (ya da boşta) ızgaraya döner
    window.setTimeout(() => {
      if (document.querySelector(".oh-dialog")) return;
      const a = document.activeElement;
      if (a === t || a === document.body || a === null) ctl.focusGrid();
    }, 0);
  };
  window.addEventListener("keydown", onKeyDown, true);
  return () => window.removeEventListener("keydown", onKeyDown, true);
}

/** Etkin öğe bir yazı alanı mı (ad kutusu, formül çubuğu, sayfa adı kutusu vb.)? */
function typingTarget(a: Element | null): boolean {
  if (!(a instanceof HTMLElement) || a === document.body) return false;
  if (a.isContentEditable || a instanceof HTMLTextAreaElement) return true;
  return a instanceof HTMLInputElement && !NON_TEXT.has(a.type);
}

/**
 * Pencere/menü kapandıktan sonra odağı ızgaraya döndürür; kullanıcı bu arada
 * bir yazı alanına (ör. Ad Kutusu) geçtiyse odağını çalmaz.
 */
export function refocusGridSoon(ctl: Controller): void {
  requestAnimationFrame(() => {
    if (!typingTarget(document.activeElement)) ctl.focusGrid();
  });
}
