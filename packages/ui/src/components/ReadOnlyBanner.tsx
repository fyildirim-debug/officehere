// Görüntüleyici kipi (masaüstü): Seçenekler'de açıksa İndirilenler'den ya da e-posta ekinden gelen
// dosya salt okunur açılır ve şeridin altında sarı şerit gösterilir. Kaynak tespiti ana süreçtedir
// (apps/desktop/src/main/fileOrigin.ts); sayfa yalnız "bu yol salt okunur mu" diye sorar. Pencere
// başına tek belge olduğundan durum modül düzeyindedir. Bu kipte otomatik kaydetme kapalıdır
// (autoSaveCore); düzenlemeyi her uygulama kendi kilidiyle kapatır (Belge'de Düzenlemeyi Kısıtla'nın
// salt okunur kipi). "Düzenlemeyi etkinleştir" sonrası her şey normaldir.
import { useEffect, useSyncExternalStore } from "react";
import { ShieldError20Regular } from "@fluentui/react-icons";
import { getNative } from "@officehere/core";
import { setAutoSaveReadOnly } from "../hooks/autoSaveCore";

interface ViewerState {
  /** Ana sürece sorulan belge yolu. */
  path?: string;
  readOnly: boolean;
}

let state: ViewerState = { readOnly: false };
/** Son sorgu: yanıt gelmeden yol değiştiyse eski yanıt yok sayılır. */
let seq = 0;
const listeners = new Set<() => void>();

function set(next: ViewerState) {
  state = next;
  listeners.forEach((l) => l());
}

function subscribe(cb: () => void): () => void {
  listeners.add(cb);
  return () => listeners.delete(cb);
}

const getState = () => state;

/**
 * Belgenin yolu değişince (dosya açıldı, Farklı Kaydet) ana sürece sorar; belge salt okunur
 * açıldıysa true döner. Uygulama bu değerle kendi düzenleme kilidini kurar.
 */
export function useReadOnlyView(path: string | undefined): boolean {
  useEffect(() => {
    const id = ++seq;
    if (state.path && state.path !== path) setAutoSaveReadOnly(state.path, false);
    set({ path, readOnly: false });
    const viewer = getNative()?.viewer;
    if (!path || !viewer) return;
    viewer.check(path).then(
      (on) => {
        if (id !== seq || !on) return;
        setAutoSaveReadOnly(path, true);
        set({ path, readOnly: true });
      },
      () => undefined,
    );
  }, [path]);
  const s = useSyncExternalStore(subscribe, getState, getState);
  return s.readOnly && s.path === path;
}

/** "Düzenlemeyi etkinleştir": dosya bu oturum boyunca düzenlenebilir; otomatik kaydetme sürer. */
export function enableEditing(): void {
  const { path, readOnly } = state;
  if (!path || !readOnly) return;
  setAutoSaveReadOnly(path, false);
  set({ path, readOnly: false });
  void getNative()?.viewer?.allowEditing(path).catch(() => undefined);
}

/** Şeridin altındaki sarı şerit; belge salt okunur açılmadıysa hiçbir şey çizmez. */
export function ReadOnlyBanner() {
  const s = useSyncExternalStore(subscribe, getState, getState);
  if (!s.readOnly) return null;
  const enable = () => {
    // Klavyeyle etkinleştirildiyse odak kaybolan düğmeden belgeye döner (fareyle odak zaten belgededir).
    const fromKeyboard = document.activeElement?.closest(".oh-readonly-banner");
    enableEditing();
    if (fromKeyboard) requestAnimationFrame(() => document.querySelector<HTMLElement>(".oh-shell-main [contenteditable='true'], .oh-shell-main input, .oh-shell-main [tabindex='0']")?.focus());
  };
  return (
    <div className="oh-messagebar is-warning oh-readonly-banner" role="status">
      <ShieldError20Regular className="oh-readonly-banner-icon" aria-hidden="true" />
      <span className="oh-messagebar-text">
        <strong>Salt okunur</strong> · Bu dosya internetten geldi; güvenliğiniz için salt okunur açıldı.
      </span>
      <button type="button" className="oh-btn is-primary" onMouseDown={(e) => e.preventDefault()} onClick={enable}>
        Düzenlemeyi etkinleştir
      </button>
    </div>
  );
}
