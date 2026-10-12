import { useCallback, useEffect, useRef, useState, type ReactNode } from "react";
import { getNative } from "@officehere/core";
import { SaveChangesDialog, type ConfirmChoice } from "../components/Dialog";

export interface DocumentLifecycleOptions {
  /** Pencere başlığında görünen belge adı. */
  title: string;
  dirty: boolean;
  /** Belgenin diskteki tam yolu (varsa); masaüstü kabuğu oturum ve pencere eşlemesi için kullanır. */
  path?: string;
  /**
   * Masaüstünden bir dosya yolu açılmak istendiğinde (başlangıçta `?file=`
   * parametresi, dosya ilişkilendirme, son kullanılanlar).
   */
  onOpenPath: (path: string) => void;
  /** "Kaydet" seçilince çağrılır; kayıt başarılıysa true dönmeli. */
  onSave: () => Promise<boolean>;
}

export interface DocumentLifecycle {
  /** Bileşende bir kez çizilmesi gereken "Değişiklikler kaydedilsin mi?" iletişim kutusu. */
  dialog: ReactNode;
  /**
   * Belge kirliyse kullanıcıya sorar ve gerekirse kaydeder.
   * Devam edilebilirse true döner (kaydedildi ya da "Kaydetme" seçildi).
   */
  confirmDiscard: () => Promise<boolean>;
}

/**
 * Köprüye giden belge durumu. Geçici yerel tip: ortak sözleşmede köprü
 * `setDocumentState(state: DocumentState)` olarak genişliyor; birleştirmeden
 * sonra `@officehere/core`'daki DocumentState kullanılabilir.
 */
interface DocumentState {
  title: string;
  dirty: boolean;
  path?: string;
}

/**
 * Masaüstü kabuğuyla belge yaşam döngüsünü bağlar: başlık/kirli durumu,
 * kapatmadan önce sorma, dosya açma istekleri; tarayıcıda sayfadan ayrılma uyarısı.
 */
export function useDocumentLifecycle({ title, dirty, path, onOpenPath, onSave }: DocumentLifecycleOptions): DocumentLifecycle {
  const [asking, setAsking] = useState(false);
  const resolver = useRef<((c: ConfirmChoice) => void) | null>(null);
  const dirtyRef = useRef(dirty);
  const openRef = useRef(onOpenPath);
  const saveRef = useRef(onSave);
  dirtyRef.current = dirty;
  openRef.current = onOpenPath;
  saveRef.current = onSave;

  const ask = useCallback(
    () =>
      new Promise<ConfirmChoice>((resolve) => {
        resolver.current = resolve;
        setAsking(true);
      }),
    [],
  );

  const confirmDiscard = useCallback(async () => {
    if (!dirtyRef.current) return true;
    const choice = await ask();
    if (choice === "cancel") return false;
    if (choice === "save") return saveRef.current();
    return true;
  }, [ask]);

  useEffect(() => {
    const state: DocumentState = { title, dirty, path };
    getNative()?.setDocumentState(state);
  }, [title, dirty, path]);

  useEffect(() => {
    const native = getNative();
    if (!native) return;
    const offOpen = native.onOpenPath((p) => openRef.current(p));
    const offClose = native.onBeforeClose(() => confirmDiscard());
    return () => {
      offOpen();
      offClose();
    };
  }, [confirmDiscard]);

  // Başlangıçta ?file=<yol> ile açılan pencere.
  useEffect(() => {
    const file = new URLSearchParams(window.location.search).get("file");
    if (file && getNative()) openRef.current(file);
  }, []);

  // Tarayıcıda kaydedilmemiş değişiklik varken sayfadan ayrılma uyarısı.
  useEffect(() => {
    if (getNative()) return;
    const onUnload = (e: BeforeUnloadEvent) => {
      if (!dirtyRef.current) return;
      e.preventDefault();
      e.returnValue = "";
    };
    window.addEventListener("beforeunload", onUnload);
    return () => window.removeEventListener("beforeunload", onUnload);
  }, []);

  const dialog = (
    <SaveChangesDialog
      open={asking}
      fileName={title}
      onChoice={(c) => {
        setAsking(false);
        resolver.current?.(c);
        resolver.current = null;
      }}
    />
  );

  return { dialog, confirmDiscard };
}
