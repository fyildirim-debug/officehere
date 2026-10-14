import { useCallback, useEffect, useRef, useState, type ReactNode } from "react";
import { getNative, type DocumentState } from "@officehere/core";
import { SaveChangesDialog, type ConfirmChoice } from "../components/Dialog";
import { installErrorReporting, logPageEvent } from "./errorReporting";
import { saveStatus } from "./saveStatus";
import { disableAutoSave } from "./autoSaveCore";

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
  /**
   * Masaüstünde pencere kullanıcı isteğiyle kapanırken (kaydedildikten ya da
   * "Kaydetme" seçildikten sonra) çağrılır; kurtarma taslağını silmek için.
   * Pencere bu işin bitmesini bekler.
   */
  onClose?: () => void | Promise<void>;
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
 * Pencerenin belgesiyle ilgili, uygulamadan bağımsız bileşenlerin (ör. Dosya › Bilgi ›
 * Önceki sürümler) kullandığı kancalar. useDocumentLifecycle günceller.
 */
export const documentLifecycle = {
  /** Belgede kaydedilmemiş değişiklik var mı. */
  isDirty: (): boolean => false,
  /** Bu pencerenin kurtarma taslağını siler (belge diskten yeniden açılmadan önce). */
  forgetDraft: async (): Promise<void> => {},
};

/** Aynı dosya başka bir pencerede de açıkken otomatik kaydetme duraklatılır. */
export const PATH_CONFLICT_REASON = "Bu dosya başka bir pencerede de açık";

/**
 * Masaüstü kabuğuyla belge yaşam döngüsünü bağlar: başlık/kirli durumu,
 * kapatmadan önce sorma, dosya açma istekleri; tarayıcıda sayfadan ayrılma uyarısı.
 */
export function useDocumentLifecycle({ title, dirty, path, onOpenPath, onSave, onClose }: DocumentLifecycleOptions): DocumentLifecycle {
  const [asking, setAsking] = useState(false);
  const resolver = useRef<((c: ConfirmChoice) => void) | null>(null);
  const dirtyRef = useRef(dirty);
  const openRef = useRef(onOpenPath);
  const saveRef = useRef(onSave);
  const closeRef = useRef(onClose);
  closeRef.current = onClose;
  dirtyRef.current = dirty;
  openRef.current = onOpenPath;
  saveRef.current = onSave;

  useEffect(() => {
    installErrorReporting();
    documentLifecycle.isDirty = () => dirtyRef.current;
    documentLifecycle.forgetDraft = async () => {
      await closeRef.current?.();
    };
  }, []);

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
    if (choice === "save") {
      // Kaydetme başarısızsa (ya da hata verirse) belge açık kalır; neden rozette görünür.
      try {
        return await saveRef.current();
      } catch (err) {
        logPageEvent({ level: "hata", message: `Kapatırken kaydetme: ${(err as Error)?.message ?? String(err)}`, stack: (err as Error)?.stack, kind: "confirm-save" });
        return false;
      }
    }
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
    const offClose = native.onBeforeClose(async () => {
      const ok = await confirmDiscard();
      if (ok) {
        try {
          await closeRef.current?.();
        } catch {
          // taslak silinemese de pencere kapanır
        }
      }
      return ok;
    });
    // Aynı dosya başka pencerede de açık: bu pencere otomatik kaydetmez ve kullanıcı uyarılır.
    const offConflict =
      native.onPathConflict?.((p) => {
        disableAutoSave(p, PATH_CONFLICT_REASON);
        saveStatus.pause(PATH_CONFLICT_REASON, "off");
        saveStatus.alert({
          title: "Dosya başka bir pencerede açık",
          message:
            "Bu dosya başka bir OfficeHere penceresinde de açık. Değişikliklerin birbirinin üzerine yazılmaması için bu pencerede otomatik kaydetme yapılmaz. Diğer pencereyi kullanın ya da bu kopyayı Farklı Kaydet ile yeni bir dosyaya kaydedin.",
        });
      }) ?? (() => {});
    return () => {
      offOpen();
      offClose();
      offConflict();
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
