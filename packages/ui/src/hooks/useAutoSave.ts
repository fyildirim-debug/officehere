// Otomatik kaydetme (yalnız masaüstü): diskte yolu olan, kendi biçimimizdeki
// belge düzenlendikçe boşta kalma süresinden sonra aynı dosyaya yazılır. İlk
// yazmadan önce özgün dosyanın yedeği alınır; rozet saveStatus üzerinden güncellenir.
import { useCallback, useEffect, useRef, useSyncExternalStore } from "react";
import { getNative } from "@officehere/core";
import {
  AUTO_SAVE_EVENT,
  AUTO_SAVE_KEY,
  canAutoSavePath,
  createIdleScheduler,
  readAutoSavePref,
  runAutoSave,
  writeAutoSavePref,
  type AutoSaveIo,
  type PreparedSave,
} from "./autoSaveCore";
import { saveStatus } from "./saveStatus";

export interface AutoSaveOptions {
  /** Belgenin diskteki tam yolu (yoksa otomatik kaydetme yapılmaz). */
  path: string | undefined;
  /** Uygulamanın kendi biçimi (ör. ["docx"]); biçim kaybettiren dosyalar kaydedilmez. */
  formats: readonly string[];
  isDirty: () => boolean;
  /**
   * Belge içeriği her değiştiğinde çağrılacak dinleyiciyi kaydeder, ayırıcı döner.
   * Yalnız ilk çizimde okunur (kararlı bir kaynağa bağlanmalı).
   */
  subscribe: (onChange: () => void) => () => void;
  /** Belgeyi kendi biçiminde üretir. */
  prepare: () => Promise<PreparedSave>;
  /** Yazma başarılı ve belge değişmemişse: kirli durumu temizle (kurtarma taslağı silinir). */
  markClean: () => void;
  /** Normal Kaydet ("Kaydedilemedi" rozetine tıklanınca). */
  save: () => void;
}

function subscribePref(cb: () => void): () => void {
  const onStorage = (e: StorageEvent) => {
    if (e.key === AUTO_SAVE_KEY || e.key === null) cb();
  };
  window.addEventListener("storage", onStorage);
  window.addEventListener(AUTO_SAVE_EVENT, cb);
  return () => {
    window.removeEventListener("storage", onStorage);
    window.removeEventListener(AUTO_SAVE_EVENT, cb);
  };
}

/** Seçenekler'deki "Otomatik kaydet" (pencereler arası eşit). */
export function useAutoSavePref(): [boolean, (on: boolean) => void] {
  const on = useSyncExternalStore(subscribePref, () => readAutoSavePref(), () => true);
  return [on, writeAutoSavePref];
}

/**
 * Belgeyi kendiliğinden diske kaydeder. Dönen `touch`, `subscribe` ile
 * yakalanmayan değişikliklerde (ör. belge ayarları) elle çağrılabilir.
 */
export function useAutoSave(options: AutoSaveOptions): { touch: () => void } {
  const ref = useRef(options);
  ref.current = options;
  const touchRef = useRef<() => void>(() => {});

  // "Kaydedilemedi" rozeti normal Kaydet'i çağırır (tarayıcıda da).
  useEffect(() => {
    saveStatus.setRetry(() => ref.current.save());
    return () => saveStatus.setRetry(null);
  }, []);

  useEffect(() => {
    const native = getNative();
    if (!native) return;
    const io: AutoSaveIo = {
      read: async (p) => (await native.readFile(p))?.data ?? null,
      write: (p, data) => native.writeFile(p, data),
    };
    let busy = false;
    let again = false;
    let disposed = false;

    const run = async () => {
      // Önceki tur sürüyorsa bittiğinde yeniden planlanır.
      if (busy) {
        again = true;
        return;
      }
      busy = true;
      const path = ref.current.path;
      try {
        await runAutoSave(
          {
            path,
            formats: ref.current.formats,
            enabled: readAutoSavePref(),
            isDirty: () => ref.current.isDirty() && ref.current.path === path,
            prepare: () => ref.current.prepare(),
            // Bu arada başka belge açıldıysa ya da Farklı Kaydet yapıldıysa dokunulmaz.
            markClean: () => {
              if (ref.current.path === path) ref.current.markClean();
            },
          },
          io,
          saveStatus,
        );
      } finally {
        busy = false;
        if (again && !disposed) {
          again = false;
          scheduler.touch();
        }
      }
    };

    const scheduler = createIdleScheduler(() => void run());
    const touch = () => {
      if (readAutoSavePref() && canAutoSavePath(ref.current.path, ref.current.formats)) scheduler.touch();
    };
    touchRef.current = touch;
    const offChange = ref.current.subscribe(touch);
    // Seçenek açılınca bekleyen değişiklikler kaydedilir, kapanınca bekleyen kayıt iptal edilir.
    const offPref = subscribePref(() => {
      if (!readAutoSavePref()) scheduler.cancel();
      else if (ref.current.isDirty()) touch();
    });
    return () => {
      disposed = true;
      touchRef.current = () => {};
      offChange();
      offPref();
      scheduler.cancel();
    };
  }, []);

  const touch = useCallback(() => touchRef.current(), []);
  return { touch };
}
