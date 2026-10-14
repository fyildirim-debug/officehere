// Otomatik kaydetme (yalnız masaüstü): diskte yolu olan, kendi biçimimizdeki
// belge düzenlendikçe boşta kalma süresinden sonra aynı dosyaya yazılır. Yazma
// ana sürecin güvenli yolundan geçer; ilk yazmadan önce özgün dosyanın yedeği
// alınır (yalnız yoksa); yazılan dosya uygulamanın okuyucusuyla doğrulanır.
// Rozet saveStatus üzerinden güncellenir.
import { useCallback, useEffect, useRef, useSyncExternalStore } from "react";
import { getNative, type WriteOutcome } from "@officehere/core";
import {
  AUTO_SAVE_EVENT,
  AUTO_SAVE_KEY,
  autoSaveDelays,
  autoSaveDisabledReason,
  canAutoSavePath,
  createIdleScheduler,
  readAutoSavePref,
  runAutoSaveTurn,
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
  /**
   * Dosya açılırken uyarı verildiyse (okunamayan öğeler, basitleştirilmiş/uyumluluk modu)
   * kısa nedeni: ilk ELLE kayda kadar otomatik kaydetme yapılmaz, rozet nedeni gösterir.
   */
  holdReason?: string | null;
  /**
   * Yazılan baytlar uygulamanın kendi okuyucusuyla açılabiliyor mu (ör. openDocument).
   * Boşta kalma anında çalışır; false/hata: önceki içerik geri yüklenir ve otomatik kaydetme kapanır.
   */
  verify?: (data: ArrayBuffer) => Promise<boolean>;
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

type IdleWindow = Window & { requestIdleCallback?: (cb: () => void, o?: { timeout: number }) => number };

/** Tarayıcı boşta kalınca (en geç 2 sn sonra) devam eder. */
function whenIdle(): Promise<void> {
  return new Promise((resolve) => {
    const w = window as IdleWindow;
    if (w.requestIdleCallback) w.requestIdleCallback(() => resolve(), { timeout: 2000 });
    else setTimeout(resolve, 50);
  });
}

/** Bekletme (açılış uyarısı, kurtarılan taslak): bu belge yüklemesinde henüz elle kaydedilmediyse. */
function activeHold(reason: string | null | undefined): string | null {
  const s = saveStatus.get();
  if (s.manualGen === s.gen) return null;
  if (s.hold && s.holdGen === s.gen) return s.hold;
  return reason || null;
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

  // Bekletme nedeni rozete yansır; yeni belge yüklenince (reset) ya da elle kaydedilince yeniden değerlendirilir.
  useEffect(() => {
    if (!getNative()) return;
    const sync = () => {
      const { path, formats, holdReason } = ref.current;
      if (!readAutoSavePref() || !canAutoSavePath(path, formats)) return saveStatus.pause(null);
      const off = autoSaveDisabledReason(path);
      if (off) saveStatus.pause(off, "off");
      else saveStatus.pause(activeHold(holdReason), "hold");
    };
    sync();
    return saveStatus.subscribe(sync);
  }, [options.holdReason, options.path]);

  useEffect(() => {
    const native = getNative();
    if (!native) return;
    const failed = (code: string): WriteOutcome => ({ ok: false, code });
    const io: AutoSaveIo = {
      read: async (p) => (await native.readFile(p))?.data ?? null,
      // Eski kabuk: güvenli yazma köprüsü yoksa yalnız başarı bilgisi gelir.
      write: async (p, data, opts) => (native.writeFileSafe ? native.writeFileSafe(p, data, opts) : (await native.writeFile(p, data)) ? { ok: true } : failed("EUNKNOWN")),
      revert: async (p) => (native.revertFile ? native.revertFile(p) : failed("ENOTSUP")),
    };
    let busy = false;
    let again = false;
    let disposed = false;
    /** Son turun maliyeti: büyük dosyada aralık seyreltilir. */
    let lastCost = 0;

    const run = async () => {
      // Önceki tur (doğrulama dahil) sürüyorsa bittiğinde yeniden planlanır.
      if (busy) {
        again = true;
        return;
      }
      busy = true;
      const path = ref.current.path;
      try {
        const r = await runAutoSaveTurn(
          {
            path,
            formats: ref.current.formats,
            enabled: readAutoSavePref(),
            isDirty: () => ref.current.isDirty() && ref.current.path === path,
            prepare: () => ref.current.prepare(),
            hold: activeHold(ref.current.holdReason),
            verify: ref.current.verify,
            idle: whenIdle,
            // Bu arada başka belge açıldıysa ya da Farklı Kaydet yapıldıysa dokunulmaz.
            markClean: () => {
              if (ref.current.path === path) ref.current.markClean();
            },
          },
          io,
          {
            saving: saveStatus.saving,
            saved: () => saveStatus.saved(),
            failed: (reason, detail) => saveStatus.failed(reason, detail),
            paused: (reason, kind) => saveStatus.pause(reason, kind),
            alert: (title, message) => saveStatus.alert({ title, message }),
          },
        );
        if (r.result === "saved" || r.result === "changed") lastCost = r.costMs;
      } finally {
        busy = false;
        if (again && !disposed) {
          again = false;
          scheduler.touch();
        }
      }
    };

    const scheduler = createIdleScheduler(
      () => void run(),
      () => autoSaveDelays(lastCost).idleMs,
      () => autoSaveDelays(lastCost).maxWaitMs,
    );
    const touch = () => {
      if (readAutoSavePref() && canAutoSavePath(ref.current.path, ref.current.formats)) scheduler.touch();
    };
    touchRef.current = touch;
    const offChange = ref.current.subscribe(touch);
    // Seçenek açılınca bekleyen değişiklikler kaydedilir, kapanınca bekleyen kayıt iptal edilir.
    const offPref = subscribePref(() => {
      if (!readAutoSavePref()) {
        scheduler.cancel();
        saveStatus.pause(null);
      } else if (ref.current.isDirty()) touch();
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
