/**
 * Otomatik Kurtarma: çalışma kitabı değiştikçe (kaydedilmemişken) birkaç
 * saniyelik boşluktan sonra bu pencerenin taslağı IndexedDB'ye yazılır.
 * Serileştirme boşta kalma süresinde parça parça yapılır; bu sırada kitap
 * değişirse yarıda bırakılıp yeniden planlanır. Kaydedince ya da yeni kitap
 * açılınca bu pencerenin taslağı silinir. Açılışta sahibi kapanmış
 * pencerelerin taslakları sırayla "Kurtarılan çalışma kitabı" çubuğuyla
 * önerilir; geri yüklenen taslak bu pencereye devredilir (ya da pencere
 * kullanılıyorsa yeni pencerede açılır). Çubuğun Kapat (×) düğmesi önerilen
 * taslakları siler; kapatılan taslak sonraki açılışlarda yeniden önerilmez.
 */
import { useCallback, useEffect, useState } from "react";
import { MessageBar, saveStatus } from "@officehere/ui";
import type { Controller } from "../controller";
import { serializeDraft, workbookFromDraft, type WorkbookDraft } from "../io/draft";
import { SESSION_ID, adoptDraft, announceLive, deleteDraft, lastWriteQuota, orphanDrafts, pruneHeartbeats, readDraft, writeDraft, type DraftMeta, type DraftRecord } from "./draftStore";

export { SESSION_ID };
/** Son değişiklikten sonra beklenen boşluk. */
const IDLE_MS = 2500;
/** Sürekli düzenlemede en geç bu sürede bir yazılır. */
const MAX_WAIT_MS = 15000;
/** Yeni pencerede açılacak taslağın adres parametresi. */
const RESTORE_PARAM = "kurtar";

type IdleDeadline = { timeRemaining: () => number; didTimeout: boolean };
type IdleWin = Window & { requestIdleCallback?: (cb: (d: IdleDeadline) => void, o?: { timeout: number }) => number; cancelIdleCallback?: (h: number) => void };

function idle(cb: (d: IdleDeadline) => void): () => void {
  const w = window as IdleWin;
  if (w.requestIdleCallback) {
    const h = w.requestIdleCallback(cb, { timeout: 2000 });
    return () => w.cancelIdleCallback?.(h);
  }
  const h = window.setTimeout(() => cb({ timeRemaining: () => 12, didTimeout: true }), 16);
  return () => window.clearTimeout(h);
}

/** Taslağı bu pencereye yükler ve devralır. */
async function loadInto(ctl: Controller, rec: DraftRecord): Promise<void> {
  const wb = workbookFromDraft(rec.draft);
  ctl.load(wb, rec.fileName, rec.path ? { name: rec.fileName, path: rec.path } : null);
  ctl.dirty = true;
  ctl.pristine = false;
  ctl.infoBar = { kind: "info", text: "Kurtarılan çalışma kitabı geri yüklendi. Saklamak için kaydedin." };
  ctl.emit();
  // Kurtarılan taslak diskteki dosyanın üzerine kendiliğinden yazılmaz: ilk kaydı kullanıcı yapar.
  if (rec.path) saveStatus.hold("Kurtarılan taslak; ilk kaydı siz onaylayın");
  await adoptDraft(rec);
}

export function useAutoRecover(ctl: Controller) {
  const [offers, setOffers] = useState<DraftMeta[]>([]);

  useEffect(() => announceLive(SESSION_ID), []);

  // Açılışta: yeni pencerede açılması istenen taslak ya da sahipsiz taslaklar
  useEffect(() => {
    let alive = true;
    void (async () => {
      const params = new URLSearchParams(window.location.search);
      const wanted = params.get(RESTORE_PARAM);
      if (wanted) {
        params.delete(RESTORE_PARAM);
        const q = params.toString();
        window.history.replaceState(null, "", `${window.location.pathname}${q ? `?${q}` : ""}${window.location.hash}`);
        const rec = await readDraft(wanted);
        if (rec && alive) {
          try {
            await loadInto(ctl, rec);
          } catch (e) {
            ctl.alert(`Kurtarılan çalışma kitabı açılamadı: ${(e as Error).message}`);
          }
        }
      }
      pruneHeartbeats();
      const list = await orphanDrafts();
      if (alive) setOffers(list);
    })();
    return () => {
      alive = false;
    };
  }, [ctl]);

  // Taslak yazıcı
  useEffect(() => {
    let timer: number | undefined;
    let cancelIdle: (() => void) | null = null;
    let firstChange = 0;
    let lastVersion = -1;
    let wasDirty = ctl.dirty;
    let disposed = false;
    let writing = false;

    const flush = () => {
      timer = undefined;
      if (disposed || writing || !ctl.dirty) return;
      const startVersion = ctl.version;
      const wbAtStart = ctl.wb;
      const gen = serializeDraft(ctl.wb);
      const step = (dl: IdleDeadline) => {
        cancelIdle = null;
        if (disposed) return;
        // Serileştirme sürerken kitap değiştiyse yeniden planla
        if (ctl.version !== startVersion || ctl.wb !== wbAtStart) {
          schedule();
          return;
        }
        let r = gen.next();
        while (!r.done && (dl.timeRemaining() > 4 || dl.didTimeout)) r = gen.next();
        if (!r.done) {
          cancelIdle = idle(step);
          return;
        }
        void save(r.value, startVersion);
      };
      cancelIdle = idle(step);
    };

    let warned = false;
    const save = async (draft: WorkbookDraft, version: number) => {
      writing = true;
      try {
        const ok = await writeDraft({ session: SESSION_ID, fileName: ctl.fileName, path: ctl.fileRef?.path, savedAt: Date.now(), sheetCount: draft.sheets.length, cellCount: draft.cellCount, draft });
        // Taslak yazılamadı (depolama dolu): kullanıcı bir kez uyarılır, sessizce yutulmaz.
        if (!ok && !warned && !disposed) {
          warned = true;
          ctl.infoBar = {
            kind: "warning",
            text: lastWriteQuota
              ? "Kurtarma taslağı kaydedilemedi: depolama alanı dolu. Çökme durumunda değişiklikler kurtarılamayabilir; çalışma kitabını kaydedin."
              : "Kurtarma taslağı kaydedilemedi. Çökme durumunda değişiklikler kurtarılamayabilir; çalışma kitabını kaydedin.",
          };
          ctl.emit();
        } else if (ok) warned = false;
        firstChange = 0;
      } catch {
        /* taslak yazılamazsa sessizce geç */
      } finally {
        writing = false;
      }
      // Yazarken kaydedildiyse taslak silinir; değiştiyse yeniden planlanır
      if (!ctl.dirty) void deleteDraft(SESSION_ID);
      else if (ctl.version !== version) schedule();
    };

    const schedule = () => {
      const now = Date.now();
      if (!firstChange) firstChange = now;
      window.clearTimeout(timer);
      timer = window.setTimeout(flush, now - firstChange >= MAX_WAIT_MS ? 0 : IDLE_MS);
    };

    const unsub = ctl.subscribe(() => {
      if (!ctl.dirty) {
        window.clearTimeout(timer);
        timer = undefined;
        cancelIdle?.();
        cancelIdle = null;
        firstChange = 0;
        lastVersion = ctl.version;
        // Kaydedildi (ya da yeni kitap açıldı): bu pencerenin taslağı her durumda silinir
        if (wasDirty) void deleteDraft(SESSION_ID);
        wasDirty = false;
        return;
      }
      wasDirty = true;
      if (ctl.version === lastVersion) return;
      lastVersion = ctl.version;
      schedule();
    });
    const onHide = () => {
      if (document.visibilityState === "hidden" && timer !== undefined) {
        window.clearTimeout(timer);
        flush();
      }
    };
    // Sayfada yakalanmamış bir hata oldu ("officehere:autosave"): bekleyen taslak hemen yazılır.
    const onFlush = () => {
      if (!ctl.dirty || writing || cancelIdle) return;
      window.clearTimeout(timer);
      flush();
    };
    document.addEventListener("visibilitychange", onHide);
    window.addEventListener("officehere:autosave", onFlush);
    return () => {
      disposed = true;
      unsub();
      window.clearTimeout(timer);
      cancelIdle?.();
      document.removeEventListener("visibilitychange", onHide);
      window.removeEventListener("officehere:autosave", onFlush);
    };
  }, [ctl]);

  const offer = offers[0] ?? null;
  const next = () => setOffers((list) => list.slice(1));

  const restore = useCallback(async () => {
    if (!offer) return;
    next();
    // Bu pencere kullanılıyorsa taslak yeni pencerede açılır (geçerli kitap ezilmez)
    if (!ctl.pristine) {
      const url = new URL(window.location.href);
      url.searchParams.set(RESTORE_PARAM, offer.session);
      if (window.open(url.toString(), "_blank")) return;
      ctl.alert("Kurtarılan çalışma kitabı yeni pencerede açılamadı. Açılır pencerelere izin verip yeniden deneyin.");
      setOffers((list) => [offer, ...list]);
      return;
    }
    const rec = await readDraft(offer.session);
    if (!rec) {
      ctl.alert("Kurtarılan çalışma kitabı okunamadı.");
      return;
    }
    try {
      await loadInto(ctl, rec);
    } catch (e) {
      ctl.alert(`Kurtarılan çalışma kitabı açılamadı: ${(e as Error).message}`);
    }
  }, [ctl, offer]);

  const discard = useCallback(() => {
    if (!offer) return;
    void deleteDraft(offer.session);
    next();
  }, [offer]);

  // Kapat (×): çubukta anılan taslakların hepsi silinir; sonraki açılışta yeniden önerilmez.
  const dismiss = useCallback(() => {
    for (const o of offers) void deleteDraft(o.session);
    setOffers([]);
  }, [offers]);

  const bar = offer ? (
    <MessageBar
      kind="warning"
      onClose={dismiss}
      closeLabel={offers.length > 1 ? "Taslakları sil ve kapat" : "Taslağı sil ve kapat"}
      actions={
        <span className="tb-recover-actions">
          <button type="button" className="oh-btn is-primary" onClick={() => void restore()}>
            Geri Yükle
          </button>
          <button type="button" className="oh-btn" onClick={discard}>
            Sil
          </button>
        </span>
      }
    >
      <span className="tb-recover">
        <strong>Kurtarılan çalışma kitabı</strong> “{offer.fileName}” için kaydedilmemiş değişiklikler bulundu ({new Date(offer.savedAt).toLocaleString("tr-TR")}, {offer.sheetCount} sayfa).
        {offers.length > 1 && ` Kurtarılmayı bekleyen ${offers.length - 1} çalışma kitabı daha var.`}
      </span>
    </MessageBar>
  ) : null;

  return { offer, offers, bar, restore, discard };
}
