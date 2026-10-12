/**
 * Otomatik Kurtarma: çalışma kitabı değiştikçe (kaydedilmemişken) birkaç
 * saniyelik boşluktan sonra taslak olarak IndexedDB'ye yazılır.
 * Serileştirme boşta kalma süresinde parça parça yapılır; bu sırada kitap
 * değişirse yarıda bırakılıp yeniden planlanır. Kaydedince ya da yeni kitap
 * açılınca taslak silinir. Açılışta kaydedilmemiş taslak varsa ileti
 * çubuğuyla "Kurtarılan çalışma kitabı" önerilir.
 */
import { useCallback, useEffect, useState } from "react";
import { MessageBar } from "@officehere/ui";
import type { Controller } from "../controller";
import { serializeDraft, workbookFromDraft, type WorkbookDraft } from "../io/draft";
import { announceLive, deleteDraft, listDraftMeta, liveSessions, newSessionId, pickDraft, readDraft, writeDraft } from "./draftStore";

export const SESSION_ID = newSessionId();
/** Son değişiklikten sonra beklenen boşluk. */
const IDLE_MS = 2500;
/** Sürekli düzenlemede en geç bu sürede bir yazılır. */
const MAX_WAIT_MS = 15000;

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

interface Offer {
  session: string;
  fileName: string;
  savedAt: number;
  sheetCount: number;
}

export function useAutoRecover(ctl: Controller) {
  const [offer, setOffer] = useState<Offer | null>(null);

  useEffect(() => announceLive(SESSION_ID), []);

  // Açılışta kurtarılacak taslak var mı?
  useEffect(() => {
    let alive = true;
    void (async () => {
      const [metas, live] = await Promise.all([listDraftMeta(), liveSessions()]);
      const d = pickDraft(metas, live, SESSION_ID);
      if (alive && d) setOffer(d);
    })();
    return () => {
      alive = false;
    };
  }, []);

  // Taslak yazıcı
  useEffect(() => {
    let timer: number | undefined;
    let cancelIdle: (() => void) | null = null;
    let firstChange = 0;
    let lastVersion = -1;
    let hasDraft = false;
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

    const save = async (draft: WorkbookDraft, version: number) => {
      writing = true;
      try {
        const ok = await writeDraft({ session: SESSION_ID, fileName: ctl.fileName, path: ctl.fileRef?.path, savedAt: Date.now(), sheetCount: draft.sheets.length, cellCount: draft.cellCount, draft });
        hasDraft = hasDraft || ok;
        firstChange = 0;
      } catch {
        /* taslak yazılamazsa sessizce geç */
      } finally {
        writing = false;
      }
      // Yazarken kaydedildiyse taslağı sil; değiştiyse yeniden planla
      if (!ctl.dirty && hasDraft) {
        hasDraft = false;
        void deleteDraft(SESSION_ID);
      } else if (ctl.version !== version && ctl.dirty) schedule();
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
        if (hasDraft) {
          hasDraft = false;
          void deleteDraft(SESSION_ID);
        }
        return;
      }
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
    document.addEventListener("visibilitychange", onHide);
    return () => {
      disposed = true;
      unsub();
      window.clearTimeout(timer);
      cancelIdle?.();
      document.removeEventListener("visibilitychange", onHide);
    };
  }, [ctl]);

  const restore = useCallback(async () => {
    if (!offer) return;
    const rec = await readDraft(offer.session);
    setOffer(null);
    if (!rec) {
      ctl.alert("Kurtarılan çalışma kitabı okunamadı.");
      return;
    }
    try {
      const wb = workbookFromDraft(rec.draft);
      ctl.load(wb, rec.fileName, rec.path ? { name: rec.fileName, path: rec.path } : null);
      ctl.dirty = true;
      ctl.pristine = false;
      ctl.infoBar = { kind: "info", text: "Kurtarılan çalışma kitabı geri yüklendi. Saklamak için kaydedin." };
      ctl.emit();
      void deleteDraft(offer.session);
    } catch (e) {
      ctl.alert(`Kurtarılan çalışma kitabı açılamadı: ${(e as Error).message}`);
    }
  }, [ctl, offer]);

  const discard = useCallback(() => {
    if (offer) void deleteDraft(offer.session);
    setOffer(null);
  }, [offer]);

  const bar = offer ? (
    <MessageBar
      kind="warning"
      onClose={discard}
      actions={
        <>
          <button type="button" className="oh-btn is-primary" onClick={() => void restore()}>
            Geri Yükle
          </button>
          <button type="button" className="oh-btn" onClick={discard}>
            Sil
          </button>
        </>
      }
    >
      <strong>Kurtarılan çalışma kitabı</strong> “{offer.fileName}” için kaydedilmemiş değişiklikler bulundu ({new Date(offer.savedAt).toLocaleString("tr-TR")}, {offer.sheetCount} sayfa).
    </MessageBar>
  ) : null;

  return { offer, bar, restore, discard };
}
