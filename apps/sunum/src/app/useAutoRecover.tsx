/**
 * Otomatik Kurtarma: sunu değiştikçe (kirliyken) birkaç saniyelik boşluktan
 * sonra taslak olarak yazılır; kaydedilince/yeni belge açılınca taslak
 * silinir. Açılışta kaydedilmemiş bir taslak bulunursa ileti çubuğuyla
 * "Kurtarılan sunu" önerilir.
 */
import { useCallback, useEffect, useState } from "react";
import { MessageBar } from "@officehere/ui";
import type { Actions } from "../state/actions";
import { documentPres } from "../state/masterActions";
import { announceLive, deleteDraft, listDrafts, liveSessions, newSessionId, pickDraft, writeDraft, type Draft } from "./draftStore";
import { loadPresentation } from "./fileOps";

/** Bu pencerenin oturum kimliği. */
export const SESSION_ID = newSessionId();
/** Son değişiklikten sonra beklenen boşluk (ms). */
const IDLE_MS = 2000;
/** Sürekli düzenlemede en geç bu kadar sürede bir yazılır (ms). */
const MAX_WAIT_MS = 10000;

type IdleWindow = Window & { requestIdleCallback?: (cb: () => void, o?: { timeout: number }) => number };

export function useAutoRecover(a: Actions) {
  const [offer, setOffer] = useState<Draft | null>(null);

  useEffect(() => announceLive(SESSION_ID), []);

  // Açılışta kurtarılacak taslak var mı?
  useEffect(() => {
    let alive = true;
    void (async () => {
      const [all, live] = await Promise.all([listDrafts(), liveSessions()]);
      const d = pickDraft(all, live, SESSION_ID);
      if (alive && d) setOffer(d);
    })();
    return () => {
      alive = false;
    };
  }, []);

  // Taslak yazıcı.
  useEffect(() => {
    let timer: number | undefined;
    let firstChange = 0;
    let lastPres = a.s.pres;
    let hasDraft = false;
    let writing = false;
    let disposed = false;

    const flush = async () => {
      timer = undefined;
      if (disposed || writing || !a.s.dirty) return;
      writing = true;
      firstChange = 0;
      try {
        const s = a.s;
        const ok = await writeDraft({ session: SESSION_ID, fileName: s.fileName, path: s.fileRef?.path, slideIdx: s.masterBase?.slideIdx ?? s.slideIdx, savedAt: Date.now(), pres: documentPres(a) });
        hasDraft = hasDraft || ok;
      } catch {
        /* taslak yazılamazsa sessizce geç */
      } finally {
        writing = false;
      }
      // Yazarken belge kaydedildiyse taslağı temizle.
      if (!a.s.dirty && hasDraft) {
        hasDraft = false;
        void deleteDraft(SESSION_ID);
      }
    };

    const schedule = () => {
      const now = Date.now();
      if (!firstChange) firstChange = now;
      window.clearTimeout(timer);
      const wait = now - firstChange >= MAX_WAIT_MS ? 0 : IDLE_MS;
      timer = window.setTimeout(() => {
        const w = window as IdleWindow;
        if (w.requestIdleCallback) w.requestIdleCallback(() => void flush(), { timeout: 1500 });
        else void flush();
      }, wait);
    };

    const unsub = a.store.subscribe(() => {
      const s = a.s;
      if (!s.dirty) {
        lastPres = s.pres;
        window.clearTimeout(timer);
        timer = undefined;
        firstChange = 0;
        if (hasDraft) {
          hasDraft = false;
          void deleteDraft(SESSION_ID);
        }
        return;
      }
      if (s.pres === lastPres) return;
      lastPres = s.pres;
      schedule();
    });
    const onHide = () => {
      if (document.visibilityState === "hidden" && timer !== undefined) {
        window.clearTimeout(timer);
        void flush();
      }
    };
    document.addEventListener("visibilitychange", onHide);
    window.addEventListener("pagehide", onHide);
    return () => {
      disposed = true;
      unsub();
      window.clearTimeout(timer);
      document.removeEventListener("visibilitychange", onHide);
      window.removeEventListener("pagehide", onHide);
    };
  }, [a]);

  const restore = useCallback(() => {
    if (!offer) return;
    loadPresentation(a, offer.pres, offer.fileName, offer.path ? { name: `${offer.fileName}.pptx`, path: offer.path } : null);
    const idx = Math.max(0, Math.min(offer.pres.slides.length - 1, offer.slideIdx));
    a.store.set({ dirty: true, slideIdx: idx, slideSel: offer.pres.slides[idx] ? [offer.pres.slides[idx].id] : [] });
    void deleteDraft(offer.session);
    setOffer(null);
    a.message("Kurtarılan sunu geri yüklendi. Saklamak için kaydedin.", "info");
  }, [a, offer]);

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
          <button type="button" className="oh-btn is-primary" onClick={restore}>
            Geri Yükle
          </button>
          <button type="button" className="oh-btn" onClick={discard}>
            Sil
          </button>
        </>
      }
    >
      <strong>Kurtarılan sunu</strong> “{offer.fileName}” için kaydedilmemiş değişiklikler bulundu ({new Date(offer.savedAt).toLocaleString("tr-TR")}, {offer.pres.slides.length} slayt).
    </MessageBar>
  ) : null;

  return { offer, bar, restore, discard };
}
