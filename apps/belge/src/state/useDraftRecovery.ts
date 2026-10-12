// Açılışta belge kurtarma: başka pencereden aktarılan belge (birleştirme sonucu, kurtarılan taslak),
// kök hata sınırından kurtarma ve kapanmış pencerelerden kalan (sahipsiz) taslakların önerilmesi.
import { useCallback, useEffect, useRef, useState } from "react";
import { defaultSettings, type BelgeDocument } from "../model/document";
import { adoptDraft, orphanDrafts, pruneHeartbeats, readDraftByKey, readOwnDraft, removeDraft, type StoredDraft } from "./drafts";
import { openInNewWindow, takeHandoff } from "../mailmerge/data";
import { takeAutoRestore } from "../views/ErrorBoundary";

interface Options {
  load: (doc: BelgeDocument, name: string, ref: null) => void;
  setDirty: (d: boolean) => void;
  isPristine: () => boolean;
  onPrint: () => void;
  notify: (t: string, kind?: "info" | "warning" | "error" | "success") => void;
}

export function useDraftRecovery({ load, setDirty, isPristine, onPrint, notify }: Options) {
  const [drafts, setDrafts] = useState<StoredDraft[]>([]);
  const started = useRef(false);

  useEffect(() => {
    // Kök hata sınırından "Belgeyi Kurtar": bu pencerenin taslağı kendiliğinden geri yüklenir.
    if (takeAutoRestore()) {
      const d = readOwnDraft();
      if (d) {
        load({ content: d.content, settings: { ...defaultSettings(), ...d.settings } }, d.fileName, null);
        setDirty(true);
        return;
      }
    }
    if (started.current) return;
    started.current = true;
    const h = takeHandoff();
    if (h) {
      load({ content: h.doc.content, settings: { ...defaultSettings(), ...h.doc.settings } }, h.name, null);
      setDirty(true);
      // Kurtarılan taslak bu pencereye devredilir (önce kendi anahtarına yazılır, sonra eskisi silinir).
      const d = h.restoreKey ? readDraftByKey(h.restoreKey) : null;
      if (d) adoptDraft(d);
      if (h.then === "print") onPrint();
      return;
    }
    pruneHeartbeats();
    void orphanDrafts().then(setDrafts);
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const restore = useCallback(
    (d: StoredDraft) => {
      setDrafts((ds) => ds.filter((x) => x.key !== d.key));
      if (isPristine()) {
        load({ content: d.content, settings: { ...defaultSettings(), ...d.settings } }, d.fileName, null);
        setDirty(true);
        adoptDraft(d);
      } else if (!openInNewWindow({ content: d.content, settings: d.settings }, d.fileName, undefined, d.key)) {
        notify(`“${d.fileName}” yeni pencerede açılamadı.`, "error");
        setDrafts((ds) => [d, ...ds]);
      }
    },
    [isPristine, load, setDirty, notify],
  );

  const discard = useCallback((d: StoredDraft) => {
    removeDraft(d.key);
    setDrafts((ds) => ds.filter((x) => x.key !== d.key));
  }, []);

  const close = useCallback(() => setDrafts([]), []);

  return { drafts, restore, discard, close };
}
