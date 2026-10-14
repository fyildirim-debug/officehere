// Açılışta belge kurtarma: başka pencereden aktarılan belge (birleştirme sonucu, kurtarılan taslak),
// kök hata sınırından kurtarma ve kapanmış pencerelerden kalan (sahipsiz) taslakların önerilmesi.
// Taslaklar IndexedDB'dedir (eşzamansız okunur).
import { useCallback, useEffect, useRef, useState } from "react";
import type { FileRef } from "@officehere/core";
import { saveStatus } from "@officehere/ui";
import { defaultSettings, type BelgeDocument } from "../model/document";
import { adoptDraft, orphanDrafts, pruneHeartbeats, readDraftByKey, readOwnDraft, removeDraft, type StoredDraft } from "./drafts";
import { openInNewWindow, takeHandoff } from "../mailmerge/data";
import { takeAutoRestore } from "../views/ErrorBoundary";

interface Options {
  load: (doc: BelgeDocument, name: string, ref: FileRef | null) => void;
  setDirty: (d: boolean) => void;
  isPristine: () => boolean;
  onPrint: () => void;
  notify: (t: string, kind?: "info" | "warning" | "error" | "success") => void;
}

/** Kurtarılan taslağın dosya bağlantısı: yalnız .docx ise aynı dosyaya kaydedilebilir (diğerlerinde Farklı Kaydet). */
function refOf(d: Pick<StoredDraft, "path">): FileRef | null {
  if (!d.path || !/\.docx$/i.test(d.path)) return null;
  const name = d.path.split(/[\\/]/).pop() ?? d.path;
  return { name, path: d.path };
}

/** Kurtarılan taslak diskteki dosyanın üzerine kendiliğinden yazılmaz: ilk kaydı kullanıcı yapar. */
const RECOVERED_HOLD = "Kurtarılan taslak; ilk kaydı siz onaylayın";

export function useDraftRecovery({ load, setDirty, isPristine, onPrint, notify }: Options) {
  const [drafts, setDrafts] = useState<StoredDraft[]>([]);
  const started = useRef(false);

  useEffect(() => {
    let alive = true;
    void (async () => {
      // Kök hata sınırından "Belgeyi Kurtar": bu pencerenin taslağı kendiliğinden geri yüklenir.
      if (takeAutoRestore()) {
        const d = await readOwnDraft();
        if (d && alive) {
          load({ content: d.content, settings: { ...defaultSettings(), ...d.settings } }, d.fileName, refOf(d));
          saveStatus.hold(RECOVERED_HOLD);
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
        const d = h.restoreKey ? await readDraftByKey(h.restoreKey) : null;
        if (d) await adoptDraft(d);
        if (h.then === "print") onPrint();
        return;
      }
      pruneHeartbeats();
      const list = await orphanDrafts();
      if (alive) setDrafts(list);
    })();
    return () => {
      alive = false;
    };
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const restore = useCallback(
    (d: StoredDraft) => {
      setDrafts((ds) => ds.filter((x) => x.key !== d.key));
      if (isPristine()) {
        load({ content: d.content, settings: { ...defaultSettings(), ...d.settings } }, d.fileName, refOf(d));
        saveStatus.hold(RECOVERED_HOLD);
        setDirty(true);
        void adoptDraft(d);
      } else if (!openInNewWindow({ content: d.content, settings: d.settings }, d.fileName, undefined, d.key)) {
        notify(`“${d.fileName}” yeni pencerede açılamadı. Yeni, boş bir belge penceresinde yeniden deneyin.`, "error");
        setDrafts((ds) => [d, ...ds]);
      }
    },
    [isPristine, load, setDirty, notify],
  );

  const discard = useCallback((d: StoredDraft) => {
    void removeDraft(d.key);
    setDrafts((ds) => ds.filter((x) => x.key !== d.key));
  }, []);

  const close = useCallback(() => setDrafts([]), []);

  return { drafts, restore, discard, close };
}
