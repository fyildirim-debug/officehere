// Düzenlemeyi Kısıtla: korumanın düzenleyicilere bildirilmesi ve ayar değişikliklerinin süzülmesi.
import { useEffect } from "react";
import type { Editor } from "@tiptap/core";
import { getUserName } from "@officehere/ui";
import type { DocSettings } from "../model/document";
import { protectionSource } from "../editor/protection";
import { isTracking, setTracking } from "../editor/trackChanges";

/** Koruma altında izin verilmeyen ayar değişikliği yok sayılır (açıklama kipinde açıklamalar serbest). */
export function guardSettings(prev: DocSettings, next: DocSettings): DocSettings {
  const p = prev.protection;
  if (!p?.enforced || p.mode === "trackedChanges") return next;
  const allowed = new Set<string>(["protection", ...(p.mode === "comments" ? ["comments"] : [])]);
  const changed = (Object.keys({ ...prev, ...next }) as (keyof DocSettings)[]).filter((k) => next[k] !== prev[k]);
  return changed.every((k) => allowed.has(k)) ? next : prev;
}

export function useProtection(editors: (Editor | null)[], settings: DocSettings) {
  protectionSource.get = () => settings.protection ?? null;
  useEffect(() => {
    // Düzenlenebilirlik durumunun yenilenmesi için boş işlem.
    for (const ed of editors) if (ed && !ed.isDestroyed) ed.view.dispatch(ed.state.tr.setMeta("belgeProtection", true));
    const p = settings.protection;
    const main = editors[0];
    if (main && !main.isDestroyed && p?.enforced && p.mode === "trackedChanges" && !isTracking(main.state)) setTracking(main.view, true, getUserName());
  }, [settings.protection, ...editors]); // eslint-disable-line react-hooks/exhaustive-deps
}
