// Adres mektup birleştirme durumu: alıcı listesi, hariç tutulan kayıtlar, sonuç önizlemesi.
import { createContext, useCallback, useContext, useEffect, useState } from "react";
import type { Editor } from "@tiptap/core";
import { setFieldContext } from "../editor/fields";
import type { RecipientList } from "./data";
import type { DocSettings, MailMergeData } from "../model/document";

export interface MergeState {
  list: RecipientList | null;
  excluded: number[];
  preview: boolean;
  index: number;
  highlight: boolean;
}

export interface MergeApi extends MergeState {
  setList: (l: RecipientList | null) => void;
  setExcluded: (e: number[]) => void;
  setPreview: (v: boolean) => void;
  setIndex: (i: number) => void;
  setHighlight: (v: boolean) => void;
  /** Birleştirmeye girecek kayıtlar (hariç tutulanlar dışında). */
  included: () => Record<string, string>[];
}

/**
 * Veri kaynağı (alıcı listesi, hariç tutulanlar) belge ayarlarında saklanır: kaydet/aç, otomatik kurtarma
 * ve sayfa yenileme sonrası korunur. Önizleme durumu yalnız bu oturumdadır.
 */
export function useMailMerge(editor: Editor, remeasure: () => void, data: MailMergeData | null, updateSettings: (fn: (s: DocSettings) => DocSettings) => void): MergeApi {
  const [ui, setUi] = useState<Omit<MergeState, "list" | "excluded">>({ preview: false, index: 0, highlight: false });
  const s: MergeState = { ...ui, list: data?.list ?? null, excluded: data?.excluded ?? [] };
  const patch = useCallback((p: Partial<Omit<MergeState, "list" | "excluded">>) => setUi((x) => ({ ...x, ...p })), []);
  const setData = useCallback((d: MailMergeData | null) => updateSettings((x) => ({ ...x, mailMerge: d })), [updateSettings]);
  // Önizlemede birleştirme alanları geçerli kaydın değerlerini gösterir.
  useEffect(() => {
    const rec = s.preview && s.list ? (s.list.records[s.index] ?? null) : null;
    setFieldContext(null, { record: rec });
    // Değerler değişince sayfalama yeniden ölçülür.
    if (!editor.isDestroyed) requestAnimationFrame(remeasure);
  }, [editor, s.preview, s.index, s.list]); // eslint-disable-line react-hooks/exhaustive-deps
  return {
    ...s,
    setList: (list) => {
      setData(list ? { list, excluded: [] } : null);
      patch({ index: 0, preview: false });
    },
    setExcluded: (excluded) => s.list && setData({ list: s.list, excluded }),
    setPreview: (preview) => patch({ preview }),
    setIndex: (index) => patch({ index: Math.max(0, Math.min((s.list?.records.length ?? 1) - 1, index)) }),
    setHighlight: (highlight) => patch({ highlight }),
    included: () => (s.list?.records ?? []).filter((_, i) => !s.excluded.includes(i)),
  };
}

export const MergeContext = createContext<MergeApi | null>(null);

export function useMerge(): MergeApi {
  const c = useContext(MergeContext);
  if (!c) throw new Error("MergeContext yok");
  return c;
}
