// Adres mektup birleştirme durumu: alıcı listesi, hariç tutulan kayıtlar, sonuç önizlemesi.
import { createContext, useCallback, useContext, useEffect, useState } from "react";
import type { Editor } from "@tiptap/core";
import { setFieldContext } from "../editor/fields";
import type { RecipientList } from "./data";

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

export function useMailMerge(editor: Editor, remeasure: () => void): MergeApi {
  const [s, setS] = useState<MergeState>({ list: null, excluded: [], preview: false, index: 0, highlight: false });
  const patch = useCallback((p: Partial<MergeState>) => setS((x) => ({ ...x, ...p })), []);
  // Önizlemede birleştirme alanları geçerli kaydın değerlerini gösterir.
  useEffect(() => {
    const rec = s.preview && s.list ? (s.list.records[s.index] ?? null) : null;
    setFieldContext(null, { record: rec });
    // Değerler değişince sayfalama yeniden ölçülür.
    if (!editor.isDestroyed) requestAnimationFrame(remeasure);
  }, [editor, s.preview, s.index, s.list]); // eslint-disable-line react-hooks/exhaustive-deps
  return {
    ...s,
    setList: (list) => patch({ list, excluded: [], index: 0, preview: false }),
    setExcluded: (excluded) => patch({ excluded }),
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
