// Durum çubuğu: sayfa, sözcük sayısı, dil, izleme durumu, görünüm düğmeleri ve yakınlaştırma.
import { DocumentOnePage20Regular, Globe20Regular, ReadingModeMobile20Regular, TextProofingTools20Regular, ScanText20Regular } from "@fluentui/react-icons";
import { StatusBar, StatusItem, ViewSwitcher, ZoomSlider, tipProps } from "@officehere/ui";
import { useCallback, useSyncExternalStore } from "react";
import { useEditorState } from "@tiptap/react";
import type { Editor } from "@tiptap/core";
import type { ViewOptions } from "../state/context";
import type { PaginationController } from "../editor/pagination";
import { countWords } from "../editor/commands";

interface Props {
  editor: Editor;
  controller: PaginationController;
  /** Yazdırma Düzeni'nde sayfa bilgisi gösterilir. */
  paged: boolean;
  words: number;
  tracking: boolean;
  view: ViewOptions;
  setView: (p: Partial<ViewOptions>) => void;
  onPageClick: () => void;
  onWordsClick: () => void;
  onZoomClick: () => void;
  lang: string;
}

const LANG_NAMES: Record<string, string> = { tr: "Türkçe (Türkiye)", "en-US": "İngilizce (ABD)", de: "Almanca (Almanya)" };

export function BelgeStatusBar({ editor, controller, paged, words, tracking, view, setView, onPageClick, onWordsClick, onZoomClick, lang }: Props) {
  const n = (x: number) => x.toLocaleString("tr-TR");
  // İmleç ve seçim burada izlenir; böylece her tuşta yalnız durum çubuğu yeniden çizilir.
  const sel = useEditorState({
    editor,
    selector: ({ editor: ed }) => {
      const { from, to, empty } = ed.state.selection;
      return { from, words: empty ? -1 : countWords(ed.state.doc.textBetween(from, to, "\n", " ")) };
    },
    equalityFn: (a, b) => !!a && !!b && a.from === b.from && a.words === b.words,
  });
  const info = useSyncExternalStore(
    useCallback((cb: () => void) => controller.subscribe(cb), [controller]),
    () => controller.info,
  );
  const pages = paged ? info.pageCount : 1;
  const page = paged ? controller.pageOf(sel?.from ?? 0) + 1 : 1;
  const selWords = sel?.words ?? -1;
  return (
    <StatusBar
      left={
        <>
          <StatusItem onClick={onPageClick} tip="Sayfa numarası. Gezinti bölmesini açmak için tıklayın.">
            Sayfa {n(Math.min(page, pages))} / {n(pages)}
          </StatusItem>
          <StatusItem onClick={onWordsClick} tip="Sözcük sayısı. Sözcük Sayısı iletişim kutusunu açmak için tıklayın.">
            {selWords >= 0 ? `${n(selWords)} / ${n(words)} sözcük` : `${n(words)} sözcük`}
          </StatusItem>
          <StatusItem tip="Yazım denetimi dili">
            <TextProofingTools20Regular className="bl-sb-icon" /> {LANG_NAMES[lang] ?? lang}
          </StatusItem>
          {tracking && <StatusItem tip="Değişiklikleri İzle açık">Değişiklikleri İzle: Açık</StatusItem>}
        </>
      }
      right={
        <>
          <button type="button" className={`oh-status-icon bl-focus-btn${view.focusMode ? " is-checked" : ""}`} onClick={() => setView({ focusMode: !view.focusMode })} aria-label="Odak" {...tipProps("Odak", undefined, "Dikkat dağıtıcı öğeleri gizleyin.")}>
            <ScanText20Regular />
            <span>Odak</span>
          </button>
          <ViewSwitcher
            active={view.mode}
            onChange={(id) => setView({ mode: id as ViewOptions["mode"] })}
            views={[
              { id: "read", icon: <ReadingModeMobile20Regular />, tip: "Okuma Modu" },
              { id: "print", icon: <DocumentOnePage20Regular />, tip: "Yazdırma Düzeni" },
              { id: "web", icon: <Globe20Regular />, tip: "Web Düzeni" },
            ]}
          />
          <ZoomSlider value={view.zoom} onChange={(z) => setView({ zoom: z })} onPercentClick={onZoomClick} />
        </>
      }
    />
  );
}
