/** Durum çubuğu: slayt numarası, dil, Notlar, Açıklamalar, görünüm düğmeleri, yakınlaştırma. */
import { BookOpen20Regular, Comment20Regular, Globe20Regular, Notepad20Regular, SlideGrid20Regular, SlidePlay20Regular, SlideText20Regular, ZoomFit20Regular } from "@fluentui/react-icons";
import { StatusBar, StatusItem, ViewSwitcher, ZoomSlider, tipProps } from "@officehere/ui";
import type { Actions } from "../state/actions";
import { langLabel } from "../panes/ToolPanes";
import type { AppState } from "../state/store";

export function StatusArea({ state, actions: a, onSetView, onStartShow, onComments, onZoomDialog }: { state: AppState; actions: Actions; onSetView: (v: AppState["view"]) => void; onStartShow: () => void; onComments: () => void; onZoomDialog: () => void }) {
  const n = state.pres.slides.length;
  return (
    <StatusBar
      left={
        <>
          <StatusItem>{state.view === "master" ? "Asıl Slayt" : n ? `Slayt ${state.slideIdx + 1} / ${n}` : "Slayt yok"}</StatusItem>
          <StatusItem tip="Dil (yazım denetimi)" onClick={() => a.store.set((s) => ({ pane: s.pane === "spelling" ? null : "spelling" }))}>
            <span className="status-btn status-lang">
              <Globe20Regular /> {langLabel(state.proofLang)}
            </span>
          </StatusItem>
        </>
      }
      right={
        <>
          <StatusItem
            onClick={() => {
              if (state.view !== "normal") onSetView("normal");
              a.store.set({ showNotes: !state.showNotes });
            }}
            tip="Notlar"
          >
            <span className="status-btn" style={{ fontWeight: state.showNotes ? 600 : undefined }}>
              <Notepad20Regular /> Notlar
            </span>
          </StatusItem>
          <StatusItem onClick={onComments} tip="Açıklamalar">
            <span className="status-btn" style={{ fontWeight: state.pane === "comments" ? 600 : undefined }}>
              <Comment20Regular /> Açıklamalar
            </span>
          </StatusItem>
          <ViewSwitcher
            active={state.view}
            onChange={(v) => {
              if (v === "show") onStartShow();
              else onSetView(v as AppState["view"]);
            }}
            views={[
              { id: "normal", icon: <SlideText20Regular />, tip: "Normal" },
              { id: "sorter", icon: <SlideGrid20Regular />, tip: "Slayt Sıralayıcısı" },
              { id: "reading", icon: <BookOpen20Regular />, tip: "Okuma Görünümü" },
              { id: "show", icon: <SlidePlay20Regular />, tip: "Slayt Gösterisi" },
            ]}
          />
          <ZoomSlider value={state.zoom} min={10} max={400} onChange={(z) => a.store.set({ zoom: z, fit: false })} onPercentClick={onZoomDialog} />
          <button type="button" className="status-fit" aria-label="Slaytı geçerli pencereye sığdır" {...tipProps("Slaytı geçerli pencereye sığdır")} onClick={() => a.store.set({ fit: true })}>
            <ZoomFit20Regular />
          </button>
        </>
      }
    />
  );
}
