/** Uygulamanın tüm iletişim kutuları (App.tsx'i sade tutmak için ayrı). */
import { Dialog, DialogButton } from "@officehere/ui";
import type { ChartData, PicShape } from "../model/types";
import type { Actions } from "../state/actions";
import type { AppState } from "../state/store";
import type { TextInfo } from "../state/textActions";
import type { AppCommands, DialogKind } from "../ribbon/types";
import type { ShowStroke } from "../show/ShowInk";
import { applyLink, type LinkContext } from "../state/linkActions";
import { insertChart, updateChart } from "../state/mediaActions";
import { applySubtype } from "../model/chart";
import { FontDialog, ParagraphDialog } from "../dialogs/TextDialogs";
import { DateTimeDialog, HeaderFooterDialog, IconsDialog, InsertTableDialog, SymbolDialog } from "../dialogs/InsertDialogs";
import { AboutDialog, FindDialog, ReplaceFontsDialog, ShortcutsDialog, SizePosDialog, SlideSizeDialog, ZoomDialog } from "../dialogs/MiscDialogs";
import { CustomShowsDialog, RehearsalDialog, ScaleChoiceDialog, ShowSetupDialog } from "../dialogs/ShowDialogs";
import type { ShowSettings } from "../show/settings";
import type { RehearsalResult } from "../show/Rehearsal";
import { LinkDialog } from "../dialogs/LinkDialog";
import { RenameDialog, type RenameRequest } from "../dialogs/RenameDialog";
import { ChartDataDialog, InsertChartDialog } from "../dialogs/ChartDialogs";

/** Parametreli iletişim kutuları. */
export type ExtraDialog =
  | { kind: "link"; ctx: LinkContext }
  | { kind: "rename"; req: RenameRequest }
  | { kind: "ink"; ink: Record<string, ShowStroke[]> }
  | { kind: "chartType"; changeId?: string; current?: ChartData }
  | { kind: "chartData"; id: string; chart: ChartData }
  | { kind: "media"; shape: PicShape }
  | { kind: "customShows" }
  | { kind: "rehearsal"; result: RehearsalResult }
  | { kind: "scaleChoice"; w: number; h: number };

export interface AppDialogsProps {
  a: Actions;
  state: AppState;
  info: TextInfo | null;
  cmd: AppCommands;
  dialog: DialogKind | null;
  closeDialog: () => void;
  extra: ExtraDialog | null;
  setExtra: (d: ExtraDialog | null) => void;
  showSettings: ShowSettings;
  setShowSettings: (s: ShowSettings) => void;
  /** Komut aramasından gelen "Bul" metni. */
  findQuery?: string;
}

export function AppDialogs({ a, state, info, cmd, dialog, closeDialog, extra, setExtra, showSettings, setShowSettings, findQuery }: AppDialogsProps) {
  const closeExtra = () => setExtra(null);
  const insertText = (t: string) => {
    if (a.s.editing && a.textHandle) a.textHandle.insert(t);
    else a.pasteText(t);
  };
  return (
    <>
      <FontDialog open={dialog === "font"} onClose={closeDialog} actions={a} info={info} />
      <ParagraphDialog open={dialog === "paragraph"} onClose={closeDialog} actions={a} info={info} />
      <InsertTableDialog open={dialog === "insertTable"} onClose={closeDialog} actions={a} />
      <HeaderFooterDialog open={dialog === "headerFooter"} onClose={closeDialog} actions={a} />
      <SymbolDialog open={dialog === "symbol"} onClose={closeDialog} onInsert={insertText} />
      <DateTimeDialog open={dialog === "dateTime"} onClose={closeDialog} onInsert={insertText} />
      <IconsDialog open={dialog === "icons"} onClose={closeDialog} actions={a} />
      <SlideSizeDialog open={dialog === "slideSize"} onClose={closeDialog} actions={a} />
      <FindDialog open={dialog === "find" || dialog === "replace"} replaceMode={dialog === "replace"} onClose={closeDialog} actions={a} initialQuery={dialog === "find" ? findQuery : undefined} />
      <ReplaceFontsDialog open={dialog === "replaceFonts"} onClose={closeDialog} actions={a} />
      <ZoomDialog open={dialog === "zoom"} onClose={closeDialog} zoom={state.zoom} onZoom={cmd.zoomTo} />
      <SizePosDialog open={dialog === "sizePos"} onClose={closeDialog} actions={a} />
      <ShowSetupDialog open={dialog === "showSetup"} onClose={closeDialog} value={showSettings} onChange={setShowSettings} pres={state.pres} />
      <CustomShowsDialog open={extra?.kind === "customShows"} onClose={closeExtra} actions={a} onShow={cmd.startCustomShow} />
      <RehearsalDialog result={extra?.kind === "rehearsal" ? extra.result : null} onClose={closeExtra} actions={a} />
      <ScaleChoiceDialog req={extra?.kind === "scaleChoice" ? { w: extra.w, h: extra.h } : null} onClose={closeExtra} actions={a} />
      <ShortcutsDialog open={dialog === "shortcuts"} onClose={closeDialog} />
      <AboutDialog open={dialog === "about"} onClose={closeDialog} />
      <Dialog
        open={extra?.kind === "ink"}
        title="OfficeHere Sunum"
        onClose={closeExtra}
        width={380}
        footer={
          <>
            <DialogButton
              primary
              onClick={() => {
                if (extra?.kind === "ink") a.keepShowInk(extra.ink);
                closeExtra();
              }}
            >
              Tut
            </DialogButton>
            <DialogButton onClick={closeExtra}>At</DialogButton>
          </>
        }
      >
        <p style={{ margin: "4px 0 8px" }}>Mürekkep ek açıklamalarınızı tutmak istiyor musunuz?</p>
      </Dialog>
      <RenameDialog request={extra?.kind === "rename" ? extra.req : null} onClose={closeExtra} />
      <LinkDialog context={extra?.kind === "link" ? extra.ctx : null} pres={state.pres} onApply={(c, l, t) => applyLink(a, c, l, t)} onClose={closeExtra} />
      <InsertChartDialog
        open={extra?.kind === "chartType"}
        current={extra?.kind === "chartType" ? extra.current : undefined}
        theme={state.pres.theme}
        onClose={closeExtra}
        onPick={(t, patch) => {
          if (extra?.kind !== "chartType") return;
          if (extra.changeId) {
            updateChart(a, extra.changeId, (c) => applySubtype(c, patch.type ?? t, patch));
            closeExtra();
            return;
          }
          const id = insertChart(a, t, patch);
          const sh = id ? a.selectedShapes()[0] : undefined;
          // PowerPoint gibi: ekledikten sonra veri tablosu açılır.
          if (id && sh?.kind === "chart") setExtra({ kind: "chartData", id, chart: sh.chart });
          else closeExtra();
        }}
      />
      <ChartDataDialog chart={extra?.kind === "chartData" ? extra.chart : null} onClose={closeExtra} onApply={(c) => extra?.kind === "chartData" && updateChart(a, extra.id, () => c)} />
      <Dialog open={extra?.kind === "media"} title={extra?.kind === "media" && extra.shape.media?.kind === "audio" ? "Ses Önizleme" : "Video Önizleme"} onClose={closeExtra} width={extra?.kind === "media" && extra.shape.media?.kind === "audio" ? 380 : 680} footer={<DialogButton onClick={closeExtra}>Kapat</DialogButton>}>
        {extra?.kind === "media" && extra.shape.media && (extra.shape.media.kind === "video" ? <video src={extra.shape.media.src} poster={extra.shape.src} controls autoPlay style={{ width: "100%", maxHeight: 420, background: "#000" }} /> : <audio src={extra.shape.media.src} controls autoPlay style={{ width: "100%" }} />)}
      </Dialog>
    </>
  );
}
