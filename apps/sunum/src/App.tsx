/** OfficeHere Sunum — uygulama kökü. */
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { ArrowRedo20Regular, ArrowUndo20Regular, Save20Regular, SlidePlay20Regular } from "@fluentui/react-icons";
import { AppShell, CommandSearch, MessageBar, QatButton, Ribbon, TitleBar, useContextMenu, useDocumentLifecycle, type OfficeBackstagePage } from "@officehere/ui";
import { downloadBlob } from "@officehere/core";
import { blankPresentation } from "./model/presentation";
import { Canvas } from "./editor/Canvas";
import { ThumbnailPane } from "./panes/ThumbnailPane";
import { NotesPane, Splitter } from "./panes/NotesPane";
import { BackgroundPane, ShapePane } from "./panes/FormatPanes";
import { AnimationPane, CommentsPane, SelectionPane } from "./panes/ListPanes";
import { ReusePane, SpellingPane } from "./panes/ToolPanes";
import { AccessibilityPane } from "./panes/AccessibilityPane";
import { useStore, type AppState } from "./state/store";
import { textInfo } from "./state/textActions";
import { useKeyboard } from "./app/keyboard";
import { insertImageFiles, useClipboard } from "./app/clipboard";
import { loadPresentation, newFromTemplate, openBytes, openDialog, openFromPath, openRecent, save } from "./app/fileOps";
import { exportPdfNative, exportPngs, outlineText, PrintRoot } from "./app/exporting";
import { masterMenu, sectionMenu, shapeMenu, slideAreaMenu, thumbMenu } from "./app/menus";
import { tableCellMenu, textEditMenu } from "./app/textMenus";
import { MasterThumbPane } from "./panes/MasterThumbPane";
import { documentPres } from "./state/masterActions";
import { StatusArea } from "./app/StatusArea";
import { FileBackstage } from "./app/FileBackstage";
import { OutlinePane } from "./panes/OutlinePane";
import type { DialogKind, RibbonCtx } from "./ribbon/types";
import { SlideShow } from "./show/SlideShow";
import { NotesPageView, SorterView } from "./views/SorterView";
import { DEFAULT_SHOW_SETTINGS, type ShowSettings } from "./show/settings";
import type { ShowStroke } from "./show/ShowInk";
import type { RehearsalResult } from "./show/Rehearsal";
import { AppDialogs, type ExtraDialog } from "./app/AppDialogs";
import { appSession, recoveryHandlers, type ShowState } from "./app/appState";
import { useAppCommands } from "./app/useAppCommands";
import { buildRibbonTabs } from "./app/ribbonTabs";
import { ErrorBoundary } from "./app/ErrorBoundary";
import { forgetOwnDraft, useAutoRecover } from "./app/useAutoRecover";
import { useFileAutoSave } from "./app/useFileAutoSave";
import { extraCommands } from "./app/extraCommands";
import "./styles/slide.css";
import "./styles/editor.css";
import "./styles/app.css";
import "./styles/show.css";
import "./styles/print.css";
import "./styles/extra.css";
import "./styles/theme.css";
import "./styles/round4.css";

export function App() {
  const { store, actions: a } = useMemo(() => appSession(), []);
  const state = useStore(store);
  const [backstage, setBackstage] = useState<false | true | OfficeBackstagePage>(false);
  const [dialog, setDialog] = useState<DialogKind | null>(null);
  const [show, setShow] = useState<ShowState | null>(null);
  const [printing, setPrinting] = useState<null | "print" | "pdf">(null);
  const [collapsed, setCollapsed] = useState(false);
  const [guides, setGuides] = useState({ h: [360], v: [640] });
  const [extra, setExtra] = useState<ExtraDialog | null>(null);
  const [showSettings, setShowSettings] = useState<ShowSettings>(DEFAULT_SHOW_SETTINGS);
  const [findQuery, setFindQuery] = useState("");
  const ctxMenu = useContextMenu();
  const blockedRef = useRef(false);
  blockedRef.current = !!dialog || !!extra || !!backstage || (!!show && !show.preview);
  const recovery = useMemo(() => recoveryHandlers(), []);
  const autoRecover = useAutoRecover(a);
  useFileAutoSave(a, state.fileRef?.path);

  const lifecycle = useDocumentLifecycle({
    title: state.fileName,
    dirty: state.dirty,
    path: state.fileRef?.path,
    onOpenPath: (p) => void openFromPath(a, p),
    onSave: () => save(a),
    onClose: forgetOwnDraft,
  });

  // Eylemlerin istediği iletişim kutusu (ör. eylem düğmesi çizilince Eylem Ayarları).
  useEffect(() => {
    if (!state.dialogRequest) return;
    setDialog(state.dialogRequest as DialogKind);
    a.store.set({ dialogRequest: null });
  }, [state.dialogRequest, a]);

  // Kılavuzlar slayt boyutu değişince ortalanır.
  useEffect(() => setGuides({ h: [state.pres.height / 2], v: [state.pres.width / 2] }), [state.pres.width, state.pres.height]);

  const openDialogFn = useCallback((d: DialogKind) => setDialog(d), []);
  const cmd = useAppCommands({ a, setDialog, setExtra, setShow, setBackstage, setPrinting, setFindQuery, showSettings });

  const newDoc = useCallback(async () => {
    if (!(await lifecycle.confirmDiscard())) return;
    newFromTemplate(a, "blank");
    setBackstage(false);
  }, [a, lifecycle]);

  useKeyboard(a, {
    save: () => void save(a),
    saveAs: () => void save(a, true),
    open: async () => {
      if (await lifecycle.confirmDiscard()) await openDialog(a);
    },
    newDoc: () => void newDoc(),
    print: cmd.print,
    showFromStart: () => cmd.startShow("start"),
    showFromCurrent: () => cmd.startShow("current"),
    find: () => cmd.find(),
    replace: () => setDialog("replace"),
    link: () => cmd.insertLink(),
    spelling: () => cmd.spelling(),
    isBlocked: () => blockedRef.current,
  });
  useClipboard(a, useCallback(() => blockedRef.current, []));

  // Metin bilgisi (şerit durumları için).
  const info = useMemo(() => textInfo(a), [a, state.pres, state.sel, state.editing, state.textTick, state.cellSel]); // eslint-disable-line react-hooks/exhaustive-deps

  const searchExtra = useMemo(
    () =>
      extraCommands(
        a,
        {
          save: () => void save(a),
          saveAs: () => void save(a, true),
          print: cmd.print,
          exportPdf: cmd.exportPdf,
          newDoc: () => void newDoc(),
          open: async () => {
            if (await lifecycle.confirmDiscard()) await openDialog(a);
          },
          backstage: (page) => {
            a.stopEditing();
            setBackstage(page);
          },
          dialog: (k) => setDialog(k),
        },
        cmd.previewTransition,
      ),
    [a, cmd, newDoc, lifecycle],
  );
  const ctx: RibbonCtx = { state, actions: a, text: info, openDialog: openDialogFn, cmd, show: { settings: showSettings, set: setShowSettings } };
  const tabs = buildRibbonTabs(ctx);
  const editView = state.view === "normal" || state.view === "outline" || state.view === "master";
  const masterView = state.view === "master";
  const activeTab = tabs.find((t) => t.id === state.ribbonTab && !t.hidden) ? state.ribbonTab : "home";

  const closeDialog = () => setDialog(null);

  const onDrop = (e: React.DragEvent) => {
    const files = Array.from(e.dataTransfer.files ?? []);
    if (!files.length) return;
    e.preventDefault();
    const pptx = files.find((f) => f.name.toLowerCase().endsWith(".pptx"));
    if (pptx) {
      void (async () => {
        if (await lifecycle.confirmDiscard()) await openBytes(a, await pptx.arrayBuffer(), pptx.name, null);
      })();
      return;
    }
    void insertImageFiles(a, files);
  };

  const sidePane = (() => {
    const close = () => a.store.set({ pane: null });
    switch (state.pane) {
      case "background":
        return <BackgroundPane state={state} actions={a} onClose={close} />;
      case "shape":
        return <ShapePane state={state} actions={a} onClose={close} />;
      case "animation":
        return <AnimationPane state={state} actions={a} onClose={close} onPlay={() => cmd.previewAnimations()} />;
      case "selection":
        return <SelectionPane state={state} actions={a} onClose={close} />;
      case "comments":
        return <CommentsPane state={state} actions={a} onClose={close} />;
      case "spelling":
        return <SpellingPane state={state} actions={a} onClose={close} />;
      case "reuse":
        return <ReusePane state={state} actions={a} onClose={close} />;
      case "accessibility":
        return <AccessibilityPane state={state} actions={a} onClose={close} />;
      default:
        return null;
    }
  })();

  // Kip ipuçları slayt alanının üstünde yüzer (yerleşimi kaydırmaz).
  const canvasHint = state.insert
    ? state.insert.kind === "textbox"
      ? "Metin kutusu eklemek için slayta tıklayın ya da sürükleyin."
      : "Şekli çizmek için slaytta sürükleyin (Shift: oranı koru). Vazgeçmek için Esc."
    : state.cropId
      ? "Kırpmak için kenar tutamaçlarını, resmi kaydırmak için içini sürükleyin. Bitirmek için Esc."
      : state.ink
        ? state.ink.tool === "eraser"
          ? "Silmek için mürekkep çizimlerinin üzerinden geçin. Bitirmek için Esc."
          : "Çizmek için slaytta sürükleyin. Bitirmek için Esc."
        : null;

  const canvasMenu = (e: React.MouseEvent, kind: "shape" | "slide" | "text" | "cell") => {
    const one = a.selectedShapes()[0];
    if (kind === "text") ctxMenu.open(e, textEditMenu(a, cmd, openDialogFn));
    else if (kind === "cell" && one?.kind === "table") ctxMenu.open(e, tableCellMenu(a, cmd, openDialogFn, one));
    else if (kind === "shape" || kind === "cell") ctxMenu.open(e, shapeMenu(a, cmd, () => setDialog("sizePos")));
    else ctxMenu.open(e, slideAreaMenu(a, cmd, () => a.store.set({ showRuler: !a.s.showRuler }), a.s.showRuler));
  };

  const bound = (region: string, node: React.ReactNode, onReset?: () => void) => (
    <ErrorBoundary region={region} recovery={recovery} onReset={onReset}>
      {node}
    </ErrorBoundary>
  );

  let main: React.ReactNode;
  if (state.view === "sorter") {
    main = bound(
      "Slayt Sıralayıcısı",
      <SorterView
        state={state}
        actions={a}
        onOpenSlide={(i) => {
          a.goToSlide(i);
          a.store.set({ view: "normal" });
        }}
        onContextMenu={(e) => ctxMenu.open(e, thumbMenu(a, cmd, true))}
        onSectionMenu={(e, id) => ctxMenu.open(e, sectionMenu(a, cmd, id))}
        onRenameSection={cmd.renameSection}
      />,
    );
  } else if (state.view === "notes") {
    main = bound("Not Sayfası", <NotesPageView state={state} actions={a} />);
  } else if (state.view === "reading") {
    main = (
      <div className="app-center" style={{ position: "relative" }}>
        {bound(
          "Okuma Görünümü",
          <SlideShow
            key={`read-${state.slideIdx}`}
            pres={state.pres}
            startIndex={state.slideIdx}
            mode="window"
            onExit={(i, ink) => {
              a.goToSlide(Math.min(i, a.s.pres.slides.length - 1));
              a.store.set({ view: "normal" });
              if (ink && Object.keys(ink).length) setExtra({ kind: "ink", ink });
            }}
            onViewChange={(v) => a.store.set({ view: v })}
          />,
          () => a.store.set({ view: "normal" }),
        )}
      </div>
    );
  } else {
    main = (
      <>
        {bound(
          "Küçük resim bölmesi",
          masterView ? (
            <MasterThumbPane state={state} actions={a} width={state.thumbWidth} onContextMenu={(e) => ctxMenu.open(e, masterMenu(a, cmd))} />
          ) : state.view === "outline" ? (
            <OutlinePane state={state} actions={a} width={Math.max(300, state.thumbWidth * 1.5)} />
          ) : (
            <ThumbnailPane state={state} actions={a} width={state.thumbWidth} onContextMenu={(e, onSlide) => ctxMenu.open(e, thumbMenu(a, cmd, onSlide))} onSectionMenu={(e, id) => ctxMenu.open(e, sectionMenu(a, cmd, id))} onRenameSection={cmd.renameSection} />
          ),
        )}
        <Splitter dir="v" onDrag={(d) => a.store.set((s) => ({ thumbWidth: Math.max(110, Math.min(420, s.thumbWidth + (s.view === "outline" ? d / 1.5 : d))) }))} />
        <div className="app-center" style={{ position: "relative" }} onDragOver={(e) => e.preventDefault()} onDrop={onDrop}>
          {bound(
            "Slayt alanı",
            <Canvas
              state={state}
              actions={a}
              onContextMenu={canvasMenu}
              onPlaceholderAction={(action, id) => {
                a.pendingPlaceholder = id;
                a.select([id]);
                if (action === "table") setDialog("insertTable");
                else if (action === "chart") cmd.insertChart();
                else if (action === "video") cmd.insertMedia("video");
                else if (action === "icons") setDialog("icons");
                else cmd.insertPictures();
              }}
              onShapeDoubleClick={(s) => {
                if (s.kind === "chart") cmd.editChartData();
                else if (s.kind === "pic" && s.media) a.store.set({ ribbonTab: "playback" });
                else if (s.kind === "pic") a.store.set({ ribbonTab: "pictureFormat" });
                else if (s.kind === "sp" && !s.text) a.store.set({ ribbonTab: "shapeFormat" });
              }}
              showAnimBadges={activeTab === "animations" || state.pane === "animation"}
              guides={guides}
              onGuidesChange={setGuides}
              onZoomChange={(zoom, fit) => a.store.set({ zoom, fit })}
            />,
            () => a.store.set({ editing: null, insert: null, cropId: null }),
          )}
          {state.showNotes && !masterView && (
            <>
              <Splitter dir="h" onDrag={(d) => a.store.set((s) => ({ notesHeight: Math.max(30, Math.min(500, s.notesHeight - d)) }))} />
              {bound("Notlar bölmesi", <NotesPane state={state} actions={a} height={state.notesHeight} spellcheck={state.spellcheck} />)}
            </>
          )}
          {canvasHint && (
            <div className="canvas-hint" role="status">
              {canvasHint}
            </div>
          )}
          {show?.preview && bound("Önizleme", <SlideShow pres={state.pres} startIndex={show.start} mode="window" preview={show.preview} frame={show.frame} onExit={() => setShow(null)} />, () => setShow(null))}
        </div>
      </>
    );
  }

  const endShow = (i: number, ink?: Record<string, ShowStroke[]>, rehearsal?: RehearsalResult) => {
    show?.link?.end();
    setShow(null);
    if (i >= 0 && i < a.s.pres.slides.length) a.goToSlide(i);
    if (rehearsal) setExtra({ kind: "rehearsal", result: rehearsal });
    else if (ink && Object.keys(ink).length) setExtra({ kind: "ink", ink });
  };

  return (
    <>
      <AppShell
        app="sunum"
        titleBar={
          <TitleBar
            app="sunum"
            title={state.fileName}
            dirty={state.dirty}
            quickAccess={
              <>
                <QatButton icon={<Save20Regular />} tip="Kaydet" shortcut="Ctrl+S" onClick={() => void save(a)} />
                <QatButton icon={<ArrowUndo20Regular />} tip="Geri Al" shortcut="Ctrl+Z" disabled={!store.canUndo} onClick={() => store.undo()} />
                <QatButton icon={<ArrowRedo20Regular />} tip="Yinele" shortcut="Ctrl+Y" disabled={!store.canRedo} onClick={() => store.redo()} />
                <QatButton icon={<SlidePlay20Regular />} tip="Baştan" shortcut="F5" onClick={() => cmd.startShow("start")} />
              </>
            }
          />
        }
        ribbon={bound(
          "Şerit",
          <Ribbon
            tabs={tabs}
            activeTab={activeTab}
            onTabChange={(id) => a.store.set({ ribbonTab: id })}
            onFileClick={() => {
              a.stopEditing();
              setBackstage(true);
            }}
            collapsed={collapsed}
            onCollapsedChange={setCollapsed}
          />,
          () => a.store.set({ ribbonTab: "home" }),
        )}
        belowRibbon={
          autoRecover.bar || state.message ? (
            <>
              {autoRecover.bar}
              {state.message && (
                <MessageBar kind={state.message.kind} onClose={() => a.store.set({ message: null })}>
                  {state.message.text}
                </MessageBar>
              )}
            </>
          ) : undefined
        }
        statusBar={<StatusArea state={state} actions={a} onSetView={cmd.setView} onStartShow={() => cmd.startShow("current")} onComments={() => cmd.togglePane("comments")} onZoomDialog={() => setDialog("zoom")} />}
        backstage={
          backstage ? (
            <FileBackstage
              initialPage={backstage === true ? undefined : backstage}
              state={state}
              actions={a}
              onBack={() => setBackstage(false)}
              onNew={async (id) => {
                if (!(await lifecycle.confirmDiscard())) return;
                newFromTemplate(a, id);
                setBackstage(false);
              }}
              onOpen={async () => {
                if (!(await lifecycle.confirmDiscard())) return;
                await openDialog(a);
                setBackstage(false);
              }}
              onOpenRecent={async (f) => {
                if (!(await lifecycle.confirmDiscard())) return;
                await openRecent(a, f);
                setBackstage(false);
              }}
              onSave={async () => {
                if (await save(a)) setBackstage(false);
              }}
              onSaveAs={async () => {
                if (await save(a, true)) setBackstage(false);
              }}
              onPrint={cmd.print}
              onExport={async (id) => {
                const base = state.fileName;
                const doc = documentPres(a);
                if (id === "pdf") cmd.exportPdf();
                else if (id === "png") await exportPngs(doc, base, "current", Math.min(state.masterBase?.slideIdx ?? state.slideIdx, doc.slides.length - 1));
                else if (id === "pngAll") await exportPngs(doc, base, "all", 0);
                else if (id === "outline") downloadBlob(new Blob([outlineText(doc)], { type: "text/plain;charset=utf-8" }), `${base}-AnaHat.txt`);
                else if (id === "pptx") void save(a, true);
                setBackstage(false);
              }}
              onClose={async () => {
                if (!(await lifecycle.confirmDiscard())) return;
                loadPresentation(a, blankPresentation(), "Sunu1", null);
                setBackstage(false);
              }}
            />
          ) : undefined
        }
      >
        <div className={`app-main${editView && state.colorMode !== "color" ? ` is-${state.colorMode}` : ""}`} onDragOver={(e) => e.preventDefault()}>
          {main}
          {editView && sidePane && bound("Görev bölmesi", sidePane, () => a.store.set({ pane: null }))}
        </div>
      </AppShell>
      <CommandSearch
        app="sunum"
        tabs={tabs}
        collapsed={collapsed}
        onExpand={() => setCollapsed(false)}
        onTabChange={(id) => {
          setBackstage(false);
          a.store.set({ ribbonTab: id });
        }}
        extra={searchExtra}
        find={{ heading: "Sunuda bul", label: (q) => `“${q}” için Bul`, path: "Giriş › Düzenleme", shortcut: "Ctrl+F", run: (q) => cmd.find(q) }}
        onMessage={(t) => a.message(t, "info")}
      />
      {show &&
        !show.preview &&
        bound(
          "Slayt gösterisi",
          <SlideShow pres={state.pres} startIndex={show.start} mode={show.mode} opts={show.opts} link={show.link ?? null} onExit={endShow} />,
          () => endShow(-1),
        )}
      {printing && (
        <PrintRoot
          pres={state.pres}
          onDone={() => setPrinting(null)}
          onReady={
            printing === "pdf"
              ? async () => {
                  try {
                    const name = await exportPdfNative(a.s.pres, a.s.fileName);
                    if (name) a.message(`PDF olarak kaydedildi: ${name}`, "success");
                  } catch (e) {
                    a.message(`PDF oluşturulamadı: ${(e as Error).message}`, "error");
                  }
                }
              : undefined
          }
        />
      )}
      {ctxMenu.element}
      {lifecycle.dialog}
      <AppDialogs a={a} state={state} info={info} cmd={cmd} dialog={dialog} closeDialog={closeDialog} extra={extra} setExtra={setExtra} showSettings={showSettings} setShowSettings={setShowSettings} findQuery={findQuery} />
    </>
  );
}

export type { AppState };
