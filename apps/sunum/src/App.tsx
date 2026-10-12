/** OfficeHere Sunum — uygulama kökü. */
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { ArrowRedo20Regular, ArrowUndo20Regular, Save20Regular, SlidePlay20Regular } from "@fluentui/react-icons";
import { AppShell, MessageBar, QatButton, Ribbon, TitleBar, useContextMenu, useDocumentLifecycle, type RibbonTabDef } from "@officehere/ui";
import { makeTextBox } from "./model/factory";
import { newPresentation } from "./model/presentation";
import { Canvas } from "./editor/Canvas";
import { ThumbnailPane } from "./panes/ThumbnailPane";
import { NotesPane, Splitter } from "./panes/NotesPane";
import { BackgroundPane, ShapePane } from "./panes/FormatPanes";
import { AnimationPane, CommentsPane, SelectionPane } from "./panes/ListPanes";
import { Actions } from "./state/actions";
import { Store, useStore, type AppState, type PaneKind, type ViewMode } from "./state/store";
import { textInfo } from "./state/textActions";
import { useKeyboard } from "./app/keyboard";
import { insertImageFiles, pasteFromRibbon, useClipboard } from "./app/clipboard";
import { loadPresentation, newFromTemplate, openBytes, openDialog, openFromPath, openRecent, save } from "./app/fileOps";
import { exportPngs, outlineText, PrintRoot } from "./app/exporting";
import { shapeMenu, slideAreaMenu, thumbMenu } from "./app/menus";
import { StatusArea } from "./app/StatusArea";
import { FileBackstage } from "./app/FileBackstage";
import { HomeTab } from "./ribbon/HomeTab";
import { InsertTab } from "./ribbon/InsertTab";
import { DrawTab } from "./ribbon/DrawTab";
import { OutlinePane } from "./panes/OutlinePane";
import { DesignTab } from "./ribbon/DesignTab";
import { TransitionsTab } from "./ribbon/TransitionsTab";
import { AnimationsTab } from "./ribbon/AnimationsTab";
import { HelpTab, ReviewTab, SlideShowTab, ViewTab } from "./ribbon/MiscTabs";
import { PictureFormatTab, ShapeFormatTab } from "./ribbon/FormatTabs";
import { TableDesignTab, TableLayoutTab } from "./ribbon/TableTabs";
import { WORDART_PRESETS } from "./ribbon/presets";
import type { AppCommands, DialogKind, RibbonCtx } from "./ribbon/types";
import { SlideShow, type ShowMode } from "./show/SlideShow";
import { NotesPageView, SorterView } from "./views/SorterView";
import { FontDialog, ParagraphDialog } from "./dialogs/TextDialogs";
import { DateTimeDialog, HeaderFooterDialog, IconsDialog, InsertTableDialog, SymbolDialog } from "./dialogs/InsertDialogs";
import { AboutDialog, FindDialog, ReplaceFontsDialog, ShortcutsDialog, ShowSetupDialog, SizePosDialog, SlideSizeDialog, ZoomDialog, type ShowSettings } from "./dialogs/MiscDialogs";
import { downloadBlob } from "@officehere/core";
import "./styles/slide.css";
import "./styles/editor.css";
import "./styles/app.css";
import "./styles/show.css";
import "./styles/print.css";

function initialState(): AppState {
  const pres = newPresentation("varsayilan", 0);
  return {
    pres,
    slideIdx: 0,
    slideSel: [pres.slides[0].id],
    sel: [],
    editing: null,
    fileName: "Sunu1",
    fileRef: null,
    dirty: false,
    view: "normal",
    zoom: 60,
    fit: true,
    showNotes: true,
    notesHeight: 64,
    thumbWidth: 220,
    showRuler: false,
    showGrid: false,
    showGuides: false,
    pane: null,
    insert: null,
    focus: "canvas",
    cellSel: null,
    painter: null,
    spellcheck: true,
    cropId: null,
    ink: null,
    textTick: 0,
    ribbonTab: "home",
    message: null,
  };
}

interface ShowState {
  mode: ShowMode;
  start: number;
  preview?: "transition" | "animations";
  frame?: { x: number; y: number; w: number; h: number };
}

/** Önizleme için slayt alanının .app-center içindeki konumu. */
function stageFrame(): ShowState["frame"] {
  const stage = document.querySelector(".app-center .stage");
  const center = document.querySelector(".app-center");
  if (!stage || !center) return undefined;
  const a = stage.getBoundingClientRect();
  const c = center.getBoundingClientRect();
  const x = Math.max(a.left, c.left);
  const y = Math.max(a.top, c.top);
  const r = Math.min(a.right, c.right);
  const b = Math.min(a.bottom, c.bottom);
  if (r - x < 40 || b - y < 40 || a.width > c.width || a.height > c.height) return undefined;
  return { x: x - c.left, y: y - c.top, w: r - x, h: b - y };
}

export function App() {
  const store = useMemo(() => new Store(initialState()), []);
  const actions = useMemo(() => new Actions(store), [store]);
  const state = useStore(store);
  const a = actions;
  const [backstage, setBackstage] = useState(false);
  const [dialog, setDialog] = useState<DialogKind | null>(null);
  const [show, setShow] = useState<ShowState | null>(null);
  const [printing, setPrinting] = useState(false);
  const [collapsed, setCollapsed] = useState(false);
  const [guides, setGuides] = useState({ h: [360], v: [640] });
  const [showSettings, setShowSettings] = useState<ShowSettings>({ loop: false, noAnimations: false, presenter: false });
  const ctxMenu = useContextMenu();
  const blockedRef = useRef(false);
  blockedRef.current = !!dialog || backstage || !!show;

  const lifecycle = useDocumentLifecycle({
    title: state.fileName,
    dirty: state.dirty,
    onOpenPath: (p) => void openFromPath(a, p),
    onSave: () => save(a),
  });

  // Kılavuzlar slayt boyutu değişince ortalanır.
  useEffect(() => setGuides({ h: [state.pres.height / 2], v: [state.pres.width / 2] }), [state.pres.width, state.pres.height]);

  const openDialogFn = useCallback((d: DialogKind) => setDialog(d), []);

  const cmd: AppCommands = useMemo(
    () => ({
      paste: (mode) => {
        if (mode === "text" && store.clipboard?.kind === "shapes") {
          const text = (store.clipboard.shapes ?? []).map((s) => (s.kind === "sp" && s.text ? s.text.paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join("\n") : "")).filter(Boolean).join("\n");
          if (text) a.pasteText(text);
          return;
        }
        void pasteFromRibbon(a);
      },
      cut: () => {
        if (a.s.editing) document.execCommand("cut");
        else if (a.s.focus === "slides") a.copySlides(true);
        else a.copySelection(true);
      },
      copy: () => {
        if (a.s.editing) document.execCommand("copy");
        else if (a.s.focus === "slides" || !a.s.sel.length) a.copySlides(false);
        else a.copySelection(false);
      },
      insertPictures: () => {
        const input = document.createElement("input");
        input.type = "file";
        input.accept = "image/*";
        input.multiple = true;
        input.onchange = () => {
          void insertImageFiles(a, Array.from(input.files ?? [])).then(() => a.store.set({ ribbonTab: "pictureFormat" }));
        };
        input.click();
      },
      startShow: (from) => {
        const s = a.s;
        if (!s.pres.slides.length) return;
        a.stopEditing();
        // Tam ekran isteği kullanıcı hareketi içinde yapılmalı.
        document.documentElement.requestFullscreen?.().catch(() => undefined);
        setShow({ mode: showSettings.presenter ? "presenter" : "full", start: from === "start" ? 0 : s.slideIdx });
      },
      presenterView: () => {
        document.documentElement.requestFullscreen?.().catch(() => undefined);
        setShow({ mode: "presenter", start: a.s.slideIdx });
      },
      print: () => {
        a.stopEditing();
        setBackstage(false);
        setPrinting(true);
      },
      previewTransition: () => a.s.pres.slides.length && setShow({ mode: "window", start: a.s.slideIdx, preview: "transition", frame: stageFrame() }),
      previewAnimations: () => a.s.pres.slides.length && setShow({ mode: "window", start: a.s.slideIdx, preview: "animations", frame: stageFrame() }),
      setView: (v: ViewMode) => {
        a.stopEditing();
        a.store.set({ view: v, sel: [] });
      },
      togglePane: (p: PaneKind) => a.store.set((s) => ({ pane: s.pane === p ? null : p })),
      pickShape: (item) => {
        if (item.geom === "textbox") a.store.set({ insert: { kind: "textbox" } });
        else a.store.set({ insert: { kind: "shape", geom: item.geom, head: item.head, tail: item.tail } });
      },
      insertWordArt: (i) => {
        const w = WORDART_PRESETS[i];
        const p = a.pres;
        const tb = makeTextBox(p.width * 0.2, p.height * 0.38, p.width * 0.6, 120, "Metninizi buraya yazın");
        tb.name = `Sanatsal Metin ${a.nextShapeNumber()}`;
        tb.text = { ...tb.text!, wrap: true, levels: [{ para: { align: "ctr" }, run: { size: 54, color: w.color, outline: w.outline, shadow: w.shadow || undefined, glow: w.glow ? w.color : undefined, bold: w.bold } }] };
        const i0 = a.s.slideIdx;
        a.store.commit((pp) => {
          const slides = pp.slides.slice();
          slides[i0] = { ...slides[i0], shapes: [...slides[i0].shapes, tb] };
          return { ...pp, slides };
        }, { sel: [tb.id], editing: { shapeId: tb.id, selectAll: true } });
        a.store.set({ ribbonTab: "shapeFormat" });
      },
      insertDateTime: () => setDialog("dateTime"),
      insertSlideNumber: () => {
        if (a.s.editing && a.textHandle) a.textHandle.insertRun("‹#›", { field: "slidenum" });
        else setDialog("headerFooter");
      },
      spelling: () => {
        const on = !a.s.spellcheck;
        a.store.set({ spellcheck: on, message: { text: on ? "Yazım denetimi açık: hatalı sözcükler metin düzenlenirken kırmızı dalgalı çizgiyle gösterilir." : "Yazım denetimi kapatıldı.", kind: "info" } });
      },
      zoomTo: (z, fit) => a.store.set({ zoom: Math.max(10, Math.min(400, z)), fit: !!fit }),
    }),
    [a, store, showSettings.presenter],
  );

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
    find: () => setDialog("find"),
    replace: () => setDialog("replace"),
    isBlocked: () => blockedRef.current,
  });
  useClipboard(a, useCallback(() => blockedRef.current, []));

  // Metin bilgisi (şerit durumları için).
  const info = useMemo(() => textInfo(a), [a, state.pres, state.sel, state.editing, state.textTick, state.cellSel]); // eslint-disable-line react-hooks/exhaustive-deps

  const ctx: RibbonCtx = { state, actions: a, text: info, openDialog: openDialogFn, cmd };
  const shapes = a.selectedShapes();
  const one = shapes.length === 1 ? shapes[0] : undefined;
  const hasSpLike = shapes.some((s) => s.kind === "sp" || s.kind === "group");
  const hasPic = !!one && one.kind === "pic";
  const hasTable = !!one && one.kind === "table";
  const editView = state.view === "normal" || state.view === "outline";
  const tabs: RibbonTabDef[] = [
    { id: "home", label: "Giriş", content: <HomeTab ctx={ctx} /> },
    { id: "insert", label: "Ekle", content: <InsertTab ctx={ctx} /> },
    { id: "draw", label: "Çizim", content: <DrawTab ctx={ctx} /> },
    { id: "design", label: "Tasarım", content: <DesignTab ctx={ctx} /> },
    { id: "transitions", label: "Geçişler", content: <TransitionsTab ctx={ctx} /> },
    { id: "animations", label: "Animasyonlar", content: <AnimationsTab ctx={ctx} /> },
    { id: "slideshow", label: "Slayt Gösterisi", content: <SlideShowTab ctx={ctx} /> },
    { id: "review", label: "Gözden Geçir", content: <ReviewTab ctx={ctx} /> },
    { id: "view", label: "Görünüm", content: <ViewTab ctx={ctx} /> },
    { id: "help", label: "Yardım", content: <HelpTab ctx={ctx} /> },
    { id: "shapeFormat", label: "Şekil Biçimi", contextual: true, hidden: !hasSpLike || !editView, content: <ShapeFormatTab ctx={ctx} /> },
    { id: "pictureFormat", label: "Resim Biçimi", contextual: true, hidden: !hasPic || !editView, content: <PictureFormatTab ctx={ctx} /> },
    { id: "tableDesign", label: "Tablo Tasarımı", contextual: true, hidden: !hasTable || !editView, content: <TableDesignTab ctx={ctx} /> },
    { id: "tableLayout", label: "Düzen", contextual: true, hidden: !hasTable || !editView, content: <TableLayoutTab ctx={ctx} /> },
  ];
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

  let main: React.ReactNode;
  if (state.view === "sorter") {
    main = (
      <SorterView
        state={state}
        actions={a}
        onOpenSlide={(i) => {
          a.goToSlide(i);
          a.store.set({ view: "normal" });
        }}
        onContextMenu={(e) => ctxMenu.open(e, thumbMenu(a, cmd, true))}
      />
    );
  } else if (state.view === "notes") {
    main = <NotesPageView state={state} actions={a} />;
  } else if (state.view === "reading") {
    main = (
      <div className="app-center" style={{ position: "relative" }}>
        <SlideShow
          key={`read-${state.slideIdx}`}
          pres={state.pres}
          startIndex={state.slideIdx}
          mode="window"
          onExit={(i) => {
            a.goToSlide(Math.min(i, a.s.pres.slides.length - 1));
            a.store.set({ view: "normal" });
          }}
          onViewChange={(v) => a.store.set({ view: v })}
        />
      </div>
    );
  } else {
    main = (
      <>
        {state.view === "outline" ? (
          <OutlinePane state={state} actions={a} width={Math.max(300, state.thumbWidth * 1.5)} />
        ) : (
          <ThumbnailPane state={state} actions={a} width={state.thumbWidth} onContextMenu={(e, onSlide) => ctxMenu.open(e, thumbMenu(a, cmd, onSlide))} />
        )}
        <Splitter dir="v" onDrag={(d) => a.store.set((s) => ({ thumbWidth: Math.max(110, Math.min(420, s.thumbWidth + (s.view === "outline" ? d / 1.5 : d))) }))} />
        <div className="app-center" style={{ position: "relative" }} onDragOver={(e) => e.preventDefault()} onDrop={onDrop}>
          <Canvas
            state={state}
            actions={a}
            onContextMenu={(e, kind) => ctxMenu.open(e, kind === "shape" ? shapeMenu(a, cmd, () => setDialog("sizePos")) : slideAreaMenu(a, cmd, () => a.store.set({ showRuler: !a.s.showRuler }), a.s.showRuler))}
            onPlaceholderAction={(action, id) => {
              a.pendingPlaceholder = id;
              a.select([id]);
              if (action === "table") setDialog("insertTable");
              else cmd.insertPictures();
            }}
            onShapeDoubleClick={(s) => {
              if (s.kind === "pic") a.store.set({ ribbonTab: "pictureFormat" });
              else if (s.kind === "sp" && !s.text) a.store.set({ ribbonTab: "shapeFormat" });
            }}
            showAnimBadges={activeTab === "animations" || state.pane === "animation"}
            guides={guides}
            onGuidesChange={setGuides}
            onZoomChange={(zoom, fit) => a.store.set({ zoom, fit })}
          />
          {state.showNotes && (
            <>
              <Splitter dir="h" onDrag={(d) => a.store.set((s) => ({ notesHeight: Math.max(30, Math.min(500, s.notesHeight - d)) }))} />
              <NotesPane state={state} actions={a} height={state.notesHeight} spellcheck={state.spellcheck} />
            </>
          )}
          {canvasHint && (
            <div className="canvas-hint" role="status">
              {canvasHint}
            </div>
          )}
          {show?.preview && <SlideShow pres={state.pres} startIndex={show.start} mode="window" preview={show.preview} frame={show.frame} onExit={() => setShow(null)} />}
        </div>
      </>
    );
  }

  return (
    <>
      <AppShell
        app="sunum"
        titleBar={
          <TitleBar
            app="sunum"
            title={state.fileName}
            dirty={state.dirty}
            onSearch={(q) => {
              if (q.trim()) setDialog("find");
            }}
            searchPlaceholder="Ara (Alt+Q)"
            quickAccess={
              <>
                <QatButton icon={<Save20Regular />} tip="Kaydet" shortcut="Ctrl+S" onClick={() => void save(a)} />
                <QatButton icon={<ArrowUndo20Regular />} tip="Geri Al" shortcut="Ctrl+Z" disabled={!store.canUndo} onClick={() => store.undo()} />
                <QatButton icon={<ArrowRedo20Regular />} tip="Yinele" shortcut="Ctrl+Y" disabled={!store.canRedo} onClick={() => store.redo()} />
                <QatButton icon={<SlidePlay20Regular />} tip="Baştan Başla" shortcut="F5" onClick={() => cmd.startShow("start")} />
              </>
            }
          />
        }
        ribbon={
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
          />
        }
        belowRibbon={
          state.message ? (
            <MessageBar kind={state.message.kind} onClose={() => a.store.set({ message: null })}>
              {state.message.text}
            </MessageBar>
          ) : undefined
        }
        statusBar={<StatusArea state={state} actions={a} onStartShow={() => cmd.startShow("current")} onComments={() => cmd.togglePane("comments")} onZoomDialog={() => setDialog("zoom")} />}
        backstage={
          backstage ? (
            <FileBackstage
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
                if (id === "pdf") cmd.print();
                else if (id === "png") await exportPngs(state.pres, base, "current", state.slideIdx);
                else if (id === "pngAll") await exportPngs(state.pres, base, "all", state.slideIdx);
                else if (id === "outline") downloadBlob(new Blob([outlineText(state.pres)], { type: "text/plain;charset=utf-8" }), `${base}-AnaHat.txt`);
                else if (id === "pptx") void save(a, true);
                setBackstage(false);
              }}
              onClose={async () => {
                if (!(await lifecycle.confirmDiscard())) return;
                loadPresentation(a, newPresentation("varsayilan", 0), "Sunu1", null);
                setBackstage(false);
              }}
            />
          ) : undefined
        }
      >
        <div className="app-main" onDragOver={(e) => e.preventDefault()}>
          {main}
          {editView && sidePane}
        </div>
      </AppShell>
      {show && !show.preview && (
        <SlideShow
          pres={state.pres}
          startIndex={show.start}
          mode={show.mode}
          loop={showSettings.loop}
          noAnimations={showSettings.noAnimations}
          onExit={(i) => {
            setShow(null);
            if (i >= 0 && i < a.s.pres.slides.length) a.goToSlide(i);
          }}
        />
      )}
      {printing && <PrintRoot pres={state.pres} onDone={() => setPrinting(false)} />}
      {ctxMenu.element}
      {lifecycle.dialog}
      <FontDialog open={dialog === "font"} onClose={closeDialog} actions={a} info={info} />
      <ParagraphDialog open={dialog === "paragraph"} onClose={closeDialog} actions={a} info={info} />
      <InsertTableDialog open={dialog === "insertTable"} onClose={closeDialog} actions={a} />
      <HeaderFooterDialog open={dialog === "headerFooter"} onClose={closeDialog} actions={a} />
      <SymbolDialog
        open={dialog === "symbol"}
        onClose={closeDialog}
        onInsert={(ch) => {
          if (a.s.editing && a.textHandle) a.textHandle.insert(ch);
          else a.pasteText(ch);
        }}
      />
      <DateTimeDialog
        open={dialog === "dateTime"}
        onClose={closeDialog}
        onInsert={(t) => {
          if (a.s.editing && a.textHandle) a.textHandle.insert(t);
          else a.pasteText(t);
        }}
      />
      <IconsDialog open={dialog === "icons"} onClose={closeDialog} actions={a} />
      <SlideSizeDialog open={dialog === "slideSize"} onClose={closeDialog} actions={a} />
      <FindDialog open={dialog === "find" || dialog === "replace"} replaceMode={dialog === "replace"} onClose={closeDialog} actions={a} />
      <ReplaceFontsDialog open={dialog === "replaceFonts"} onClose={closeDialog} actions={a} />
      <ZoomDialog open={dialog === "zoom"} onClose={closeDialog} zoom={state.zoom} onZoom={cmd.zoomTo} />
      <SizePosDialog open={dialog === "sizePos"} onClose={closeDialog} actions={a} />
      <ShowSetupDialog open={dialog === "showSetup"} onClose={closeDialog} value={showSettings} onChange={setShowSettings} />
      <ShortcutsDialog open={dialog === "shortcuts"} onClose={closeDialog} />
      <AboutDialog open={dialog === "about"} onClose={closeDialog} />
    </>
  );
}

