// OfficeHere Belge: uygulama kabuğu (başlık, şerit, çalışma alanı, durum çubuğu, Dosya görünümü).
import { useCallback, useEffect, useMemo, useRef, useState, useSyncExternalStore } from "react";
import { useEditor, useEditorState } from "@tiptap/react";
import type { Editor } from "@tiptap/core";
import { NodeSelection, TextSelection } from "@tiptap/pm/state";
import { ArrowRedo20Regular, ArrowUndo20Regular, Save20Regular, Print20Regular } from "@fluentui/react-icons";
import { AppShell, MessageBar, QatButton, Ribbon, TitleBar, useContextMenu, useDocumentLifecycle, useHotkeys, type OfficeBackstagePage } from "@officehere/ui";
import { defaultSettings, emptyDoc, type DocSettings } from "./model/document";
import { setActiveNoteEditor, useActiveNoteEditor } from "./state/activeNote";
import { columnWidthAt as colWidthAt } from "./state/pageSetup";
import { renderNoteHtml } from "./views/noteHtml";
import { resolveStyles, styleSheetCss } from "./model/styles";
import { editorExtensions } from "./editor/extensions";
import { PaginationController } from "./editor/pagination";
import { computeFormat, sameFormat } from "./editor/formatState";
import { collectHeadings, docStats } from "./editor/commands";
import { setShowMarks } from "./editor/showMarks";
import { readImageFile } from "./editor/image";
import { annotateNotesHtml } from "./editor/notes";
import { BelgeContext, type BelgeCtx, type DialogKind, type DialogState, type ViewOptions } from "./state/context";
import { clearDraft, readDraft, useFileOps, type Draft } from "./state/useFileOps";
import { DocumentView, TILE_GAP } from "./views/DocumentView";
import { BelgeBackstage } from "./views/BelgeBackstage";
import { DraftBar } from "./views/DraftBar";
import { usePrint } from "./state/usePrint";
import { useComments } from "./state/useComments";
import { useEditorClicks } from "./state/useEditorClicks";
import { useFormatPainter } from "./state/useFormatPainter";
import { ReadMode } from "./views/ReadMode";
import { BelgeStatusBar } from "./views/BelgeStatusBar";
import { buildContextMenu } from "./views/contextMenu";
import { DialogHost } from "./dialogs/DialogHost";
import type { ZoomTarget } from "./dialogs/MiscDialogs";
import { HomeTab } from "./ribbon/HomeTab";
import { InsertTab } from "./ribbon/InsertTab";
import { DesignTab } from "./ribbon/DesignTab";
import { LayoutTab } from "./ribbon/LayoutTab";
import { ReferencesTab } from "./ribbon/ReferencesTab";
import { ReviewTab } from "./ribbon/ReviewTab";
import { HeaderFooterTab, HelpTab, ViewTab } from "./ribbon/ViewTab";
import { TableDesignTab, TableLayoutTab } from "./ribbon/TableTabs";
import { PictureTab } from "./ribbon/PictureTab";
import "./styles/belge.css";
import "./styles/document.css";

type Note = { text: string; kind: "info" | "warning" | "error" | "success"; id: number };

export function App() {
  const controller = useMemo(() => new PaginationController(), []);
  const [settings, setSettingsState] = useState<DocSettings>(defaultSettings);
  const [dirty, setDirty] = useState(false);
  const [view, setViewState] = useState<ViewOptions>({ mode: "print", zoom: 100, ruler: true, gridlines: false, navPane: false, focusMode: false, showMarks: false, showComments: true });
  const [tab, setTab] = useState("home");
  const [ribbonCollapsed, setRibbonCollapsed] = useState(false);
  const [dialog, setDialog] = useState<DialogState | null>(null);
  const [backstage, setBackstage] = useState<OfficeBackstagePage | null>(null);
  const [hfEditing, setHfEditingState] = useState<"header" | "footer" | null>(null);
  const [hfPage, setHfPage] = useState(0);
  const setHfEditing = useCallback((w: "header" | "footer" | null, page?: number) => {
    if (page !== undefined) setHfPage(page);
    setHfEditingState(w);
  }, []);
  const [hfEditor, setHfEditor] = useState<Editor | null>(null);
  const [note, setNote] = useState<Note | null>(null);
  const [draft, setDraft] = useState<Draft | null>(() => readDraft());
  const [version, setVersion] = useState(0);
  const [navFocus, setNavFocus] = useState(0);
  const [spellcheck, setSpellcheck] = useState(true);
  const [lang, setLang] = useState("tr");
  const [autosave, setAutosave] = useState(true);
  const scrollRef = useRef<HTMLDivElement>(null);
  const insertImageFileRef = useRef<(f: File, pos: number | null, view: import("@tiptap/pm/view").EditorView) => Promise<void>>(async () => {});
  const imageInput = useRef<HTMLInputElement>(null);
  const openFileRef = useRef<(f: File) => Promise<void>>(async () => {});
  const ctxMenu = useContextMenu();

  const styles = useMemo(() => resolveStyles(settings.styleSet, settings.theme, settings.fontTheme, settings.paragraphSpacing), [settings.styleSet, settings.theme, settings.fontTheme, settings.paragraphSpacing]);
  const stylesRef = useRef(styles);
  stylesRef.current = styles;

  const notify = useCallback((text: string, kind: Note["kind"] = "info") => setNote({ text, kind, id: Date.now() }), []);
  useEffect(() => {
    if (!note || note.kind === "error" || note.kind === "warning") return;
    const id = window.setTimeout(() => setNote((n) => (n?.id === note.id ? null : n)), 4000);
    return () => window.clearTimeout(id);
  }, [note]);

  const updateSettings = useCallback((fn: (s: DocSettings) => DocSettings) => {
    setSettingsState((s) => fn(s));
    setDirty(true);
  }, []);

  const setView = useCallback((patch: Partial<ViewOptions>) => setViewState((v) => ({ ...v, ...patch })), []);

  // Kısayolların erişeceği güncel işleyiciler.
  const hooksRef = useRef({ openDialog: (_k: "font" | "paragraph" | "link") => {}, toggleShowMarks: () => {}, newComment: () => {}, toggleTrack: () => {}, insertNote: (_k: "footnote" | "endnote") => {} });

  const editor = useEditor({
    extensions: editorExtensions(controller, {
      styles: () => stylesRef.current,
      openDialog: (k) => hooksRef.current.openDialog(k),
      toggleShowMarks: () => hooksRef.current.toggleShowMarks(),
      newComment: () => hooksRef.current.newComment(),
      toggleTrack: () => hooksRef.current.toggleTrack(),
      insertNote: (k) => hooksRef.current.insertNote(k),
    }),
    content: emptyDoc(),
    // Otomatik odak aşağıdaki etkide yapılır (tiptap'in gecikmeli "start" odağı imleci başa atabiliyor).
    autofocus: false,
    editorProps: {
      attributes: { class: "bl-editor", spellcheck: "true", lang: "tr", "aria-label": "Belge gövdesi" },
      // Resim dosyalarını sürükle-bırak ve panodan yapıştırma.
      handlePaste: (view, event) => {
        const files = Array.from(event.clipboardData?.files ?? []).filter((f) => f.type.startsWith("image/"));
        if (!files.length) return false;
        event.preventDefault();
        for (const f of files) void insertImageFileRef.current(f, null, view);
        return true;
      },
      handleDrop: (view, event, _slice, moved) => {
        if (moved) return false;
        // Belge dosyası bırakılırsa açılır (Word'deki gibi).
        const doc = Array.from(event.dataTransfer?.files ?? []).find((f) => /\.(docx|docm|dotx|html?|txt)$/i.test(f.name));
        if (doc) {
          event.preventDefault();
          void openFileRef.current(doc);
          return true;
        }
        const files = Array.from(event.dataTransfer?.files ?? []).filter((f) => f.type.startsWith("image/"));
        if (!files.length) return false;
        event.preventDefault();
        const pos = view.posAtCoords({ left: event.clientX, top: event.clientY })?.pos ?? null;
        for (const f of files) void insertImageFileRef.current(f, pos, view);
        return true;
      },
    },
    onUpdate: () => setDirty(true),
  });

  // Geliştirme sırasında otomatik testler için erişim.
  useEffect(() => {
    if (import.meta.env.DEV) (window as unknown as { __belge: unknown }).__belge = { editor, controller };
  }, [editor, controller]);

  // Pencereye bırakılan belge dosyası açılır (editör dışına bırakılsa da).
  useEffect(() => {
    const isFiles = (e: DragEvent) => Array.from(e.dataTransfer?.types ?? []).includes("Files");
    const over = (e: DragEvent) => {
      if (isFiles(e)) e.preventDefault();
    };
    const drop = (e: DragEvent) => {
      if (!isFiles(e) || (e.target as HTMLElement).closest?.(".ProseMirror")) return;
      const f = Array.from(e.dataTransfer?.files ?? []).find((x) => /\.(docx|docm|dotx|html?|txt)$/i.test(x.name));
      e.preventDefault();
      if (f) void openFileRef.current(f);
    };
    window.addEventListener("dragover", over);
    window.addEventListener("drop", drop);
    return () => {
      window.removeEventListener("dragover", over);
      window.removeEventListener("drop", drop);
    };
  }, []);

  // Açılışta belgeye hemen odaklan.
  useEffect(() => {
    const ae = document.activeElement;
    if (!ae || ae === document.body) editor.view.focus();
  }, [editor]);

  // Yazım denetimi ve dil özniteliği
  useEffect(() => {
    editor.view.dom.setAttribute("spellcheck", String(spellcheck));
    editor.view.dom.setAttribute("lang", lang);
  }, [editor, spellcheck, lang]);

  // Belge sürümü (gezinti, açıklamalar, sözcük sayısı için; seyreltilmiş).
  useEffect(() => {
    let t = 0;
    const bump = () => {
      window.clearTimeout(t);
      t = window.setTimeout(() => setVersion((v) => v + 1), 200);
    };
    editor.on("update", bump);
    return () => {
      editor.off("update", bump);
      window.clearTimeout(t);
    };
  }, [editor]);

  // Şeridin hedefi: üst/alt bilgi, odaktaki not düzenleyicisi ya da belge gövdesi.
  const noteEditor = useActiveNoteEditor();
  const active = hfEditing && hfEditor ? hfEditor : (noteEditor ?? editor);
  useEffect(() => {
    const off = () => setActiveNoteEditor(null);
    editor.on("focus", off);
    return () => {
      editor.off("focus", off);
    };
  }, [editor]);
  const fmt = useEditorState({
    editor: active,
    selector: ({ editor: ed }) => computeFormat(ed, stylesRef.current),
    equalityFn: (a, b) => !!a && !!b && sameFormat(a, b),
  });

  const stats = useMemo(() => docStats(editor.state.doc), [editor, version]); // eslint-disable-line react-hooks/exhaustive-deps
  // Yalnız sayfa sayısı izlenir (sayfa başlangıç konumları her tuşta kayar; tüm kabuğu yeniden çizmeye gerek yok).
  const pageCount = useSyncExternalStore(
    useCallback((cb: () => void) => controller.subscribe(cb), [controller]),
    () => controller.info.pageCount,
  );

  // Tümünü Göster
  useEffect(() => {
    setShowMarks(editor.view, view.showMarks);
  }, [editor, view.showMarks, version === 0]); // eslint-disable-line react-hooks/exhaustive-deps

  // Açıklamalar ve değişiklik izleme
  const { activeComment, setActiveComment, newComment, deleteComment, deleteAllComments, toggleTrack } = useComments({ editor, updateSettings, setView, notify });

  const onLoaded = useCallback(() => {
    setHfEditing(null);
    setActiveComment(null);
    setVersion((v) => v + 1);
    requestAnimationFrame(() => setShowMarks(editor.view, view.showMarks));
  }, [editor, view.showMarks]);

  // Belge yaşam döngüsü (masaüstü kabuğu ile kaydetme sorusu, dosya açma istekleri).
  const confirmRef = useRef<() => Promise<boolean>>(async () => true);
  const confirmStable = useCallback(() => confirmRef.current(), []);
  const files = useFileOps({ editor, settings, setSettings: setSettingsState, dirty, setDirty, notify, confirmDiscard: confirmStable, autosave, onLoaded });
  const { dialog: lifecycleDialog, confirmDiscard } = useDocumentLifecycle({
    title: files.fileName,
    dirty,
    onOpenPath: (p) => void files.openPath(p),
    onSave: files.save,
  });
  confirmRef.current = confirmDiscard;

  // Şerit/menü etkileşiminden sonra odak boşta kalırsa belgeye dön (Word'deki gibi).
  const activeRef = useRef(editor);
  activeRef.current = active;
  useEffect(() => {
    const onUp = (e: PointerEvent) => {
      const t = e.target as HTMLElement;
      if (t.closest?.(".bl-body, .bl-hf-edit, input, textarea, select, .oh-dialog, .bl-nav, .bl-comments, .oh-titlebar, .oh-backstage, .bl-read")) return;
      window.setTimeout(() => {
        const ae = document.activeElement as HTMLElement | null;
        // Şerit sekmesine tıklamak da klavye odağını belgeden almaz (Word'deki gibi).
        const idle = !ae || ae === document.body || !!ae.closest?.(".oh-ribbon-tabs [role=tab]");
        if (idle && !document.querySelector(".oh-dialog, [data-oh-popover]")) activeRef.current.view.focus();
      }, 0);
    };
    document.addEventListener("pointerup", onUp, true);
    return () => document.removeEventListener("pointerup", onUp, true);
  }, []);

  // İletişim kutusu kapanınca odak belgeye döner (kutuyu açan şerit düğmesinde kalmaz).
  const hadDialog = useRef(false);
  useEffect(() => {
    if (hadDialog.current && !dialog) {
      window.setTimeout(() => {
        const ae = document.activeElement as HTMLElement | null;
        if (!ae || !ae.closest?.(".ProseMirror, input, textarea, select, .oh-dialog, .oh-backstage, .bl-fn-text, .bl-note-text")) activeRef.current.view.focus();
      }, 0);
    }
    hadDialog.current = !!dialog;
  }, [dialog]);

  const openDialog = useCallback((kind: DialogKind, arg?: string) => setDialog({ kind, arg }), []);
  const openFind = useCallback(() => {
    setView({ navPane: true });
    setNavFocus((n) => n + 1);
  }, [setView]);

  hooksRef.current = {
    openDialog: (k) => openDialog(k),
    toggleShowMarks: () => setView({ showMarks: !view.showMarks }),
    newComment,
    toggleTrack,
    insertNote: (k) => insertNoteRef.current(k),
  };
  const insertNoteRef = useRef<(k: "footnote" | "endnote") => void>(() => {});

  const { painter, startFormatPainter } = useFormatPainter(editor);

  /* ---------------- Ekle yardımcıları ---------------- */

  const insertImageFromFile = useCallback(() => imageInput.current?.click(), []);

  const updateToc = useCallback(() => {
    const headings = collectHeadings(editor.state.doc).filter((h) => h.level <= 3);
    const entries = headings.map((h) => ({ level: h.level, text: h.text, page: controller.pageOf(h.pos) + 1, pos: h.pos }));
    const tr = editor.state.tr;
    let found = false;
    editor.state.doc.descendants((n, pos) => {
      if (n.type.name === "toc") {
        tr.setNodeMarkup(pos, undefined, { ...n.attrs, entries });
        found = true;
      }
      return true;
    });
    if (found) editor.view.dispatch(tr);
  }, [editor, controller]);

  const insertNote = useCallback(
    (kind: "footnote" | "endnote") => {
      if (hfEditing) return;
      editor.chain().focus().insertNote(kind).run();
      const before = editor.state.doc.resolve(editor.state.selection.from).nodeBefore;
      const id = before?.type.name === "note" ? (before.attrs.id as string) : null;
      if (!id) return;
      // Not metni alanı sayfalama sonrası çizilir; hazır olunca odaklan.
      let tries = 0;
      const focus = () => {
        const el = document.querySelector<HTMLElement>(`.ProseMirror[data-note-id="${id}"]`);
        if (el) el.focus();
        else if (tries++ < 90) requestAnimationFrame(focus);
      };
      requestAnimationFrame(focus);
    },
    [editor, hfEditing],
  );

  insertNoteRef.current = insertNote;

  /* ---------------- Yakınlaştırma ve yazdırma ---------------- */

  const zoomTo = useCallback(
    (z: ZoomTarget | "100" | "onePage" | "multi" | "width") => {
      const sc = scrollRef.current;
      // İmleçteki sayfanın geometrisi (bölümler farklı boyutta olabilir).
      const info = controller.info;
      const pg = info.pages[Math.min(info.pages.length - 1, controller.pageOf(editor.state.selection.from))] ?? info.pages[0];
      const w = info.bodyW || pg.w;
      const availW = (sc?.clientWidth ?? 1000) - 48;
      const availH = (sc?.clientHeight ?? 700) - 48;
      let pct = view.zoom;
      const target = typeof z === "string" ? (z === "100" ? { pct: 100 } : { fit: z === "onePage" ? "page" : z }) : z;
      if ("pct" in target) pct = target.pct;
      else if (target.fit === "width") pct = (availW / w) * 100;
      else if (target.fit === "text") pct = (availW / (pg.w - pg.ml - pg.mr)) * 100;
      else if (target.fit === "page") pct = Math.min(availW / w, availH / pg.h) * 100;
      // Birden Çok Sayfa: iki sayfa yan yana tam görünür.
      else if (target.fit === "multi") pct = Math.min(availW / (2 * w + TILE_GAP + 8), availH / pg.h) * 100;
      setView({ zoom: Math.max(10, Math.min(500, Math.floor(pct))) });
    },
    [controller, editor, view.zoom, setView],
  );

  const leaveModes = useCallback(() => {
    setBackstage(null);
    setHfEditing(null);
    setView({ mode: "print", focusMode: false });
  }, [setHfEditing, setView]);
  const { printing, print } = usePrint({ controller, settings, zoom: view.zoom, fileName: files.fileName, leaveModes, notify });

  /* ---------------- Kısayollar ---------------- */

  useHotkeys({
    "Mod+S": () => void files.save(),
    F12: () => void files.saveAs(),
    "Mod+O": () => void files.openDialog(),
    "Mod+N": () => void files.newDocument("blank"),
    "Mod+P": () => void print(),
    "Mod+F": openFind,
    "Mod+H": () => openDialog("findReplace", "replace"),
    "Mod+G": () => openDialog("findReplace", "goto"),
    "Mod+Shift+G": () => openDialog("wordCount"),
    Escape: () => {
      if (view.focusMode) setView({ focusMode: false });
      else if (hfEditing) setHfEditing(null);
    },
  });

  // Ctrl+tıklama ile köprüyü aç.
  const openLink = useEditorClicks(editor);

  // Odak modu ve Okuma Modu'nda şerit gizlenir.
  const onContextMenu = (e: React.MouseEvent) => {
    if ((e.target as HTMLElement).closest(".bl-comments")) return;
    const t = e.target as HTMLElement;
    if (t.closest(".ProseMirror")) {
      // Sağ tık imleci seçimin dışındaysa oraya taşı.
      const hit = editor.view.posAtCoords({ left: e.clientX, top: e.clientY });
      const { from, to } = editor.state.selection;
      if (hit && (hit.pos < from || hit.pos > to) && !(editor.state.selection instanceof NodeSelection)) {
        editor.view.dispatch(editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(hit.pos))));
      }
    }
    ctxMenu.open(e, buildContextMenu({ editor: active, openDialog, newComment, notify, openLink }));
  };

  const editHeaderFooter = useCallback(
    (which: "header" | "footer" | null) => {
      setHfEditing(null);
      if (which) {
        // Belge gövdesi hemen odağı bırakır: üst bilgi açılana dek basılan tuşlar gövdeye yazılmasın.
        editor.view.dom.blur();
        setView({ mode: "print" });
        // İmleçteki sayfanın üst/alt bilgisi açılır.
        const page = controller.pageOf(editor.state.selection.from);
        requestAnimationFrame(() => setHfEditing(which, page));
      }
    },
    [setView, setHfEditing, controller, editor],
  );

  const ctx: BelgeCtx = {
    editor: active,
    mainEditor: editor,
    fmt: fmt ?? computeFormat(active, styles),
    settings,
    updateSettings,
    styles,
    controller,
    view,
    setView,
    openDialog,
    openFind,
    editHeaderFooter,
    headerFooterEditing: hfEditing,
    headerFooterPage: hfPage,
    newComment,
    toggleTrack,
    tracking: fmt?.tracking ?? false,
    formatPainter: painter,
    startFormatPainter,
    insertImageFromFile,
    updateToc,
    insertNote,
    notify,
    openBackstage: (p) => setBackstage(p ?? "home"),
  };

  const tabs = [
    { id: "home", label: "Giriş", content: <HomeTab /> },
    { id: "insert", label: "Ekle", content: <InsertTab /> },
    { id: "design", label: "Tasarım", content: <DesignTab /> },
    { id: "layout", label: "Düzen", content: <LayoutTab /> },
    { id: "references", label: "Başvurular", content: <ReferencesTab /> },
    { id: "review", label: "Gözden Geçir", content: <ReviewTab spellcheck={spellcheck} setSpellcheck={setSpellcheck} lang={lang} setLang={setLang} deleteComment={deleteComment} deleteAllComments={deleteAllComments} /> },
    { id: "view", label: "Görünüm", content: <ViewTab zoomTo={zoomTo} /> },
    { id: "help", label: "Yardım", content: <HelpTab onFeedback={() => window.open("https://github.com/fyildirim-debug/officehere/issues", "_blank", "noopener")} /> },
    { id: "hf", label: "Üst Bilgi ve Alt Bilgi", content: <HeaderFooterTab />, contextual: true, hidden: !hfEditing },
    { id: "tableDesign", label: "Tablo Tasarımı", content: <TableDesignTab />, contextual: true, hidden: !fmt?.inTable || !!hfEditing },
    { id: "tableLayout", label: "Düzen", content: <TableLayoutTab />, contextual: true, hidden: !fmt?.inTable || !!hfEditing },
    { id: "picture", label: "Resim Biçimi", content: <PictureTab />, contextual: true, hidden: !fmt?.image },
  ];

  // Üst bilgi düzenlemesi başlayınca bağlamsal sekmeye geç.
  useEffect(() => {
    if (hfEditing) setTab("hf");
    else setTab((t) => (t === "hf" ? "home" : t));
  }, [hfEditing]);

  /** Konumdaki bölümün sütun genişliği (resimler buna sığdırılır). */
  const columnWidthAt = (pos: number) => colWidthAt(editor, settings.page, pos);

  insertImageFileRef.current = async (file, pos, view) => {
    try {
      const img = await readImageFile(file, columnWidthAt(pos ?? view.state.selection.from));
      const node = view.state.schema.nodes.image.create({ src: img.src, width: img.width, height: img.height, alt: file.name.replace(/\.[^.]+$/, "") });
      const at = pos ?? view.state.selection.from;
      view.dispatch(view.state.tr.insert(Math.min(at, view.state.doc.content.size), node));
    } catch (err) {
      notify((err as Error).message, "error");
    }
  };

  const openLocalFile = async (file: File) => {
    if (!(await confirmDiscard())) return;
    setBackstage(null);
    await files.openData(file.name, await file.arrayBuffer(), { name: file.name });
  };
  const onOpenFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (file) await openLocalFile(file);
  };
  openFileRef.current = openLocalFile;

  const onImageFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    try {
      const img = await readImageFile(file, columnWidthAt(editor.state.selection.from));
      active.chain().focus().setImage({ src: img.src, width: img.width, height: img.height, alt: file.name.replace(/\.[^.]+$/, "") }).run();
    } catch (err) {
      notify((err as Error).message, "error");
    }
  };

  const backstageEl =
    backstage !== null ? (
      <BelgeBackstage
        page={backstage}
        files={files}
        settings={settings}
        pageCount={pageCount}
        stats={stats}
        previewHtml={backstage === "print" ? editor.getHTML() : ""}
        autosave={autosave}
        setAutosave={setAutosave}
        spellcheck={spellcheck}
        setSpellcheck={setSpellcheck}
        onPrint={() => void print()}
        onPdf={() => void print("pdf")}
        onClose={() => setBackstage(null)}
      />
    ) : undefined;

  const belowRibbon = (
    <>
      {draft && (
        <DraftBar
          draft={draft}
          onRestore={() => {
            files.load({ content: draft.content, settings: { ...defaultSettings(), ...draft.settings } }, draft.fileName, null);
            setDirty(true);
            setDraft(null);
          }}
          onDiscard={() => {
            clearDraft();
            setDraft(null);
          }}
        />
      )}
      {note && (
        <MessageBar kind={note.kind} onClose={() => setNote(null)}>
          {note.text}
        </MessageBar>
      )}
    </>
  );

  const hideChrome = view.focusMode;

  return (
    <BelgeContext.Provider value={ctx}>
      <style>{styleSheetCss(styles)}</style>
      {view.mode === "read" && !printing && (
        <ReadMode
          html={annotateNotesHtml(editor.getHTML(), renderNoteHtml)}
          title={files.fileName}
          onExit={() => setView({ mode: "print" })}
          onFind={() => {
            setView({ mode: "print" });
            openFind();
          }}
        />
      )}
      <div className={`bl-app${hideChrome ? " is-focus" : ""}${printing ? " is-printing" : ""}${view.mode === "read" ? " is-read" : ""}`}>
        <AppShell
          app="belge"
          titleBar={
            <TitleBar
              app="belge"
              title={files.fileName}
              dirty={dirty}
              onSearch={(q) => {
                openFind();
                requestAnimationFrame(() => {
                  const input = document.querySelector<HTMLInputElement>(".bl-nav-search input");
                  if (input) {
                    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")?.set;
                    setter?.call(input, q);
                    input.dispatchEvent(new Event("input", { bubbles: true }));
                  }
                });
              }}
              searchPlaceholder="Ara (Alt+Q)"
              quickAccess={
                <>
                  <QatButton icon={<Save20Regular />} tip="Kaydet" shortcut="Ctrl+S" onClick={() => void files.save()} />
                  <QatButton icon={<ArrowUndo20Regular />} tip="Geri Al" shortcut="Ctrl+Z" disabled={!fmt?.canUndo} onClick={() => active.chain().focus().undo().run()} />
                  <QatButton icon={<ArrowRedo20Regular />} tip="Yinele" shortcut="Ctrl+Y" disabled={!fmt?.canRedo} onClick={() => active.chain().focus().redo().run()} />
                  <QatButton icon={<Print20Regular />} tip="Yazdır" shortcut="Ctrl+P" onClick={() => void print()} />
                </>
              }
            />
          }
          ribbon={hideChrome ? undefined : <Ribbon tabs={tabs} activeTab={tab} onTabChange={setTab} onFileClick={() => setBackstage("home")} collapsed={ribbonCollapsed} onCollapsedChange={setRibbonCollapsed} />}
          belowRibbon={hideChrome ? undefined : belowRibbon}
          statusBar={
            hideChrome ? undefined : (
              <BelgeStatusBar
                editor={editor}
                controller={controller}
                paged={view.mode === "print"}
                words={stats.words}
                tracking={fmt?.tracking ?? false}
                view={view}
                setView={setView}
                onPageClick={() => {
                  setView({ navPane: true });
                }}
                onWordsClick={() => openDialog("wordCount")}
                onZoomClick={() => openDialog("zoom")}
                lang={lang}
              />
            )
          }
          backstage={backstageEl}
        >
          <DocumentView
            editor={editor}
            settings={settings}
            updateSettings={updateSettings}
            controller={controller}
            view={view}
            setView={setView}
            hfEditing={hfEditing}
            hfPage={hfPage}
            setHfEditing={setHfEditing}
            onHfEditor={setHfEditor}
            activeComment={activeComment}
            setActiveComment={setActiveComment}
            deleteComment={deleteComment}
            version={version}
            navFocus={navFocus}
            onPageSetup={() => openDialog("pageSetup")}
            onContextMenu={onContextMenu}
            scrollRef={scrollRef}
            printing={false}
            formatPainter={painter.active}
          />
        </AppShell>
      </div>
      <DialogHost dialog={dialog} onClose={() => setDialog(null)} zoom={view.zoom} onZoom={zoomTo} pages={pageCount} />
      {ctxMenu.element}
      {lifecycleDialog}
      <input ref={imageInput} type="file" accept="image/*" className="oh-hidden-input" onChange={onImageFile} />
      {/* Dosya açma girişi (sürükle-bırak ve dosya seçicisi olmayan ortamlar için de kullanılır). */}
      <input id="bl-open-input" type="file" accept=".docx,.docm,.dotx,.html,.htm,.txt" className="oh-hidden-input" aria-label="Belge aç" onChange={onOpenFile} />
      {view.focusMode && (
        <button type="button" className="bl-focus-exit" onClick={() => setView({ focusMode: false })}>
          Odaktan çık (Esc)
        </button>
      )}
    </BelgeContext.Provider>
  );
}
