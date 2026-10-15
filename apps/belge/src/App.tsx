// OfficeHere Belge: uygulama kabuğu (başlık, şerit, çalışma alanı, durum çubuğu, Dosya görünümü).
import { useCallback, useEffect, useMemo, useRef, useState, useSyncExternalStore } from "react";
import { useEditor, useEditorState } from "@tiptap/react";
import type { Editor } from "@tiptap/core";
import { NodeSelection } from "@tiptap/pm/state";
import { AppShell, MessageBar, Ribbon, getUserName, useContextMenu, useDocumentLifecycle, useHotkeys, type OfficeBackstagePage } from "@officehere/ui";
import { defaultSettings, emptyDoc, type DocSettings } from "./model/document";
import { getActiveNoteEditor, setActiveNoteEditor, useActiveNoteEditor } from "./state/activeNote";
import { setTrackAuthorSource } from "./editor/trackChanges";
import { renderNoteHtml } from "./views/noteHtml";
import { resolveStyles, styleSheetCss } from "./model/styles";
import { editorExtensions } from "./editor/extensions";
import { PaginationController } from "./editor/pagination";
import { computeFormat, sameFormat } from "./editor/formatState";
import { docStats } from "./editor/commands";
import { setShowMarks } from "./editor/showMarks";
import { annotateNotesHtml } from "./editor/notes";
import { BelgeContext, type BelgeCtx, type DialogKind, type DialogState, type ViewOptions } from "./state/context";
import { useFileOps } from "./state/useFileOps";
import { useFormatPrompt } from "./state/useFormatPrompt";
import { OPEN_ACCEPT } from "./io/formats";
import { useDraftRecovery } from "./state/useDraftRecovery";
import { forgetOwnDraft } from "./state/drafts";
import { guardSettings, useProtection } from "./state/useProtection";
import { ProtectionBar } from "./dialogs/ProtectDialog";
import { DocumentView } from "./views/DocumentView";
import { BelgeBackstage } from "./views/BelgeBackstage";
import { DraftBar } from "./views/DraftBar";
import { AppTitleBar } from "./views/AppTitleBar";
import { DuplexDialog } from "./print/DuplexDialog";
import { MiniToolbarHost, type MiniBarAnchor } from "./views/MiniToolbar";
import { ProofPane } from "./views/ProofPane";
import { StylesPane } from "./views/StylesPane";
import { ErrorBoundary, recovery } from "./views/ErrorBoundary";
import { usePrint } from "./state/usePrint";
import { useFocusReturn, useVersionBump, useWindowFileDrop } from "./state/useShellEffects";
import { useZoom } from "./state/useZoom";
import { useFieldSync, useHeaderFooterReturn } from "./state/useFieldSync";
import { useImageInsert } from "./state/useImageInsert";
import { useComments } from "./state/useComments";
import { useEditorClicks } from "./state/useEditorClicks";
import { useFormatPainter } from "./state/useFormatPainter";
import { ReadMode } from "./views/ReadMode";
import { BelgeStatusBar } from "./views/BelgeStatusBar";
import { ReadAloudBar } from "./views/ReadAloudBar";
import { ContextMenuContent, menuPointForSelection, selectContextTarget, showsMiniBar, targetOfSelection } from "./views/contextMenu";
import { DialogHost } from "./dialogs/DialogHost";
import { HomeTab } from "./ribbon/HomeTab";
import { InsertTab } from "./ribbon/InsertTab";
import { DesignTab } from "./ribbon/DesignTab";
import { LayoutTab } from "./ribbon/LayoutTab";
import { ReferencesTab } from "./ribbon/ReferencesTab";
import { ReviewTab } from "./ribbon/ReviewTab";
import { HeaderFooterTab, HelpTab, ViewTab } from "./ribbon/ViewTab";
import { TableDesignTab, TableLayoutTab } from "./ribbon/TableTabs";
import { PictureTab } from "./ribbon/PictureTab";
import { ShapeTab } from "./ribbon/ShapeTab";
import { MailingsTab } from "./ribbon/MailingsTab";
import { OutlineTab } from "./ribbon/OutlineTab";
import { MergeContext, useMailMerge } from "./mailmerge/useMailMerge";
import { EQUATION_EVENT } from "./editor/equation";
import { updateFields } from "./editor/fields";
import { FIELD_REFRESH_EVENT } from "./editor/nodes";
import { CHART_EVENT } from "./editor/chart";
import "./styles/workspace.css";
import "./styles/panes.css";
import "./styles/ribbon.css";
import "./styles/dialogs.css";
import "./styles/document.css";
import "./styles/document-marks.css";
import "./styles/ui-extras.css";
import "./styles/doc-extras.css";
import "./styles/views.css";
import "./styles/features.css";
import "./styles/features5.css";

type Note = { text: string; kind: "info" | "warning" | "error" | "success"; id: number };

// Değişiklik izleme ve açıklamalar Seçenekler'deki kullanıcı adını kullanır.
setTrackAuthorSource(getUserName);

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
  const [version, setVersion] = useState(0);
  const [navFocus, setNavFocus] = useState(0);
  const [spellcheck, setSpellcheck] = useState(true);
  const [lang, setLang] = useState("tr");
  const [autosave, setAutosave] = useState(true);
  const [readAloud, setReadAloud] = useState(false);
  const toggleReadAloud = useCallback(() => setReadAloud((r) => !r), []);
  const scrollRef = useRef<HTMLDivElement>(null);
  const insertImageFileRef = useRef<(f: File, pos: number | null, view: import("@tiptap/pm/view").EditorView) => Promise<void>>(async () => {});
  const imageInput = useRef<HTMLInputElement>(null);
  const openFileRef = useRef<(f: File) => Promise<void>>(async () => {});
  const ctxMenu = useContextMenu();
  const miniBar = useRef<((a: MiniBarAnchor | null) => void) | null>(null);

  const styles = useMemo(
    () => resolveStyles(settings.styleSet, settings.theme, settings.fontTheme, settings.paragraphSpacing, settings.customStyles),
    [settings.styleSet, settings.theme, settings.fontTheme, settings.paragraphSpacing, settings.customStyles],
  );
  const stylesRef = useRef(styles);
  stylesRef.current = styles;

  const notify = useCallback((text: string, kind: Note["kind"] = "info") => setNote({ text, kind, id: Date.now() }), []);
  useEffect(() => {
    if (!note || note.kind === "error" || note.kind === "warning") return;
    const id = window.setTimeout(() => setNote((n) => (n?.id === note.id ? null : n)), 4000);
    return () => window.clearTimeout(id);
  }, [note]);

  // Düzenlemeyi Kısıtla zorlanıyorsa izin verilmeyen ayar değişiklikleri yok sayılır.
  const updateSettings = useCallback((fn: (s: DocSettings) => DocSettings) => {
    setSettingsState((s) => guardSettings(s, fn(s)));
    setDirty(true);
  }, []);

  // Yakınlaştırma değişince "Tek Sayfa" kipi biter (Word'deki gibi kaydırıcıyla yeniden yan yana dizilir).
  const setView = useCallback((patch: Partial<ViewOptions>) => setViewState((v) => ({ ...v, ...(patch.zoom !== undefined && patch.onePage === undefined ? { onePage: false } : {}), ...patch })), []);

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

  useWindowFileDrop(openFileRef);

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

  useVersionBump(editor, useCallback(() => setVersion((v) => v + 1), []));

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
    // Önceki belgenin iletisi (ör. uyumluluk modu) yeni belgeye taşınmaz.
    setNote(null);
    setHfEditing(null);
    setActiveComment(null);
    setVersion((v) => v + 1);
    requestAnimationFrame(() => setShowMarks(editor.view, view.showMarks));
  }, [editor, view.showMarks]);

  // Belge yaşam döngüsü (masaüstü kabuğu ile kaydetme sorusu, dosya açma istekleri).
  const confirmRef = useRef<() => Promise<boolean>>(async () => true);
  const confirmStable = useCallback(() => confirmRef.current(), []);
  const formatPrompt = useFormatPrompt();
  const files = useFileOps({ editor, settings, setSettings: setSettingsState, dirty, setDirty, notify, confirmDiscard: confirmStable, autosave, onLoaded, askFormat: formatPrompt.ask });
  const { dialog: lifecycleDialog, confirmDiscard } = useDocumentLifecycle({
    title: files.title,
    dirty,
    path: files.fileRef?.path,
    onOpenPath: (p) => void files.openPath(p),
    onSave: files.save,
    onClose: forgetOwnDraft,
  });
  confirmRef.current = confirmDiscard;

  // Hata sınırları için güncel belge anlık görüntüsü ve kaydetme işlevi.
  useProtection([editor, hfEditor], settings);
  recovery.snapshot = () => (editor.isDestroyed ? null : { content: editor.getJSON(), settings, fileName: files.fileName, savedAt: Date.now() });
  recovery.save = files.save;
  // Açılışta: aktarılan belge (birleştirme sonucu / kurtarılan taslak), çökme kurtarması ve sahipsiz taslaklar.
  const recoveryUi = useDraftRecovery({ load: files.load, setDirty, isPristine: files.isPristine, onPrint: () => setBackstage("print"), notify });
  const merge = useMailMerge(editor, () => controller.request(), settings.mailMerge ?? null, updateSettings);

  useFocusReturn(active, !!dialog);

  const openDialog = useCallback((kind: DialogKind, arg?: string) => setDialog({ kind, arg }), []);
  // Denkleme çift tıklama: denklem seçilir ve Denklem iletişim kutusu açılır.
  useEffect(() => {
    const on = (e: Event) => {
      const { pos, editor: ed } = (e as CustomEvent<{ pos?: number; editor: Editor }>).detail;
      if (typeof pos === "number" && !ed.isDestroyed) ed.view.dispatch(ed.state.tr.setSelection(NodeSelection.create(ed.state.doc, pos)));
      openDialog("equation");
    };
    // Grafiğe çift tıklama: Verileri Düzenle.
    const onChart = (e: Event) => {
      const { pos, editor: ed } = (e as CustomEvent<{ pos?: number; editor: Editor }>).detail;
      if (typeof pos === "number" && !ed.isDestroyed) ed.view.dispatch(ed.state.tr.setSelection(NodeSelection.create(ed.state.doc, pos)));
      openDialog("chart");
    };
    window.addEventListener(EQUATION_EVENT, on);
    window.addEventListener(CHART_EVENT, onChart);
    return () => {
      window.removeEventListener(EQUATION_EVENT, on);
      window.removeEventListener(CHART_EVENT, onChart);
    };
  }, [openDialog]);
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

  const { onImageFile, pickImage } = useImageInsert({ editor, active, settings, notify, setTab, insertRef: insertImageFileRef, input: imageInput });
  const insertImageFromFile = useCallback(() => pickImage(null), [pickImage]);

  const updateToc = useFieldSync(editor, controller, settings);

  // F9: alanları güncelleştir (resim yazısı/çapraz başvuru/sayfa/tarih); seçimde içindekiler varsa o da yenilenir.
  const updateAllFields = useCallback(() => {
    const ed = active;
    const tr = updateFields(ed.state, (p) => controller.pageOf(p));
    if (tr) ed.view.dispatch(tr);
    controller.request();
    window.dispatchEvent(new CustomEvent(FIELD_REFRESH_EVENT));
    if (ed === editor) {
      const { from, to } = editor.state.selection;
      let toc = false;
      editor.state.doc.nodesBetween(Math.max(0, from - 1), Math.min(editor.state.doc.content.size, Math.max(to, from + 1)), (n) => {
        if (n.type.name === "toc") toc = true;
        return !toc;
      });
      if (toc) updateToc();
    }
  }, [active, editor, controller, updateToc]);

  const insertNote = useCallback(
    (kind: "footnote" | "endnote") => {
      if (hfEditing) return;
      editor.chain().focus().insertNote(kind).run();
      const before = editor.state.doc.resolve(editor.state.selection.from).nodeBefore;
      const id = before?.type.name === "note" ? (before.attrs.id as string) : null;
      if (!id) return;
      // Not metni alanı sayfalama sonrası çizilir; hazır olunca odaklan. Bu arada basılan karakterler gövdeye
      // yazılmaz: gövde odağı hemen bırakır, tuşlar biriktirilir ve not alanına aktarılır (Word'deki gibi).
      editor.view.dom.blur();
      let pending = "";
      const buffer = (e: KeyboardEvent) => {
        if (e.ctrlKey || e.metaKey || e.altKey || e.isComposing) return;
        if (e.key.length === 1) pending += e.key;
        else if (e.key === "Backspace") pending = pending.slice(0, -1);
        else return;
        e.preventDefault();
      };
      window.addEventListener("keydown", buffer, true);
      let tries = 0;
      const focus = () => {
        const el = document.querySelector<HTMLElement>(`.ProseMirror[data-note-id="${id}"]`);
        if (el || tries++ >= 90) window.removeEventListener("keydown", buffer, true);
        if (el) {
          el.focus();
          const ne = getActiveNoteEditor();
          if (pending && ne) ne.commands.insertContent(pending);
          else if (pending) document.execCommand("insertText", false, pending);
        } else if (tries < 90) requestAnimationFrame(focus);
        else if (pending) editor.chain().focus().insertContent(pending).run();
      };
      requestAnimationFrame(focus);
    },
    [editor, hfEditing],
  );

  insertNoteRef.current = insertNote;

  /* ---------------- Yakınlaştırma ve yazdırma ---------------- */

  const zoomTo = useZoom(controller, editor, scrollRef, view.zoom, setView);

  const leaveModes = useCallback(() => {
    setBackstage(null);
    setHfEditing(null);
    setView({ mode: "print", focusMode: false });
  }, [setHfEditing, setView]);
  const currentPage = useCallback(() => controller.pageOf(editor.state.selection.from), [controller, editor]);
  const { printing, print, options: printOptions, setOptions: setPrintOptions, duplex } = usePrint({ controller, settings, zoom: view.zoom, fileName: files.fileName, currentPage, leaveModes, notify });
  const offscreenRef = useRef<HTMLDivElement>(null);

  /* ---------------- Kısayollar ---------------- */

  useHotkeys({
    "Mod+S": () => void files.save(),
    F12: () => void files.saveAs(),
    "Mod+O": () => void files.openDialog(),
    "Mod+N": () => void files.newDocument("blank"),
    // Word'deki gibi Ctrl+P, Dosya › Yazdır sayfasını (ayarlar ve önizleme) açar.
    "Mod+P": () => setBackstage("print"),
    "Mod+F": openFind,
    F7: () => setView({ proofPane: true }),
    F9: updateAllFields,
    "Mod+Alt+Shift+S": () => setView({ stylesPane: !view.stylesPane }),
    "Mod+H": () => openDialog("findReplace", "replace"),
    "Mod+G": () => openDialog("findReplace", "goto"),
    "Mod+Shift+G": () => openDialog("wordCount"),
    "Alt+=": () => openDialog("equation"),
    "Alt+Shift+X": () => openDialog("markEntry"),
    "Mod+Alt+Space": toggleReadAloud,
    Escape: () => {
      if (view.focusMode) setView({ focusMode: false });
      else if (hfEditing) setHfEditing(null);
    },
  });

  // Ctrl+tıklama ile köprüyü aç.
  const openLink = useEditorClicks(editor);

  /**
   * Sağ tık (ya da Shift+F10 / Menü tuşu): tıklanan nesne seçilir ve ona özgü menü açılır; metinde imleç
   * seçimin dışındaysa tıklanan yere taşınır. Klavyeyle açılınca (olay düzenleyici köküne gelir) hedef
   * geçerli seçimdir ve menü seçimin yanında belirir.
   */
  const onContextMenu = (e: React.MouseEvent) => {
    const t = e.target as HTMLElement;
    if (t.closest(".bl-comments")) return;
    const ed = active.view.dom.contains(t) ? active : editor.view.dom.contains(t) ? editor : null;
    // Klavyeyle açılan menünün olayı odaktaki düzenleyici köküne gelir (fare düğmesi yok).
    const keyboard = !!ed && t === ed.view.dom && e.button !== 2;
    const target = !ed ? targetOfSelection(active.state) : keyboard ? targetOfSelection(ed.state) : selectContextTarget(ed.view, t, { left: e.clientX, top: e.clientY });
    const at = keyboard && ed ? menuPointForSelection(ed.view) : { x: e.clientX, y: e.clientY };
    const deps = { editor: ed ?? active, openLink, replaceImage: () => pickImage(ed ?? active) };
    ctxMenu.open({ clientX: at.x, clientY: at.y, preventDefault: () => e.preventDefault() }, <ContextMenuContent target={target} deps={deps} />);
    // Word'deki gibi metin menüsünün üstünde mini araç çubuğu (nesne menülerinde yok).
    if (ed && showsMiniBar(target.kind)) miniBar.current?.({ x: at.x, y: at.y, context: true });
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
    inBody: active === editor,
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
    editHeaderFooterAt: (which, page) => {
      setHfEditing(null);
      requestAnimationFrame(() => setHfEditing(which, page));
    },
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
    showTab: setTab,
    readAloud,
    toggleReadAloud,
  };

  const tabs = [
    { id: "outline", label: "Anahat", content: <OutlineTab />, hidden: view.mode !== "outline" },
    { id: "home", label: "Giriş", content: <HomeTab /> },
    { id: "insert", label: "Ekle", content: <InsertTab /> },
    { id: "design", label: "Tasarım", content: <DesignTab /> },
    { id: "layout", label: "Düzen", content: <LayoutTab /> },
    { id: "references", label: "Başvurular", content: <ReferencesTab /> },
    { id: "mailings", label: "Postalar", content: <MailingsTab /> },
    { id: "review", label: "Gözden Geçir", content: <ReviewTab lang={lang} setLang={setLang} deleteComment={deleteComment} deleteAllComments={deleteAllComments} /> },
    { id: "view", label: "Görünüm", content: <ViewTab zoomTo={zoomTo} /> },
    { id: "help", label: "Yardım", content: <HelpTab onFeedback={() => window.open("https://github.com/fyildirim-debug/officehere/issues", "_blank", "noopener")} /> },
    { id: "hf", label: "Üst Bilgi ve Alt Bilgi", content: <HeaderFooterTab />, contextual: true, hidden: !hfEditing },
    { id: "tableDesign", label: "Tablo Tasarımı", content: <TableDesignTab />, contextual: true, hidden: !fmt?.inTable || !!hfEditing },
    { id: "tableLayout", label: "Düzen", content: <TableLayoutTab />, contextual: true, hidden: !fmt?.inTable || !!hfEditing },
    { id: "picture", label: "Resim Biçimi", content: <PictureTab />, contextual: true, hidden: !fmt?.image },
    { id: "shape", label: "Şekil Biçimi", content: <ShapeTab />, contextual: true, hidden: !fmt?.textBox || !!hfEditing },
  ];

  // Üst bilgi düzenlemesi başlayınca bağlamsal sekmeye geç; kapanınca imleç gövdedeki eski yerine döner.
  const hfReturn = useRef<{ from: number; to: number } | null>(null);
  // Anahat görünümüne geçince Anahat sekmesi açılır.
  useEffect(() => setTab((t) => (view.mode === "outline" ? "outline" : t === "outline" ? "view" : t)), [view.mode]);
  useHeaderFooterReturn(editor, hfEditing, setTab, hfReturn);


  const openLocalFile = async (file: File) => {
    if (!(await confirmDiscard())) return;
    setBackstage(null);
    let data: ArrayBuffer;
    try {
      data = await file.arrayBuffer();
    } catch (err) {
      return notify(`"${file.name}" okunamadı: ${(err as Error).message}`, "error");
    }
    await files.openData(file.name, data, { name: file.name });
  };
  const onOpenFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (file) await openLocalFile(file);
  };
  openFileRef.current = openLocalFile;

  const backstageEl =
    backstage !== null ? (
      <ErrorBoundary name="backstage" label="Dosya görünümü" onReset={() => setBackstage(null)}>
      <BelgeBackstage
        page={backstage}
        files={files}
        settings={settings}
        pageCount={pageCount}
        stats={docStats(editor.state.doc)}
        printOptions={printOptions}
        setPrintOptions={setPrintOptions}
        autosave={autosave}
        setAutosave={setAutosave}
        spellcheck={spellcheck}
        setSpellcheck={setSpellcheck}
        onPrint={(o) => void print("print", o)}
        onPdf={(o) => void print("pdf", o)}
        onClose={() => setBackstage(null)}
        onDialog={(k) => openDialog(k)}
      />
      </ErrorBoundary>
    ) : undefined;

  const belowRibbon = (
    <>
      <DraftBar drafts={recoveryUi.drafts} onRestore={recoveryUi.restore} onDiscard={recoveryUi.discard} onClose={recoveryUi.close} />
      {readAloud && (
        <ErrorBoundary name="readAloud" label="Sesli Oku" onReset={() => setReadAloud(false)}>
          <ReadAloudBar editor={editor} notify={notify} onClose={() => setReadAloud(false)} />
        </ErrorBoundary>
      )}
      {settings.protection?.enforced && <ProtectionBar mode={settings.protection.mode} onStop={() => openDialog("protect")} />}
      {note && (
        <MessageBar kind={note.kind} onClose={() => setNote(null)}>
          {note.text}
        </MessageBar>
      )}
    </>
  );

  const hideChrome = view.focusMode;
  // Ana ve ekran dışı (Yazdır önizlemesi) belge görünümlerinin ortak özellikleri.
  const docProps = { editor, settings, updateSettings, controller, setView, setHfEditing, onHfEditor: setHfEditor, setActiveComment, deleteComment, version };

  return (
    <BelgeContext.Provider value={ctx}>
      <MergeContext.Provider value={merge}>
      <style>{styleSheetCss(styles)}</style>
      {view.mode === "read" && !printing && (
        <ReadMode
          html={annotateNotesHtml(editor.getHTML(), renderNoteHtml)}
          title={files.title}
          onExit={() => setView({ mode: "print" })}
          onFind={() => {
            setView({ mode: "print" });
            openFind();
          }}
        />
      )}
      <div className={`bl-app${hideChrome ? " is-focus" : ""}${printing ? " is-printing" : ""}${view.mode === "read" ? " is-read" : ""}${merge.highlight ? " is-merge-hl" : ""}`}>
        <AppShell
          app="belge"
          titleBar={
            <AppTitleBar
              title={files.title}
              dirty={dirty}
              active={active}
              canUndo={!!fmt?.canUndo}
              canRedo={!!fmt?.canRedo}
              onSave={() => void files.save()}
              onPrint={() => setBackstage("print")}
              onFind={openFind}
              search={{
                tabs,
                collapsed: ribbonCollapsed,
                onExpand: () => setRibbonCollapsed(false),
                onTab: (id) => {
                  setBackstage(null);
                  setTab(id);
                },
                onBackstage: setBackstage,
                onNew: () => void files.newDocument("blank"),
                onOpen: () => void files.openDialog(),
                onSaveAs: () => void files.saveAs(),
                onPdf: () => void print("pdf", printOptions),
              }}
            />
          }
          ribbon={
            hideChrome ? undefined : (
              <ErrorBoundary name="ribbon" label="Şerit">
                <Ribbon tabs={tabs} activeTab={tab} onTabChange={setTab} onFileClick={() => setBackstage("home")} collapsed={ribbonCollapsed} onCollapsedChange={setRibbonCollapsed} />
              </ErrorBoundary>
            )
          }
          belowRibbon={hideChrome ? undefined : belowRibbon}
          statusBar={
            hideChrome ? undefined : (
              <ErrorBoundary name="status" label="Durum çubuğu">
              <BelgeStatusBar
                editor={editor}
                controller={controller}
                paged={view.mode === "print"}
                tracking={fmt?.tracking ?? false}
                view={view}
                setView={setView}
                onPageClick={() => {
                  setView({ navPane: true });
                }}
                onWordsClick={() => openDialog("wordCount")}
                onZoomClick={() => openDialog("zoom")}
                lang={lang}
                progress={files.progress}
              />
              </ErrorBoundary>
            )
          }
          backstage={backstageEl}
        >
          <ErrorBoundary name="document" label="Belge görünümü">
          <DocumentView
            {...docProps}
            view={view}
            hfEditing={hfEditing}
            hfPage={hfPage}
            activeComment={activeComment}
            navFocus={navFocus}
            onPageSetup={() => openDialog("pageSetup")}
            onContextMenu={onContextMenu}
            scrollRef={scrollRef}
            printing={false}
            formatPainter={painter.active}
            rightPane={
              <>
                {view.stylesPane && (
                  <ErrorBoundary name="styles" label="Stiller bölmesi" onReset={() => setView({ stylesPane: false })}>
                    <StylesPane onClose={() => setView({ stylesPane: false })} onEdit={(id) => openDialog("style", id)} />
                  </ErrorBoundary>
                )}
                {view.proofPane && (
                  <ErrorBoundary name="proof" label="Düzenleyici" onReset={() => setView({ proofPane: false })}>
                    <ProofPane editor={editor} version={version} spellcheck={spellcheck} setSpellcheck={setSpellcheck} onClose={() => setView({ proofPane: false })} />
                  </ErrorBoundary>
                )}
              </>
            }
          />
          </ErrorBoundary>
        </AppShell>
      </div>
      <ErrorBoundary name="dialog" label="İletişim kutusu" onReset={() => setDialog(null)}>
        <DialogHost dialog={dialog} onClose={() => setDialog(null)} zoom={view.zoom} onZoom={zoomTo} pages={pageCount} />
      </ErrorBoundary>
      {backstage !== null && (
        // Dosya görünümü açıkken belge ekran dışında yazdırma düzeninde çizilir (Yazdır önizlemesi için).
        <div className="bl-offscreen" aria-hidden="true">
          <DocumentView {...docProps} view={{ ...view, mode: "print", zoom: 100, focusMode: false, navPane: false, showComments: false, ruler: false, onePage: true }} hfEditing={null} hfPage={0} activeComment={null} navFocus={0} onPageSetup={() => {}} onContextMenu={() => {}} scrollRef={offscreenRef} printing formatPainter={false} />
        </div>
      )}
      <DuplexDialog open={duplex.open} onNext={duplex.next} onCancel={duplex.cancel} />
      {!backstage && !hideChrome && view.mode !== "read" && (
        <ErrorBoundary name="minibar" label="Mini araç çubuğu">
          <MiniToolbarHost controlRef={miniBar} />
        </ErrorBoundary>
      )}
      {ctxMenu.element}
      {lifecycleDialog}
      {formatPrompt.dialog}
      <input ref={imageInput} type="file" accept="image/*" className="oh-hidden-input" onChange={onImageFile} />
      {/* Dosya açma girişi (sürükle-bırak ve dosya seçicisi olmayan ortamlar için de kullanılır). */}
      <input id="bl-open-input" type="file" accept={OPEN_ACCEPT} className="oh-hidden-input" aria-label="Belge aç" onChange={onOpenFile} />
      {view.focusMode && (
        <button type="button" className="bl-focus-exit" onClick={() => setView({ focusMode: false })}>
          Odaktan çık (Esc)
        </button>
      )}
      </MergeContext.Provider>
    </BelgeContext.Provider>
  );
}
