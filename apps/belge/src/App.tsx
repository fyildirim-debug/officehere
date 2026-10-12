// OfficeHere Belge: uygulama kabuğu (başlık, şerit, çalışma alanı, durum çubuğu, Dosya görünümü).
import { useCallback, useEffect, useMemo, useRef, useState, useSyncExternalStore } from "react";
import { flushSync } from "react-dom";
import { useEditor, useEditorState } from "@tiptap/react";
import type { Editor } from "@tiptap/core";
import { NodeSelection, TextSelection } from "@tiptap/pm/state";
import { ArrowRedo20Regular, ArrowUndo20Regular, Save20Regular, Print20Regular, DocumentPdf24Regular, Code24Regular, DocumentText24Regular, Document24Regular } from "@fluentui/react-icons";
import { AppShell, MessageBar, OfficeBackstage, QatButton, Ribbon, TitleBar, getUserName, useContextMenu, useDocumentLifecycle, useHotkeys, type OfficeBackstagePage } from "@officehere/ui";
import { defaultSettings, emptyDoc, pageMetrics, type DocSettings } from "./model/document";
import { resolveStyles, styleSheetCss } from "./model/styles";
import { editorExtensions } from "./editor/extensions";
import { PaginationController } from "./editor/pagination";
import { computeFormat, sameFormat } from "./editor/formatState";
import { collectHeadings, copyFormat, countWords, docStats, pasteFormat, wordAt, type CopiedFormat } from "./editor/commands";
import { setShowMarks } from "./editor/showMarks";
import { isTracking, setTracking } from "./editor/trackChanges";
import { readImageFile } from "./editor/image";
import { annotateNotesHtml } from "./editor/notes";
import { BelgeContext, type BelgeCtx, type DialogKind, type DialogState, type ViewOptions } from "./state/context";
import { clearDraft, readDraft, useFileOps, type Draft } from "./state/useFileOps";
import { DocumentView, PAGE_GAP } from "./views/DocumentView";
import { buildPrintRoot, removePrintRoot } from "./views/printRender";
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
import { TEMPLATES } from "./templates/templates";
import { initialsOf } from "./io/docx/reader";
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
  const [hfEditing, setHfEditing] = useState<"header" | "footer" | null>(null);
  const [hfEditor, setHfEditor] = useState<Editor | null>(null);
  const [activeComment, setActiveComment] = useState<string | null>(null);
  const [note, setNote] = useState<Note | null>(null);
  const [draft, setDraft] = useState<Draft | null>(() => readDraft());
  const [printing, setPrinting] = useState(false);
  const [painter, setPainter] = useState<{ active: boolean; locked: boolean; fmt: CopiedFormat | null }>({ active: false, locked: false, fmt: null });
  const [version, setVersion] = useState(0);
  const [navFocus, setNavFocus] = useState(0);
  const [spellcheck, setSpellcheck] = useState(true);
  const [lang, setLang] = useState("tr");
  const [autosave, setAutosave] = useState(true);
  const scrollRef = useRef<HTMLDivElement>(null);
  const insertImageFileRef = useRef<(f: File, pos: number | null, view: import("@tiptap/pm/view").EditorView) => Promise<void>>(async () => {});
  const imageInput = useRef<HTMLInputElement>(null);
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
    autofocus: "start",
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

  const active = hfEditing && hfEditor ? hfEditor : editor;
  const fmt = useEditorState({
    editor: active,
    selector: ({ editor: ed }) => computeFormat(ed, stylesRef.current),
    equalityFn: (a, b) => !!a && !!b && sameFormat(a, b),
  });

  const sel = useEditorState({
    editor,
    selector: ({ editor: ed }) => {
      const { from, to, empty } = ed.state.selection;
      return { from, words: empty ? -1 : countWords(ed.state.doc.textBetween(from, to, "\n", " ")) };
    },
    equalityFn: (a, b) => !!a && !!b && a.from === b.from && a.words === b.words,
  });

  const stats = useMemo(() => docStats(editor.state.doc), [editor, version]); // eslint-disable-line react-hooks/exhaustive-deps
  const layout = useSyncExternalStore(
    useCallback((cb: () => void) => controller.subscribe(cb), [controller]),
    () => controller.info,
  );

  // Tümünü Göster
  useEffect(() => {
    setShowMarks(editor.view, view.showMarks);
  }, [editor, view.showMarks, version === 0]); // eslint-disable-line react-hooks/exhaustive-deps

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
  activeRef.current = hfEditing && hfEditor ? hfEditor : editor;
  useEffect(() => {
    const onUp = (e: PointerEvent) => {
      const t = e.target as HTMLElement;
      if (t.closest?.(".bl-body, .bl-hf-edit, input, textarea, select, .oh-dialog, .bl-nav, .bl-comments, .oh-titlebar, .oh-backstage, .bl-read")) return;
      window.setTimeout(() => {
        if ((document.activeElement === document.body || document.activeElement === null) && !document.querySelector(".oh-dialog, [data-oh-popover]")) activeRef.current.view.focus();
      }, 0);
    };
    document.addEventListener("pointerup", onUp, true);
    return () => document.removeEventListener("pointerup", onUp, true);
  }, []);

  /* ---------------- Açıklamalar ve değişiklik izleme ---------------- */

  const newComment = useCallback(() => {
    let { from, to } = editor.state.selection;
    if (from === to) {
      const w = wordAt(editor.state.doc, from);
      if (!w) return notify("Açıklama eklemek için metin seçin.");
      from = w.from;
      to = w.to;
    }
    const id = `c${Date.now().toString(36)}`;
    const author = getUserName();
    editor.chain().setTextSelection({ from, to }).setComment(id).run();
    updateSettings((s) => ({ ...s, comments: { ...s.comments, [id]: { id, author, initials: initialsOf(author), date: new Date().toISOString(), text: "" } } }));
    setActiveComment(id);
    setView({ showComments: true });
  }, [editor, notify, updateSettings, setView]);

  const deleteComment = useCallback(
    (id: string) => {
      editor.commands.unsetComment(id);
      updateSettings((s) => {
        const c = { ...s.comments };
        delete c[id];
        return { ...s, comments: c };
      });
      setActiveComment(null);
    },
    [editor, updateSettings],
  );

  const deleteAllComments = useCallback(() => {
    const type = editor.state.schema.marks.comment;
    const tr = editor.state.tr.removeMark(0, editor.state.doc.content.size, type);
    editor.view.dispatch(tr);
    updateSettings((s) => ({ ...s, comments: {} }));
  }, [editor, updateSettings]);

  const toggleTrack = useCallback(() => {
    const on = !isTracking(editor.state);
    setTracking(editor.view, on, getUserName());
  }, [editor]);

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

  /* ---------------- Biçim Boyacısı ---------------- */

  const startFormatPainter = useCallback(
    (lock: boolean) => {
      if (painter.active && !lock) return setPainter({ active: false, locked: false, fmt: null });
      setPainter({ active: true, locked: lock, fmt: copyFormat(editor) });
    },
    [editor, painter.active],
  );

  useEffect(() => {
    if (!painter.active || !painter.fmt) return;
    const dom = editor.view.dom;
    const onUp = () => {
      requestAnimationFrame(() => {
        pasteFormat(editor, painter.fmt!);
        if (!painter.locked) setPainter({ active: false, locked: false, fmt: null });
      });
    };
    const onKey = (e: KeyboardEvent) => e.key === "Escape" && setPainter({ active: false, locked: false, fmt: null });
    dom.addEventListener("mouseup", onUp);
    window.addEventListener("keydown", onKey);
    return () => {
      dom.removeEventListener("mouseup", onUp);
      window.removeEventListener("keydown", onKey);
    };
  }, [editor, painter]);

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
        const el = document.querySelector<HTMLTextAreaElement>(`textarea[data-note-id="${id}"]`);
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
      const m = pageMetrics(settings.page);
      const availW = (sc?.clientWidth ?? 1000) - 48;
      const availH = (sc?.clientHeight ?? 700) - 48;
      let pct = view.zoom;
      const target = typeof z === "string" ? (z === "100" ? { pct: 100 } : { fit: z === "onePage" ? "page" : z }) : z;
      if ("pct" in target) pct = target.pct;
      else if (target.fit === "width") pct = (availW / m.w) * 100;
      else if (target.fit === "text") pct = (availW / m.contentW) * 100;
      else if (target.fit === "page") pct = Math.min(availW / m.w, availH / m.h) * 100;
      else if (target.fit === "multi") pct = Math.min(availW / m.w, availH / (m.h * 2 + 16)) * 100;
      setView({ zoom: Math.max(10, Math.min(500, Math.round(pct))) });
    },
    [settings.page, view.zoom, setView],
  );

  const preparePrint = useCallback(() => {
    const pagesEl = document.querySelector<HTMLElement>(".bl-pages");
    if (!pagesEl) return false;
    const m = pageMetrics(settings.page);
    let style = document.getElementById("bl-page-style");
    if (!style) {
      style = document.createElement("style");
      style.id = "bl-page-style";
      document.head.appendChild(style);
    }
    style.textContent = `@page { size: ${m.w}px ${m.h}px; margin: 0; }`;
    const info = controller.info;
    return !!buildPrintRoot(pagesEl, { pageCount: info.pageCount, pageH: m.h, pageW: m.w, stride: m.h + PAGE_GAP, scale: view.zoom / 100, columns: settings.page.columns, pageColor: settings.pageColor });
  }, [settings.page, settings.pageColor, controller, view.zoom]);

  const finishPrint = useCallback(() => {
    removePrintRoot();
    document.getElementById("bl-page-style")?.remove();
  }, []);

  const print = useCallback(async () => {
    flushSync(() => {
      setBackstage(null);
      setHfEditing(null);
      setView({ mode: "print", focusMode: false });
    });
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
    controller.measureNow?.();
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
    if (!preparePrint()) return;
    setPrinting(true);
    try {
      window.print();
    } finally {
      finishPrint();
      setPrinting(false);
    }
  }, [controller, setView, preparePrint, finishPrint]);

  useEffect(() => {
    if (import.meta.env.DEV) Object.assign((window as unknown as { __belge: object }).__belge ?? {}, { preparePrint, finishPrint });
  }, [preparePrint, finishPrint]);

  // Tarayıcı menüsünden yazdırma.
  useEffect(() => {
    const before = () => {
      if (!document.getElementById("bl-print-root")) preparePrint();
    };
    window.addEventListener("beforeprint", before);
    window.addEventListener("afterprint", finishPrint);
    return () => {
      window.removeEventListener("beforeprint", before);
      window.removeEventListener("afterprint", finishPrint);
    };
  }, [preparePrint, finishPrint]);

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
  const openLink = useCallback(
    (href: string) => {
      if (href.startsWith("#")) {
        const name = href.slice(1);
        let target = -1;
        editor.state.doc.descendants((n, pos) => {
          if (target < 0 && n.type.name === "bookmark" && n.attrs.name === name) target = pos;
          return target < 0;
        });
        if (name === "_top") target = 0;
        if (target >= 0) editor.view.dispatch(editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(target))).scrollIntoView());
        return;
      }
      window.open(href, "_blank", "noopener");
    },
    [editor],
  );

  useEffect(() => {
    const dom = editor.view.dom;
    const onClick = (e: MouseEvent) => {
      const a = (e.target as HTMLElement).closest("a[href]");
      if (a && (e.ctrlKey || e.metaKey)) {
        e.preventDefault();
        openLink(a.getAttribute("href") ?? "");
      } else if (a) e.preventDefault();
      const toc = (e.target as HTMLElement).closest(".bl-toc-entry");
      if (toc && (e.ctrlKey || e.metaKey)) {
        const pos = Number(toc.getAttribute("data-pos"));
        const heading = collectHeadings(editor.state.doc).find((h) => h.pos === pos) ?? collectHeadings(editor.state.doc).find((h) => h.text === toc.querySelector(".bl-toc-text")?.textContent);
        if (heading) editor.view.dispatch(editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(heading.pos + 1))).scrollIntoView());
      }
    };
    dom.addEventListener("click", onClick);
    return () => dom.removeEventListener("click", onClick);
  }, [editor, openLink]);

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
        setView({ mode: "print" });
        requestAnimationFrame(() => setHfEditing(which));
      }
    },
    [setView],
  );

  const ctx: BelgeCtx = {
    editor: active,
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

  insertImageFileRef.current = async (file, pos, view) => {
    try {
      const m = pageMetrics(settings.page);
      const img = await readImageFile(file, m.contentW);
      const node = view.state.schema.nodes.image.create({ src: img.src, width: img.width, height: img.height, alt: file.name.replace(/\.[^.]+$/, "") });
      const at = pos ?? view.state.selection.from;
      view.dispatch(view.state.tr.insert(Math.min(at, view.state.doc.content.size), node));
    } catch (err) {
      notify((err as Error).message, "error");
    }
  };

  const onImageFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    try {
      const m = pageMetrics(settings.page);
      const img = await readImageFile(file, m.contentW);
      active.chain().focus().setImage({ src: img.src, width: img.width, height: img.height, alt: file.name.replace(/\.[^.]+$/, "") }).run();
    } catch (err) {
      notify((err as Error).message, "error");
    }
  };

  const backstageEl =
    backstage !== null ? (
      <OfficeBackstage
        app="belge"
        initialPage={backstage}
        fileName={files.fileName}
        filePath={files.fileRef?.path}
        templates={TEMPLATES.map((t) => ({ id: t.id, label: t.label, preview: t.preview }))}
        onNew={async (id) => {
          if (await files.newDocument(id)) setBackstage(null);
        }}
        onOpen={async () => {
          setBackstage(null);
          await files.openDialog();
        }}
        onOpenRecent={async (f) => {
          setBackstage(null);
          if (f.path !== f.name) await files.openPath(f.path);
          else await files.openDialog();
        }}
        onSave={async () => {
          if (await files.save()) setBackstage(null);
        }}
        onSaveAs={async () => {
          if (await files.saveAs()) setBackstage(null);
        }}
        onPrint={() => void print()}
        exports={[
          { id: "pdf", title: "PDF Belgesi Oluştur", description: "Yazdır iletişim kutusunda hedef olarak “PDF olarak kaydet” seçin.", icon: <DocumentPdf24Regular /> },
          { id: "docx", title: "Belge (.docx)", description: "Office Open XML belgesi olarak indirin.", icon: <Document24Regular /> },
          { id: "html", title: "Web Sayfası (.html)", description: "Stilleriyle birlikte tek dosyalık web sayfası.", icon: <Code24Regular /> },
          { id: "txt", title: "Düz Metin (.txt)", description: "Biçimlendirme olmadan yalnızca metin (UTF-8).", icon: <DocumentText24Regular /> },
        ]}
        onExport={async (id) => {
          if (id === "pdf") return void print();
          await files.exportAs(id);
          setBackstage(null);
        }}
        onClose={async () => {
          await files.closeDocument();
          setBackstage(null);
        }}
        onBack={() => setBackstage(null)}
        info={[
          ["Sayfalar", String(layout.pageCount)],
          ["Sözcükler", stats.words.toLocaleString("tr-TR")],
          ["Karakterler", stats.chars.toLocaleString("tr-TR")],
          ["Başlık", settings.properties.title || "Başlık ekleyin"],
          ["Yazar", settings.properties.author || getUserName()],
          ["Oluşturulma", new Date(settings.properties.created).toLocaleString("tr-TR")],
          ["Son Değiştirme", new Date(settings.properties.modified).toLocaleString("tr-TR")],
        ]}
        printPreview={<div className="bl-print-preview bl-doc" dangerouslySetInnerHTML={{ __html: editor.getHTML() }} />}
        extraOptions={
          <>
            <label className="oh-field is-inline">
              <span className="oh-field-label">Otomatik kurtarma bilgilerini her 5 saniyede bir kaydet</span>
              <input type="checkbox" checked={autosave} onChange={(e) => setAutosave(e.target.checked)} />
            </label>
            <label className="oh-field is-inline">
              <span className="oh-field-label">Yazarken yazım denetimi yap</span>
              <input type="checkbox" checked={spellcheck} onChange={(e) => setSpellcheck(e.target.checked)} />
            </label>
          </>
        }
      />
    ) : undefined;

  const belowRibbon = (
    <>
      {draft && (
        <MessageBar
          kind="warning"
          onClose={() => {
            clearDraft();
            setDraft(null);
          }}
          actions={
            <>
              <button
                type="button"
                className="oh-btn is-primary"
                onClick={() => {
                  files.load({ content: draft.content, settings: { ...defaultSettings(), ...draft.settings } }, draft.fileName, null);
                  setDirty(true);
                  setDraft(null);
                }}
              >
                Geri Yükle
              </button>
              <button
                type="button"
                className="oh-btn"
                onClick={() => {
                  clearDraft();
                  setDraft(null);
                }}
              >
                Sil
              </button>
            </>
          }
        >
          <strong>Kurtarılan belge</strong> “{draft.fileName}” için kaydedilmemiş değişiklikler bulundu ({new Date(draft.savedAt).toLocaleString("tr-TR")}).
        </MessageBar>
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
          html={annotateNotesHtml(editor.getHTML())}
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
                page={view.mode === "print" ? controller.pageOf(sel?.from ?? 0) + 1 : 1}
                pages={view.mode === "print" ? layout.pageCount : 1}
                words={stats.words}
                selWords={sel?.words ?? -1}
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
      <DialogHost dialog={dialog} onClose={() => setDialog(null)} zoom={view.zoom} onZoom={zoomTo} pages={layout.pageCount} />
      {ctxMenu.element}
      {lifecycleDialog}
      <input ref={imageInput} type="file" accept="image/*" className="oh-hidden-input" onChange={onImageFile} />
      {view.focusMode && (
        <button type="button" className="bl-focus-exit" onClick={() => setView({ focusMode: false })}>
          Odaktan çık (Esc)
        </button>
      )}
    </BelgeContext.Provider>
  );
}
