// Dosya işlemleri: yeni, aç, kaydet, farklı kaydet, dışa aktar, otomatik kurtarma ve otomatik kaydetme.
import { useCallback, useEffect, useRef, useState } from "react";
import type { Editor } from "@tiptap/core";
import { EditorState, TextSelection } from "@tiptap/pm/state";
import { addRecentFile, baseName, extensionOf, getNative, openFile, readPath, saveFile, saveFileAs, type FileRef, type SaveData } from "@officehere/core";
import { getUserName, isDirectSave, saveStatus, trackSave, useAutoSave } from "@officehere/ui";
import { defaultSettings, normalizeDoc, type BelgeDocument, type DocSettings } from "../model/document";
import { FORMATS, LoadWarnings, OPEN_FILTERS, docToText, formatOfExt, htmlDocument, isWritable, openDocument, readDocx, saveFilters, writeDocx, writeMarkdown, writeOdt, writeRtf, type FormatId } from "../io";
import { TEMPLATES } from "../templates/templates";
import { annotateNotesHtml } from "../editor/notes";
import { clearOwnDraft, writeDraft, type Draft } from "./drafts";

/** Açılışta okunamayan içerik varsa otomatik kaydetme ilk elle kayda kadar bekler (rozet nedeni gösterir). */
export const LOAD_WARNING_HOLD = "Bu dosyada desteklenmeyen öğeler var; ilk kaydı siz onaylayın";

/**
 * Otomatik kaydedilen .docx kendi okuyucumuzla yeniden açılabiliyor mu: okuyucu hata vermemeli ve
 * paketin bir kısmını bozuk/eksik saymamalı (desteklenmeyen öğe uyarıları doğrulamayı bozmaz).
 */
export async function canReopenDocx(data: ArrayBuffer): Promise<boolean> {
  const warnings = new LoadWarnings();
  const doc = await readDocx(data, warnings);
  if (!doc?.content || (["corrupt", "unclosed", "simplified", "flattened"] as const).some((k) => warnings.has(k))) return false;
  return true;
}
import type { FormatChoice } from "./useFormatPrompt";

export type { Draft } from "./drafts";

/** Geçmişi sıfırlayarak belge içeriğini yükler. */
export function loadIntoEditor(editor: Editor, doc: BelgeDocument["content"]) {
  editor.commands.setContent(normalizeDoc(doc), { emitUpdate: false });
  const state = EditorState.create({ doc: editor.state.doc, plugins: editor.state.plugins, selection: TextSelection.atStart(editor.state.doc) });
  editor.view.updateState(state);
}

interface Options {
  editor: Editor;
  settings: DocSettings;
  setSettings: (s: DocSettings) => void;
  dirty: boolean;
  setDirty: (d: boolean) => void;
  notify: (t: string, kind?: "info" | "warning" | "error" | "success") => void;
  confirmDiscard: () => Promise<boolean>;
  autosave: boolean;
  onLoaded: () => void;
  /** Yerel olmayan biçimde ilk kayıtta sorulan soru. */
  askFormat?: (label: string, ext: string, fileName: string) => Promise<FormatChoice>;
}

/** Dosya adının (yolun) bulunduğu klasördeki göreli yol (masaüstü). */
function siblingPath(path: string, rel: string): string {
  const sep = path.includes("\\") && !path.includes("/") ? "\\" : "/";
  const dir = path.slice(0, Math.max(path.lastIndexOf("/"), path.lastIndexOf("\\")) + 1);
  return dir + rel.replace(/^\.\//, "").replace(/[\\/]/g, sep);
}

export function useFileOps({ editor, settings, setSettings, dirty, setDirty, notify, confirmDiscard, autosave, onLoaded, askFormat }: Options) {
  const [fileName, setFileName] = useState("Belge1");
  const [fileRef, setFileRef] = useState<FileRef | null>(null);
  /** Geçerli dosyanın biçimi ve gerçek uzantısı (yeni belgede docx). */
  const [format, setFormat] = useState<FormatId>("docx");
  const [fileExt, setFileExt] = useState("docx");
  /** Yerel olmayan biçimde kaydetme sorusu bu dosya için yanıtlandı mı. */
  const keepAsked = useRef(false);
  /** Dosya uyarıyla açıldıysa otomatik kaydetmeyi bekleten neden (ilk elle kayda kadar). */
  const [loadHold, setLoadHold] = useState<string | null>(null);
  const counter = useRef(1);
  const settingsRef = useRef(settings);
  settingsRef.current = settings;
  const fileNameRef = useRef(fileName);
  fileNameRef.current = fileName;

  // Yazar boşsa Seçenekler'deki kullanıcı adı yazılır (Word'deki gibi).
  const currentDoc = useCallback(
    (): BelgeDocument => ({ content: editor.getJSON(), settings: { ...settingsRef.current, properties: { ...settingsRef.current.properties, author: settingsRef.current.properties.author || getUserName(), modified: new Date().toISOString() } } }),
    [editor],
  );

  /** Belgeyi istenen biçimde kodlar. */
  const encode = useCallback(
    async (f: FormatId, doc = currentDoc()): Promise<SaveData> => {
      switch (f) {
        case "odt":
          return writeOdt(doc);
        case "rtf":
          return writeRtf(doc);
        case "md":
          return writeMarkdown(doc);
        case "html":
          return htmlDocument(annotateNotesHtml(editor.getHTML()), doc.settings, doc.settings.properties.title || baseName(fileNameRef.current));
        case "txt":
          return "﻿" + docToText(doc.content);
        default:
          return writeDocx(doc);
      }
    },
    [currentDoc, editor],
  );

  const load = useCallback(
    (doc: BelgeDocument, name: string, ref: FileRef | null, fmt: FormatId = "docx", ext = "docx") => {
      setSettings(doc.settings);
      loadIntoEditor(editor, doc.content);
      setFileName(name);
      setFileRef(ref);
      setFormat(fmt);
      setFileExt(ext);
      keepAsked.current = false;
      setLoadHold(null);
      setDirty(false);
      saveStatus.reset();
      onLoaded();
    },
    [editor, setSettings, setDirty, onLoaded],
  );

  const newDocument = useCallback(
    async (templateId = "blank") => {
      if (!(await confirmDiscard())) return false;
      const t = TEMPLATES.find((x) => x.id === templateId) ?? TEMPLATES[0];
      counter.current += 1;
      const doc = t.build();
      load(doc, templateId === "blank" ? `Belge${counter.current}` : `${t.label}`, null);
      clearOwnDraft();
      return true;
    },
    [confirmDiscard, load],
  );

  const openData = useCallback(
    async (name: string, data: ArrayBuffer, ref: FileRef | null) => {
      try {
        const native = getNative();
        const path = ref?.path;
        // Markdown'daki göreli resimler masaüstünde belgenin klasöründen okunur.
        const resolveImage =
          native && path
            ? async (rel: string) => {
                const f = await readPath(siblingPath(path, rel));
                return f ? new Uint8Array(f.data) : null;
              }
            : undefined;
        const opened = await openDocument(name, data, { resolveImage });
        const { doc, warning, format: fmt } = opened;
        load(doc, baseName(name), ref, fmt, extensionOf(name) || FORMATS[fmt].extensions[0]);
        // Okunamayan (kaydedilince kaybolacak) içerik: özgün dosyanın üzerine sessizce yazılmaz.
        if ((opened.warnings ?? []).length > 0) setLoadHold(LOAD_WARNING_HOLD);
        // Önceki (vazgeçilen) belgenin taslağı bu pencereden kaldırılır.
        clearOwnDraft();
        if (warning) notify(warning, "warning");
        await addRecentFile({ path: ref?.path ?? name, name, app: "belge", openedAt: Date.now() });
      } catch (err) {
        notify(`"${name}" açılamadı: ${(err as Error).message}`, "error");
      }
    },
    [load, notify],
  );

  const openDialog = useCallback(async () => {
    if (!(await confirmDiscard())) return;
    const f = await openFile(OPEN_FILTERS);
    if (!f) return;
    await openData(f.name, f.data, { name: f.name, path: f.path, handle: f.handle });
  }, [confirmDiscard, openData]);

  const isPristine = useCallback(() => !dirty && editor.state.doc.textContent.trim() === "" && editor.state.doc.childCount <= 1 && (editor.state.doc.firstChild?.childCount ?? 0) <= 1, [dirty, editor]);

  const openPath = useCallback(
    async (path: string) => {
      if (!isPristine()) {
        const n = getNative();
        if (n) return void n.newWindow("belge", path);
      }
      const f = await readPath(path);
      if (f) await openData(f.name, f.data, { name: f.name, path: f.path });
    },
    [isPristine, openData],
  );

  /** Kayıt sonrası durum: dosya, biçim, son kullanılanlar. */
  const saved = useCallback(
    async (ref: FileRef, fmt: FormatId) => {
      setFileRef(ref);
      setFileName(baseName(ref.name));
      setFormat(fmt);
      setFileExt(extensionOf(ref.name) || FORMATS[fmt].saveExt || "docx");
      setDirty(false);
      clearOwnDraft();
      await addRecentFile({ path: ref.path ?? ref.name, name: ref.name, app: "belge", openedAt: Date.now() });
    },
    [setDirty],
  );

  /** Farklı Kaydet: biçim listesi sunulur; seçilen uzantının yazıcısı kullanılır. */
  const saveAsFormat = useCallback(
    async (first: FormatId): Promise<boolean> => {
      try {
        const ext = FORMATS[first].saveExt ?? "docx";
        const ref = await trackSave(false, async () => saveFileAs(`${baseName(fileNameRef.current)}.${ext}`, saveFilters(first), await encode(first)));
        if (!ref) return false;
        // İletişim kutusunda başka bir biçim seçildiyse dosya o biçimde yeniden yazılır.
        const chosen = formatOfExt(extensionOf(ref.name));
        let fmt: FormatId = first;
        if (chosen && chosen !== first && isWritable(chosen)) {
          const data = await encode(chosen);
          const native = getNative();
          if (ref.path && native) await saveFile(ref, ref.name, saveFilters(chosen), data);
          else if (ref.handle) await saveFile(ref, ref.name, saveFilters(chosen), data);
          fmt = chosen;
        }
        keepAsked.current = true;
        await saved(ref, fmt);
        return true;
      } catch (err) {
        notify(`Kaydedilemedi: ${(err as Error).message}`, "error");
        return false;
      }
    },
    [encode, notify, saved],
  );

  const saveAs = useCallback(async (): Promise<boolean> => saveAsFormat(isWritable(format) ? format : "docx"), [format, saveAsFormat]);

  const save = useCallback(async (): Promise<boolean> => {
    // Eski .doc ya da yazılamayan biçim: özgün dosyanın üzerine yazılmaz, .docx olarak Farklı Kaydet.
    if (!isWritable(format)) return saveAsFormat("docx");
    if (format !== "docx") {
      // Yerel olmayan biçim: ilk kayıtta bir kez sorulur.
      if (!keepAsked.current && askFormat) {
        const choice = await askFormat(FORMATS[format].short, fileExt, `${fileNameRef.current}.${fileExt}`);
        if (choice === "cancel") return false;
        if (choice === "docx") return saveAsFormat("docx");
        keepAsked.current = true;
      }
      try {
        const target = fileRef && formatOfExt(extensionOf(fileRef.name)) === format ? fileRef : null;
        const ref = await trackSave(isDirectSave(target), async () => saveFile(target, `${fileNameRef.current}.${fileExt}`, saveFilters(format), await encode(format)));
        if (!ref) return false;
        await saved(ref, format);
        notify(`"${ref.name}" kaydedildi.`, "success");
        return true;
      } catch (err) {
        notify(`Kaydedilemedi: ${(err as Error).message}`, "error");
        return false;
      }
    }
    // Yeni belge (dosyası yok): biçim listesiyle Farklı Kaydet.
    if (!fileRef) return saveAsFormat("docx");
    try {
      const target = /\.docx$/i.test(fileRef.name) ? fileRef : null;
      const name = `${baseName(fileNameRef.current)}.docx`;
      const ref = await trackSave(isDirectSave(target), async () => saveFile(target, name, target ? saveFilters("docx").slice(0, 1) : saveFilters("docx"), await writeDocx(currentDoc())));
      if (!ref) return false;
      const chosen = formatOfExt(extensionOf(ref.name));
      if (chosen && chosen !== "docx" && isWritable(chosen) && (ref.path || ref.handle)) await saveFile(ref, ref.name, saveFilters(chosen), await encode(chosen));
      keepAsked.current = chosen !== "docx";
      await saved(ref, chosen && isWritable(chosen) ? chosen : "docx");
      notify(`"${ref.name}" kaydedildi.`, "success");
      return true;
    } catch (err) {
      notify(`Kaydedilemedi: ${(err as Error).message}`, "error");
      return false;
    }
  }, [askFormat, currentDoc, encode, fileExt, fileRef, format, notify, saveAsFormat, saved]);

  const exportAs = useCallback(
    async (kind: string) => {
      const base = baseName(fileName);
      const fmt = kind as FormatId;
      if (!FORMATS[fmt]?.saveExt) return;
      const ext = FORMATS[fmt].saveExt!;
      await saveFileAs(`${base}.${ext}`, [{ name: FORMATS[fmt].label, extensions: [ext] }], await encode(fmt));
    },
    [encode, fileName],
  );

  const closeDocument = useCallback(async () => {
    if (!(await confirmDiscard())) return;
    counter.current += 1;
    load({ content: { type: "doc", content: [{ type: "paragraph" }] }, settings: defaultSettings() }, `Belge${counter.current}`, null);
    clearOwnDraft();
  }, [confirmDiscard, load]);

  // Otomatik kurtarma: kirli belgeyi birkaç saniyede bir bu pencerenin taslağına yaz.
  const dirtyRef = useRef(dirty);
  dirtyRef.current = dirty;
  const saveRef = useRef(save);
  saveRef.current = save;
  const pathRef = useRef<string | undefined>(undefined);
  pathRef.current = fileRef?.path;
  const notifyRef = useRef(notify);
  notifyRef.current = notify;
  useEffect(() => {
    if (!autosave) return;
    /** Kota uyarısı oturumda bir kez (taslak yazılamazsa kullanıcı bilmeli). */
    let warned = false;
    let writing = false;
    const write = () => {
      if (!dirtyRef.current || editor.isDestroyed || writing) return;
      const d: Draft = { content: editor.getJSON(), settings: settingsRef.current, fileName: fileNameRef.current, savedAt: Date.now(), ...(pathRef.current ? { path: pathRef.current } : {}) };
      writing = true;
      void writeDraft(d)
        .then((r) => {
          if ((r === "quota" || r === "error") && !warned) {
            warned = true;
            notifyRef.current(
              r === "quota"
                ? "Kurtarma taslağı kaydedilemedi: tarayıcı depolama alanı dolu. Çökme durumunda değişiklikler kurtarılamayabilir; belgeyi kaydedin."
                : "Kurtarma taslağı kaydedilemedi. Çökme durumunda değişiklikler kurtarılamayabilir; belgeyi kaydedin.",
              "warning",
            );
          } else if (r === "ok") warned = false;
        })
        .finally(() => {
          writing = false;
        });
    };
    const id = window.setInterval(write, 5000);
    // Sayfa kapanırken/gizlenirken son durum da yazılır (5 sn beklemeden).
    const onHide = () => write();
    window.addEventListener("pagehide", onHide);
    window.addEventListener("officehere:autosave", onHide);
    const onVis = () => document.visibilityState === "hidden" && write();
    document.addEventListener("visibilitychange", onVis);
    return () => {
      window.clearInterval(id);
      window.removeEventListener("pagehide", onHide);
      window.removeEventListener("officehere:autosave", onHide);
      document.removeEventListener("visibilitychange", onVis);
    };
  }, [editor, autosave]);

  // Otomatik kaydetme (masaüstü, yalnız yerel .docx): yazma durunca aynı dosyaya kaydedilir; taslak da silinir.
  const autoSave = useAutoSave({
    path: format === "docx" ? fileRef?.path : undefined,
    formats: ["docx"],
    holdReason: loadHold,
    verify: canReopenDocx,
    isDirty: () => dirtyRef.current,
    subscribe: (onChange) => {
      editor.on("update", onChange);
      return () => {
        editor.off("update", onChange);
      };
    },
    prepare: async () => {
      const doc = editor.state.doc;
      const s = settingsRef.current;
      return { data: await writeDocx(currentDoc()), isCurrent: () => !editor.isDestroyed && editor.state.doc === doc && settingsRef.current === s };
    },
    markClean: () => {
      // Taslak yazıcı yeniden çizimi beklemeden temiz görsün.
      dirtyRef.current = false;
      setDirty(false);
      clearOwnDraft();
    },
    save: () => void saveRef.current(),
  });
  // Belge ayarları (sayfa yapısı, üst/alt bilgi, açıklamalar…) düzenleyici olayı üretmez.
  const touchAutoSave = autoSave.touch;
  useEffect(() => touchAutoSave(), [settings, touchAutoSave]);

  /** Başlık çubuğu: yerel olmayan biçimde uzantı görünür; .doc uyumluluk modunda. */
  const title = format === "docx" ? fileName : format === "doc" ? `${fileName}.${fileExt} [Uyumluluk Modu]` : `${fileName}.${fileExt}`;

  return { isPristine, fileName, setFileName, fileRef, format, title, newDocument, openDialog, openPath, openData, save, saveAs, exportAs, closeDocument, load, currentDoc };
}
