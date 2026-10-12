// Dosya işlemleri: yeni, aç, kaydet, farklı kaydet, dışa aktar, otomatik kurtarma ve otomatik kaydetme.
import { useCallback, useEffect, useRef, useState } from "react";
import type { Editor } from "@tiptap/core";
import { EditorState, TextSelection } from "@tiptap/pm/state";
import { addRecentFile, baseName, downloadBlob, getNative, openFile, readPath, saveFile, saveFileAs, type FileRef } from "@officehere/core";
import { getUserName, isDirectSave, saveStatus, trackSave, useAutoSave } from "@officehere/ui";
import { defaultSettings, normalizeDoc, type BelgeDocument, type DocSettings } from "../model/document";
import { DOCX_FILTER, OPEN_FILTERS, docToText, htmlDocument, openDocument, writeDocx } from "../io";
import { TEMPLATES } from "../templates/templates";
import { annotateNotesHtml } from "../editor/notes";
import { clearOwnDraft, writeDraft, type Draft } from "./drafts";

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
}

export function useFileOps({ editor, settings, setSettings, dirty, setDirty, notify, confirmDiscard, autosave, onLoaded }: Options) {
  const [fileName, setFileName] = useState("Belge1");
  const [fileRef, setFileRef] = useState<FileRef | null>(null);
  const counter = useRef(1);
  const settingsRef = useRef(settings);
  settingsRef.current = settings;

  // Yazar boşsa Seçenekler'deki kullanıcı adı yazılır (Word'deki gibi).
  const currentDoc = useCallback(
    (): BelgeDocument => ({ content: editor.getJSON(), settings: { ...settingsRef.current, properties: { ...settingsRef.current.properties, author: settingsRef.current.properties.author || getUserName(), modified: new Date().toISOString() } } }),
    [editor],
  );

  const load = useCallback(
    (doc: BelgeDocument, name: string, ref: FileRef | null) => {
      setSettings(doc.settings);
      loadIntoEditor(editor, doc.content);
      setFileName(name);
      setFileRef(ref);
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
        const { doc, warning } = await openDocument(name, data);
        load(doc, baseName(name), ref);
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

  const docxName = (n: string) => (n.toLowerCase().endsWith(".docx") ? n : `${baseName(n)}.docx`);

  const save = useCallback(async (): Promise<boolean> => {
    try {
      const target = fileRef && /\.docx$/i.test(fileRef.name) ? fileRef : null;
      const ref = await trackSave(isDirectSave(target), async () => saveFile(target, docxName(fileName), DOCX_FILTER, await writeDocx(currentDoc())));
      if (!ref) return false;
      setFileRef(ref);
      setFileName(baseName(ref.name));
      setDirty(false);
      clearOwnDraft();
      await addRecentFile({ path: ref.path ?? ref.name, name: ref.name, app: "belge", openedAt: Date.now() });
      notify(`"${ref.name}" kaydedildi.`, "success");
      return true;
    } catch (err) {
      notify(`Kaydedilemedi: ${(err as Error).message}`, "error");
      return false;
    }
  }, [currentDoc, fileRef, fileName, setDirty, notify]);

  const saveAs = useCallback(async (): Promise<boolean> => {
    try {
      const ref = await trackSave(false, async () => saveFileAs(docxName(fileName), DOCX_FILTER, await writeDocx(currentDoc())));
      if (!ref) return false;
      setFileRef(ref);
      setFileName(baseName(ref.name));
      setDirty(false);
      clearOwnDraft();
      await addRecentFile({ path: ref.path ?? ref.name, name: ref.name, app: "belge", openedAt: Date.now() });
      return true;
    } catch (err) {
      notify(`Kaydedilemedi: ${(err as Error).message}`, "error");
      return false;
    }
  }, [currentDoc, fileName, setDirty, notify]);

  const exportAs = useCallback(
    async (kind: string) => {
      const doc = currentDoc();
      const base = baseName(fileName);
      if (kind === "docx") {
        const bytes = await writeDocx(doc);
        downloadBlob(new Blob([bytes as BlobPart], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }), `${base}.docx`);
      } else if (kind === "html") {
        const html = htmlDocument(annotateNotesHtml(editor.getHTML()), doc.settings, doc.settings.properties.title || base);
        await saveFileAs(`${base}.html`, [{ name: "Web Sayfası", extensions: ["html"] }], html);
      } else if (kind === "txt") {
        await saveFileAs(`${base}.txt`, [{ name: "Düz Metin", extensions: ["txt"] }], "﻿" + docToText(doc.content));
      }
    },
    [currentDoc, editor, fileName],
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
  const nameRef = useRef(fileName);
  nameRef.current = fileName;
  useEffect(() => {
    if (!autosave) return;
    const write = () => {
      if (!dirtyRef.current || editor.isDestroyed) return;
      const d: Draft = { content: editor.getJSON(), settings: settingsRef.current, fileName: nameRef.current, savedAt: Date.now() };
      writeDraft(d);
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

  // Otomatik kaydetme (masaüstü, .docx): yazma durunca aynı dosyaya kaydedilir; taslak da silinir.
  const autoSave = useAutoSave({
    path: fileRef?.path,
    formats: ["docx"],
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

  return { isPristine, fileName, setFileName, fileRef, newDocument, openDialog, openPath, openData, save, saveAs, exportAs, closeDocument, load, currentDoc };
}
