// Dosya işlemleri: yeni, aç, kaydet, farklı kaydet, dışa aktar, otomatik kurtarma ve otomatik kaydetme.
// Ayrıştırma ve yazma Worker'da (io/client.ts); büyük belgede ilk ekran önce yüklenir, kalanı parça parça eklenir.
import { useCallback, useEffect, useRef, useState } from "react";
import type { Editor, JSONContent } from "@tiptap/core";
import { Fragment } from "@tiptap/pm/model";
import { EditorState, TextSelection, type Transaction } from "@tiptap/pm/state";
import { addRecentFile, baseName, extensionOf, getNative, getSource, isAbortError, openFile, openSource, readPath, releaseSource, saveFile, saveFileAs, type FileFilter, type FileRef, type SaveData } from "@officehere/core";
import { getUserName, isDirectSave, saveStatus, trackSave, useAutoSave } from "@officehere/ui";
import { defaultSettings, normalizeDoc, type BelgeDocument, type DocSettings } from "../model/document";
import { FORMATS, OPEN_FILTERS, docToText, finishOpen, formatOfExt, htmlDocument, isWritable, openDocument, saveFilters, type FormatId, type OpenResult } from "../io";
import { encodeInWorker, readInWorker, releaseInWorker, verifyInWorker, writeDocxInWorker, writeOdtInWorker } from "../io/client";
import { LoadWarnings, ODF_OBJECTS_FALLBACK, OPAQUE_PLACEHOLDERS, lossMessage } from "../io/warnings";
import { hasOdfObjects, hasOpaque } from "../editor/opaque";
import { protectionSource } from "../editor/protection";
import type { PaginationController } from "../editor/pagination";
import { TEMPLATES } from "../templates/templates";
import { annotateNotesHtml } from "../editor/notes";
import { clearOwnDraft, ensureSource, writeDraft, type Draft } from "./drafts";

/** Açılışta okunamayan içerik varsa otomatik kaydetme ilk elle kayda kadar bekler (rozet nedeni gösterir). */
export const LOAD_WARNING_HOLD = "Bu dosyada desteklenmeyen öğeler var; ilk kaydı siz onaylayın";

/** Dijital imzalı belge: kaydetmek imzayı kaldırır; otomatik kaydetme ilk elle kayda kadar bekler. */
export const SIGNATURE_HOLD = "Dijital imzalı belge; kaydederseniz imza kaldırılır. İlk kaydı siz onaylayın";

/** Makrolu belge (.docm/.dotm): makrolar korunur, çalıştırılmaz. */
const isMacroName = (name: string) => /\.(docm|dotm)$/i.test(name);
const DOCM_FILTER: FileFilter[] = [{ name: "Makro İçerebilen Belge (.docm)", extensions: ["docm"] }];

/**
 * Otomatik kaydedilen .docx kendi okuyucumuzla yeniden açılabiliyor mu: okuyucu hata vermemeli ve
 * paketin bir kısmını bozuk/eksik saymamalı (desteklenmeyen öğe uyarıları doğrulamayı bozmaz).
 * Ayrı bir Worker'da çalışır (açmanın iptali doğrulamayı düşürmez; ana iş parçacığı donmaz).
 */
export function canReopenDocx(data: ArrayBuffer): Promise<boolean> {
  return verifyInWorker(data);
}
import type { FormatChoice } from "./useFormatPrompt";

export type { Draft } from "./drafts";

/**
 * Koruma (salt okunur) yüklemeyi süzmesin: yüklenen içerik düzenleme değildir; önceki belgenin koruması da yeni
 * belgenin açılmasını engellememeli.
 */
function unprotected(fn: () => void) {
  const get = protectionSource.get;
  protectionSource.get = () => null;
  try {
    fn();
  } finally {
    protectionSource.get = get;
  }
}

/** Geçmişi sıfırlayarak belge içeriğini yükler (önceki belgede değişiklik izleme açıksa yeni içerik izlenmez). */
export function loadIntoEditor(editor: Editor, doc: BelgeDocument["content"]) {
  unprotected(() => editor.chain().setMeta("belgeNoTrack", true).setContent(normalizeDoc(doc), { emitUpdate: false }).run());
  const state = EditorState.create({ doc: editor.state.doc, plugins: editor.state.plugins, selection: TextSelection.atStart(editor.state.doc) });
  editor.view.updateState(state);
}

/* ------------------------------------------------------------------ */
/* Büyük belge: ilk ekran önce, kalanı parça parça                      */
/* ------------------------------------------------------------------ */

/** Bölümün son bloğu olabilenler (editor/sections.ts): parça sınırı yalnız bunlardan sonra (araya boş paragraf girmesin). */
const SECTION_END_OK = new Set(["paragraph", "heading", "bulletList", "orderedList"]);
/** Bu boyuttan (yaklaşık karakter) büyük belge parça parça yüklenir (~25 sayfa). */
export const PROGRESSIVE_MIN = 80_000;
/** İlk parça: ilk ekran ve birkaç sayfa. */
const FIRST_PART = 24_000;
/** Parça başına hedef en uzun görev (ms); uzun görev gözlemcisi yoksa parça başına geçen süre hedefi. */
const PART_TASK_MS = 140;
const PART_WALL_MS = 200;
/** Parça, yüklenmiş boyutun bu oranından küçük olmaz. */
const PART_MIN_SHARE = 0.04;
/** Yükleme işlemlerinin işareti. */
const LOADING_META = "belgeYukleme";
/** Bundan kısa süren açmada ilerleme göstergesi çıkmaz. */
const PROGRESS_DELAY_MS = 300;

export interface ProgressiveLoad {
  /** Kalan parçaları hemen ekler (kaydetme, klavye girdisi, şerit tıklaması). */
  flush(): void;
  /** Eklemeyi durdurur (eklenenler kalır). */
  stop(): void;
  /** Tamamlanınca true, durdurulunca false. */
  readonly done: Promise<boolean>;
}

/** Bloğun yaklaşık boyutu (metin ve düğüm sayısı). */
function sizeOf(n: JSONContent): number {
  let s = 2 + (n.text?.length ?? 0);
  if (n.content) for (const c of n.content) s += sizeOf(c);
  return s;
}

/** Yükleme parçası: geri alınamaz, kirli saymaz, izlenmez, korumaya takılmaz. */
function dispatchLoading(editor: Editor, tr: Transaction) {
  unprotected(() => editor.view.dispatch(tr.setMeta("addToHistory", false).setMeta("preventUpdate", true).setMeta("belgeNoTrack", true).setMeta(LOADING_META, true)));
}

/** Bir sonraki çizimden sonra (sekme arka plandaysa zamanlayıcıyla) çalıştırır. */
function afterFrame(fn: () => void): () => void {
  let done = false;
  const run = () => {
    if (done) return;
    done = true;
    fn();
  };
  const t = window.setTimeout(run, 100);
  const raf = requestAnimationFrame(() => window.setTimeout(run, 0));
  return () => {
    done = true;
    clearTimeout(t);
    cancelAnimationFrame(raf);
  };
}

/**
 * Büyük belgeyi parça parça yükler: ilk parça (ilk ekran) hemen ve geçmiş sıfırlanarak, kalanı her çizimden sonra
 * belgenin sonuna eklenir (geri alınamaz, kirli saymaz, izlenmez). Parça boyutu, bir önceki parçanın süresine göre
 * ayarlanır; artımlı sayfalama yalnız eklenen sayfaları ölçer. Kullanıcı yazarsa ya da arayüzde bir şeye basarsa
 * kalan hemen eklenir. Küçük belgede null (çağıran tek seferde yükler).
 */
export function loadProgressive(editor: Editor, json: JSONContent, onProgress?: (ratio: number) => void): ProgressiveLoad | null {
  const doc = normalizeDoc(json);
  const sections = doc.content ?? [];
  const units: { sec: number; block: JSONContent; size: number; cut: boolean }[] = [];
  let total = 0;
  sections.forEach((sec, si) => {
    const blocks = sec.content ?? [];
    blocks.forEach((block, bi) => {
      const size = sizeOf(block);
      total += size;
      units.push({ sec: si, block, size, cut: bi === blocks.length - 1 || SECTION_END_OK.has(block.type ?? "") });
    });
  });
  if (total < PROGRESSIVE_MIN) return null;

  /** `from`dan başlayıp `budget` boyutuna ulaşan ve bir parça sınırında biten aralığın sonu. */
  const take = (from: number, budget: number) => {
    let to = from;
    let size = 0;
    while (to < units.length && (size < budget || !units[to - 1]!.cut)) size += units[to++]!.size;
    return to;
  };
  /** [from, to) bloklarının bölümleri (bölüm öznitelikleriyle). */
  const sectionsOf = (from: number, to: number): JSONContent[] => {
    const out: JSONContent[] = [];
    for (let i = from; i < to; i++) {
      const u = units[i]!;
      if (i === from || units[i - 1]!.sec !== u.sec) out.push({ ...sections[u.sec], content: [] });
      out[out.length - 1]!.content!.push(u.block);
    }
    return out;
  };

  let next = take(0, FIRST_PART);
  let loaded = units.slice(0, next).reduce((s, u) => s + u.size, 0);
  loadIntoEditor(editor, { ...doc, content: sectionsOf(0, next) });
  onProgress?.(loaded / total);

  const append = (to: number) => {
    const schema = editor.schema;
    const tr = editor.state.tr;
    let i = next;
    // Son yüklenen bölümün kalan blokları o bölümün sonuna, sonraki bölümler belgenin sonuna.
    const open: JSONContent[] = [];
    while (i < to && units[i]!.sec === units[next - 1]!.sec) open.push(units[i++]!.block);
    if (open.length) tr.insert(tr.doc.content.size - 1, Fragment.fromJSON(schema, open));
    if (i < to) tr.insert(tr.doc.content.size, Fragment.fromJSON(schema, sectionsOf(i, to)));
    for (let k = next; k < to; k++) loaded += units[k]!.size;
    next = to;
    dispatchLoading(editor, tr);
  };

  let resolve!: (v: boolean) => void;
  const done = new Promise<boolean>((r) => (resolve = r));
  let cancelFrame: (() => void) | null = null;
  let finished = false;
  let budget = FIRST_PART * 2;

  // Sayfalama ölçümü (yavaşken gecikmeli çalışır) bir sonraki parçadan önce beklenir: her ölçüm yalnız bir parçayı
  // kapsar, maliyeti doğru parçaya yazılır.
  const pager = (editor.extensionManager.extensions.find((e) => e.name === "belgePagination")?.options as { controller?: PaginationController } | undefined)?.controller;
  const afterLayout = (fn: () => void): (() => void) => {
    const before = pager?.lastRun;
    const t0 = performance.now();
    let cancel = () => {};
    const poll = () => {
      cancel = !pager?.geometry.enabled || pager.lastRun !== before || performance.now() - t0 > 300 ? afterFrame(fn) : afterFrame(poll);
    };
    cancel = afterFrame(poll);
    return () => cancel();
  };
  // Parçanın maliyeti: sürdüğü en uzun görev (gözlemci yoksa geçen süre).
  let longest = 0;
  let observer: PerformanceObserver | null = null;
  try {
    if (PerformanceObserver.supportedEntryTypes?.includes("longtask")) {
      observer = new PerformanceObserver((l) => {
        for (const e of l.getEntries()) longest = Math.max(longest, e.duration);
      });
      observer.observe({ type: "longtask" });
    }
  } catch {
    observer = null;
  }

  const onKey = (e: KeyboardEvent) => {
    if (!["Control", "Shift", "Alt", "Meta", "AltGraph", "CapsLock"].includes(e.key)) flush();
  };
  // Belge alanında tıklamak (imleç, kaydırma, okuma) ve ilerleme çubuğu yüklemeyi bitirmez; şerit, menüler,
  // iletişim kutuları bitirir.
  const onPointer = (e: Event) => {
    const t = e.target as Element | null;
    if (!t?.closest?.(".bl-scroll, .ProseMirror, .oh-progress")) flush();
  };
  const onInput = () => flush();
  const onTransaction = ({ transaction }: { transaction: Transaction }) => {
    if (transaction.docChanged && !transaction.getMeta(LOADING_META) && transaction.getMeta("addToHistory") !== false) queueMicrotask(flush);
  };
  const dom = editor.view.dom;
  document.addEventListener("keydown", onKey, true);
  document.addEventListener("pointerdown", onPointer, true);
  for (const t of ["beforeinput", "paste", "drop"]) dom.addEventListener(t, onInput, true);
  editor.on("transaction", onTransaction);

  const finish = (complete: boolean) => {
    if (finished) return;
    finished = true;
    cancelFrame?.();
    observer?.disconnect();
    document.removeEventListener("keydown", onKey, true);
    document.removeEventListener("pointerdown", onPointer, true);
    for (const t of ["beforeinput", "paste", "drop"]) dom.removeEventListener(t, onInput, true);
    editor.off("transaction", onTransaction);
    resolve(complete);
  };

  const step = () => {
    cancelFrame = null;
    if (finished || editor.isDestroyed) return finish(false);
    const t0 = performance.now();
    observer?.takeRecords();
    longest = 0;
    append(take(next, budget));
    onProgress?.(loaded / total);
    if (next >= units.length) return finish(true);
    cancelFrame = afterLayout(() => {
      // Ekleme, çizim ve sayfalama maliyetine göre sonraki parça büyütülür/küçültülür (sıçramasın diye sınırlı).
      // Her parçanın belge boyutuyla büyüyen sabit bir maliyeti var (bölüm kutusunun yeniden yerleşimi, sayfa
      // sayısının yeniden çizimi): parça, yüklenenin belli bir oranından küçülmez; yoksa parça sayısı patlar.
      for (const e of observer?.takeRecords() ?? []) longest = Math.max(longest, e.duration);
      const k = observer ? PART_TASK_MS / Math.max(1, longest) : PART_WALL_MS / Math.max(1, performance.now() - t0);
      budget = Math.round(Math.min(FIRST_PART * 8, Math.max(FIRST_PART / 3, loaded * PART_MIN_SHARE, budget * Math.min(1.5, Math.max(0.5, k)))));
      step();
    });
  };

  function flush() {
    if (finished || editor.isDestroyed) return;
    if (next < units.length) append(units.length);
    onProgress?.(1);
    finish(true);
  }

  cancelFrame = afterFrame(step);
  return { flush, stop: () => finish(false), done };
}

/** Durum çubuğundaki dosya işlemi ilerlemesi (yalnız ilerleme göstergesi yeniden çizilir). */
export interface FileProgress {
  label: string;
  /** 0..1; null: belirsiz. */
  value: number | null;
  detail?: string;
  onCancel?: () => void;
}

export class ProgressStore {
  private value: FileProgress | null = null;
  private listeners = new Set<() => void>();
  readonly subscribe = (fn: () => void) => {
    this.listeners.add(fn);
    return () => void this.listeners.delete(fn);
  };
  readonly get = () => this.value;
  set(v: FileProgress | null) {
    this.value = v;
    for (const fn of this.listeners) fn();
  }
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
  const fileRefRef = useRef(fileRef);
  fileRefRef.current = fileRef;
  const formatRef = useRef({ format, fileExt, loadHold });
  formatRef.current = { format, fileExt, loadHold };
  /** Sürmekte olan parça parça yükleme (büyük belge). */
  const loading = useRef<ProgressiveLoad | null>(null);
  /** Sürmekte olan açma (yeni açma ya da İptal onu durdurur). */
  const opening = useRef<AbortController | null>(null);
  const [progress] = useState(() => new ProgressStore());

  // Yazar boşsa Seçenekler'deki kullanıcı adı yazılır (Word'deki gibi). Parça parça yükleme sürüyorsa önce tamamlanır.
  const currentDoc = useCallback((): BelgeDocument => {
    loading.current?.flush();
    return { content: editor.getJSON(), settings: { ...settingsRef.current, properties: { ...settingsRef.current.properties, author: settingsRef.current.properties.author || getUserName(), modified: new Date().toISOString() } } };
  }, [editor]);

  /**
   * .docx/.docm baytları Worker'da yazılır: açılan belgenin kaynak paketi oradadır, yazıcı modelin tanımadığı
   * parçaları ondan taşır; Worker'da olmayan kaynak (Worker yeniden kuruldu, kurtarılan taslak) bu iş
   * parçacığından ya da kurtarma deposundan gönderilir. Makrolu hedefte makrolar korunur (çalıştırılmaz).
   * `notes`: kaydetmede bildirilecekler (ör. makrolu belge .docx'e kaydedilince makroların kaldırılması).
   */
  const docxData = useCallback(async (doc: BelgeDocument, macro = false, notes?: string[], format: "docx" | "odt" = "docx"): Promise<Uint8Array> => {
    const provide = async (ids: string[]) => {
      const out: Record<string, Uint8Array | null> = {};
      for (const id of ids) {
        await ensureSource(id);
        out[id] = getSource(id)?.bytes ?? null;
      }
      return out;
    };
    // .odt de aynı yoldan: açılan .odt'nin ayarları ve gömülü nesneleri kaynak paketten taşınır.
    const r = format === "odt" ? await writeOdtInWorker(doc, provide) : await writeDocxInWorker(doc, macro, provide);
    notes?.push(...r.notes);
    return r.bytes;
  }, []);

  /** Belgeyi istenen biçimde kodlar (`macro`: .docx biçimi makrolu .docm olarak). */
  const encode = useCallback(
    async (f: FormatId, doc = currentDoc(), macro = false, notes?: string[]): Promise<SaveData> => {
      // .docx'e özgü (opak) öğeler bu biçimlerde yer tutucu metin olarak yazılır; .odt'den gelen gömülü nesneler
      // .odt'de özgün hâliyle, diğer metin biçimlerinde yedek görünümleriyle.
      if ((f === "odt" || f === "rtf" || f === "md" || f === "html") && hasOpaque(doc, true)) notes?.push(OPAQUE_PLACEHOLDERS);
      if ((f === "rtf" || f === "md" || f === "html") && hasOdfObjects(doc)) notes?.push(ODF_OBJECTS_FALLBACK);
      switch (f) {
        case "odt":
          return docxData(doc, false, notes, "odt");
        case "rtf":
        case "md":
          return encodeInWorker(f, doc);
        case "html":
          return htmlDocument(annotateNotesHtml(editor.getHTML()), doc.settings, doc.settings.properties.title || baseName(fileNameRef.current));
        case "txt":
          return "﻿" + docToText(doc.content);
        default:
          return docxData(doc, macro, notes);
      }
    },
    [currentDoc, docxData, editor],
  );

  /** Dosya okunurken düzenleme kapalı (sahibi, kapatan açma işlemi). */
  const editLock = useRef<{ owner: AbortController; focus: boolean } | null>(null);
  const lockEditing = useCallback(
    (owner: AbortController) => {
      if (!editLock.current) {
        editLock.current = { owner, focus: editor.view.hasFocus() };
        if (editor.isEditable) editor.setEditable(false, false);
      } else editLock.current.owner = owner;
    },
    [editor],
  );
  const unlockEditing = useCallback(
    (owner?: AbortController) => {
      const l = editLock.current;
      if (!l || (owner && l.owner !== owner)) return;
      editLock.current = null;
      if (editor.isDestroyed) return;
      editor.setEditable(true, false);
      if (l.focus) editor.view.focus();
    },
    [editor],
  );

  /** Sürmekte olan açmayı ve parça parça yüklemeyi bırakır (yerine başka belge yükleniyor). */
  const supersede = useCallback(() => {
    const c = opening.current;
    opening.current = null;
    c?.abort();
    unlockEditing();
    loading.current?.stop();
    loading.current = null;
    progress.set(null);
  }, [progress, unlockEditing]);

  /** Yüklenen belgenin dosya, biçim ve kaydetme durumu. */
  const fileState = useCallback(
    (name: string, ref: FileRef | null, fmt: FormatId, ext: string) => {
      setFileName(name);
      setFileRef(ref);
      setFormat(fmt);
      setFileExt(ext);
      keepAsked.current = false;
      setLoadHold(null);
      setDirty(false);
      saveStatus.reset();
    },
    [setDirty],
  );

  /** Belgenin özgün paketi bellekten ve Worker'dan bırakılır (kurtarma deposundaki kopyası taslağıyla yaşar). */
  const release = useCallback((id: string | null | undefined, keep: string | null | undefined) => {
    if (!id || id === keep) return;
    releaseSource(id);
    releaseInWorker(id);
  }, []);

  const load = useCallback(
    (doc: BelgeDocument, name: string, ref: FileRef | null, fmt: FormatId = "docx", ext = "docx") => {
      supersede();
      release(settingsRef.current.kaynak, doc.settings.kaynak);
      setSettings(doc.settings);
      settingsRef.current = doc.settings;
      loadIntoEditor(editor, doc.content);
      fileState(name, ref, fmt, ext);
      onLoaded();
    },
    [editor, setSettings, onLoaded, supersede, release, fileState],
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

  /**
   * Dosyayı açar: ayrıştırma Worker'da (HTML ve Markdown ana iş parçacığında), büyük belgede ilk ekran önce.
   * Durum çubuğunda ilerleme ve İptal: okuma sürerken iptal açık belgeye dokunmaz; sayfalar yerleştirilirken
   * iptal önceki belgeyi geri getirir.
   */
  const openData = useCallback(
    async (name: string, data: ArrayBuffer, ref: FileRef | null) => {
      // Önceki belge hâlâ parça parça yükleniyorsa tamamlanır: bu açma iptal edilirse eksiksiz geri gelmeli.
      loading.current?.flush();
      supersede();
      const ctrl = new AbortController();
      opening.current = ctrl;
      const { signal } = ctrl;
      let shown: FileProgress = { label: "Açılıyor…", value: null, detail: "Dosya okunuyor", onCancel: () => ctrl.abort() };
      let visible = false;
      const show = (p: Partial<FileProgress>) => {
        shown = { ...shown, ...p };
        if (visible && opening.current === ctrl) progress.set(shown);
      };
      // Kısa sürenler için gösterge çıkmaz.
      const timer = window.setTimeout(() => {
        visible = true;
        show({});
      }, PROGRESS_DELAY_MS);
      /** Bu açmanın kaydettiği kaynak paket (iptalde bırakılır). */
      let registered: string | null = null;
      // Dosya okunurken görünen (yerini yeni belgeye bırakacak) belgeye yazılanlar kaybolmasın: düzenleme kapalı.
      lockEditing(ctrl);
      const unlock = () => unlockEditing(ctrl);
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
        const reply = await readInWorker(name, data, { signal, onProgress: ({ step }) => show(step ? { detail: step } : {}) });
        let opened: OpenResult;
        if (reply.main) opened = await openDocument(name, data, { resolveImage });
        else {
          // Ana iş parçacığı da özgün paketi tutar: taslakla kurtarma deposuna yazılır, Worker yeniden kurulursa oradan gönderilir.
          const id = reply.doc.settings.kaynak;
          if (id && !getSource(id)) {
            await openSource(data, { id, name }).catch(() => undefined);
            registered = id;
          }
          opened = await finishOpen({ doc: reply.doc, format: reply.format, notes: reply.notes, warnings: LoadWarnings.fromPlain(reply.warnings) }, { resolveImage });
        }
        if (signal.aborted) throw signal.reason;
        const { doc, warning, format: fmt } = opened;
        const ext = extensionOf(name) || FORMATS[fmt].extensions[0];
        show({ detail: "Sayfalar yerleştiriliyor", value: 0 });

        // Büyük belge parça parça yüklenir; iptal edilirse önceki belge geri gelir.
        const prev = { state: editor.state, settings: settingsRef.current, fileName: fileNameRef.current, fileRef: fileRefRef.current, ...formatRef.current, keepAsked: keepAsked.current, dirty: dirtyRef.current };
        unlock();
        setSettings(doc.settings);
        settingsRef.current = doc.settings;
        const job = loadProgressive(editor, doc.content, (value) => show({ value }));
        if (!job) loadIntoEditor(editor, doc.content);
        fileState(baseName(name), ref, fmt, ext);
        onLoaded();
        if (job) {
          loading.current = job;
          const stop = () => job.stop();
          signal.addEventListener("abort", stop);
          const complete = await job.done;
          signal.removeEventListener("abort", stop);
          if (loading.current === job) loading.current = null;
          if (!complete) {
            // Başka bir belge yükleniyor: o devralır. İptal: önceki belge geri gelir.
            if (opening.current !== ctrl || editor.isDestroyed) return;
            editor.view.updateState(prev.state);
            setSettings(prev.settings);
            settingsRef.current = prev.settings;
            setFileName(prev.fileName);
            setFileRef(prev.fileRef);
            setFormat(prev.format);
            setFileExt(prev.fileExt);
            keepAsked.current = prev.keepAsked;
            setLoadHold(prev.loadHold);
            setDirty(prev.dirty);
            onLoaded();
            throw signal.reason;
          }
          // Sayfalar tamam: gezinti, açıklamalar ve satır numaraları yenilenir.
          onLoaded();
        }
        registered = null;
        release(prev.settings.kaynak, doc.settings.kaynak);
        // Okunamayan (kaydedilince kaybolacak) içerik ya da dijital imza: özgün dosyanın üzerine sessizce yazılmaz.
        const lw = opened.warnings ?? [];
        if (lw.length > 0) setLoadHold(lw.includes(lossMessage("signature")) ? SIGNATURE_HOLD : LOAD_WARNING_HOLD);
        // Önceki (vazgeçilen) belgenin taslağı bu pencereden kaldırılır.
        clearOwnDraft();
        if (warning) notify(warning, "warning");
        await addRecentFile({ path: ref?.path ?? name, name, app: "belge", openedAt: Date.now() });
      } catch (err) {
        if (registered) release(registered, settingsRef.current.kaynak);
        if (!isAbortError(err)) notify(`"${name}" açılamadı: ${(err as Error).message}`, "error");
      } finally {
        unlock();
        window.clearTimeout(timer);
        if (opening.current === ctrl) {
          opening.current = null;
          progress.set(null);
        }
      }
    },
    [editor, notify, setSettings, onLoaded, supersede, release, fileState, progress, lockEditing, unlockEditing],
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
        // Makrolu belge önce .docm olarak önerilir (makrolar korunur); .docx seçilirse makrolar kaldırılır.
        const macro = first === "docx" && isMacroName(`.${fileExt}`);
        const ext = macro ? "docm" : (FORMATS[first].saveExt ?? "docx");
        const filters = macro ? [...DOCM_FILTER, ...saveFilters(first)] : saveFilters(first);
        let notes: string[] = [];
        const ref = await trackSave(false, async () => saveFileAs(`${baseName(fileNameRef.current)}.${ext}`, filters, await encode(first, undefined, macro, notes)));
        if (!ref) return false;
        // İletişim kutusunda başka bir biçim seçildiyse dosya o biçimde yeniden yazılır.
        const chosen = formatOfExt(extensionOf(ref.name));
        let fmt: FormatId = first;
        const rewrite = async (f: FormatId, data: SaveData) => {
          const native = getNative();
          if (ref.path && native) await saveFile(ref, ref.name, saveFilters(f), data);
          else if (ref.handle) await saveFile(ref, ref.name, saveFilters(f), data);
        };
        if (chosen && chosen !== first && isWritable(chosen)) {
          await rewrite(chosen, await encode(chosen));
          fmt = chosen;
          notes = [];
        } else if (chosen === "docx" && isMacroName(ref.name) !== macro) {
          notes = [];
          await rewrite(chosen, await encode(chosen, undefined, !macro, notes));
        }
        keepAsked.current = true;
        await saved(ref, fmt);
        if (notes.length) notify(notes.join(" "), "warning");
        return true;
      } catch (err) {
        notify(`Kaydedilemedi: ${(err as Error).message}`, "error");
        return false;
      }
    },
    [encode, fileExt, notify, saved],
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
        const notes: string[] = [];
        const ref = await trackSave(isDirectSave(target), async () => saveFile(target, `${fileNameRef.current}.${fileExt}`, saveFilters(format), await encode(format, undefined, false, notes)));
        if (!ref) return false;
        await saved(ref, format);
        if (notes.length) notify(notes.join(" "), "warning");
        else notify(`"${ref.name}" kaydedildi.`, "success");
        return true;
      } catch (err) {
        notify(`Kaydedilemedi: ${(err as Error).message}`, "error");
        return false;
      }
    }
    // Yeni belge (dosyası yok): biçim listesiyle Farklı Kaydet.
    if (!fileRef) return saveAsFormat("docx");
    try {
      // .docm .docm olarak kaydedilir: makrolar korunur (çalıştırılmaz); .dotm'den yeni .docm.
      const macro = isMacroName(fileRef.name);
      const target = /\.(docx|docm)$/i.test(fileRef.name) ? fileRef : null;
      const name = `${baseName(fileNameRef.current)}.${macro ? "docm" : "docx"}`;
      const filters = macro ? DOCM_FILTER : target ? saveFilters("docx").slice(0, 1) : saveFilters("docx");
      const ref = await trackSave(isDirectSave(target), async () => saveFile(target, name, filters, await docxData(currentDoc(), macro)));
      if (!ref) return false;
      const chosen = formatOfExt(extensionOf(ref.name));
      if (chosen && chosen !== "docx" && isWritable(chosen) && (ref.path || ref.handle)) await saveFile(ref, ref.name, saveFilters(chosen), await encode(chosen));
      else if (chosen === "docx" && isMacroName(ref.name) !== macro && (ref.path || ref.handle)) await saveFile(ref, ref.name, saveFilters("docx"), await docxData(currentDoc(), !macro));
      keepAsked.current = chosen !== "docx";
      await saved(ref, chosen && isWritable(chosen) ? chosen : "docx");
      notify(`"${ref.name}" kaydedildi.`, "success");
      return true;
    } catch (err) {
      notify(`Kaydedilemedi: ${(err as Error).message}`, "error");
      return false;
    }
  }, [askFormat, currentDoc, docxData, encode, fileExt, fileRef, format, notify, saveAsFormat, saved]);

  const exportAs = useCallback(
    async (kind: string) => {
      const base = baseName(fileName);
      const fmt = kind as FormatId;
      if (!FORMATS[fmt]?.saveExt) return;
      const ext = FORMATS[fmt].saveExt!;
      const notes: string[] = [];
      await saveFileAs(`${base}.${ext}`, [{ name: FORMATS[fmt].label, extensions: [ext] }], await encode(fmt, undefined, false, notes));
      if (notes.length) notify(notes.join(" "), "warning");
    },
    [encode, fileName, notify],
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
      loading.current?.flush();
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
      return { data: await docxData(currentDoc()), isCurrent: () => !editor.isDestroyed && editor.state.doc === doc && settingsRef.current === s };
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

  return { isPristine, fileName, setFileName, fileRef, format, title, newDocument, openDialog, openPath, openData, save, saveAs, exportAs, closeDocument, load, currentDoc, progress };
}
