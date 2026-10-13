/** Dosya işlemleri: Yeni, Aç, Kaydet, Farklı Kaydet, son kullanılanlar ve dosya biçimleri. */
import { addRecentFile, baseName, extensionOf, getNative, openFile, readPath, saveFile, saveFileAs, type FileFilter, type FileRef, type RecentFile } from "@officehere/core";
import { isDirectSave, saveStatus, trackSave } from "@officehere/ui";
import type { Presentation } from "../model/types";
import { FORMATS, isSaveFormat, isSunuFormat, OPEN_FILTERS, readPresentationFile, saveFilters, writePresentationFormat, type SaveFormat, type SunuFormat } from "../io/formats";
import type { Actions } from "../state/actions";
import type { AppState } from "../state/store";
import { abortMasterView, documentPres } from "../state/masterActions";
import { TEMPLATES } from "../templates/templates";
import { askKeepFormat } from "./formatAsk";

/** Aç iletişim kutusu süzgeçleri (tüm sunu biçimleri). */
export const PPTX_FILTERS: FileFilter[] = OPEN_FILTERS;

let untitled = 1;

export function nextUntitled(): string {
  untitled += 1;
  return `Sunu${untitled}`;
}

export function loadPresentation(a: Actions, pres: Presentation, fileName: string, fileRef: FileRef | null, format: string = "pptx") {
  abortMasterView(a);
  a.store.set({
    pres,
    slideIdx: 0,
    slideSel: pres.slides[0] ? [pres.slides[0].id] : [],
    sel: [],
    editing: null,
    cellSel: null,
    fileName,
    fileRef,
    dirty: false,
    view: "normal",
    pane: null,
    insert: null,
    cropId: null,
    ink: null,
    fit: true,
    message: null,
    ribbonTab: "home",
    collapsedSections: [],
    masterBase: null,
    fileFormat: format,
    formatAsked: false,
  });
  a.store.resetHistory();
  saveStatus.reset();
}

export function isPristine(a: Actions): boolean {
  const s = a.s;
  return !s.dirty && !s.fileRef && s.pres.slides.length <= 1;
}

/** Belgenin şu anki dosya biçimi: diskteki yolun uzantısı, yoksa açılırken tanınan biçim. */
export function currentFormat(s: Pick<AppState, "fileRef" | "fileFormat">): string {
  const p = s.fileRef?.path ?? (s.fileRef?.handle ? s.fileRef.name : undefined);
  return (p ? extensionOf(p) : "") || s.fileFormat || "pptx";
}

/** Açılış iletisi (eski biçim, makrolar, gösteri dosyası). */
function openNotice(a: Actions, format: SunuFormat) {
  const info = FORMATS[format];
  if (info.legacy) a.message("Bu sunu 97-2003 biçiminde (uyumluluk modu). Bazı öğeler gösterilemeyebilir; kaydedince .pptx olur.", "warning");
  else if (info.macro) a.message("Bu dosyadaki makrolar yok sayıldı. Kaydedince makrosuz .pptx olur.", "info", info.show ? "startShow" : undefined);
  else if (info.show) a.message("Bu dosya bir slayt gösterisi. Düzenleyebilir ya da gösteriyi başlatabilirsiniz.", "info", "startShow");
}

export async function openBytes(a: Actions, data: ArrayBuffer, name: string, ref: FileRef | null): Promise<boolean> {
  try {
    const { pres, format } = await readPresentationFile(data, name);
    if (FORMATS[format].template) {
      // Şablon: adsız yeni sunu olarak açılır (özgün dosyanın üzerine yazılmaz).
      loadPresentation(a, pres, nextUntitled(), null, "pptx");
      a.message(`“${name}” şablonundan yeni sunu oluşturuldu.`, "info");
    } else {
      loadPresentation(a, pres, baseName(name), ref, format);
      openNotice(a, format);
    }
    void addRecentFile({ path: ref?.path ?? name, name, app: "sunum", openedAt: Date.now() });
    return true;
  } catch (e) {
    a.message(`Dosya açılamadı: ${(e as Error).message}`, "error");
    return false;
  }
}

export async function openDialog(a: Actions): Promise<void> {
  const f = await openFile(OPEN_FILTERS);
  if (!f) return;
  await openBytes(a, f.data, f.name, f.handle || f.path ? { name: f.name, path: f.path, handle: f.handle } : null);
}

export async function openFromPath(a: Actions, path: string): Promise<void> {
  const native = getNative();
  if (!isPristine(a) && native) {
    await native.newWindow("sunum", path);
    return;
  }
  const f = await readPath(path);
  if (f) await openBytes(a, f.data, f.name, { name: f.name, path: f.path });
}

export async function openRecent(a: Actions, file: RecentFile): Promise<void> {
  if (getNative()) {
    await openFromPath(a, file.path);
    return;
  }
  a.message("Tarayıcıda son kullanılan dosyalar yeniden açılamaz; lütfen Gözat ile seçin.", "info");
  await openDialog(a);
}

/** Sunuyu verilen biçimde baytlara yazar (asıl görünümdeyken de belgenin kendisi). */
export async function presentationBytes(a: Actions, format: SaveFormat = "pptx"): Promise<Uint8Array> {
  const doc = documentPres(a);
  return writePresentationFormat({ ...doc, props: { ...doc.props, title: doc.props.title || a.s.fileName } }, format);
}

/**
 * Farklı Kaydet: biçim listesi sunulur; kullanıcı süzgeçten başka bir biçim
 * seçtiyse dosya o biçimin yazıcısıyla yeniden yazılır.
 */
async function saveAsFormat(a: Actions, name: string, format: SaveFormat, bytes: Uint8Array): Promise<FileRef | null> {
  const ref = await saveFileAs(name, saveFilters(format), bytes);
  if (!ref) return null;
  const ext = extensionOf(ref.path ?? ref.name);
  if (isSaveFormat(ext) && ext !== format) {
    const other = await presentationBytes(a, ext);
    const buf = other.buffer.slice(other.byteOffset, other.byteOffset + other.byteLength) as ArrayBuffer;
    const native = getNative();
    if (ref.path && native) {
      if (!(await native.writeFile(ref.path, buf))) throw new Error("Dosya yazılamadı.");
    } else if (ref.handle) {
      const w = await ref.handle.createWritable();
      await w.write(buf);
      await w.close();
    }
  }
  return ref;
}

/**
 * Kaydet / Farklı Kaydet. `as` verilirse o biçimde farklı kaydedilir.
 * Yazamadığımız biçimler (.ppt, makrolu) Farklı Kaydet → .pptx olur; .odp
 * belgesinde ilk kayıtta bir kez biçim sorusu sorulur.
 */
export async function save(a: Actions, forceAs = false, as?: SaveFormat): Promise<boolean> {
  const s = a.s;
  a.stopEditing();
  const cur = currentFormat(s);
  const writable = isSunuFormat(cur) ? FORMATS[cur].writable : isSaveFormat(cur) ? cur : undefined;
  let format: SaveFormat = as ?? writable ?? "pptx";
  let inPlace = !forceAs && !as && !!writable;
  let asked = !!s.formatAsked;
  if (inPlace && format === "odp" && !asked) {
    const choice = await askKeepFormat("odp");
    if (choice === "cancel") return false;
    asked = true;
    if (choice === "native") {
      format = "pptx";
      inPlace = false;
    }
  }
  try {
    const name = `${s.fileName}.${format}`;
    const target = inPlace ? s.fileRef : null;
    const ref = await trackSave(isDirectSave(target), async () => {
      const bytes = await presentationBytes(a, format);
      return target ? saveFile(target, name, saveFilters(format), bytes) : saveAsFormat(a, name, format, bytes);
    });
    if (!ref) return false;
    const ext = extensionOf(ref.path ?? ref.name);
    a.store.set({ fileRef: ref.path || ref.handle ? ref : null, fileName: baseName(ref.name), dirty: false, fileFormat: isSaveFormat(ext) ? ext : format, formatAsked: asked });
    void addRecentFile({ path: ref.path ?? ref.name, name: ref.name, app: "sunum", openedAt: Date.now() });
    a.message("Sunu kaydedildi.", "success");
    window.setTimeout(() => {
      if (a.s.message?.kind === "success") a.store.set({ message: null });
    }, 2500);
    return true;
  } catch (e) {
    a.message(`Kaydedilemedi: ${(e as Error).message}`, "error");
    return false;
  }
}

export function newFromTemplate(a: Actions, id: string) {
  const t = TEMPLATES.find((x) => x.id === id) ?? TEMPLATES[0];
  const pres = t.build();
  loadPresentation(a, pres, nextUntitled(), null);
}
