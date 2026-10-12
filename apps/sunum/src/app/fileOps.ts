/** Dosya işlemleri: Yeni, Aç, Kaydet, Farklı Kaydet, son kullanılanlar. */
import { addRecentFile, baseName, getNative, openFile, readPath, saveFile, saveFileAs, type FileFilter, type FileRef, type RecentFile } from "@officehere/core";
import type { Presentation } from "../model/types";
import { readPptx } from "../io/pptxReader";
import { writePptx } from "../io/pptxWriter";
import type { Actions } from "../state/actions";
import { abortMasterView, documentPres } from "../state/masterActions";
import { TEMPLATES } from "../templates/templates";

export const PPTX_FILTERS: FileFilter[] = [{ name: "Sunu", extensions: ["pptx"] }];

let untitled = 1;

export function nextUntitled(): string {
  untitled += 1;
  return `Sunu${untitled}`;
}

export function loadPresentation(a: Actions, pres: Presentation, fileName: string, fileRef: FileRef | null) {
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
  });
  a.store.resetHistory();
}

export function isPristine(a: Actions): boolean {
  const s = a.s;
  return !s.dirty && !s.fileRef && s.pres.slides.length <= 1;
}

export async function openBytes(a: Actions, data: ArrayBuffer, name: string, ref: FileRef | null): Promise<boolean> {
  try {
    const pres = await readPptx(data);
    loadPresentation(a, pres, baseName(name), ref);
    void addRecentFile({ path: ref?.path ?? name, name, app: "sunum", openedAt: Date.now() });
    return true;
  } catch (e) {
    a.message(`Dosya açılamadı: ${(e as Error).message}`, "error");
    return false;
  }
}

export async function openDialog(a: Actions): Promise<void> {
  const f = await openFile(PPTX_FILTERS);
  if (!f) return;
  if (!f.name.toLowerCase().endsWith(".pptx")) {
    a.message("Yalnızca .pptx sunuları açılabilir.", "warning");
    return;
  }
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

export async function save(a: Actions, forceAs = false): Promise<boolean> {
  const s = a.s;
  a.stopEditing();
  try {
    const doc = documentPres(a);
    const bytes = await writePptx({ ...doc, props: { ...doc.props, title: doc.props.title || s.fileName } });
    const name = `${s.fileName}.pptx`;
    const ref = forceAs || !s.fileRef ? await saveFileAs(name, PPTX_FILTERS, bytes) : await saveFile(s.fileRef, name, PPTX_FILTERS, bytes);
    if (!ref) return false;
    a.store.set({ fileRef: ref.path || ref.handle ? ref : null, fileName: baseName(ref.name), dirty: false });
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
