import { getNative, type AppKind, type FileFilter, type NativeSaveResult, type NativeWriteFailure, type OfficehereNative, type RecentFile } from "./bridge";
import { SaveError, saveErrorFrom } from "./saveErrors";
import { orderRecent, withRecentAdded, withRecentPinned, withRecentRemoved } from "./recent";

/** Açılmış ya da kaydedilmiş bir dosyanın tanıtıcısı. */
export interface FileRef {
  name: string;
  /** Masaüstünde tam yol. */
  path?: string;
  /** Tarayıcıda File System Access API tanıtıcısı (destekleniyorsa). */
  handle?: FileSystemFileHandleLike;
}

export interface OpenedFile extends FileRef {
  data: ArrayBuffer;
}

interface FileSystemWritableLike {
  write(data: BufferSource | Blob): Promise<void>;
  close(): Promise<void>;
}

export interface FileSystemFileHandleLike {
  name: string;
  getFile(): Promise<File>;
  createWritable(): Promise<FileSystemWritableLike>;
}

interface PickerWindow {
  showOpenFilePicker?: (options: unknown) => Promise<FileSystemFileHandleLike[]>;
  showSaveFilePicker?: (options: unknown) => Promise<FileSystemFileHandleLike>;
}

const MIME: Record<string, string> = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  csv: "text/csv",
  txt: "text/plain",
  html: "text/html",
  json: "application/json",
  pdf: "application/pdf",
  odt: "application/vnd.oasis.opendocument.text",
  ods: "application/vnd.oasis.opendocument.spreadsheet",
  odp: "application/vnd.oasis.opendocument.presentation",
  ott: "application/vnd.oasis.opendocument.text-template",
  ots: "application/vnd.oasis.opendocument.spreadsheet-template",
  otp: "application/vnd.oasis.opendocument.presentation-template",
  rtf: "application/rtf",
  md: "text/markdown",
  markdown: "text/markdown",
  tsv: "text/tab-separated-values",
  htm: "text/html",
  doc: "application/msword",
  xls: "application/vnd.ms-excel",
  ppt: "application/vnd.ms-powerpoint",
  ppsx: "application/vnd.openxmlformats-officedocument.presentationml.slideshow",
  potx: "application/vnd.openxmlformats-officedocument.presentationml.template",
  dotx: "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
  xltx: "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
};

export function mimeFor(name: string): string {
  const ext = extensionOf(name);
  return MIME[ext] ?? "application/octet-stream";
}

export function extensionOf(name: string): string {
  const i = name.lastIndexOf(".");
  return i < 0 ? "" : name.slice(i + 1).toLowerCase();
}

export function baseName(name: string): string {
  const slash = Math.max(name.lastIndexOf("/"), name.lastIndexOf("\\"));
  const file = name.slice(slash + 1);
  const dot = file.lastIndexOf(".");
  return dot <= 0 ? file : file.slice(0, dot);
}

function pickerTypes(filters: FileFilter[]) {
  return filters.map((f) => ({
    description: f.name,
    accept: { [mimeFor(`x.${f.extensions[0]}`)]: f.extensions.map((e) => `.${e}`) },
  }));
}

/** Kullanıcıya dosya seçtirir ve içeriğini okur. İptalde null döner. */
export async function openFile(filters: FileFilter[]): Promise<OpenedFile | null> {
  const native = getNative();
  if (native) {
    const r = await native.openFile(filters);
    return r ? { name: r.name, path: r.path, data: r.data } : null;
  }
  const w = window as unknown as PickerWindow;
  if (w.showOpenFilePicker) {
    try {
      const [handle] = await w.showOpenFilePicker({ types: pickerTypes(filters), multiple: false });
      const file = await handle.getFile();
      return { name: file.name, handle, data: await file.arrayBuffer() };
    } catch (err) {
      if ((err as DOMException)?.name === "AbortError") return null;
      // Bazı tarayıcılar API'yi gösterip izin vermez; klasik yönteme düş.
    }
  }
  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = filters.flatMap((f) => f.extensions.map((e) => `.${e}`)).join(",");
    input.style.display = "none";
    input.addEventListener("change", async () => {
      const file = input.files?.[0];
      input.remove();
      resolve(file ? { name: file.name, data: await file.arrayBuffer() } : null);
    });
    input.addEventListener("cancel", () => {
      input.remove();
      resolve(null);
    });
    document.body.appendChild(input);
    input.click();
  });
}

/** Masaüstünde verilen yoldaki dosyayı okur. */
export async function readPath(path: string): Promise<OpenedFile | null> {
  const native = getNative();
  if (!native) return null;
  const r = await native.readFile(path);
  return r ? { name: r.name, path: r.path, data: r.data } : null;
}

function toArrayBuffer(data: ArrayBuffer | Uint8Array | Blob | string): Promise<ArrayBuffer> {
  if (data instanceof ArrayBuffer) return Promise.resolve(data);
  if (data instanceof Uint8Array) {
    return Promise.resolve(data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer);
  }
  if (typeof data === "string") return Promise.resolve(new TextEncoder().encode(data).buffer as ArrayBuffer);
  return data.arrayBuffer();
}

export type SaveData = ArrayBuffer | Uint8Array | Blob | string;

/**
 * Masaüstünde dosyayı ana sürecin güvenli yazma yoluyla yazar. Yazılamazsa
 * hata kodu ve Türkçe nedeniyle SaveError fırlatır ("Disk dolu" vb.).
 */
export async function writeNativeFile(native: OfficehereNative, path: string, data: ArrayBuffer): Promise<void> {
  if (native.writeFileSafe) {
    const r = await native.writeFileSafe(path, data);
    if (!r.ok) throw saveErrorFrom(r);
    return;
  }
  if (!(await native.writeFile(path, data))) throw new SaveError("EUNKNOWN");
}

function isWriteFailure(r: NativeSaveResult | NativeWriteFailure | null): r is NativeWriteFailure {
  return !!r && typeof (r as NativeWriteFailure).error === "object" && !!(r as NativeWriteFailure).error;
}

/**
 * "Kaydet": dosyanın bir konumu varsa üzerine yazar, yoksa "Farklı Kaydet" açar.
 * Başarıda güncel dosya tanıtıcısını, iptalde null döner; yazılamazsa SaveError fırlatır.
 */
export async function saveFile(ref: FileRef | null, suggestedName: string, filters: FileFilter[], data: SaveData): Promise<FileRef | null> {
  const native = getNative();
  if (ref?.path && native) {
    await writeNativeFile(native, ref.path, await toArrayBuffer(data));
    return ref;
  }
  if (ref?.handle) {
    try {
      const writable = await ref.handle.createWritable();
      await writable.write(await toArrayBuffer(data));
      await writable.close();
      return ref;
    } catch {
      // İzin geri alınmış olabilir; Farklı Kaydet'e düş.
    }
  }
  return saveFileAs(suggestedName, filters, data);
}

/** "Farklı Kaydet": her zaman konum sorar (tarayıcıda gerekirse indirir). Yazılamazsa SaveError fırlatır. */
export async function saveFileAs(suggestedName: string, filters: FileFilter[], data: SaveData): Promise<FileRef | null> {
  const native = getNative();
  const buffer = await toArrayBuffer(data);
  if (native) {
    const r = await native.saveFileAs(suggestedName, filters, buffer);
    if (isWriteFailure(r)) throw saveErrorFrom(r.error);
    return r ? { name: r.name, path: r.path } : null;
  }
  const w = window as unknown as PickerWindow;
  if (w.showSaveFilePicker) {
    try {
      const handle = await w.showSaveFilePicker({ suggestedName, types: pickerTypes(filters) });
      const writable = await handle.createWritable();
      await writable.write(buffer);
      await writable.close();
      return { name: handle.name, handle };
    } catch (err) {
      if ((err as DOMException)?.name === "AbortError") return null;
    }
  }
  downloadBlob(new Blob([buffer], { type: mimeFor(suggestedName) }), suggestedName);
  return { name: suggestedName };
}

/** Veriyi tarayıcı indirmesi olarak verir (dışa aktarma için). */
export function downloadBlob(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  a.style.display = "none";
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30_000);
}

const RECENT_KEY = "officehere.recent";
const WEB_RECENT_MAX = 25;

const recentListeners = new Set<() => void>();

/** Bu penceredeki son kullanılanlar değişikliklerini (ekleme, sabitleme, kaldırma) bildirir. */
export function onRecentFilesChange(callback: () => void): () => void {
  recentListeners.add(callback);
  return () => recentListeners.delete(callback);
}

function notifyRecent() {
  for (const cb of recentListeners) {
    try {
      cb();
    } catch {
      // Bir dinleyicinin hatası diğerlerini durdurmasın.
    }
  }
}

function readWebRecent(): RecentFile[] {
  try {
    const list = JSON.parse(localStorage.getItem(RECENT_KEY) ?? "[]") as RecentFile[];
    return Array.isArray(list) ? list.filter((f) => f && typeof f.path === "string") : [];
  } catch {
    return [];
  }
}

function writeWebRecent(list: RecentFile[]) {
  try {
    localStorage.setItem(RECENT_KEY, JSON.stringify(list));
  } catch {
    // Yerel depo kapalıysa son kullanılanlar tutulmaz.
  }
}

/** Son kullanılan dosyalar (masaüstünde ana süreçten, tarayıcıda yerel depodan); sabitlenenler başta. */
export async function getRecentFiles(app?: AppKind): Promise<RecentFile[]> {
  const native = getNative();
  const list = orderRecent(native ? await native.getRecent() : readWebRecent());
  return app ? list.filter((f) => f.app === app) : list;
}

export async function addRecentFile(file: RecentFile): Promise<void> {
  const native = getNative();
  if (native) await native.addRecent(file);
  else writeWebRecent(withRecentAdded(readWebRecent(), file, WEB_RECENT_MAX));
  notifyRecent();
}

/** Dosyayı son kullanılanlarda sabitler ya da sabitlemeyi kaldırır. */
export async function pinRecentFile(file: Pick<RecentFile, "app" | "path">, pinned: boolean): Promise<void> {
  const native = getNative();
  if (native) await native.pinRecent?.(file.path, pinned);
  else writeWebRecent(withRecentPinned(readWebRecent(), file, pinned));
  notifyRecent();
}

/** Dosyayı son kullanılanlar listesinden çıkarır. */
export async function removeRecentFile(file: Pick<RecentFile, "app" | "path">): Promise<void> {
  const native = getNative();
  if (native) await native.removeRecent?.(file.path);
  else writeWebRecent(withRecentRemoved(readWebRecent(), file));
  notifyRecent();
}

/** Yollardaki dosyalar hâlâ var mı (masaüstü). Tarayıcıda denetlenemez; hepsi var sayılır. */
export async function getRecentFileStatus(paths: string[]): Promise<Record<string, boolean>> {
  const native = getNative();
  if (native?.recentStatus && paths.length) {
    try {
      return await native.recentStatus(paths);
    } catch {
      // Denetlenemezse dosyalar var sayılır.
    }
  }
  return Object.fromEntries(paths.map((p) => [p, true]));
}

/** "Klasörde göster" yalnız masaüstünde kullanılabilir. */
export function canShowInFolder(): boolean {
  return typeof getNative()?.showInFolder === "function";
}

export async function showInFolder(path: string): Promise<void> {
  await getNative()?.showInFolder?.(path);
}
