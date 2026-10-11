import { getNative, type AppKind, type FileFilter, type RecentFile } from "./bridge";

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
 * "Kaydet": dosyanın bir konumu varsa üzerine yazar, yoksa "Farklı Kaydet" açar.
 * Başarıda güncel dosya tanıtıcısını, iptalde null döner.
 */
export async function saveFile(ref: FileRef | null, suggestedName: string, filters: FileFilter[], data: SaveData): Promise<FileRef | null> {
  const native = getNative();
  if (ref?.path && native) {
    const ok = await native.writeFile(ref.path, await toArrayBuffer(data));
    return ok ? ref : null;
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

/** "Farklı Kaydet": her zaman konum sorar (tarayıcıda gerekirse indirir). */
export async function saveFileAs(suggestedName: string, filters: FileFilter[], data: SaveData): Promise<FileRef | null> {
  const native = getNative();
  const buffer = await toArrayBuffer(data);
  if (native) {
    const r = await native.saveFileAs(suggestedName, filters, buffer);
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

/** Son kullanılan dosyalar (masaüstünde ana süreçten, tarayıcıda yerel depodan). */
export async function getRecentFiles(app?: AppKind): Promise<RecentFile[]> {
  const native = getNative();
  let list: RecentFile[] = [];
  if (native) {
    list = await native.getRecent();
  } else {
    try {
      list = JSON.parse(localStorage.getItem(RECENT_KEY) ?? "[]") as RecentFile[];
    } catch {
      list = [];
    }
  }
  return app ? list.filter((f) => f.app === app) : list;
}

export async function addRecentFile(file: RecentFile): Promise<void> {
  const native = getNative();
  if (native) return native.addRecent(file);
  try {
    const list = (JSON.parse(localStorage.getItem(RECENT_KEY) ?? "[]") as RecentFile[]).filter(
      (f) => !(f.path === file.path && f.app === file.app),
    );
    list.unshift(file);
    localStorage.setItem(RECENT_KEY, JSON.stringify(list.slice(0, 25)));
  } catch {
    // Yerel depo kapalıysa son kullanılanlar tutulmaz.
  }
}
