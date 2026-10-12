/**
 * Uygulama durumu: sunu + görünüm durumu. Geri al/yinele, sunu ve seçimin
 * anlık görüntüleriyle (core History) yapılır; ardışık yazma birleştirilir.
 */
import { useSyncExternalStore } from "react";
import { History, type FileRef } from "@officehere/core";
import type { Paragraph, Presentation, RunProps, Shape, Slide } from "../model/types";

export interface EditTarget {
  shapeId: string;
  cell?: [number, number];
  /** Metnin tamamı seçili başlasın. */
  selectAll?: boolean;
  /** Mevcut metin yazılan karakterle değiştirilsin. */
  replaceWith?: string;
  /** Tıklama noktası (imleci oraya koymak için). */
  point?: { x: number; y: number };
  /** Seçilecek metin aralığı (Bul). */
  range?: [{ p: number; o: number }, { p: number; o: number }];
}

export type ViewMode = "normal" | "outline" | "sorter" | "notes" | "reading" | "master";
export type PaneKind = "background" | "shape" | "animation" | "selection" | "comments" | null;

export interface InsertMode {
  kind: "shape" | "textbox" | "table";
  geom?: string;
  head?: "triangle";
  tail?: "triangle";
  vertical?: boolean;
}

/** Çizim sekmesi: etkin kalem ya da silgi. */
export interface InkTool {
  tool: "pen" | "pencil" | "highlighter" | "eraser";
  /** "#RRGGBB" */
  color: string;
  /** pt */
  width: number;
}

export interface AppState {
  pres: Presentation;
  slideIdx: number;
  /** Küçük resim bölmesinde seçili slaytlar. */
  slideSel: string[];
  /** Geçerli slaytta seçili şekiller. */
  sel: string[];
  editing: EditTarget | null;
  fileName: string;
  fileRef: FileRef | null;
  dirty: boolean;
  view: ViewMode;
  zoom: number;
  fit: boolean;
  showNotes: boolean;
  notesHeight: number;
  thumbWidth: number;
  showRuler: boolean;
  showGrid: boolean;
  showGuides: boolean;
  pane: PaneKind;
  insert: InsertMode | null;
  focus: "slides" | "canvas" | "notes";
  /** Seçili tablo hücre aralığı. */
  cellSel: { r1: number; c1: number; r2: number; c2: number } | null;
  /** Biçim Boyacısı: kopyalanan biçim. */
  painter: { shape?: Partial<Shape>; run?: RunProps; para?: Paragraph["props"]; sticky: boolean } | null;
  spellcheck: boolean;
  /** Kırpma modundaki resim. */
  cropId: string | null;
  /** Çizim modu (null: seçim). */
  ink: InkTool | null;
  /** Seçim değişimini şeride bildirmek için sayaç. */
  textTick: number;
  ribbonTab: string;
  /** Daraltılmış bölümlerin kimlikleri. */
  collapsedSections: string[];
  /** Asıl Slayt Görünümü açıkken: girişteki sunu ve slayt dizini. */
  masterBase: { pres: Presentation; slideIdx: number; dirty: boolean; tab: string } | null;
  message: { text: string; kind: "info" | "warning" | "error" | "success" } | null;
}

interface Snapshot {
  pres: Presentation;
  slideIdx: number;
  sel: string[];
}

export interface CommitOptions {
  sel?: string[];
  slideIdx?: number;
  editing?: EditTarget | null;
  /** Aynı anahtarla gelen ardışık değişiklikler tek geri alma adımıdır. */
  merge?: string;
  /** Geçmişe eklenmez (ölçüm sonucu boyut güncellemesi vb.). */
  silent?: boolean;
}

export interface ClipboardData {
  kind: "shapes" | "slides" | "text";
  shapes?: Shape[];
  slides?: Slide[];
  paragraphs?: Paragraph[];
  sourceTheme?: string;
}

export class Store {
  private state: AppState;
  private listeners = new Set<() => void>();
  private history = new History<Snapshot>(300);
  private lastMerge: { key: string; time: number } | null = null;
  /** Ayrı geri alma geçmişiyle yürüyen oturum (Asıl Slayt Görünümü) sırasında asıl geçmiş. */
  private savedHistory: History<Snapshot> | null = null;
  clipboard: ClipboardData | null = null;
  /** Her değişiklikten sonra sunuya uygulanır (asıl görünümde asıl slaytı eşitler). */
  normalize: ((p: Presentation) => Presentation) | null = null;

  constructor(initial: AppState) {
    this.state = initial;
  }

  get = () => this.state;

  subscribe = (fn: () => void) => {
    this.listeners.add(fn);
    return () => {
      this.listeners.delete(fn);
    };
  };

  private emit() {
    for (const l of this.listeners) l();
  }

  set(patch: Partial<AppState> | ((s: AppState) => Partial<AppState>)) {
    const p = typeof patch === "function" ? patch(this.state) : patch;
    this.state = { ...this.state, ...p };
    this.emit();
  }

  commit(fn: (p: Presentation, s: AppState) => Presentation, opts: CommitOptions = {}) {
    const s = this.state;
    let next = fn(s.pres, s);
    if (next !== s.pres && this.normalize) next = this.normalize(next);
    if (next === s.pres && opts.sel === undefined && opts.slideIdx === undefined) return;
    if (!opts.silent && next !== s.pres) {
      const now = Date.now();
      const merging = opts.merge && this.lastMerge?.key === opts.merge && now - this.lastMerge.time < 2500;
      if (!merging) this.history.push({ pres: s.pres, slideIdx: s.slideIdx, sel: s.sel });
      this.lastMerge = opts.merge ? { key: opts.merge, time: now } : null;
    }
    const slideIdx = Math.max(0, Math.min(next.slides.length - 1, opts.slideIdx ?? s.slideIdx));
    const patch: Partial<AppState> = { pres: next, dirty: s.dirty || (next !== s.pres && !opts.silent), slideIdx };
    if (opts.sel) patch.sel = opts.sel;
    if (opts.editing !== undefined) patch.editing = opts.editing;
    if (opts.slideIdx !== undefined) patch.slideSel = [next.slides[slideIdx]?.id].filter(Boolean) as string[];
    this.set(patch);
  }

  /** Birleştirme zincirini keser (yeni geri alma adımı başlatır). */
  breakMerge() {
    this.lastMerge = null;
  }

  get canUndo() {
    return this.history.canUndo;
  }

  get canRedo() {
    return this.history.canRedo;
  }

  undo() {
    const s = this.state;
    const prev = this.history.undo({ pres: s.pres, slideIdx: s.slideIdx, sel: s.sel });
    if (!prev) return;
    this.lastMerge = null;
    this.restore(prev);
  }

  redo() {
    const s = this.state;
    const next = this.history.redo({ pres: s.pres, slideIdx: s.slideIdx, sel: s.sel });
    if (!next) return;
    this.lastMerge = null;
    this.restore(next);
  }

  private restore(snap: Snapshot) {
    const idx = Math.max(0, Math.min(snap.pres.slides.length - 1, snap.slideIdx));
    const editing = this.state.editing;
    const keepEditing = editing && snap.pres.slides[idx]?.shapes.some((sh) => sh.id === editing.shapeId) ? { shapeId: editing.shapeId, cell: editing.cell } : null;
    this.set({ pres: snap.pres, slideIdx: idx, sel: snap.sel, editing: keepEditing, dirty: true, slideSel: [snap.pres.slides[idx]?.id].filter(Boolean) as string[] });
  }

  resetHistory() {
    this.history.clear();
    this.lastMerge = null;
  }

  /** Ayrı geçmişli oturum başlatır (asıl görünüm). */
  beginSession() {
    if (this.savedHistory) return;
    this.savedHistory = this.history;
    this.history = new History<Snapshot>(300);
    this.lastMerge = null;
  }

  /** Oturumu bitirir; oturumda değişiklik yapıldıysa true. */
  endSession(): boolean {
    if (!this.savedHistory) return false;
    const changed = this.history.canUndo;
    this.history = this.savedHistory;
    this.savedHistory = null;
    this.lastMerge = null;
    this.normalize = null;
    return changed;
  }

  get inSession(): boolean {
    return !!this.savedHistory;
  }
}

export function useStore(store: Store): AppState {
  return useSyncExternalStore(store.subscribe, store.get, store.get);
}
