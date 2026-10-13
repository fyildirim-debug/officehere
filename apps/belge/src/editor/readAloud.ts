// Sesli Oku (Gözden Geçir › Konuşma): belge imleçten (ya da seçimin başından) itibaren paragraf paragraf
// tarayıcının konuşma sentezleyicisiyle (Web Speech API) okunur; okunan sözcük ve paragraf vurgulanır.
// Türkçe ses yoksa varsayılan sesle okunur ve kullanıcı uyarılır.
import { Extension } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import { Plugin, PluginKey, type EditorState } from "@tiptap/pm/state";
import { Decoration, DecorationSet, type EditorView } from "@tiptap/pm/view";

/* ------------------------------------------------------------------ */
/* Okunacak birimler                                                     */
/* ------------------------------------------------------------------ */

export interface ReadUnit {
  /** Paragrafın belgedeki konumu (düğümün önü). */
  pos: number;
  /** Paragrafın sonu. */
  end: number;
  text: string;
  /** Metindeki her karakterin belge konumu. */
  map: number[];
}

/** Metin bloğunun okunacak metni ve konum eşlemesi (satır içi atomlar boşluk sayılır, silinmiş metin atlanır). */
export function unitOf(node: PMNode, pos: number): ReadUnit {
  let text = "";
  const map: number[] = [];
  node.forEach((child, off) => {
    const at = pos + 1 + off;
    if (child.isText) {
      if (child.marks.some((m) => m.type.name === "deletion")) return;
      const t = child.text ?? "";
      for (let i = 0; i < t.length; i++) map.push(at + i);
      text += t;
    } else {
      map.push(at);
      text += " ";
    }
  });
  return { pos, end: pos + node.nodeSize, text, map };
}

/** from konumundan başlayarak okunacak paragraflar (boş paragraflar atlanır). */
export function readUnits(doc: PMNode, from = 0): ReadUnit[] {
  const out: ReadUnit[] = [];
  doc.descendants((n, pos) => {
    if (!n.isTextblock) return true;
    if (pos + n.nodeSize > from) {
      const u = unitOf(n, pos);
      if (u.text.trim()) out.push(u);
    }
    return false;
  });
  return out;
}

/** Karakter dizinindeki sözcüğün sonu (sözcük sınırı olayında uzunluk verilmezse). */
export function wordEnd(text: string, index: number): number {
  const m = /^[\p{L}\p{N}'’\-.,]*[\p{L}\p{N}]/u.exec(text.slice(index));
  return index + (m ? m[0].length : 1);
}

/** Okumanın başlayacağı yer: seçim varsa başı, yoksa imleçteki paragrafın başı. */
export function readStart(state: EditorState): number {
  const { from, empty, $from } = state.selection;
  if (!empty) return from;
  return $from.parent.isTextblock ? $from.before() : from;
}

/** Türkçe ses önce; ses adına göre sıralı. */
export function sortVoices(voices: SpeechSynthesisVoice[]): SpeechSynthesisVoice[] {
  const tr = (v: SpeechSynthesisVoice) => (v.lang.toLowerCase().startsWith("tr") ? 0 : 1);
  return [...voices].sort((a, b) => tr(a) - tr(b) || a.name.localeCompare(b.name, "tr"));
}

/* ------------------------------------------------------------------ */
/* Vurgu                                                                */
/* ------------------------------------------------------------------ */

interface Highlight {
  para: { from: number; to: number } | null;
  word: { from: number; to: number } | null;
}

export const readAloudKey = new PluginKey<{ hl: Highlight; deco: DecorationSet }>("belgeReadAloud");

function build(doc: PMNode, hl: Highlight): DecorationSet {
  const decos: Decoration[] = [];
  if (hl.para && hl.para.to <= doc.content.size) decos.push(Decoration.node(hl.para.from, hl.para.to, { class: "bl-ra-para" }));
  if (hl.word && hl.word.from < hl.word.to && hl.word.to <= doc.content.size) decos.push(Decoration.inline(hl.word.from, hl.word.to, { class: "bl-ra-word" }));
  return DecorationSet.create(doc, decos);
}

export const ReadAloudHighlight = Extension.create({
  name: "readAloudHighlight",
  addProseMirrorPlugins() {
    return [
      new Plugin<{ hl: Highlight; deco: DecorationSet }>({
        key: readAloudKey,
        state: {
          init: () => ({ hl: { para: null, word: null }, deco: DecorationSet.empty }),
          apply(tr, prev) {
            const meta = tr.getMeta(readAloudKey) as Highlight | undefined;
            if (meta) return { hl: meta, deco: build(tr.doc, meta) };
            if (!tr.docChanged || (!prev.hl.para && !prev.hl.word)) return prev;
            const map = (r: { from: number; to: number } | null) => (r ? { from: tr.mapping.map(r.from, 1), to: tr.mapping.map(r.to, -1) } : null);
            const hl = { para: map(prev.hl.para), word: map(prev.hl.word) };
            return { hl, deco: prev.deco.map(tr.mapping, tr.doc) };
          },
        },
        props: {
          decorations: (state) => readAloudKey.getState(state)?.deco,
        },
      }),
    ];
  },
});

function setHighlight(view: EditorView, hl: Highlight) {
  if (view.isDestroyed) return;
  view.dispatch(view.state.tr.setMeta(readAloudKey, hl).setMeta("addToHistory", false));
}

/* ------------------------------------------------------------------ */
/* Okuyucu                                                              */
/* ------------------------------------------------------------------ */

export type ReadStatus = "idle" | "playing" | "paused";

export interface ReadState {
  status: ReadStatus;
  /** Okunan paragrafın sırası / toplam. */
  index: number;
  count: number;
}

/** Konuşma sentezleyici var mı (Electron/Chromium ve çoğu tarayıcı). */
export function speechSupported(): boolean {
  return typeof window !== "undefined" && "speechSynthesis" in window && typeof window.SpeechSynthesisUtterance === "function";
}

/**
 * Paragraf paragraf okuyan denetleyici. Her paragraf ayrı bir söyleyiş olarak kuyruğa verilir (uzun
 * metinlerde bazı sentezleyiciler yarıda kesildiği için); sözcük sınırı olaylarıyla vurgu güncellenir.
 */
export class ReadAloudPlayer {
  private units: ReadUnit[] = [];
  private index = 0;
  private status: ReadStatus = "idle";
  /** Birimlerin hesaplandığı belge: düzenleme olduysa birimler yeniden hesaplanır. */
  private doc: PMNode | null = null;
  /** Her söyleyişe kimlik: iptal edilen söyleyişin geç gelen olayları yok sayılır. */
  private token = 0;
  private listeners = new Set<() => void>();
  rate = 1;
  voice: SpeechSynthesisVoice | null = null;

  constructor(private view: EditorView) {}

  subscribe(fn: () => void): () => void {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  }

  get state(): ReadState {
    return { status: this.status, index: this.index, count: this.units.length };
  }

  private emit() {
    this.listeners.forEach((l) => l());
  }

  /** from konumundan okumaya başlar. */
  start(from: number) {
    this.doc = this.view.state.doc;
    this.units = readUnits(this.doc, from);
    this.index = 0;
    // Seçim paragrafın ortasında başlıyorsa ilk paragraf seçimin başından okunur.
    const first = this.units[0];
    if (first && from > first.pos + 1) {
      const k = first.map.findIndex((p) => p >= from);
      if (k > 0) this.units[0] = { ...first, text: first.text.slice(k), map: first.map.slice(k) };
    }
    if (!this.units.length) return this.stop();
    this.speak();
  }

  /** Okuma sırasında belge düzenlendiyse birimler yeniden hesaplanır; okunan paragraf korunur. */
  private resync() {
    const doc = this.view.state.doc;
    if (doc === this.doc) return;
    const para = readAloudKey.getState(this.view.state)?.hl.para;
    this.doc = doc;
    this.units = readUnits(doc, 0);
    const at = para ? this.units.findIndex((u) => u.end > para.from) : 0;
    this.index = Math.max(0, at);
  }

  private speak() {
    const synth = window.speechSynthesis;
    const id = ++this.token;
    synth.cancel();
    const unit = this.units[this.index];
    if (!unit) return this.stop();
    this.status = "playing";
    setHighlight(this.view, { para: { from: unit.pos, to: unit.end }, word: null });
    this.scrollTo(unit.map[0] ?? unit.pos);
    const utter = new SpeechSynthesisUtterance(unit.text);
    utter.lang = this.voice?.lang ?? "tr-TR";
    if (this.voice) utter.voice = this.voice;
    utter.rate = this.rate;
    utter.onboundary = (e) => {
      if (id !== this.token || e.name !== "word" || this.view.state.doc !== this.doc) return;
      const a = e.charIndex;
      const b = e.charLength ? a + e.charLength : wordEnd(unit.text, a);
      const from = unit.map[a];
      if (from === undefined) return;
      const to = (unit.map[Math.min(b, unit.map.length) - 1] ?? from) + 1;
      setHighlight(this.view, { para: { from: unit.pos, to: unit.end }, word: { from, to } });
      this.scrollTo(from);
    };
    utter.onend = () => {
      if (id !== this.token) return;
      this.resync();
      if (this.index < this.units.length - 1) {
        this.index++;
        this.speak();
      } else this.stop();
    };
    utter.onerror = (e) => {
      if (id !== this.token || e.error === "interrupted" || e.error === "canceled") return;
      this.stop();
    };
    synth.speak(utter);
    this.emit();
  }

  private scrollTo(pos: number) {
    try {
      const dom = this.view.domAtPos(pos).node;
      const el = dom instanceof HTMLElement ? dom : dom.parentElement;
      el?.scrollIntoView?.({ block: "nearest" });
    } catch {
      // konum çizilmemiş
    }
  }

  pause() {
    if (this.status !== "playing") return;
    window.speechSynthesis.pause();
    this.status = "paused";
    this.emit();
  }

  resume() {
    if (this.status !== "paused") return;
    window.speechSynthesis.resume();
    this.status = "playing";
    this.emit();
  }

  /** Önceki / sonraki paragraf. */
  skip(dir: 1 | -1) {
    if (!this.units.length) return;
    this.resync();
    this.index = Math.max(0, Math.min(this.units.length - 1, this.index + dir));
    this.speak();
  }

  /** Hız ya da ses değişince okunan paragraf baştan, yeni ayarla okunur. */
  restartCurrent() {
    if (this.status !== "playing") return;
    this.resync();
    this.speak();
  }

  stop() {
    this.token++;
    if (speechSupported()) window.speechSynthesis.cancel();
    this.status = "idle";
    setHighlight(this.view, { para: null, word: null });
    this.emit();
  }
}
