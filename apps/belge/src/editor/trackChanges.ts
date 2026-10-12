// Değişiklikleri İzle: eklenen metin "insertion", silinen metin "deletion" işaretiyle belgede kalır.
// Kapsam: paragraf içi ekleme/silme (yazma, Geri Sil, Sil, yapıştırma, seçimi değiştirme) ve karakter
// biçimi değişiklikleri ("formatChange": önceki biçim saklanır; reddedilince geri yüklenir).
import { Extension } from "@tiptap/core";
import { Fragment, Slice, type Mark, type Node as PMNode, type Schema } from "@tiptap/pm/model";
import { Plugin, PluginKey, TextSelection, type EditorState, type Transaction } from "@tiptap/pm/state";
import { AddMarkStep, RemoveMarkStep, ReplaceAroundStep, ReplaceStep, type Step } from "@tiptap/pm/transform";
import type { EditorView } from "@tiptap/pm/view";

interface TrackState {
  on: boolean;
  author: string;
}

export const trackKey = new PluginKey<TrackState>("belgeTrack");

let lastKey = "";

/** Değişikliklerin yazarı: Seçenekler'deki kullanıcı adı (her değişiklikte güncel değer okunur). */
let authorSource: (() => string) | null = null;
export function setTrackAuthorSource(fn: () => string) {
  authorSource = fn;
}

function nowIso() {
  return new Date().toISOString().replace(/\.\d+Z$/, "Z");
}

/** Silinen dilimi silme işaretiyle yeniden kurar; aynı yazarın eklemeleri gerçekten silinir. */
function markDeleted(frag: Fragment, delMark: Mark, insType: string): Fragment {
  const out: PMNode[] = [];
  frag.forEach((n) => {
    if (n.isText) {
      if (n.marks.some((m) => m.type.name === insType)) return;
      out.push(n.mark(delMark.addToSet(n.marks)));
    } else if (n.isInline) {
      out.push(n.mark(delMark.addToSet(n.marks)));
    }
  });
  return Fragment.from(out);
}

/** Biçim sayılmayan işaretler (izleme, açıklama, köprü). */
const NOT_FORMAT = new Set(["insertion", "deletion", "formatChange", "comment", "link"]);

const formatMarks = (marks: readonly Mark[]) => marks.filter((m) => !NOT_FORMAT.has(m.type.name));
const marksKey = (marks: readonly Mark[]) => JSON.stringify(formatMarks(marks).map((m) => m.toJSON()));

/** Biçim değişikliği adımlarının aralıkları; konum değiştiren bir adım varsa null. */
function formatRanges(t: Transaction): { from: number; to: number }[] | null {
  const out: { from: number; to: number }[] = [];
  for (const step of t.steps as Step[]) {
    if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) {
      const s = step as unknown as { from: number; to: number; mark: Mark };
      if (!NOT_FORMAT.has(s.mark.type.name)) out.push({ from: s.from, to: s.to });
    } else if (step instanceof ReplaceAroundStep) {
      // Düğüm işaretlemesi (paragraf stili) konumları değiştirmez; izlenmez.
      const s = step as unknown as { from: number; to: number; gapFrom: number; gapTo: number; insert: number; slice: Slice };
      if (s.slice.size + s.gapTo - s.gapFrom !== s.to - s.from || s.insert !== s.gapFrom - s.from) return null;
    } else return null;
  }
  return out;
}

/** Biçim değişikliklerini işaretler: metnin değişiklik öncesi biçimi "old" özniteliğinde saklanır. */
function trackFormat(t: Transaction, ranges: { from: number; to: number }[], newState: EditorState, attrs: { author: string; date: string }): Transaction | null {
  const fc = newState.schema.marks.formatChange;
  const insType = newState.schema.marks.insertion;
  if (!fc || !ranges.length) return null;
  const out = newState.tr;
  let touched = false;
  for (const r of ranges) {
    newState.doc.nodesBetween(r.from, r.to, (node, pos) => {
      if (!node.isInline) return true;
      const a = Math.max(pos, r.from);
      const b = Math.min(pos + node.nodeSize, r.to);
      if (b <= a || node.marks.some((m) => m.type === insType)) return false;
      const prev = node.marks.find((m) => m.type === fc);
      if (prev) {
        // Biçim eski hâline döndüyse değişiklik kaydı kalkar.
        if (marksKey(node.marks) === JSON.stringify(JSON.parse(String(prev.attrs.old || "[]")))) {
          out.removeMark(a, b, fc);
          touched = true;
        }
        return false;
      }
      // Önceki biçim, eski belgedeki parçalara göre (parça sınırları değişmiş olabilir).
      const now = marksKey(node.marks);
      t.before.nodesBetween(a, b, (on, opos) => {
        if (!on.isInline) return true;
        const sa = Math.max(a, opos);
        const sb = Math.min(b, opos + on.nodeSize);
        const oldKey = marksKey(on.marks);
        if (sb > sa && oldKey !== now) {
          out.addMark(sa, sb, fc.create({ ...attrs, old: oldKey }));
          touched = true;
        }
        return false;
      });
      return false;
    });
  }
  if (!touched) return null;
  out.setMeta(trackKey, { applied: true });
  return out;
}

/** "old" özniteliğindeki biçim işaretlerini kurar. */
function oldMarks(schema: Schema, json: string): Mark[] {
  try {
    return (JSON.parse(json) as { type: string; attrs?: Record<string, unknown> }[]).filter((m) => schema.marks[m.type]).map((m) => schema.mark(m.type, m.attrs));
  } catch {
    return [];
  }
}

export const TrackChanges = Extension.create({
  name: "belgeTrackChanges",
  addProseMirrorPlugins() {
    return [
      new Plugin<TrackState>({
        key: trackKey,
        state: {
          init: () => ({ on: false, author: "Kullanıcı" }),
          apply(tr, value) {
            const meta = tr.getMeta(trackKey) as Partial<TrackState> | undefined;
            return meta && ("on" in meta || "author" in meta) ? { ...value, ...meta } : value;
          },
        },
        props: {
          handleKeyDown(_view, event) {
            lastKey = event.key;
            return false;
          },
          handleTextInput() {
            lastKey = "";
            return false;
          },
        },
        appendTransaction(trs, _oldState, newState) {
          const st = trackKey.getState(newState);
          if (!st?.on || trs.length !== 1) return null;
          const t = trs[0];
          if (!t.docChanged || t.getMeta(trackKey) || t.getMeta("history$") || t.getMeta("belgeNoTrack")) return null;
          const schema = newState.schema;
          const insType = schema.marks.insertion;
          const delType = schema.marks.deletion;
          if (!insType || !delType) return null;
          const author = authorSource?.() || st.author;
          const attrs = { author, date: nowIso() };
          const fmt = formatRanges(t);
          if (fmt) return trackFormat(t, fmt, newState, attrs);
          const ins = insType.create(attrs);
          const del = delType.create(attrs);
          const out = newState.tr;
          let touched = false;
          let cursor: number | null = null;
          for (let i = 0; i < t.steps.length; i++) {
            const step = t.steps[i];
            if (!(step instanceof ReplaceStep)) return null;
            const s = step as unknown as { from: number; to: number; slice: Slice };
            const before = t.docs[i];
            const rest = t.mapping.slice(i + 1);
            // Eklenen içerik
            if (s.slice.size > 0) {
              const a = out.mapping.map(rest.map(s.from, -1), -1);
              const b = out.mapping.map(rest.map(s.from + s.slice.size, 1), 1);
              if (b > a) {
                // Aynı yazarın bitişik eklemesi sürer: tek değişiklik olarak kalır (her tuşta yeni tarih yok).
                const $a = out.doc.resolve(a);
                const near = [$a.nodeBefore, out.doc.resolve(b).nodeAfter].flatMap((n) => n?.marks ?? []).find((m) => m.type === insType && m.attrs.author === author);
                out.addMark(a, b, near ?? ins);
                out.removeMark(a, b, delType);
                touched = true;
              }
            }
            // Silinen içerik: yalnız tek paragraf içi silmeler geri konur.
            if (s.to > s.from) {
              const $a = before.resolve(s.from);
              const $b = before.resolve(s.to);
              if ($a.parent !== $b.parent || !$a.parent.isTextblock) continue;
              const slice = before.slice(s.from, s.to);
              const frag = markDeleted(slice.content, del, insType.name);
              if (frag.size === 0) continue;
              const at = out.mapping.map(rest.map(s.from, -1), -1);
              out.insert(at, frag);
              touched = true;
              if (s.slice.size === 0) cursor = lastKey === "Delete" ? at + frag.size : at;
            }
          }
          if (!touched || (!out.docChanged && cursor === null)) return null;
          if (cursor !== null && t.selection.empty) out.setSelection(TextSelection.create(out.doc, Math.min(cursor, out.doc.content.size)));
          out.setMeta(trackKey, { applied: true });
          return out;
        },
      }),
    ];
  },
});

export function isTracking(state: EditorState) {
  return !!trackKey.getState(state)?.on;
}

export function setTracking(view: EditorView, on: boolean, author: string) {
  view.dispatch(view.state.tr.setMeta(trackKey, { on, author }).setMeta("addToHistory", false));
}

export interface TrackedChange {
  kind: "insertion" | "deletion" | "formatChange";
  from: number;
  to: number;
  author: string;
  date: string;
  text: string;
}

/** Belgedeki izlenen değişiklikleri (bitişik aralıklar birleştirilmiş) listeler. */
export function listChanges(doc: PMNode): TrackedChange[] {
  const out: TrackedChange[] = [];
  doc.descendants((node, pos) => {
    if (!node.isInline) return true;
    for (const m of node.marks) {
      if (m.type.name !== "insertion" && m.type.name !== "deletion" && m.type.name !== "formatChange") continue;
      const kind = m.type.name as TrackedChange["kind"];
      const last = out[out.length - 1];
      const text = node.isText ? (node.text ?? "") : "";
      if (last && last.kind === kind && last.to === pos && last.author === m.attrs.author) {
        last.to = pos + node.nodeSize;
        last.text += text;
      } else out.push({ kind, from: pos, to: pos + node.nodeSize, author: m.attrs.author, date: m.attrs.date, text });
    }
    return false;
  });
  return out;
}

/** Aralıktaki değişiklikleri kabul eder ya da reddeder. */
export function resolveChanges(state: EditorState, from: number, to: number, accept: boolean): Transaction {
  const tr = state.tr.setMeta("belgeNoTrack", true);
  const insType = state.schema.marks.insertion;
  const delType = state.schema.marks.deletion;
  const fcType = state.schema.marks.formatChange;
  const ranges: { from: number; to: number; kind: string; old?: string }[] = [];
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isInline) return true;
    const a = Math.max(pos, from);
    const b = Math.min(pos + node.nodeSize, to);
    if (node.marks.some((m) => m.type === insType)) ranges.push({ from: a, to: b, kind: "ins" });
    if (node.marks.some((m) => m.type === delType)) ranges.push({ from: a, to: b, kind: "del" });
    const fc = node.marks.find((m) => m.type === fcType);
    if (fc) ranges.push({ from: a, to: b, kind: "fmt", old: String(fc.attrs.old || "[]") });
    return false;
  });
  for (let i = ranges.length - 1; i >= 0; i--) {
    const r = ranges[i];
    if (r.kind === "fmt") {
      // Kabul: kayıt kalkar. Ret: önceki biçim geri yüklenir.
      if (!accept) {
        for (const mt of Object.values(state.schema.marks)) if (!NOT_FORMAT.has(mt.name)) tr.removeMark(r.from, r.to, mt);
        for (const m of oldMarks(state.schema, r.old ?? "[]")) tr.addMark(r.from, r.to, m);
      }
      tr.removeMark(r.from, r.to, fcType);
      continue;
    }
    const removeText = (r.kind === "ins" && !accept) || (r.kind === "del" && accept);
    if (removeText) tr.delete(r.from, r.to);
    else tr.removeMark(r.from, r.to, r.kind === "ins" ? insType : delType);
  }
  return tr;
}
