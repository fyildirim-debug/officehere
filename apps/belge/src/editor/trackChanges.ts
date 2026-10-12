// Değişiklikleri İzle: eklenen metin "insertion", silinen metin "deletion" işaretiyle belgede kalır.
// Kapsam: paragraf içi ekleme/silme (yazma, Geri Sil, Sil, yapıştırma, seçimi değiştirme).
import { Extension } from "@tiptap/core";
import { Fragment, Slice, type Mark, type Node as PMNode } from "@tiptap/pm/model";
import { Plugin, PluginKey, TextSelection, type EditorState, type Transaction } from "@tiptap/pm/state";
import { ReplaceStep } from "@tiptap/pm/transform";
import type { EditorView } from "@tiptap/pm/view";

interface TrackState {
  on: boolean;
  author: string;
}

export const trackKey = new PluginKey<TrackState>("belgeTrack");

let lastKey = "";

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
          const attrs = { author: st.author, date: nowIso() };
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
                out.addMark(a, b, ins);
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
          if (!touched) return null;
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
  kind: "insertion" | "deletion";
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
      if (m.type.name !== "insertion" && m.type.name !== "deletion") continue;
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
  const ranges: { from: number; to: number; kind: string }[] = [];
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isInline) return true;
    const a = Math.max(pos, from);
    const b = Math.min(pos + node.nodeSize, to);
    if (node.marks.some((m) => m.type === insType)) ranges.push({ from: a, to: b, kind: "ins" });
    if (node.marks.some((m) => m.type === delType)) ranges.push({ from: a, to: b, kind: "del" });
    return false;
  });
  for (let i = ranges.length - 1; i >= 0; i--) {
    const r = ranges[i];
    const removeText = (r.kind === "ins" && !accept) || (r.kind === "del" && accept);
    if (removeText) tr.delete(r.from, r.to);
    else tr.removeMark(r.from, r.to, r.kind === "ins" ? insType : delType);
  }
  return tr;
}
