// Dipnot ve son not: başvuru işareti satır içi düğümdür, not metni düğüm özniteliğinde tutulur.
// Dipnotlar sayfanın altında (sayfalama alan ayırır), son notlar belgenin sonunda gösterilir.
import { Extension, Node, mergeAttributes } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import { Decoration, DecorationSet, type EditorView } from "@tiptap/pm/view";

export type NoteKind = "footnote" | "endnote";

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    note: { insertNote: (kind: NoteKind, text?: string) => ReturnType };
  }
}

export interface NoteInfo {
  id: string;
  kind: NoteKind;
  text: string;
  pos: number;
  label: string;
}

const ROMAN: [number, string][] = [
  [1000, "m"],
  [900, "cm"],
  [500, "d"],
  [400, "cd"],
  [100, "c"],
  [90, "xc"],
  [50, "l"],
  [40, "xl"],
  [10, "x"],
  [9, "ix"],
  [5, "v"],
  [4, "iv"],
  [1, "i"],
];

export function lowerRoman(n: number): string {
  let out = "";
  for (const [v, s] of ROMAN) {
    while (n >= v) {
      out += s;
      n -= v;
    }
  }
  return out;
}

/** Belgedeki notlar, sıra numaralarıyla (dipnot 1,2,3…; son not i, ii, iii…). */
export function collectNotes(doc: PMNode): NoteInfo[] {
  const out: NoteInfo[] = [];
  let f = 0;
  let e = 0;
  doc.descendants((node, pos) => {
    if (node.type.name !== "note") return true;
    const kind = node.attrs.kind as NoteKind;
    const label = kind === "endnote" ? lowerRoman(++e) : String(++f);
    out.push({ id: node.attrs.id as string, kind, text: (node.attrs.text as string) ?? "", pos, label });
    return false;
  });
  return out;
}

export const Note = Node.create({
  name: "note",
  group: "inline",
  inline: true,
  atom: true,
  selectable: true,
  addAttributes() {
    return {
      id: { default: "", parseHTML: (el) => el.getAttribute("data-note-id") ?? "", renderHTML: (a) => ({ "data-note-id": a.id }) },
      kind: { default: "footnote", parseHTML: (el) => el.getAttribute("data-note") ?? "footnote", renderHTML: (a) => ({ "data-note": a.kind }) },
      text: { default: "", parseHTML: (el) => el.getAttribute("data-text") ?? el.getAttribute("title") ?? "", renderHTML: (a) => ({ "data-text": a.text, title: a.text }) },
    };
  },
  parseHTML() {
    return [{ tag: "sup[data-note]" }];
  },
  renderHTML({ HTMLAttributes }) {
    return ["sup", mergeAttributes(HTMLAttributes, { class: "bl-noteref" })];
  },
  renderText() {
    return "";
  },
  addCommands() {
    return {
      insertNote:
        (kind, text = "") =>
        ({ commands }) =>
          commands.insertContent({ type: "note", attrs: { kind, text, id: `n${Date.now().toString(36)}${Math.random().toString(36).slice(2, 5)}` } }),
    };
  },
});

export const notesKey = new PluginKey<DecorationSet>("belgeNotes");

/** Not numaralandırması ve belge sonundaki son notlar bölümü. */
export function createNotesPlugin() {
  const build = (doc: PMNode): DecorationSet => {
    const notes = collectNotes(doc);
    const decos: Decoration[] = notes.map((n) => Decoration.node(n.pos, n.pos + 1, { "data-num": n.label }));
    const ends = notes.filter((n) => n.kind === "endnote");
    if (ends.length) {
      const key = `endnotes:${ends.map((n) => `${n.id}=${n.label}`).join(",")}`;
      decos.push(
        Decoration.widget(
          doc.content.size,
          (view: EditorView) => {
            const wrap = document.createElement("div");
            wrap.className = "bl-endnotes";
            wrap.contentEditable = "false";
            const sep = document.createElement("div");
            sep.className = "bl-note-sep";
            wrap.appendChild(sep);
            for (const n of collectNotes(view.state.doc).filter((x) => x.kind === "endnote")) {
              const row = document.createElement("div");
              row.className = "bl-note-row";
              const num = document.createElement("sup");
              num.textContent = n.label;
              const input = document.createElement("textarea");
              input.className = "bl-note-text";
              input.rows = 1;
              input.value = n.text;
              input.placeholder = "Son not metni";
              input.dataset.noteId = n.id;
              input.addEventListener("input", () => setNoteText(view, n.id, input.value));
              row.append(num, input);
              wrap.appendChild(row);
            }
            return wrap;
          },
          { side: 1, key, stopEvent: () => true, ignoreSelection: true },
        ),
      );
    }
    return DecorationSet.create(doc, decos);
  };
  return new Plugin<DecorationSet>({
    key: notesKey,
    state: {
      init: (_c, state) => build(state.doc),
      apply: (tr, old, _o, state) => (tr.docChanged ? build(state.doc) : old),
    },
    props: {
      decorations: (state) => notesKey.getState(state),
    },
  });
}

/** Not metnini kimliğe göre günceller. */
export function setNoteText(view: EditorView, id: string, text: string) {
  let target = -1;
  view.state.doc.descendants((node, pos) => {
    if (target >= 0) return false;
    if (node.type.name === "note" && node.attrs.id === id) target = pos;
    return true;
  });
  if (target < 0) return;
  const node = view.state.doc.nodeAt(target)!;
  view.dispatch(view.state.tr.setNodeMarkup(target, undefined, { ...node.attrs, text }).setMeta("belgeNoteEdit", true));
}

export const Notes = Extension.create({
  name: "belgeNotes",
  addProseMirrorPlugins() {
    return [createNotesPlugin()];
  },
});

/** Statik HTML'de (Okuma Modu, HTML dışa aktarma) not numaralarını yazar ve notları sona ekler. */
export function annotateNotesHtml(html: string): string {
  if (!html.includes("data-note")) return html;
  const doc = new DOMParser().parseFromString(`<div id="r">${html}</div>`, "text/html");
  const root = doc.getElementById("r")!;
  const rows: { label: string; text: string }[] = [];
  let f = 0;
  let e = 0;
  const sups = Array.from(root.querySelectorAll("sup[data-note]"));
  // Önce dipnotlar, sonra son notlar listelenir.
  for (const kind of ["footnote", "endnote"]) {
    for (const s of sups) {
      if (s.getAttribute("data-note") !== kind) continue;
      const label = kind === "endnote" ? lowerRoman(++e) : String(++f);
      s.textContent = label;
      rows.push({ label, text: s.getAttribute("data-text") ?? "" });
    }
  }
  const box = doc.createElement("div");
  box.className = "bl-endnotes";
  box.appendChild(Object.assign(doc.createElement("div"), { className: "bl-note-sep" }));
  for (const r of rows) {
    const row = doc.createElement("div");
    row.className = "bl-note-row";
    const sup = doc.createElement("sup");
    sup.textContent = r.label;
    const span = doc.createElement("span");
    span.className = "bl-note-text";
    span.textContent = r.text;
    row.append(sup, span);
    box.appendChild(row);
  }
  root.appendChild(box);
  return root.innerHTML;
}
