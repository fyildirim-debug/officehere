// Anahat görünümü: üst düzey bloklar anahat düzeyine göre girintilenir (başlıklar ⊕, gövde metni ○);
// Düzey Göster süzgeci, başlık daraltma/genişletme, Yükselt/Düşür ve Yukarı/Aşağı Taşı komutları.
import { Extension, type Editor } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import { Plugin, PluginKey, TextSelection, type EditorState } from "@tiptap/pm/state";
import { Decoration, DecorationSet } from "@tiptap/pm/view";
import { syncDomSelection } from "./pagination";

export interface OutlineState {
  on: boolean;
  /** Gösterilecek en yüksek düzey (0: tüm düzeyler). */
  level: number;
  /** Daraltılmış başlıkların konumları. */
  collapsed: number[];
  firstLineOnly: boolean;
}

export const outlineKey = new PluginKey<{ s: OutlineState; deco: DecorationSet }>("belgeOutline");

/** Bloğun anahat düzeyi: başlık düzeyi, anahat düzeyli paragraf ya da gövde metni (10). */
export function outlineLevelOf(n: PMNode): number {
  if (n.type.name === "heading") return n.attrs.level as number;
  if (n.type.name === "paragraph" && n.attrs.outlineLevel) return n.attrs.outlineLevel as number;
  return 10;
}

interface Block {
  pos: number;
  node: PMNode;
  level: number;
}

function blocks(doc: PMNode): Block[] {
  const out: Block[] = [];
  doc.forEach((sec, secOff) => {
    sec.forEach((n, off) => out.push({ pos: secOff + 1 + off, node: n, level: outlineLevelOf(n) }));
  });
  return out;
}

function build(doc: PMNode, s: OutlineState): DecorationSet {
  if (!s.on) return DecorationSet.empty;
  const decos: Decoration[] = [];
  let parent = 0;
  let hideBelow = 99;
  const list = blocks(doc);
  list.forEach((b, i) => {
    const heading = b.level < 10;
    if (heading && b.level <= hideBelow) hideBelow = 99;
    const depth = heading ? b.level - 1 : parent;
    const filtered = s.level > 0 && b.level > s.level;
    const hidden = filtered || b.level > hideBelow;
    // Alt öğesi olan başlık: ⊕; olmayan: ⊖.
    const hasKids = heading && i + 1 < list.length && list[i + 1].level > b.level;
    const cls = ["bl-ol", heading ? `bl-ol-h${hasKids ? " has-kids" : ""}` : "bl-ol-body", hidden ? "bl-ol-hidden" : "", s.firstLineOnly && !heading ? "bl-ol-first" : ""].filter(Boolean).join(" ");
    decos.push(Decoration.node(b.pos, b.pos + b.node.nodeSize, { class: cls, style: `--bl-ol:${depth}` }));
    if (heading) {
      parent = b.level;
      if (s.collapsed.includes(b.pos) && hideBelow === 99) hideBelow = b.level;
    }
  });
  return DecorationSet.create(doc, decos);
}

export const Outline = Extension.create({
  name: "belgeOutline",
  addKeyboardShortcuts() {
    // Word kısayolları: Alt+Shift+Yukarı/Aşağı paragrafı taşır, Alt+Shift+Sol/Sağ yükseltir/düşürür.
    return {
      "Alt-Shift-ArrowUp": () => moveBlock(this.editor, -1),
      "Alt-Shift-ArrowDown": () => moveBlock(this.editor, 1),
      "Alt-Shift-ArrowLeft": () => (shiftOutline(this.editor, -1), true),
      "Alt-Shift-ArrowRight": () => (shiftOutline(this.editor, 1), true),
    };
  },
  addProseMirrorPlugins() {
    return [
      new Plugin<{ s: OutlineState; deco: DecorationSet }>({
        key: outlineKey,
        state: {
          init: (_c, state) => {
            const s = { on: false, level: 0, collapsed: [], firstLineOnly: false };
            return { s, deco: build(state.doc, s) };
          },
          apply(tr, v, _o, state) {
            const meta = tr.getMeta(outlineKey) as Partial<OutlineState> | undefined;
            let s = v.s;
            if (tr.docChanged) s = { ...s, collapsed: s.collapsed.map((p) => tr.mapping.map(p)) };
            if (meta) s = { ...s, ...meta };
            if (!meta && !tr.docChanged) return v;
            return { s, deco: s.on ? build(state.doc, s) : DecorationSet.empty };
          },
        },
        props: { decorations: (state) => outlineKey.getState(state)?.deco },
      }),
    ];
  },
});

export function outlineState(state: EditorState): OutlineState {
  return outlineKey.getState(state)?.s ?? { on: false, level: 0, collapsed: [], firstLineOnly: false };
}

export function setOutline(editor: Editor, patch: Partial<OutlineState>) {
  editor.view.dispatch(editor.state.tr.setMeta(outlineKey, patch).setMeta("addToHistory", false));
}

/** İmleçteki üst düzey blok. */
function currentBlock(state: EditorState): Block | null {
  const $p = state.selection.$from;
  if ($p.depth < 2) return null;
  const node = $p.node(2);
  return { pos: $p.before(2), node, level: outlineLevelOf(node) };
}

/** Tıklamadan hemen sonra basılan kısayollar için DOM seçimini önce durum seçimine aktarır. */
function sync(editor: Editor) {
  if (editor.view.hasFocus()) syncDomSelection(editor.view);
}

/** Anahat düzeyini ayarlar: 1-6 başlık, 7-9 anahat düzeyli paragraf, 10 gövde metni. */
export function setOutlineLevel(editor: Editor, level: number) {
  sync(editor);
  const b = currentBlock(editor.state);
  if (!b || !b.node.isTextblock) return;
  const { schema } = editor.state;
  const tr = editor.state.tr;
  const keep = { ...b.node.attrs } as Record<string, unknown>;
  delete keep.level;
  if (level <= 6) tr.setNodeMarkup(b.pos, schema.nodes.heading, { ...keep, level, styleId: null, outlineLevel: null });
  else if (level <= 9) tr.setNodeMarkup(b.pos, schema.nodes.paragraph, { ...keep, styleId: null, outlineLevel: level });
  else tr.setNodeMarkup(b.pos, schema.nodes.paragraph, { ...keep, styleId: null, outlineLevel: null });
  editor.view.dispatch(tr);
}

/** Yükselt (-1) / Düşür (+1). */
export function shiftOutline(editor: Editor, dir: -1 | 1) {
  sync(editor);
  const b = currentBlock(editor.state);
  if (!b) return;
  if (b.level === 10) {
    if (dir > 0) return;
    // Gövde metni yükseltilince bir önceki başlığın düzeyine çıkar.
    const prev = blocks(editor.state.doc)
      .filter((x) => x.pos < b.pos && x.level < 10)
      .pop();
    setOutlineLevel(editor, prev?.level ?? 1);
    return;
  }
  setOutlineLevel(editor, Math.max(1, Math.min(9, b.level + dir)));
}

/** Yukarı Taşı / Aşağı Taşı: imleçteki üst düzey bloğu komşusuyla yer değiştirir. */
export function moveBlock(editor: Editor, dir: -1 | 1): boolean {
  sync(editor);
  const { state } = editor;
  const $p = state.selection.$from;
  if ($p.depth < 2) return false;
  const sec = $p.node(1);
  const idx = $p.index(1);
  const other = idx + dir;
  if (other < 0 || other >= sec.childCount) return false;
  const cur = sec.child(idx);
  const start = $p.before(2);
  const offsetIn = state.selection.from - start;
  const tr = state.tr;
  if (dir < 0) {
    const prev = sec.child(other);
    const prevStart = start - prev.nodeSize;
    tr.delete(start, start + cur.nodeSize);
    tr.insert(prevStart, cur);
    tr.setSelection(TextSelection.create(tr.doc, prevStart + offsetIn));
  } else {
    const next = sec.child(other);
    tr.delete(start, start + cur.nodeSize);
    const at = start + next.nodeSize;
    tr.insert(at, cur);
    tr.setSelection(TextSelection.create(tr.doc, at + offsetIn));
  }
  editor.view.dispatch(tr.scrollIntoView());
  return true;
}

/** Genişlet / Daralt: imleçteki başlığın alt öğeleri. */
export function toggleCollapse(editor: Editor, collapse: boolean) {
  sync(editor);
  const b = currentBlock(editor.state);
  if (!b || b.level >= 10) return;
  const s = outlineState(editor.state);
  const collapsed = collapse ? [...new Set([...s.collapsed, b.pos])] : s.collapsed.filter((p) => p !== b.pos);
  setOutline(editor, { collapsed });
}
