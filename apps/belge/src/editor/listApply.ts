// Madde işareti / numaralandırma uygulama (Word davranışı): birden çok paragrafa ve listeye yayılan
// seçimde hepsi tek bir hedef türde listeye dönüşür; seçimdeki her şey zaten o türdeyse liste kaldırılır.
import type { Editor } from "@tiptap/core";
import type { Node as PMNode, NodeType } from "@tiptap/pm/model";
import type { Transaction } from "@tiptap/pm/state";
import { canJoin } from "@tiptap/pm/transform";
import { wrapRangeInList } from "@tiptap/pm/schema-list";
import { sectionWise } from "./sections";

const isList = (n: PMNode) => n.type.name === "bulletList" || n.type.name === "orderedList";
const isPara = (n: PMNode) => n.type.name === "paragraph";

/**
 * Seçim bir bölümde birden çok üst düzey bloğa yayılıyorsa: diğer türdeki listeler hedef türe çevrilir,
 * ardışık paragraflar listeye sarılır ve bitişik listeler birleştirilir. Uygulanmadıysa false döner.
 */
export function applyListAcross(tr: Transaction, type: NodeType): boolean {
  const { $from, $to } = tr.selection;
  if ($from.depth < 2 || $to.depth < 2 || $from.index(0) !== $to.index(0)) return false;
  const sec = $from.node(1);
  const a = $from.index(1);
  const b = $to.index(1);
  if (a === b) return false;
  let all = true;
  for (let i = a; i <= b; i++) if (sec.child(i).type !== type && (isList(sec.child(i)) || isPara(sec.child(i)))) all = false;
  if (all) return false;
  const secPos = $from.before(1);
  const offs: number[] = [];
  let p = secPos + 1;
  sec.forEach((n) => {
    offs.push(p);
    p += n.nodeSize;
  });
  // Sondan başa: önceki blokların konumları değişmez.
  for (let i = b; i >= a; ) {
    const n = sec.child(i);
    if (isList(n)) {
      if (n.type !== type) tr.setNodeMarkup(offs[i], type);
      i--;
      continue;
    }
    if (!isPara(n)) {
      i--;
      continue;
    }
    let j = i;
    while (j - 1 >= a && isPara(sec.child(j - 1))) j--;
    const range = tr.doc.resolve(offs[j] + 1).blockRange(tr.doc.resolve(offs[i] + n.nodeSize - 1));
    if (range) wrapRangeInList(tr, range, type);
    i = j - 1;
  }
  // Bitişik hedef listeleri birleştir.
  const now = tr.doc.nodeAt(secPos)!;
  let q = secPos + 1;
  const bounds: number[] = [];
  now.forEach((n, _o, k) => {
    if (k > 0 && n.type === type && now.child(k - 1).type === type) bounds.push(q);
    q += n.nodeSize;
  });
  for (let k = bounds.length - 1; k >= 0; k--) if (canJoin(tr.doc, bounds[k])) tr.join(bounds[k]);
  return true;
}

/** Liste aç/kapat (Tümünü Seç ve bölümler arası seçimde bölüm bölüm). */
export function toggleList(editor: Editor, type: "bulletList" | "orderedList"): boolean {
  const nodeType = editor.schema.nodes[type];
  return sectionWise(editor, (c) =>
    c.command(({ tr, commands }) => applyListAcross(tr, nodeType) || (type === "bulletList" ? commands.toggleBulletList() : commands.toggleOrderedList())),
  );
}
