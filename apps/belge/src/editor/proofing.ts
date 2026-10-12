// Düzenleyici (Gözden Geçir › Yazım ve Dil Bilgisi) denetimleri: çift boşluk, yinelenen sözcük, noktalamadan
// önce boşluk, küçük harfle başlayan cümle. Yazım hataları tarayıcının denetleyicisiyle altı çizili gösterilir.
import type { Node as PMNode } from "@tiptap/pm/model";
import { defaultAutoCorrect, isSentenceStart } from "./autocorrectRules";

export type ProofKind = "doubleSpace" | "repeat" | "spaceBeforePunct" | "lowercaseStart";

export interface ProofIssue {
  from: number;
  to: number;
  kind: ProofKind;
  /** Bölmede gösterilen kategori. */
  title: string;
  /** Sorunlu metin ve önerilen düzeltme (null: düzeltme önerisi yok). */
  text: string;
  fix: string | null;
}

const lowerTr = (s: string) => s.toLocaleLowerCase("tr-TR");

/** Belgedeki dil bilgisi/biçim sorunları (en fazla `limit` tane). */
export function findIssues(doc: PMNode, limit = 300): ProofIssue[] {
  const out: ProofIssue[] = [];
  const opts = { ...defaultAutoCorrect(), capFirstSentence: true };
  doc.descendants((node, pos) => {
    if (out.length >= limit) return false;
    if (!node.isTextblock) return true;
    const text = node.textBetween(0, node.content.size, undefined, (n: PMNode) => (n.type.name === "hardBreak" ? "\n" : "￼"));
    const start = pos + 1;
    const push = (i: number, len: number, kind: ProofKind, title: string, fix: string | null) => {
      if (out.length < limit) out.push({ from: start + i, to: start + i + len, kind, title, text: text.slice(i, i + len), fix });
    };
    for (const m of text.matchAll(/(?<=\S) {2,}(?=\S)/g)) push(m.index!, m[0].length, "doubleSpace", "Boşluk", " ");
    for (const m of text.matchAll(/(?<![\p{L}\d])(\p{L}+)\s+(\p{L}+)(?![\p{L}\d])/gu)) {
      if (lowerTr(m[1]) === lowerTr(m[2]) && m[1].length > 1) push(m.index!, m[0].length, "repeat", "Yinelenen sözcük", m[1]);
    }
    for (const m of text.matchAll(/(?<=\p{L}) +([.,;:!?])(?!\S*\d)/gu)) push(m.index!, m[0].length, "spaceBeforePunct", "Noktalama", m[1]);
    for (const m of text.matchAll(/(?<![\p{L}\d])(\p{Ll})(\p{L}*)/gu)) {
      const prefix = text.slice(0, m.index!);
      if (!prefix.trim()) continue;
      if (isSentenceStart(prefix, false, opts)) push(m.index!, m[0].length, "lowercaseStart", "Büyük harf kullanımı", m[1].toLocaleUpperCase("tr-TR") + m[2]);
    }
    return false;
  });
  return out.sort((a, b) => a.from - b.from);
}
