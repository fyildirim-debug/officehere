/**
 * Slayt açıklamaları (p:cmLst) ve açıklama yazarları (commentAuthors.xml).
 * Geçiş kancası (dalga 3): modern açıklamalar (p188:cmLst, authors.xml) ve
 * açıklama yanıtları burada okunmuyor; taşınacaksa bu modüle eklenecek.
 */
import { uid } from "../../model/factory";
import type { Comment } from "../../model/types";
import { attr, kid, kids, numAttr, parseXml, REL, type Rel } from "../xml";
import type { PptxPackage } from "./package";
import type { PresPart } from "./presentation";

export type CommentAuthors = () => Promise<Map<string, string>>;

/** Açıklama yazarları (kimlik → ad); yalnız gerektiğinde, bir kez okunur. */
export function commentAuthors(pres: PresPart): CommentAuthors {
  let cache: Map<string, string> | undefined;
  return async () => {
    if (cache) return cache;
    cache = new Map();
    const path = pres.partOf(REL.commentAuthors, /^ppt\/commentAuthors\.xml$/i);
    const xml = path ? await pres.pk.text(path) : undefined;
    if (xml) for (const a of kids(parseXml(xml).documentElement, "cmAuthor")) cache.set(attr(a, "id") ?? "", attr(a, "name") ?? "Yazar");
    return cache;
  };
}

/** Slaydın açıklamaları (`n`: 1 tabanlı slayt numarası); yoksa tanımsız. */
export async function readSlideComments(pk: PptxPackage, rels: Map<string, Rel>, n: number, authors: CommentAuthors): Promise<Comment[] | undefined> {
  const crel = [...rels.values()].find((r) => r.type === REL.comments);
  if (!crel) return undefined;
  const cPart = await pk.part(crel.target, "cmLst");
  let list: Comment[] | undefined;
  if (cPart.root) {
    const names = await authors();
    list = kids(cPart.root, "cm").map((cm) => readComment(cm, names));
  }
  if (cPart.problem) pk.W.slide("comments", n, cPart.detail);
  return list?.length ? list : undefined;
}

function readComment(cm: Element, names: Map<string, string>): Comment {
  const pos = kid(cm, "pos");
  const dt = Date.parse(attr(cm, "dt") ?? "");
  return {
    id: uid("c"),
    author: names.get(attr(cm, "authorId") ?? "") ?? "Yazar",
    text: kid(cm, "text")?.textContent ?? "",
    x: (numAttr(pos, "x") ?? 0) / 6,
    y: (numAttr(pos, "y") ?? 0) / 6,
    date: Number.isFinite(dt) ? dt : Date.now(),
  };
}
