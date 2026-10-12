/**
 * Dönüştür geçişi için şekil eşleştirme (nesnelere göre): önce aynı adlı
 * şekiller ("!!" önekli adlar her zaman), sonra aynı metinli şekiller, en
 * son aynı türde/geometride ve aynı resim kaynağına sahip şekiller eşlenir.
 */
import type { Shape } from "./types";

function textOf(s: Shape): string {
  if (s.kind !== "sp" || !s.text) return "";
  return s.text.paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join("\n").trim();
}

function signature(s: Shape): string {
  if (s.kind === "sp") return `sp:${s.geom}`;
  if (s.kind === "pic") return `pic:${s.src.length}:${s.src.slice(-64)}`;
  return s.kind;
}

/** Eski ve yeni slayt şekillerini eşleştirir; [eski, yeni] çiftleri döner. */
export function matchShapes(oldShapes: Shape[], newShapes: Shape[]): [Shape, Shape][] {
  const pairs: [Shape, Shape][] = [];
  const usedOld = new Set<Shape>();
  const usedNew = new Set<Shape>();
  const take = (o: Shape, n: Shape) => {
    pairs.push([o, n]);
    usedOld.add(o);
    usedNew.add(n);
  };
  const count = (list: Shape[], key: (s: Shape) => string, k: string) => list.filter((s) => key(s) === k).length;
  const pass = (key: (s: Shape) => string, unique: boolean) => {
    for (const n of newShapes) {
      if (usedNew.has(n)) continue;
      const k = key(n);
      if (!k) continue;
      if (unique && (count(oldShapes, key, k) !== 1 || count(newShapes, key, k) !== 1)) continue;
      const o = oldShapes.find((x) => !usedOld.has(x) && key(x) === k);
      if (o) take(o, n);
    }
  };
  // 1) "!!" ile başlayan adlar her zaman; 2) benzersiz adlar; 3) aynı metin; 4) aynı tür + geometri.
  pass((s) => (s.name.startsWith("!!") ? s.name : ""), false);
  pass((s) => s.name, true);
  pass((s) => textOf(s), false);
  pass((s) => (textOf(s) ? "" : signature(s)), false);
  return pairs;
}
