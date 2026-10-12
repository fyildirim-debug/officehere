// Bölüm başına üst/alt bilgi ("Öncekine Bağla"): ilk bölümün üst/alt bilgileri belge ayarlarındadır;
// sonraki bölümler varsayılan olarak öncekine bağlıdır. Bağı kaldırılan bölümün kendi içeriği,
// bölümün kalıcı kimliğiyle (section.attrs.sid) ayarlarda saklanır (OOXML: bölümün kendi headerReference'ı).
import type { Editor, JSONContent } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import type { DocSettings, HfVariant } from "../model/document";

export type HfKey = "header" | "footer" | "headerFirst" | "footerFirst" | "headerEven" | "footerEven";
export const HF_KEYS: HfKey[] = ["header", "footer", "headerFirst", "footerFirst", "headerEven", "footerEven"];

export function hfKeyOf(which: "header" | "footer", v: HfVariant): HfKey {
  if (v === "first") return which === "header" ? "headerFirst" : "footerFirst";
  if (v === "even") return which === "header" ? "headerEven" : "footerEven";
  return which;
}

type HfSettings = Pick<DocSettings, HfKey | "sectionHf">;

/** Bölümlerin kalıcı kimlikleri (sırayla). */
export function sectionIds(doc: PMNode): (string | null)[] {
  const out: (string | null)[] = [];
  doc.forEach((sec) => out.push((sec.attrs.sid as string | null) ?? null));
  return out;
}

const own = (s: HfSettings, sid: string | null, key: HfKey) => !!sid && !!s.sectionHf?.[sid] && Object.prototype.hasOwnProperty.call(s.sectionHf[sid], key);

/**
 * İçeriğin sahibi bölüm: kendisi ya da bağlı olduğu en yakın önceki bölüm; hiçbirinin kendi içeriği yoksa
 * -1 (belge ayarlarındaki üst/alt bilgi). İlk bölüm de (bölüm sonu silinip sonraki bölümün özellikleri
 * geçtiyse) kendi içeriğine sahip olabilir.
 */
export function hfOwner(ids: (string | null)[], s: HfSettings, section: number, key: HfKey): number {
  for (let j = Math.min(section, ids.length - 1); j >= 0; j--) if (own(s, ids[j], key)) return j;
  return -1;
}

/** Bölümün etkin üst/alt bilgisi. */
export function hfContent(ids: (string | null)[], s: HfSettings, section: number, key: HfKey): JSONContent | null {
  const j = hfOwner(ids, s, section, key);
  if (j < 0) return s[key];
  return s.sectionHf?.[ids[j]!]?.[key] ?? null;
}

/** Bölüm öncekine bağlı mı (ilk bölüm hiçbir zaman). */
export function hfLinked(ids: (string | null)[], s: HfSettings, section: number, key: HfKey): boolean {
  return section > 0 && !own(s, ids[section] ?? null, key);
}

/** Düzenlenen içerik sahibine yazılır (bağlı bölümde düzenleme öncekini de değiştirir; Word gibi). */
export function withHfContent<T extends HfSettings>(s: T, ids: (string | null)[], section: number, key: HfKey, json: JSONContent | null): T {
  const j = hfOwner(ids, s, section, key);
  if (j < 0) return { ...s, [key]: json };
  const sid = ids[j]!;
  return { ...s, sectionHf: { ...s.sectionHf, [sid]: { ...s.sectionHf?.[sid], [key]: json } } };
}

let counter = 0;
const newSid = () => `s${Date.now().toString(36)}${(counter++).toString(36)}`;

/** Bölüme kalıcı kimlik verir (yoksa); kimliği döner. */
export function ensureSectionId(editor: Editor, section: number): string | null {
  const { doc } = editor.state;
  if (section <= 0 || section >= doc.childCount) return null;
  let pos = 0;
  for (let i = 0; i < section; i++) pos += doc.child(i).nodeSize;
  const node = doc.child(section);
  const cur = node.attrs.sid as string | null;
  // Kopyalanmış bölüm aynı kimliği taşıyorsa yenisi verilir.
  const dup = cur && sectionIds(doc).filter((x) => x === cur).length > 1;
  if (cur && !dup) return cur;
  const sid = newSid();
  editor.view.dispatch(editor.state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, sid }));
  return sid;
}

/** Öncekine Bağla aç/kapat. Bağ kaldırılınca bölüm önceki içeriğin bir kopyasıyla başlar. */
export function setHfLink(editor: Editor, updateSettings: (fn: (s: DocSettings) => DocSettings) => void, settings: DocSettings, section: number, key: HfKey, link: boolean) {
  if (section <= 0) return;
  if (link) {
    const sid = (editor.state.doc.child(section)?.attrs.sid as string | null) ?? null;
    if (!sid) return;
    updateSettings((s) => {
      const rest = { ...s.sectionHf?.[sid] };
      delete rest[key];
      return { ...s, sectionHf: { ...s.sectionHf, [sid]: rest } };
    });
    return;
  }
  const copy = hfContent(sectionIds(editor.state.doc), settings, section, key);
  const sid = ensureSectionId(editor, section);
  if (!sid) return;
  updateSettings((s) => ({ ...s, sectionHf: { ...s.sectionHf, [sid]: { ...s.sectionHf?.[sid], [key]: copy ? (JSON.parse(JSON.stringify(copy)) as JSONContent) : null } } }));
}
