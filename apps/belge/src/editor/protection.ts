// Düzenlemeyi Kısıtla (Gözden Geçir › Koru): salt okunur, yalnız açıklamalar ya da yalnız izlenen
// değişiklikler. Zorlama açıkken belge düzenlenemez (açıklama kipinde yalnız açıklama işaretleri
// eklenip kaldırılabilir; izlenen değişiklik kipinde değişiklikleri izleme kapatılamaz).
// İsteğe bağlı parola ISO/IEC 29500 parola karması (SHA-512, tuz, yineleme) olarak saklanır.
import { Extension } from "@tiptap/core";
import { Plugin, PluginKey, type Transaction } from "@tiptap/pm/state";
import { AddMarkStep, RemoveMarkStep } from "@tiptap/pm/transform";
import type { DocProtection } from "../model/document";
import { trackKey } from "./trackChanges";

export const PROTECTION_LABEL: Record<DocProtection["mode"], string> = {
  readOnly: "Değişiklik yok (Salt okunur)",
  comments: "Açıklamalar",
  trackedChanges: "İzlenen değişiklikler",
};

/** Geçerli koruma (App her çizimde günceller). */
export const protectionSource: { get: () => DocProtection | null } = { get: () => null };

/** Koruma zorlanıyor ve belge içeriği düzenlenemez mi. */
export function isLocked(p: DocProtection | null | undefined): boolean {
  return !!p?.enforced && p.mode !== "trackedChanges";
}

/** Korumanın izin verdiği işlem mi. */
export function allowedUnder(p: DocProtection | null | undefined, tr: Transaction): boolean {
  if (!p?.enforced) return true;
  if (p.mode === "trackedChanges") {
    // Değişiklikleri izleme kapatılamaz.
    const meta = tr.getMeta(trackKey) as { on?: boolean } | undefined;
    return !(meta && meta.on === false);
  }
  if (!tr.docChanged) return true;
  if (p.mode === "comments") return tr.steps.every((s) => (s instanceof AddMarkStep || s instanceof RemoveMarkStep) && s.mark.type.name === "comment");
  return false;
}

export const Protection = Extension.create({
  name: "belgeProtection",
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: new PluginKey("belgeProtection"),
        filterTransaction: (tr) => allowedUnder(protectionSource.get(), tr),
        // İmleç, seçim ve kısayollar çalışır; içerik değiştiren tarayıcı girdisi baştan engellenir
        // (DOM ile belge durumu ayrışmasın diye).
        props: {
          handleDOMEvents: {
            beforeinput: (_view, e) => {
              if (!isLocked(protectionSource.get())) return false;
              e.preventDefault();
              return true;
            },
          },
          handleTextInput: () => isLocked(protectionSource.get()),
          handlePaste: () => isLocked(protectionSource.get()),
          handleDrop: () => isLocked(protectionSource.get()),
          attributes: (): Record<string, string> => (isLocked(protectionSource.get()) ? { "data-protected": "" } : {}),
        },
      }),
    ];
  },
});

/* ------------------------------------------------------------------ */
/* Parola karması                                                      */
/* ------------------------------------------------------------------ */

const b64 = (bytes: Uint8Array) => {
  let s = "";
  for (const b of bytes) s += String.fromCharCode(b);
  return btoa(s);
};
const unb64 = (s: string) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));

/** H0 = SHA-512(tuz + parola[UTF-16LE]); Hn = SHA-512(Hn-1 + n[32 bit LE]). */
export async function hashPassword(password: string, saltB64: string, spinCount = 100000): Promise<string> {
  const salt = unb64(saltB64);
  const pw = new Uint8Array(password.length * 2);
  for (let i = 0; i < password.length; i++) {
    const c = password.charCodeAt(i);
    pw[i * 2] = c & 0xff;
    pw[i * 2 + 1] = c >> 8;
  }
  const first = new Uint8Array(salt.length + pw.length);
  first.set(salt);
  first.set(pw, salt.length);
  let h = new Uint8Array(await crypto.subtle.digest("SHA-512", first));
  const buf = new Uint8Array(h.length + 4);
  for (let i = 0; i < spinCount; i++) {
    buf.set(h);
    buf[h.length] = i & 0xff;
    buf[h.length + 1] = (i >> 8) & 0xff;
    buf[h.length + 2] = (i >> 16) & 0xff;
    buf[h.length + 3] = (i >>> 24) & 0xff;
    h = new Uint8Array(await crypto.subtle.digest("SHA-512", buf));
  }
  return b64(h);
}

/** Yeni koruma (parola varsa karması ile). */
export async function makeProtection(mode: DocProtection["mode"], password: string): Promise<DocProtection> {
  if (!password) return { mode, enforced: true };
  const salt = b64(crypto.getRandomValues(new Uint8Array(16)));
  const spinCount = 100000;
  return { mode, enforced: true, algorithm: "SHA-512", salt, spinCount, hash: await hashPassword(password, salt, spinCount) };
}

/** Korumayı durdurmak için parola doğru mu (parolasızsa her zaman). */
export async function checkPassword(p: DocProtection, password: string): Promise<boolean> {
  if (!p.hash || !p.salt) return true;
  if (p.algorithm && p.algorithm.toUpperCase().replace("-", "") !== "SHA512") return false;
  return (await hashPassword(password, p.salt, p.spinCount ?? 100000)) === p.hash;
}
