/**
 * Parola karmaları (ECMA-376 Bölüm 1 §18.2.29 workbookProtection):
 * - Güncel yöntem: H0 = SHA-512(tuz + parola[UTF-16LE]), Hn = SHA-512(Hn-1 + n[uint32 LE]),
 *   spinCount kez; sonuç ve tuz base64.
 * - Eski yöntem (workbookPassword): 16 bit XOR karması (onaltılık).
 * Yapı korumasının denetimleri (sayfa ekle/sil/adlandır/taşı/gizle) burada.
 */
import type { WorkbookModel } from "./workbook";
import type { WorkbookProtection } from "./types";

export const DEFAULT_SPIN = 100000;

function b64(bytes: Uint8Array): string {
  let s = "";
  for (const b of bytes) s += String.fromCharCode(b);
  return btoa(s);
}

function unb64(s: string): Uint8Array {
  const bin = atob(s);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

function utf16le(s: string): Uint8Array {
  const out = new Uint8Array(s.length * 2);
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    out[i * 2] = c & 0xff;
    out[i * 2 + 1] = c >> 8;
  }
  return out;
}

const ALGOS: Record<string, string> = { "SHA-512": "SHA-512", "SHA-384": "SHA-384", "SHA-256": "SHA-256", "SHA-1": "SHA-1" };

async function digest(algo: string, data: Uint8Array): Promise<Uint8Array> {
  return new Uint8Array(await crypto.subtle.digest(algo, data as BufferSource));
}

/** Parolanın karması (verilen tuzla ya da yeni rastgele tuzla). */
export async function hashPassword(password: string, salt?: string, spin = DEFAULT_SPIN, algorithm = "SHA-512"): Promise<NonNullable<WorkbookProtection["hash"]>> {
  const algo = ALGOS[algorithm] ?? "SHA-512";
  const saltBytes = salt ? unb64(salt) : crypto.getRandomValues(new Uint8Array(16));
  const pw = utf16le(password);
  const first = new Uint8Array(saltBytes.length + pw.length);
  first.set(saltBytes);
  first.set(pw, saltBytes.length);
  let h = await digest(algo, first);
  const buf = new Uint8Array(h.length + 4);
  for (let i = 0; i < spin; i++) {
    buf.set(h);
    buf[h.length] = i & 0xff;
    buf[h.length + 1] = (i >>> 8) & 0xff;
    buf[h.length + 2] = (i >>> 16) & 0xff;
    buf[h.length + 3] = (i >>> 24) & 0xff;
    h = await digest(algo, buf);
  }
  return { algorithm: algorithm in ALGOS ? algorithm : "SHA-512", value: b64(h), salt: b64(saltBytes), spin };
}

/** Eski 16 bit parola karması (onaltılık, büyük harf). */
export function legacyPasswordHash(password: string): string {
  let hash = 0;
  if (password.length) {
    for (let i = password.length - 1; i >= 0; i--) {
      hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
      hash ^= password.charCodeAt(i);
    }
    hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
    hash ^= password.length;
    hash ^= 0xce4b;
  }
  return hash.toString(16).toUpperCase().padStart(4, "0");
}

/** Koruma parolası doğru mu (parolasız korumada her zaman doğru). */
export async function checkProtectionPassword(p: WorkbookProtection, password: string): Promise<boolean> {
  if (p.hash) {
    const h = await hashPassword(password, p.hash.salt, p.hash.spin, p.hash.algorithm);
    return h.value === p.hash.value;
  }
  if (p.legacy) return legacyPasswordHash(password) === p.legacy.toUpperCase().padStart(4, "0");
  return true;
}

export const STRUCTURE_MESSAGE = "Çalışma kitabı korunduğu için yapısı değiştirilemez. Sayfa eklemek, silmek, taşımak, gizlemek ya da yeniden adlandırmak için Gözden Geçir > Çalışma Kitabını Koru ile korumayı kaldırın.";

/** Çalışma kitabının yapısı korumalı mı? */
export function structureLocked(wb: WorkbookModel): boolean {
  return !!wb.protection?.structure;
}
