/** Tablo adı doğrulaması (tanımlı ad kurallarıyla aynı; ek olarak boşluk ve çakışma uyarıları). */
import type { WorkbookModel } from "./workbook";

const up = (s: string) => s.toLocaleUpperCase("tr-TR");

/** Ad bir hücre başvurusu gibi okunabiliyor mu (A1, XFD1048576, R1C1, R, C). */
export function looksLikeCellRef(name: string): boolean {
  const n = name.toUpperCase();
  if (/^[A-Z]{1,3}\d{1,7}$/.test(n)) {
    const letters = /^[A-Z]+/.exec(n)![0];
    const col = [...letters].reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
    return col <= 16384 && Number(n.slice(letters.length)) <= 1048576;
  }
  return /^(R\d*C\d*|R\d*|C\d*)$/.test(n);
}

/** Geçerli değilse kullanıcıya gösterilecek ileti, geçerliyse null. */
export function tableNameError(wb: WorkbookModel, raw: string, exceptTableId?: string): string | null {
  const n = raw.trim();
  if (!n) return "Tablo adı boş bırakılamaz.";
  if (n.length > 255) return "Tablo adı en çok 255 karakter olabilir.";
  if (/\s/.test(n)) return "Tablo adında boşluk kullanılamaz. Sözcükleri ayırmak için alt çizgi (_) ya da nokta (.) kullanın.";
  if (!/^[\p{L}_\\][\p{L}\p{N}_.\\]*$/u.test(n)) return "Tablo adı bir harf, alt çizgi (_) ya da ters eğik çizgiyle (\\) başlamalı; yalnız harf, rakam, nokta ve alt çizgi içerebilir.";
  if (looksLikeCellRef(n)) return `“${n}” bir hücre başvurusuyla karıştırılabilir. Başka bir ad girin.`;
  const taken = wb.sheets.some((s) => s.tables.some((t) => t.id !== exceptTableId && up(t.name) === up(n))) || wb.engine.getNames().some((x) => up(x.name) === up(n));
  if (taken) return `“${n}” adı bu çalışma kitabında zaten kullanılıyor. Benzersiz bir ad girin.`;
  return null;
}
