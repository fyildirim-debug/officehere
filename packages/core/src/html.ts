// Dışarıdan gelen (yapıştırılan ya da açılan) HTML'in etkisiz belgede ayrıştırılması.

const PARKED = "data-oh-olay-";

/**
 * HTML'i etkisiz bir belgede (DOMParser: betik çalışmaz, resim yüklenmez) ayrıştırır; dönen belgede satır içi
 * olay öznitelikleri (onclick, onerror…) bulunmaz. Chromium 156'dan bu yana bu öznitelikler etkisiz belgede de
 * ayrıştırılırken sayfanın CSP'sine göre denetlenip konsola ihlal yazılıyor (çalıştırılmıyor). Bu yüzden etiketlerdeki
 * olay öznitelikleri ayrıştırmadan önce zararsız bir ada çevrilir, sonra atılır. Metin taraması yalnız gürültüyü
 * önler; güvenlik için çağıran yine DOM üzerinde temizlemelidir (ör. değerinde ">" geçen öznitelikten sonrası).
 */
export function parseHtml(html: string): Document {
  const parked = html.replace(/<[a-zA-Z][^>]*>/g, (tag) => tag.replace(/([\s"'/])(on[a-z]+\s*=)/gi, `$1${PARKED}$2`));
  const doc = new DOMParser().parseFromString(parked, "text/html");
  if (parked !== html)
    for (const el of Array.from(doc.querySelectorAll("*")))
      for (const a of Array.from(el.attributes)) if (a.name.startsWith(PARKED)) el.removeAttribute(a.name);
  return doc;
}
