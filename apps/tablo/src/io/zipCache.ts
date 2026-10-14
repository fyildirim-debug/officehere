/**
 * Okuma sırasında paket önbelleği. Bir XLSX dosyası açılırken aynı paket
 * birçok okuyucu (notlar, grafikler, filtreler, sayfa sonları, mini grafikler…)
 * tarafından ayrı ayrı açılıyordu; büyük bir sayfanın XML'i her seferinde
 * yeniden açılıyor (200 bin satırda ~1 sn/kez) ve süreyi katlıyordu. Burada
 * ZIP dizini ve açılmış metinler veri başına bir kez çözülür.
 * Önbellek yalnız okuma içindir: paketi değiştiren işlevler JSZip'i doğrudan kullanır.
 */
import JSZip from "jszip";

const cache = new WeakMap<ArrayBuffer, Promise<JSZip>>();
const MEMO = Symbol("ohMemo");

/** Paketteki parçaların metin olarak açılmasını (async("string")) belleğe alır. */
export function memoizeText(zip: JSZip): JSZip {
  for (const f of Object.values(zip.files)) {
    const obj = f as unknown as Record<symbol, boolean> & { async: (type: string, cb?: unknown) => Promise<unknown> };
    if (f.dir || obj[MEMO]) continue;
    obj[MEMO] = true;
    const orig = obj.async.bind(f);
    let text: Promise<unknown> | undefined;
    obj.async = (type: string, cb?: unknown) => {
      if (type !== "string" || cb) return orig(type, cb);
      if (!text) {
        text = orig("string");
        // Hata önbelleğe alınmaz (sonraki deneme yeniden açar)
        text.catch(() => (text = undefined));
      }
      return text;
    };
  }
  return zip;
}

/** Verinin ZIP paketini (önbellekten) döndürür. */
export function loadZip(data: ArrayBuffer | Uint8Array): Promise<JSZip> {
  if (!(data instanceof ArrayBuffer)) return JSZip.loadAsync(data).then(memoizeText);
  let z = cache.get(data);
  if (!z) {
    z = JSZip.loadAsync(data).then(memoizeText);
    cache.set(data, z);
    z.catch(() => cache.delete(data));
  }
  return z;
}

/** Önceden açılmış (ve gerekirse onarılmış) paketi bu veri için önbelleğe koyar. */
export function seedZip(data: ArrayBuffer, zip: JSZip): void {
  cache.set(data, Promise.resolve(memoizeText(zip)));
}

/** Okuma bitince önbelleği boşaltır (açılmış büyük metinler bellekte kalmasın). */
export function releaseZip(data: ArrayBuffer): void {
  cache.delete(data);
}
