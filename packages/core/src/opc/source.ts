// Kaynak paket kaydı: açılan dosyanın özgün paketi bellekte, kimlikle tutulur.
// Model yalnız kimliği taşır ("kaynak-3"), bayt taşımaz; böylece geri al yığını ve structuredClone
// maliyeti değişmez. Kaydederken kimlikten paket bulunur ve taşınacak parçalar oradan kopyalanır.
// Kalıcılık (taslak/kurtarma için IndexedDB) uygulamanın işidir; anahtar olarak `sha256Hex` önerilir.

import { OpcPackage } from "./package";

export interface PackageSource {
  readonly id: string;
  readonly pkg: OpcPackage;
  /** Görünen dosya adı (tanılama ve uyarılar için). */
  readonly name?: string;
  /** Kayıt zamanı (ms). */
  readonly at: number;
}

/** Bellek içi kaynak kaydı. Uygulama başına bir tane yeter (`packageSources`); testler ve Worker kendi örneğini kurabilir. */
export class SourceRegistry {
  private map = new Map<string, PackageSource>();
  private seq = 0;

  /**
   * Paketi kaydeder, kimliğini döner. `id` verilirse (ör. özgün baytların SHA-256'sı) o kullanılır ve
   * aynı kimlikli önceki kaydın yerine geçer; verilmezse "kaynak-N" üretilir.
   */
  register(pkg: OpcPackage, opts: { id?: string; name?: string } = {}): string {
    let id = opts.id;
    if (!id) {
      do id = `kaynak-${++this.seq}`;
      while (this.map.has(id));
    }
    this.map.set(id, { id, pkg, name: opts.name, at: Date.now() });
    return id;
  }

  get(id: string | null | undefined): OpcPackage | undefined {
    return id ? this.map.get(id)?.pkg : undefined;
  }

  entry(id: string | null | undefined): PackageSource | undefined {
    return id ? this.map.get(id) : undefined;
  }

  has(id: string | null | undefined): boolean {
    return !!id && this.map.has(id);
  }

  /** Kaydı bırakır (belge kapatılınca). Vardıysa true. */
  release(id: string | null | undefined): boolean {
    return !!id && this.map.delete(id);
  }

  ids(): string[] {
    return [...this.map.keys()];
  }

  clear(): void {
    this.map.clear();
  }

  /** Kayıtlı paketlerin toplam bayt boyutu (bellek göstergesi). */
  get byteSize(): number {
    let n = 0;
    for (const s of this.map.values()) n += s.pkg.bytes.byteLength;
    return n;
  }
}

/** Uygulamanın varsayılan kaynak kaydı. */
export const packageSources = new SourceRegistry();

/** Paketi varsayılan kayda ekler; kimliği döner. */
export function registerSource(pkg: OpcPackage, opts?: { id?: string; name?: string }): string {
  return packageSources.register(pkg, opts);
}

export function getSource(id: string | null | undefined): OpcPackage | undefined {
  return packageSources.get(id);
}

export function releaseSource(id: string | null | undefined): boolean {
  return packageSources.release(id);
}

/** Baytları açar ve varsayılan kayda ekler: { id, pkg }. ZIP değilse OpcError. */
export async function openSource(data: Uint8Array | ArrayBuffer, opts: { id?: string; name?: string; maxPartSize?: number } = {}): Promise<PackageSource> {
  const pkg = await OpcPackage.open(data, { maxPartSize: opts.maxPartSize });
  const id = packageSources.register(pkg, { id: opts.id, name: opts.name });
  return packageSources.entry(id)!;
}

/** Baytların SHA-256 özeti (küçük harf onaltılık); kalıcı kaynak deposunun anahtarı için. */
export async function sha256Hex(data: Uint8Array | ArrayBuffer): Promise<string> {
  const buf = data instanceof Uint8Array ? data : new Uint8Array(data);
  const d = new Uint8Array(await globalThis.crypto.subtle.digest("SHA-256", buf as BufferSource));
  let s = "";
  for (const b of d) s += b.toString(16).padStart(2, "0");
  return s;
}
