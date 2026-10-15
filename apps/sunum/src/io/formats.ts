/**
 * Sunu dosya biçimleri: uzantı + içerik imzasıyla biçim tanıma, okuyucu ve
 * yazıcı seçimi, aç/kaydet süzgeçleri. Yerel biçim .pptx'tir; .ppt ve
 * makrolu/şablon çeşitleri yalnız açılır.
 *
 * Bu modül hafiftir: biçim bilgisi (formatInfo.ts) doğrudan, ağır okuyucu ve
 * yazıcılar (PPTX/ODP/PPT, JSZip; formatsImpl.ts) ise ilk kullanımda ayrı bir
 * parçadan yüklenir. Kaydetmenin ağdan bağımsız çalışması için uygulama
 * açıldıktan sonra boşta önceden yüklenir (preloadIo). Yükleme başarısız
 * olursa Türkçe bir FormatError atılır ve sonraki çağrı yeniden dener.
 */
import type { Presentation } from "../model/types";
import { FormatError, LOAD_FAILED, type SaveFormat, type SunuFormat } from "./formatInfo";
import type { ReadOptions, ReadResult, WriteOptions } from "./formatsImpl";

export * from "./formatInfo";
export type { ReadOptions, ReadResult, WriteOptions } from "./formatsImpl";

type Impl = typeof import("./formatsImpl");
let pending: Promise<Impl> | undefined;

/** Okuyucu/yazıcı parçası; yüklenemezse Türkçe FormatError. */
export function loadIo(): Promise<Impl> {
  if (!pending) {
    pending = import("./formatsImpl").catch((e: unknown) => {
      pending = undefined;
      throw new FormatError(LOAD_FAILED, String((e as Error)?.message ?? e));
    });
  }
  return pending;
}

/** Okuyucu/yazıcı parçasını boşta önceden yükler (hata yutulur; ilk gerçek kullanımda yeniden denenir). */
export function preloadIo(): void {
  loadIo().catch(() => {});
}

/** Uzantı ve içerik imzasıyla biçimi belirler (içerik önceliklidir). */
export async function sniffFormat(data: ArrayBuffer | Uint8Array, name: string): Promise<SunuFormat> {
  return (await loadIo()).sniffFormat(data, name);
}

/** Biçimi tanıyıp uygun okuyucuyla açar. Hatalar Türkçe FormatError olarak atılır; uyarılar `warnings`'te. */
export async function readPresentationFile(data: ArrayBuffer | Uint8Array, name: string, opts?: ReadOptions): Promise<ReadResult> {
  return (await loadIo()).readPresentationFile(data, name, opts);
}

/** Sunuyu seçilen biçimde yazar; kaydetme uyarıları `opts.warnings`'e eklenir. */
export async function writePresentationFormat(pres: Presentation, format: SaveFormat, opts?: WriteOptions): Promise<Uint8Array> {
  return (await loadIo()).writePresentationFormat(pres, format, opts);
}
