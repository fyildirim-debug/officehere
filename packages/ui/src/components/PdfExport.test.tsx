// @vitest-environment jsdom
/** PDF dışa aktarma seçenekleri: sayfa aralığı ayrıştırma, iletişim kutusunun seçimi, iptal ve hatırlanan seçenekler. */
import { act } from "react";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { askPdfExport, parsePdfPageRange, type PdfExportChoice } from "./PdfExport";

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

describe("parsePdfPageRange", () => {
  it("aralıkları Chromium biçimine getirir", () => {
    expect(parsePdfPageRange("1-3, 5")).toEqual({ ranges: "1-3, 5" });
    expect(parsePdfPageRange(" 2 ;4 ")).toEqual({ ranges: "2, 4" });
    expect(parsePdfPageRange("1–3")).toEqual({ ranges: "1-3" });
    expect(parsePdfPageRange("3-3")).toEqual({ ranges: "3" });
  });

  it("sayfa sayısı biliniyorsa sona kadar kabul eder ve sınırlar", () => {
    expect(parsePdfPageRange("3-", 10)).toEqual({ ranges: "3-10" });
    expect(parsePdfPageRange("2-20", 10)).toEqual({ ranges: "2-10" });
    expect(parsePdfPageRange("12", 10)).toEqual({ error: 'Belgede 10 sayfa var; "12" belgenin dışında.' });
  });

  it("geçersiz girdide Türkçe hata verir", () => {
    expect(parsePdfPageRange("")).toEqual({ error: "Sayfa aralığını girin (ör. 1-3, 5)." });
    expect(parsePdfPageRange("abc")).toEqual({ error: 'Geçersiz sayfa aralığı: "abc"' });
    expect(parsePdfPageRange("5-2")).toEqual({ error: 'Geçersiz sayfa aralığı: "5-2"' });
    expect(parsePdfPageRange("0")).toEqual({ error: 'Geçersiz sayfa aralığı: "0"' });
    expect(parsePdfPageRange("3-")).toEqual({ error: '"3-" için son sayfayı da yazın.' });
  });
});

describe("askPdfExport", () => {
  beforeEach(() => localStorage.clear());
  afterEach(() => {
    document.body.innerHTML = "";
  });

  const dialog = () => document.querySelector<HTMLElement>("[role=dialog]")!;
  const radio = (name: string) => dialog().querySelector<HTMLInputElement>(`input[type=radio][aria-label="${name}"]`)!;
  const button = (text: string) => [...dialog().querySelectorAll("button")].find((b) => b.textContent === text)!;
  const input = (label: string) => dialog().querySelector<HTMLInputElement>(`input[aria-label="${label}"]`)!;
  const setValue = (el: HTMLInputElement | HTMLSelectElement, v: string) => {
    const proto = el instanceof HTMLSelectElement ? HTMLSelectElement.prototype : HTMLInputElement.prototype;
    Object.getOwnPropertyDescriptor(proto, "value")!.set!.call(el, v);
    el.dispatchEvent(new Event(el instanceof HTMLSelectElement ? "change" : "input", { bubbles: true }));
  };
  /** Kutuyu açar; söz bir nesne içinde döner (async işlevden söz döndürmek onu beklerdi). */
  const open = async (defaults: Parameters<typeof askPdfExport>[0]) => {
    let p!: Promise<PdfExportChoice | null>;
    await act(async () => {
      p = askPdfExport(defaults);
    });
    return { p };
  };
  const settle = async <T,>(p: Promise<T>) => {
    let v!: T;
    await act(async () => {
      v = await p;
    });
    return v;
  };

  it("seçilen biçim, aralık, bilgiler ve sonrası döner; kutu belgeden kalkar", async () => {
    const { p } = await open({ title: "Rapor", author: "Ayşe", pageCount: 12, headings: true });
    expect(dialog().getAttribute("aria-label")).toBe("PDF Olarak Dışa Aktar");
    expect(dialog().textContent).toContain("Tümü (12 sayfa)");
    expect(dialog().textContent).toContain("e-imza uygulamanızla");
    expect(input("Sayfa aralığı")).toBeTruthy();
    act(() => radio("PDF/A-2b (arşiv)").click());
    act(() => setValue(input("Sayfa aralığı"), "2-4, 9"));
    act(() => setValue(dialog().querySelector("select")!, "folder"));
    act(() => button("Dışa aktar").click());
    const choice = await settle(p);
    expect(choice).toEqual({ standard: "pdfa-2b", pageRanges: "2-4, 9", outline: true, tagged: true, after: "folder", title: "Rapor", author: "Ayşe", subject: undefined, keywords: undefined });
    expect(document.querySelector("[role=dialog]")).toBeNull();
    expect(document.querySelector(".oh-pdf-export-host")).toBeNull();
    // Biçim ve sonrası bir sonraki açılışta hatırlanır.
    const { p: p2 } = await open({ title: "İkinci" });
    expect(radio("PDF/A-2b (arşiv)").checked).toBe(true);
    expect(dialog().querySelector("select")!.value).toBe("folder");
    // Başlığı olmayan belgede yer imi seçeneği yok.
    expect(dialog().textContent).not.toContain("Başlıklardan yer imleri");
    act(() => button("İptal").click());
    expect(await settle(p2)).toBeNull();
  });

  it("geçersiz aralıkta Dışa aktar kapalıdır; kutu açıkken ikinci istek yok sayılır", async () => {
    const { p } = await open({ title: "Rapor", pageCount: 3 });
    expect(await askPdfExport({ title: "ikinci" })).toBeNull();
    act(() => setValue(input("Sayfa aralığı"), "7"));
    expect(dialog().querySelector("[role=alert]")!.textContent).toBe('Belgede 3 sayfa var; "7" belgenin dışında.');
    expect(button("Dışa aktar").disabled).toBe(true);
    act(() => setValue(input("Sayfa aralığı"), "2-3"));
    expect(button("Dışa aktar").disabled).toBe(false);
    act(() => dialog().dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true })));
    expect(await settle(p)).toBeNull();
  });
});
