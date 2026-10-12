// Görünüm › Koyu Sayfa: seçim depolamada hatırlanır ve <html data-bl-page> özniteliğine yansır.
import { describe, expect, it } from "vitest";

const saved = new Map<string, string>([["officehere.belge.darkPage", "1"]]);
(globalThis as unknown as { localStorage: Partial<Storage> }).localStorage = {
  getItem: (k: string) => saved.get(k) ?? null,
  setItem: (k: string, v: string) => void saved.set(k, v),
  removeItem: (k: string) => void saved.delete(k),
};
const attrs = new Map<string, string>();
(globalThis as unknown as { document: unknown }).document = {
  documentElement: { setAttribute: (k: string, v: string) => void attrs.set(k, v), removeAttribute: (k: string) => void attrs.delete(k) },
};

const D = await import("./darkPage");

describe("koyu sayfa", () => {
  it("açılışta kayıtlı seçim uygulanır", () => {
    expect(D.isDarkPage()).toBe(true);
    expect(attrs.get("data-bl-page")).toBe("dark");
  });

  it("kapatılınca öznitelik ve kayıt kalkar, açılınca geri gelir", () => {
    D.setDarkPage(false);
    expect(D.isDarkPage()).toBe(false);
    expect(attrs.has("data-bl-page")).toBe(false);
    expect(saved.has(D.DARK_PAGE_KEY)).toBe(false);
    D.setDarkPage(true);
    expect(D.isDarkPage()).toBe(true);
    expect(saved.get(D.DARK_PAGE_KEY)).toBe("1");
    expect(D.readDarkPage()).toBe(true);
  });
});
