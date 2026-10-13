// @vitest-environment jsdom
/** Kaydetme davranışı: eski/yazılamayan biçimler, ODP sorusu, Farklı Kaydet'te biçim seçimi (masaüstü köprüsü taklidiyle). */
import { readFileSync } from "node:fs";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { sniffFormat } from "../io/formats";
import { appSession } from "./appState";
import { currentFormat, openBytes, save } from "./fileOps";

const FIX = `${process.cwd()}/src/io/__fixtures__`;
const buf = (n: string) => {
  const b = readFileSync(`${FIX}/${n}`);
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer;
};

interface Written {
  path: string;
  data: Uint8Array;
}

function mockNative(saveAsExt?: string) {
  const written: Written[] = [];
  const dialogs: { name: string; filters: { name: string; extensions: string[] }[] }[] = [];
  const native = {
    saveFileAs: vi.fn(async (name: string, filters: { name: string; extensions: string[] }[], data: ArrayBuffer) => {
      dialogs.push({ name, filters });
      const path = `/belgeler/${saveAsExt ? name.replace(/\.[^.]+$/, `.${saveAsExt}`) : name}`;
      written.push({ path, data: new Uint8Array(data) });
      return { name: path.split("/").pop()!, path };
    }),
    writeFile: vi.fn(async (path: string, data: ArrayBuffer) => {
      written.push({ path, data: new Uint8Array(data) });
      return true;
    }),
    addRecent: vi.fn(async () => {}),
    getRecent: vi.fn(async () => []),
  };
  (window as unknown as { officehereNative: unknown }).officehereNative = native;
  return { native, written, dialogs };
}

describe("Kaydetme davranışı", () => {
  beforeEach(() => vi.useFakeTimers({ toFake: ["setTimeout"] }));
  afterEach(() => {
    delete (window as unknown as { officehereNative?: unknown }).officehereNative;
    vi.useRealTimers();
  });

  it(".ppt açıldıktan sonra Kaydet: .pptx olarak farklı kaydeder, özgün dosyaya yazmaz", async () => {
    const { written, dialogs } = mockNative();
    const { actions: a } = appSession();
    expect(await openBytes(a, buf("lo-bicim.ppt"), "lo-bicim.ppt", { name: "lo-bicim.ppt", path: "/belgeler/lo-bicim.ppt" })).toBe(true);
    expect(a.s.message?.text).toContain("uyumluluk modu");
    expect(currentFormat(a.s)).toBe("ppt");
    expect(await save(a)).toBe(true);
    expect(dialogs[0].name).toBe("lo-bicim.pptx");
    expect(dialogs[0].filters.map((f) => f.extensions[0])).toEqual(["pptx", "odp", "ppsx", "potx"]);
    expect(written.every((w) => !w.path.endsWith(".ppt"))).toBe(true);
    expect(await sniffFormat(written[0].data, written[0].path)).toBe("pptx");
    expect(a.s.fileRef?.path).toBe("/belgeler/lo-bicim.pptx");
    expect(currentFormat(a.s)).toBe("pptx");
  });

  it(".odp yerinde kaydedilir (soru bağlı değilse biçim korunur)", async () => {
    const { written, dialogs } = mockNative();
    const { actions: a } = appSession();
    await openBytes(a, buf("lo-bicim.odp"), "lo-bicim.odp", { name: "lo-bicim.odp", path: "/belgeler/lo-bicim.odp" });
    expect(await save(a)).toBe(true);
    expect(dialogs).toHaveLength(0);
    expect(written[0].path).toBe("/belgeler/lo-bicim.odp");
    expect(await sniffFormat(written[0].data, "x")).toBe("odp");
    expect(a.s.formatAsked).toBe(true);
  });

  it("Farklı Kaydet: süzgeçten başka biçim seçilirse o biçimin yazıcısıyla yeniden yazılır", async () => {
    const { written } = mockNative("odp");
    const { actions: a } = appSession();
    await openBytes(a, buf("lo-bicim.ppsx"), "lo-bicim.ppsx", { name: "lo-bicim.ppsx", path: "/belgeler/lo-bicim.ppsx" });
    expect(a.s.message?.action).toBe("startShow");
    expect(await save(a, true)).toBe(true);
    const last = written[written.length - 1];
    expect(last.path).toBe("/belgeler/lo-bicim.odp");
    expect(await sniffFormat(last.data, "x")).toBe("odp");
    expect(currentFormat(a.s)).toBe("odp");
  });

  it("şablon açılınca adsız yeni sunu olur; Kaydet konum sorar", async () => {
    const { dialogs } = mockNative();
    const { actions: a } = appSession();
    await openBytes(a, buf("lo-bicim.otp"), "lo-bicim.otp", { name: "lo-bicim.otp", path: "/belgeler/lo-bicim.otp" });
    expect(a.s.fileRef).toBeNull();
    expect(a.s.fileName).toMatch(/^Sunu\d+$/);
    await save(a);
    expect(dialogs[0].name).toMatch(/^Sunu\d+\.pptx$/);
  });

  it("Dışarı Aktar biçimi (.potx) içerik türüyle yazılır", async () => {
    const { written } = mockNative();
    const { actions: a } = appSession();
    await openBytes(a, buf("lo-bicim.odp"), "lo-bicim.odp", null);
    await save(a, true, "potx");
    expect(written[0].path).toBe("/belgeler/lo-bicim.potx");
    expect(await sniffFormat(written[0].data, "x.pptx")).toBe("potx");
  });
});
