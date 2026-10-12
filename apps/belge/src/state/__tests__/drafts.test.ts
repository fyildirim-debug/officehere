// Otomatik kurtarma taslakları: pencereye özgü anahtar, sahipsiz taslakların bulunması, devralma.
import { afterAll, beforeEach, describe, expect, it } from "vitest";

class MemoryStorage {
  private m = new Map<string, string>();
  get length() {
    return this.m.size;
  }
  key(i: number) {
    return [...this.m.keys()][i] ?? null;
  }
  getItem(k: string) {
    return this.m.has(k) ? this.m.get(k)! : null;
  }
  setItem(k: string, v: string) {
    this.m.set(k, String(v));
  }
  removeItem(k: string) {
    this.m.delete(k);
  }
  clear() {
    this.m.clear();
  }
}

const store = new MemoryStorage();
(globalThis as unknown as { localStorage: MemoryStorage }).localStorage = store;

const D = await import("../drafts");
const content = (t: string) => ({ type: "doc", content: [{ type: "section", content: [{ type: "paragraph", content: [{ type: "text", text: t }] }] }] });
const draft = (t: string, at = 1000) => ({ content: content(t), settings: {} as never, fileName: "Belge1", savedAt: at });

/** Başka bir pencereyi taklit eder: ping'lere kendi kimliğiyle yanıt verir. */
function liveWindow(id: string) {
  const ch = new BroadcastChannel("officehere.belge.windows");
  ch.onmessage = (e: MessageEvent<{ t: string; nonce: string }>) => {
    if (e.data?.t === "ping") ch.postMessage({ t: "pong", from: id, nonce: e.data.nonce });
  };
  return () => ch.close();
}

beforeEach(() => store.clear());
afterAll(() => store.clear());

describe("pencereye özgü taslaklar", () => {
  it("taslak bu pencerenin anahtarına yazılır; başka pencerenin taslağını ezmez", () => {
    store.setItem(`${D.DRAFT_PREFIX}diger`, JSON.stringify({ ...draft("DİĞER"), owner: "diger" }));
    D.writeDraft(draft("BENİM"));
    expect(store.getItem(D.draftKey())).toContain("BENİM");
    expect(store.getItem(`${D.DRAFT_PREFIX}diger`)).toContain("DİĞER");
    expect(D.readOwnDraft()?.content).toEqual(content("BENİM"));
    D.clearOwnDraft();
    expect(store.getItem(D.draftKey())).toBeNull();
    expect(store.getItem(`${D.DRAFT_PREFIX}diger`)).not.toBeNull();
  });

  it("yalnız sahipsiz taslaklar önerilir (canlı pencereninki ve bu pencereninki önerilmez)", async () => {
    D.writeDraft(draft("BENİM", 3000));
    store.setItem(`${D.DRAFT_PREFIX}canli`, JSON.stringify({ ...draft("CANLI", 2000), owner: "canli" }));
    store.setItem(`${D.DRAFT_PREFIX}kapali`, JSON.stringify({ ...draft("KAPALI", 1000), owner: "kapali" }));
    store.setItem(D.LEGACY_DRAFT_KEY, JSON.stringify(draft("ESKİ", 500)));
    const stop = liveWindow("canli");
    const orphans = await D.orphanDrafts(150);
    stop();
    expect(orphans.map((d) => d.owner)).toEqual(["kapali", ""]);
  });

  it("nabzı taze pencere yanıt vermese de canlı sayılır", async () => {
    store.setItem(`${D.DRAFT_PREFIX}mesgul`, JSON.stringify({ ...draft("MEŞGUL"), owner: "mesgul" }));
    store.setItem("officehere.belge.alive.mesgul", String(Date.now()));
    expect(await D.orphanDrafts(50)).toEqual([]);
    store.setItem("officehere.belge.alive.mesgul", String(Date.now() - 60_000));
    expect((await D.orphanDrafts(50)).map((d) => d.owner)).toEqual(["mesgul"]);
  });

  it("devralma: önce kendi anahtarına yazar, sonra eskisini siler", () => {
    store.setItem(`${D.DRAFT_PREFIX}kapali`, JSON.stringify({ ...draft("KURTAR"), owner: "kapali" }));
    const d = D.listDrafts()[0];
    D.adoptDraft(d);
    expect(store.getItem(`${D.DRAFT_PREFIX}kapali`)).toBeNull();
    expect(D.readOwnDraft()?.content).toEqual(content("KURTAR"));
    expect(D.readOwnDraft()?.savedAt).toBe(1000);
  });

  it("önizleme metni", () => {
    expect(D.draftPreview({ content: content("Merhaba dünya") as never })).toBe("Merhaba dünya");
    expect(D.draftPreview({ content: content("x".repeat(80)) as never }, 10)).toBe(`${"x".repeat(10)}…`);
  });
});
