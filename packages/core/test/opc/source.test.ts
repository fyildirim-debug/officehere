import { describe, expect, it } from "vitest";
import { OpcPackage } from "../../src/opc/package";
import { getSource, openSource, packageSources, registerSource, releaseSource, sha256Hex, SourceRegistry } from "../../src/opc/source";
import { minimalDocx } from "./paket";

describe("kaynak paket kaydı", () => {
  it("kayıt, bulma, bırakma; kimlik verilirse o kullanılır ve öncekinin yerine geçer", async () => {
    const reg = new SourceRegistry();
    const a = await OpcPackage.open(await minimalDocx());
    const b = await OpcPackage.open(await minimalDocx({ "x.xml": "<x/>" }));
    const id1 = reg.register(a, { name: "a.docx" });
    const id2 = reg.register(b);
    expect(id1).toMatch(/^kaynak-\d+$/);
    expect(id1).not.toBe(id2);
    expect(reg.get(id1)).toBe(a);
    expect(reg.entry(id1)?.name).toBe("a.docx");
    expect(reg.byteSize).toBe(a.bytes.byteLength + b.bytes.byteLength);
    expect(reg.register(b, { id: id1 })).toBe(id1);
    expect(reg.get(id1)).toBe(b);
    expect(reg.release(id1)).toBe(true);
    expect(reg.release(id1)).toBe(false);
    expect(reg.has(id1)).toBe(false);
    expect(reg.get(undefined)).toBeUndefined();
    expect(reg.ids()).toEqual([id2]);
    reg.clear();
    expect(reg.ids()).toEqual([]);
  });

  it("varsayılan kayıt: openSource ve işlev kısayolları", async () => {
    const bytes = await minimalDocx();
    const s = await openSource(bytes, { name: "b.docx" });
    expect(getSource(s.id)).toBe(s.pkg);
    expect(packageSources.has(s.id)).toBe(true);
    const id = registerSource(s.pkg, { id: await sha256Hex(bytes) });
    expect(id).toMatch(/^[0-9a-f]{64}$/);
    expect(releaseSource(id)).toBe(true);
    expect(releaseSource(s.id)).toBe(true);
    await expect(openSource(new Uint8Array(40))).rejects.toMatchObject({ name: "OpcError" });
  });

  it("SHA-256 bilinen değer", async () => {
    expect(await sha256Hex(new TextEncoder().encode("abc"))).toBe("ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad");
  });
});
