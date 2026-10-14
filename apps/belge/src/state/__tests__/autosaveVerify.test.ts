// @vitest-environment jsdom
// Otomatik kaydetme doğrulaması: yazılan .docx kendi okuyucumuzla yeniden açılabiliyor mu.
import { describe, expect, it } from "vitest";
import { writeDocx } from "../../io";
import { defaultSettings } from "../../model/document";
import { canReopenDocx } from "../useFileOps";

const ab = (u: Uint8Array) => u.buffer.slice(u.byteOffset, u.byteOffset + u.byteLength) as ArrayBuffer;

describe("otomatik kaydetme doğrulaması (Belge)", () => {
  it("kendi yazdığımız belge doğrulanır", async () => {
    const bytes = await writeDocx({ content: { type: "doc", content: [{ type: "paragraph", content: [{ type: "text", text: "Merhaba" }] }] }, settings: defaultSettings() });
    expect(await canReopenDocx(ab(bytes))).toBe(true);
  });

  it("kesik ya da bozuk dosya doğrulanamaz (hata ya da false)", async () => {
    const bytes = await writeDocx({ content: { type: "doc", content: [{ type: "paragraph", content: [{ type: "text", text: "Merhaba" }] }] }, settings: defaultSettings() });
    const cut = bytes.slice(0, Math.floor(bytes.length / 2));
    const ok = await canReopenDocx(ab(cut)).catch(() => false);
    expect(ok).toBe(false);
    expect(await canReopenDocx(new TextEncoder().encode("bozuk").buffer as ArrayBuffer).catch(() => false)).toBe(false);
  });
});
