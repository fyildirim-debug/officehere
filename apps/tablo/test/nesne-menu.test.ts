/** Nesne bağlam menüsü komutları: Resim Olarak Kaydet yardımcıları. */
import { describe, expect, it } from "vitest";
import { imageFileName, srcToBlob } from "../src/actions/imageFile";

describe("Resim Olarak Kaydet", () => {
  it("data: adresini türüyle Blob'a çevirir", async () => {
    const b = await srcToBlob("data:image/jpeg;base64,/9j/");
    expect(b.type).toBe("image/jpeg");
    expect([...new Uint8Array(await b.arrayBuffer())]).toEqual([0xff, 0xd8, 0xff]);
    expect(await (await srcToBlob("data:image/svg+xml,%3Csvg%2F%3E")).text()).toBe("<svg/>");
  });

  it("dosya adı: eski uzantı atılır, türün uzantısı eklenir", () => {
    expect(imageFileName("simge.png", "image/png")).toBe("simge.png");
    expect(imageFileName("Fotoğraf.JPEG", "image/jpeg")).toBe("Fotoğraf.jpg");
    expect(imageFileName("a:b", "image/webp")).toBe("a b.webp");
    expect(imageFileName("", "application/octet-stream")).toBe("Resim.png");
  });
});
