import { describe, expect, it } from "vitest";
import { pictureFileName, srcToBlob } from "./pictureCmds";

describe("Resim Olarak Kaydet", () => {
  it("data: adresini türüyle birlikte Blob'a çevirir", async () => {
    const b = await srcToBlob("data:image/png;base64,iVBORw==");
    expect(b.type).toBe("image/png");
    expect(new Uint8Array(await b.arrayBuffer())).toEqual(new Uint8Array([0x89, 0x50, 0x4e, 0x47]));
    const svg = await srcToBlob("data:image/svg+xml,%3Csvg%2F%3E");
    expect(await svg.text()).toBe("<svg/>");
  });

  it("dosya adı türün uzantısını alır, yasak karakterler atılır", () => {
    expect(pictureFileName("Resim 3", "image/jpeg")).toBe("Resim 3.jpg");
    expect(pictureFileName("a/b:c", "image/svg+xml")).toBe("a b c.svg");
    expect(pictureFileName("", "application/octet-stream")).toBe("Resim.png");
  });
});
