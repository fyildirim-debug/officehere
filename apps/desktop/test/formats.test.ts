import { describe, expect, it } from "vitest";
import { ALL_EXTENSIONS, FORMAT_GROUPS, appForFileName } from "../src/shared/formats";

describe("masaüstü biçim eşlemesi", () => {
  it("uzantıya göre uygulamayı bulur (büyük/küçük harf ve yol fark etmez)", () => {
    expect(appForFileName("rapor.DOCX")).toBe("belge");
    expect(appForFileName("C:\\Belgeler\\eski.doc")).toBe("belge");
    expect(appForFileName("/home/a/notlar.odt")).toBe("belge");
    expect(appForFileName("mektup.rtf")).toBe("belge");
    expect(appForFileName("BENİOKU.md")).toBe("belge");
    expect(appForFileName("bütçe.xls")).toBe("tablo");
    expect(appForFileName("veri.ods")).toBe("tablo");
    expect(appForFileName("liste.tsv")).toBe("tablo");
    expect(appForFileName("sunum.ppt")).toBe("sunum");
    expect(appForFileName("gösteri.ppsx")).toBe("sunum");
    expect(appForFileName("taslak.odp")).toBe("sunum");
  });

  it("tanınmayan ya da uzantısız dosyada null", () => {
    expect(appForFileName("resim.png")).toBeNull();
    expect(appForFileName("Makefile")).toBeNull();
    expect(appForFileName("klasör.docx/dosya")).toBeNull();
  });

  it("her uzantı tek bir uygulamaya ait", () => {
    const all = FORMAT_GROUPS.flatMap((g) => g.extensions);
    expect(new Set(all).size).toBe(all.length);
    expect(ALL_EXTENSIONS.sort()).toEqual([...all].sort());
  });
});
