// @vitest-environment jsdom
/** HTML tablolarını içe aktarma ve biçimli HTML dışa aktarma. */
import { describe, expect, it } from "vitest";
import { cellKey } from "@officehere/formula";
import { readHtml, writeHtml } from "../src/io/html";
import { openBook } from "../src/io/formats";
import { WorkbookModel } from "../src/model/workbook";

const page = `<!doctype html><html><head><meta charset="utf-8"></head><body>
<table><caption>Satışlar</caption>
<tr><th>Ürün</th><th>Tutar</th></tr>
<tr><td style="background-color:#ffff00">Çay</td><td>1.234,5</td></tr>
<tr><td colspan="2"><b>Birleşik</b></td></tr>
</table>
<p>ara metin</p>
<table><tr><td>İkinci</td><td><a href="https://ornek.com">bağlantı</a></td></tr></table>
</body></html>`;

describe("HTML içe aktarma", () => {
  it("her tablo bir sayfa; başlık, birleştirme, stil ve bağlantı", () => {
    const wb = readHtml(new TextEncoder().encode(page), "sayfa");
    expect(wb.sheets.map((s) => s.name)).toEqual(["Satışlar", "Tablo2"]);
    const [a, b] = wb.sheets;
    expect(wb.getValue(a, 0, 0)).toBe("Ürün");
    expect(wb.styleAt(a, 0, 0).bold).toBe(true);
    expect(wb.getValue(a, 1, 1)).toBe(1234.5);
    expect(wb.styleAt(a, 1, 0).fill).toBe("#FFFF00");
    expect(a.merges).toEqual([{ r1: 2, c1: 0, r2: 2, c2: 1 }]);
    expect(wb.styleAt(a, 2, 0).bold).toBe(true);
    expect(wb.getValue(b, 0, 0)).toBe("İkinci");
    expect(b.links.get(cellKey(0, 1))).toBe("https://ornek.com");
  });
  it("tablosuz sayfa anlaşılır hata verir; .html açılınca tanınır", async () => {
    expect(() => readHtml(new TextEncoder().encode("<html><body>yok</body></html>"))).toThrow(/tablo/);
    const o = await openBook("rapor.html", new TextEncoder().encode(page).buffer as ArrayBuffer);
    expect(o.format).toBe("html");
  });
});

describe("HTML dışa aktarma", () => {
  it("tüm sayfalar başlıklarıyla, biçimler satır içi", () => {
    const wb = new WorkbookModel();
    wb.setInputRaw(wb.sheets[0], 0, 0, "Merhaba <dünya>");
    wb.sheets[0].styles.set(cellKey(0, 0), wb.styles.intern({ bold: true, fill: "#FF0000" }));
    const s2 = wb.addSheet("İkinci");
    wb.setInputRaw(s2, 0, 0, 42);
    const html = writeHtml(wb, wb.sheets, "Rapor");
    expect(html).toContain("<title>Rapor</title>");
    expect(html).toContain("<h2>Sayfa1</h2>");
    expect(html).toContain("<h2>İkinci</h2>");
    expect(html).toContain("Merhaba &lt;dünya&gt;");
    expect(html).toMatch(/font-weight:\s*(bold|700)/);
    const back = readHtml(new TextEncoder().encode(html));
    expect(back.getValue(back.sheets[0], 0, 0)).toBe("Merhaba <dünya>");
    expect(back.getValue(back.sheets[1], 0, 0)).toBe(42);
  });
});
