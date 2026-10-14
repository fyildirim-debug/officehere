// @vitest-environment jsdom
// Yapıştırma güvenliği: kötü niyetli HTML etkisiz belgede ayrıştırılır, betik içeren öğe/öznitelikler atılır,
// metin ve biçim korunur.
import { describe, expect, it } from "vitest";
import { sanitizePastedHtml } from "../paste";
import { htmlToDoc } from "../../io";

const EVIL =
  '<p onclick="window.__xss=1">Merhaba <b>dünya</b> <img src="resim.png" onerror="window.__xss=2"><img src="data:image/png;base64,AAAA"><img src="data:text/html;base64,PHNjcmlwdD4="></p>' +
  '<script>window.__xss=3</script><iframe src="https://kotu.example/"></iframe><object data="x.swf"></object><embed src="x.swf">' +
  '<a href="javascript:window.__xss=4">bağ1</a> <a href=" jav&#x09;ascript:alert(1)">bağ2</a> <a href="https://ornek.com.tr/">bağ3</a>' +
  '<svg><a xlink:href="javascript:alert(1)"><text>svg</text></a><set attributeName="href" to="javascript:alert(1)"/></svg>' +
  '<form action="javascript:alert(1)"><button formaction="javascript:alert(1)">düğme</button></form>' +
  '<div style="color:red;background:url(javascript:alert(1))">stil</div><span style="color:#c00000">kırmızı</span>';

describe("yapıştırılan HTML temizliği", () => {
  it("betik, çerçeve, nesne, olay öznitelikleri ve javascript: adresleri atılır; içerik korunur", () => {
    const out = sanitizePastedHtml(EVIL);
    expect(out).not.toMatch(/<script|<iframe|<object|<embed|onerror|onclick|javascript:|formaction|data:text\/html|<set/i);
    expect(out).toContain("Merhaba <b>dünya</b>");
    expect(out).toContain('<img src="resim.png">');
    expect(out).toContain('src="data:image/png;base64,AAAA"');
    expect(out).toContain('<a href="https://ornek.com.tr/">bağ3</a>');
    expect(out).toContain(">bağ1</a>");
    expect(out).toContain(">bağ2</a>");
    expect(out).toContain("düğme");
    expect(out).toContain('<span style="color:#c00000">kırmızı</span>');
    expect(out).toContain(">stil</div>");
  });

  it("temizleme sırasında hiçbir betik çalışmaz ve resim yüklenmez", async () => {
    const w = window as unknown as { __xss?: number };
    delete w.__xss;
    sanitizePastedHtml(EVIL);
    await new Promise((r) => setTimeout(r, 50));
    expect(w.__xss).toBeUndefined();
  });

  it("açılan HTML belgesi de temizlenir", () => {
    const doc = htmlToDoc(`<html><body>${EVIL}</body></html>`);
    const json = JSON.stringify(doc);
    expect(json).not.toMatch(/javascript:|onerror|__xss/);
    expect(json).toContain("Merhaba");
    expect(json).toContain("https://ornek.com.tr/");
  });
});
