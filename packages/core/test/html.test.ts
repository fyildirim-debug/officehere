// @vitest-environment jsdom
// Etkisiz HTML ayrıştırma: olay öznitelikleri belgeye hiç girmez, içerik ve diğer öznitelikler korunur.
import { describe, expect, it } from "vitest";
import { parseHtml } from "../src/html";

describe("parseHtml", () => {
  it("olay öznitelikleri atılır (tırnaklı, tırnaksız, büyük harfli, bitişik, SVG)", () => {
    const doc = parseHtml(
      '<p onclick="a()" class="x">Bir <b ONMOUSEOVER=b()>iki</b></p><img src="r.png" alt="x"onerror=\'c()\'>' +
        '<svg><image href="y" onerror="d()"/></svg><div/onload="e()">üç</div>',
    );
    const attrs = Array.from(doc.querySelectorAll("*")).flatMap((el) => Array.from(el.attributes, (a) => a.name));
    expect(attrs.filter((n) => /^on|^data-oh-olay-/i.test(n))).toEqual([]);
    expect(doc.body.textContent).toBe("Bir ikiüç");
    expect(doc.querySelector("p")!.className).toBe("x");
    expect(doc.querySelector("img")!.getAttribute("src")).toBe("r.png");
    expect(doc.querySelector("img")!.getAttribute("alt")).toBe("x");
  });

  it("metindeki ve öznitelik değerlerindeki 'on' sözcükleri değişmez", () => {
    const doc = parseHtml('<p title="turn on light" class="onboarding">online=1 on=2 <a href="https://ornek.com.tr/?on=1">bağ</a></p>');
    const p = doc.querySelector("p")!;
    expect(p.getAttribute("title")).toBe("turn on light");
    expect(p.className).toBe("onboarding");
    expect(p.textContent).toBe("online=1 on=2 bağ");
    expect(doc.querySelector("a")!.getAttribute("href")).toBe("https://ornek.com.tr/?on=1");
  });
});
