// @vitest-environment jsdom
// Denklemler: doğrusal giriş ayrıştırma, MathML, OMML yazma/okuma gidiş-dönüşü.
import { describe, expect, it } from "vitest";
import { parseMath, toLinear } from "../math/ast";
import { fromOmml, mathElement, ommlOf, MATH_NS } from "../math/render";

const omml = (latex: string) => {
  const xml = `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="${MATH_NS}">${ommlOf(parseMath(latex), false)}</w:p>`;
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  expect(doc.getElementsByTagName("parsererror").length).toBe(0);
  return doc.getElementsByTagNameNS(MATH_NS, "oMath")[0];
};

describe("denklem ayrıştırma", () => {
  it("kesir, kök, üst simge ve işleçler", () => {
    const n = parseMath("x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}");
    const ml = mathElement(n, true);
    expect(ml).toContain('display="block"');
    expect(ml).toContain("<mfrac>");
    expect(ml).toContain("<msqrt>");
    expect(ml).toContain("<msup><mi>b</mi><mn>2</mn></msup>");
    expect(ml).toContain("<mo>±</mo>");
  });
  it("doğrusal kesir a/b ve ayraç kaldırma", () => {
    expect(parseMath("(a+b)/c")).toMatchObject({ t: "frac", n: { t: "row" }, d: { t: "id", v: "c" } });
  });
  it("büyük işleç gövdesi, işlev sınırı, matris, Yunan harfleri", () => {
    const n = parseMath("\\sum_{i=1}^{n} i^2 = \\lim_{x\\to 0} \\sin x + \\begin{pmatrix}a&b\\\\c&d\\end{pmatrix} + \\alpha");
    const ml = mathElement(n, false);
    expect(ml).toContain("<munderover><mo largeop=\"true\">∑</mo>");
    expect(ml).toContain('<munder><mi mathvariant="normal">lim</mi>');
    expect(ml).toContain("<mtable><mtr><mtd><mi>a</mi></mtd>");
    expect(ml).toContain("<mi>α</mi>");
  });
});

describe("OMML gidiş-dönüşü", () => {
  const cases = [
    "x=\\frac{-b\\pm \\sqrt{b^2-4ac}}{2a}",
    "\\sum_{k=1}^n k^2",
    "\\int_0^1 x^2dx",
    "\\sqrt[3]{x+1}",
    "\\lim_{n\\to \\infty } \\frac{1}{n}",
    "A=\\begin{pmatrix}a&b\\\\c&d\\end{pmatrix}",
    "\\hat{x}+\\bar{y}",
    "\\left[x_1,x_2\\right]",
    "E=mc^2",
  ];
  for (const c of cases) {
    it(c, () => {
      const back = toLinear(fromOmml(omml(c)));
      // Yeniden ayrıştırılan ağaç aynı MathML'i üretir.
      expect(mathElement(parseMath(back), false)).toBe(mathElement(parseMath(c), false));
    });
  }
  it("OMML yapıları", () => {
    const x = new XMLSerializer().serializeToString(omml("\\frac{1}{2}+\\sum_{i}^{n} x_i^2"));
    expect(x).toContain("<m:f>");
    expect(x).toContain('<m:chr m:val="∑"/>');
    expect(x).toContain("<m:sSubSup>");
  });
});
