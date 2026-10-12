import { describe, expect, it } from "vitest";
import { parse, tokenize, TR_SYNTAX, parseFormula, type Node } from "../src";
import { calc, err } from "./helpers";

function show(n: Node): string {
  switch (n.type) {
    case "num":
      return String(n.value);
    case "str":
      return JSON.stringify(n.value);
    case "bool":
      return n.value ? "TRUE" : "FALSE";
    case "err":
      return n.code;
    case "ref":
      return `${n.ref.sheet ? n.ref.sheet + "!" : ""}${n.ref.kind}(${n.ref.a.row},${n.ref.a.col}${n.ref.b ? `;${n.ref.b.row},${n.ref.b.col}` : ""})`;
    case "name":
      return `name:${n.name}`;
    case "func":
      return `${n.name}(${n.args.map(show).join(",")})`;
    case "unary":
      return `(${n.op}${show(n.arg)})`;
    case "postfix":
      return `(${show(n.arg)}%)`;
    case "binary":
      return `(${show(n.left)}${n.op}${show(n.right)})`;
    case "union":
      return `U(${n.items.map(show).join(",")})`;
    case "array":
      return `{${n.rows.map((r) => r.map(show).join(",")).join(";")}}`;
    case "missing":
      return "_";
  }
}

const p = (s: string) => show(parseFormula(s));

describe("sözcük çözümleyici", () => {
  it("temel belirteçler", () => {
    const t = tokenize('SUM(A1:B2, 1.5, "a""b", TRUE, #N/A)').filter((x) => x.type !== "ws");
    expect(t.map((x) => x.type)).toEqual(["func", "lparen", "ref", "sep", "number", "sep", "string", "sep", "bool", "sep", "error", "rparen"]);
    expect(t[6].value).toBe('a"b');
    expect(t[4].value).toBe(1.5);
  });
  it("başvuru biçimleri", () => {
    const refs = tokenize("$A$1+A$1+$A1+A:A+1:3+'Sayfa 1'!B2+Sayfa2!A1:B2+XFD1048576").filter((x) => x.type === "ref").map((x) => x.ref!);
    expect(refs[0].a).toMatchObject({ row: 0, col: 0, rowAbs: true, colAbs: true });
    expect(refs[1].a).toMatchObject({ rowAbs: true, colAbs: false });
    expect(refs[2].a).toMatchObject({ rowAbs: false, colAbs: true });
    expect(refs[3].kind).toBe("cols");
    expect(refs[4].kind).toBe("rows");
    expect(refs[5].sheet).toBe("Sayfa 1");
    expect(refs[6]).toMatchObject({ sheet: "Sayfa2", kind: "range" });
    expect(refs[7].a).toMatchObject({ row: 1048575, col: 16383 });
  });
  it("hücre başvurusuna benzeyen fonksiyonlar ve adlar", () => {
    const t = tokenize("LOG10(100)+A1B+XFZ1").filter((x) => x.type !== "ws");
    expect(t[0].type).toBe("func");
    expect(t.find((x) => x.text === "A1B")?.type).toBe("name");
    expect(t.find((x) => x.text === "XFZ1")?.type).toBe("name");
  });
  it("Türkçe sözdizimi", () => {
    const t = tokenize("TOPLA(A1;1,5;{1\\2;3\\4})", TR_SYNTAX).filter((x) => x.type !== "ws");
    expect(t.map((x) => x.type)).toEqual(["func", "lparen", "ref", "sep", "number", "sep", "lbrace", "number", "arrayCol", "number", "arrayRow", "number", "arrayCol", "number", "rbrace", "rparen"]);
    expect(t[4].value).toBe(1.5);
  });
  it("Türkçe hata adları", () => {
    const t = tokenize("#YOK+#SAYI/0!+#BAŞV!", TR_SYNTAX).filter((x) => x.type !== "ws" && x.type !== "op");
    expect(t[0].value).toBe("#N/A");
    expect(t[1].value).toBe("#DIV/0!");
    expect(t[2].ref?.refError).toBe(true);
  });
});

describe("ayrıştırıcı ve öncelik", () => {
  it("aritmetik öncelik", () => {
    expect(p("1+2*3")).toBe("(1+(2*3))");
    expect(p("(1+2)*3")).toBe("((1+2)*3)");
    expect(p("2^3^2")).toBe("((2^3)^2)");
    expect(p("-2^2")).toBe("((-2)^2)");
    expect(p("2^-1")).toBe("(2^(-1))");
    expect(p("10%*2")).toBe("((10%)*2)");
    expect(p('1&2="12"')).toBe('((1&2)="12")');
    expect(p("1+2&3")).toBe("((1+2)&3)");
    expect(p("1<2=TRUE")).toBe("((1<2)=TRUE)");
    expect(p("A1:B2")).toBe("range(0,0;1,1)");
    expect(p("-A1%")).toBe("((-cell(0,0))%)");
  });
  it("fonksiyonlar ve boş bağımsız değişkenler", () => {
    expect(p("IF(A1,,2)")).toBe("IF(cell(0,0),_,2)");
    expect(p("SUM()")).toBe("SUM()");
    expect(p("sum(1,max(2,3))")).toBe("SUM(1,MAX(2,3))");
    expect(p("_xlfn.XLOOKUP(1,A:A,B:B)")).toBe("XLOOKUP(1,cols(-1,0;-1,0),cols(-1,1;-1,1))");
    expect(p("{1,2;3,4}")).toBe("{1,2;3,4}");
    expect(p("{-1,\"a\"}")).toBe('{-1,"a"}');
    expect(p("SUM((A1,B1))")).toBe("SUM(U(cell(0,0),cell(0,1)))");
    expect(p("MyName*2")).toBe("(name:MyName*2)");
  });
  it("Türkçe ayrıştırma", () => {
    const n = parse("=TOPLA(1;2,5)", { syntax: TR_SYNTAX, resolveFunction: (x) => (x === "TOPLA" ? "SUM" : undefined) });
    expect(show(n)).toBe("SUM(1,2.5)");
  });
  it("sözdizimi hataları", () => {
    expect(() => parseFormula("1+")).toThrow();
    expect(() => parseFormula("SUM(1,2")).toThrow();
    expect(() => parseFormula("(1")).toThrow();
    expect(() => parseFormula("{1,2;3}")).toThrow();
    expect(() => parseFormula('"abc')).toThrow();
    expect(err(calc("=1+"))).toBe("#NAME?");
  });
});

describe("işleçler ve zorlama", () => {
  it("aritmetik", () => {
    expect(calc("=1+2*3")).toBe(7);
    expect(calc("=-2^2")).toBe(4);
    expect(calc("=2^3^2")).toBe(64);
    expect(calc("=10/4")).toBe(2.5);
    expect(calc("=50%")).toBe(0.5);
    expect(calc("=5%%")).toBe(0.0005);
    expect(calc("=0.1+0.2=0.3")).toBe(true);
    expect(err(calc("=1/0"))).toBe("#DIV/0!");
    expect(err(calc("=0^0"))).toBe("#NUM!");
    expect(calc("=(-8)^(1/3)")).toBeCloseTo(-2);
  });
  it("metin ve mantıksal zorlama", () => {
    expect(calc('="3"+4')).toBe(7);
    expect(calc('="1,5"*2')).toBe(3);
    expect(calc("=TRUE+TRUE")).toBe(2);
    expect(err(calc('="abc"+1'))).toBe("#VALUE!");
    expect(calc("=1&2")).toBe("12");
    expect(calc("=1.5&\"\"")).toBe("1,5");
    expect(calc('=TRUE&"x"')).toBe("DOĞRUx");
    expect(calc("=A1+1", {})).toBe(1);
    expect(calc('=A1&"x"', {})).toBe("x");
    expect(calc('="50%"*2')).toBe(1);
    expect(calc('="01.01.2024"+0')).toBe(45292);
  });
  it("karşılaştırma kuralları", () => {
    expect(calc('="a"="A"')).toBe(true);
    expect(calc('="abc"<"abd"')).toBe(true);
    expect(calc('=1<"a"')).toBe(true);
    expect(calc('="z"<TRUE')).toBe(true);
    expect(calc("=A1=0")).toBe(true);
    expect(calc('=A1=""')).toBe(true);
    expect(calc("=A1=FALSE")).toBe(true);
    expect(calc('="ç"<"d"')).toBe(true);
    expect(calc('="i"="İ"')).toBe(true);
    expect(calc("=1<>1")).toBe(false);
    expect(calc("=2>=2")).toBe(true);
  });
  it("hata yayılımı", () => {
    expect(err(calc("=#N/A+1"))).toBe("#N/A");
    expect(err(calc("=1+#REF!"))).toBe("#REF!");
    expect(err(calc("=NOSUCHFN(1)"))).toBe("#NAME?");
    expect(err(calc("=undefinedname+1"))).toBe("#NAME?");
    expect(err(calc("=Yok!A1"))).toBe("#REF!");
  });
  it("dizi işlemleri ve yayınım", () => {
    expect(calc("=SUM({1,2,3}*2)")).toBe(12);
    expect(calc("=SUM({1,2,3}*{1;2})")).toBe(18);
    expect(calc("=SUM((A1:A3>1)*B1:B3)", { A1: 1, A2: 2, A3: 3, B1: 10, B2: 20, B3: 30 })).toBe(50);
    expect(calc("=ROWS({1;2;3}+{1,2})")).toBe(3);
  });
});
