/**
 * Yalnız testler için en küçük .ppt üreticisi: anahat metni SlideListWithText
 * içinde (OutlineTextRefAtom ile başvurulur), başlık ve gövde yer tutucuları.
 */
import { writeCfb } from "./testCfb";

class W {
  private parts: number[] = [];
  u8(v: number) {
    this.parts.push(v & 0xff);
    return this;
  }
  u16(v: number) {
    return this.u8(v).u8(v >> 8);
  }
  u32(v: number) {
    return this.u16(v & 0xffff).u16((v >>> 16) & 0xffff);
  }
  bytes(b: Uint8Array) {
    for (const x of b) this.parts.push(x);
    return this;
  }
  get out() {
    return new Uint8Array(this.parts);
  }
}

const atom = (type: number, body: Uint8Array, inst = 0, ver = 0) => new W().u16((inst << 4) | ver).u16(type).u32(body.length).bytes(body).out;
const cont = (type: number, kids: Uint8Array[], inst = 0) => atom(type, concat(kids), inst, 0xf);
function concat(list: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(list.reduce((n, b) => n + b.length, 0));
  let o = 0;
  for (const b of list) {
    out.set(b, o);
    o += b.length;
  }
  return out;
}
const utf16 = (s: string) => {
  const w = new W();
  for (const c of s) w.u16(c.charCodeAt(0));
  return w.out;
};

/** Başlık (TextBytesAtom) ve maddeli gövde (TextCharsAtom + StyleTextPropAtom) içeren tek slaytlık sunu. */
export function buildTestPpt(title: string, body: string): Uint8Array {
  const n = body.length + 1;
  // Paragraf: madde işareti açık (•); karakter: kalın, 28 pt.
  const style = new W().u32(n).u16(0).u32(0x1 | 0x80).u16(1).u16(0x2022).u32(n).u32(0x1 | 0x20000).u16(1).u16(28).out;
  const texts = [
    atom(0x0f9f, new W().u32(0).out),
    atom(0x0fa8, new Uint8Array([...title].map((c) => c.charCodeAt(0)))),
    atom(0x0f9f, new W().u32(1).out),
    atom(0x0fa0, utf16(body)),
    atom(0x0fa1, style),
  ];
  const persist = atom(0x03f3, new W().u32(2).u32(0).u32(2).u32(256).u32(0).out);
  const docAtom = atom(0x03e9, new W().u32(5760).u32(4320).u32(4320).u32(5760).u32(1).u32(2).u32(0).u32(0).u16(1).u16(0).u8(0).u8(0).u8(0).u8(1).out, 0, 1);
  const doc = cont(0x03e8, [docAtom, cont(0x0ff0, [persist, ...texts], 0), atom(0x03ea, new Uint8Array(0))]);

  const sp = (spt: number, id: number, flags: number, extra: Uint8Array[]) => cont(0xf004, [atom(0xf00a, new W().u32(id).u32(flags).out, spt, 2), ...extra]);
  const anchor = (t: number, l: number, r: number, b: number) => atom(0xf010, new W().u16(t).u16(l).u16(r).u16(b).out);
  const ph = (id: number) => cont(0xf011, [atom(0x0bc3, new W().u32(0).u8(id).u8(0).u16(0).out)]);
  const ref = (i: number) => cont(0xf00d, [atom(0x0f9e, new W().u32(i).out)]);
  const group = cont(0xf003, [
    cont(0xf004, [atom(0xf009, new Uint8Array(16), 0, 1), atom(0xf00a, new W().u32(1024).u32(0x5).out, 0, 2)]),
    sp(1, 1025, 0xa00, [anchor(288, 576, 5184, 1008), ph(13), ref(0)]),
    sp(1, 1026, 0xa00, [anchor(1152, 576, 5184, 3888), ph(14), ref(1)]),
  ]);
  const drawing = cont(0x040c, [cont(0xf002, [atom(0xf008, new W().u32(3).u32(1026).out, 1), group])]);
  const slideAtom = atom(0x03ef, new W().u32(1).bytes(new Uint8Array(8)).u32(0).u32(0).u16(0).u16(0).out, 0, 2);
  const slide = cont(0x03ee, [slideAtom, drawing]);

  const docOff = 0;
  const slideOff = doc.length;
  const pdOff = slideOff + slide.length;
  const pd = atom(0x1772, new W().u32(1 | (2 << 20)).u32(docOff).u32(slideOff).out);
  const ueOff = pdOff + pd.length;
  const ue = atom(0x0ff5, new W().u32(256).u16(0).u8(0).u8(3).u32(0).u32(pdOff).u32(1).u32(3).u16(1).u16(0).out);
  const stream = concat([doc, slide, pd, ue]);
  const user = "Test";
  const cu = atom(0x0ff6, new W().u32(0x14).u32(0xe391c05f).u32(ueOff).u16(user.length).u16(0x03f4).u8(3).u8(0).u16(0).bytes(new Uint8Array([...user].map((c) => c.charCodeAt(0)))).u32(8).out);
  return writeCfb([
    { name: "PowerPoint Document", data: stream },
    { name: "Current User", data: cu },
  ]);
}
