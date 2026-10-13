// RTF kod sayfaları: \ansicpg ve \fcharset → TextDecoder; yazıcı için ANSI yedek baytları.

const CHARSET_CP: Record<number, number> = {
  77: 10000,
  128: 932,
  129: 949,
  134: 936,
  136: 950,
  161: 1253,
  162: 1254,
  163: 1258,
  177: 1255,
  178: 1256,
  186: 1257,
  204: 1251,
  222: 874,
  238: 1250,
};

const CP_LABEL: Record<number, string> = {
  10000: "macintosh",
  932: "shift_jis",
  949: "euc-kr",
  936: "gbk",
  950: "big5",
  874: "windows-874",
  866: "ibm866",
};

/** \fcharset değerinin kod sayfası (0/1: belgenin ANSI kod sayfası; 2: simge yazı tipi → -1). */
export function charsetCodepage(charset: number | undefined, ansi: number): number {
  if (charset === undefined || charset === 0 || charset === 1) return ansi;
  if (charset === 2) return -1;
  return CHARSET_CP[charset] ?? ansi;
}

const decoders = new Map<number, TextDecoder>();

/** Kod sayfası çözücüsü (desteklenmeyen kod sayfasında Windows-1252). */
export function decoderFor(cp: number): TextDecoder {
  let d = decoders.get(cp);
  if (d) return d;
  const label = CP_LABEL[cp] ?? (cp >= 1250 && cp <= 1258 ? `windows-${cp}` : "windows-1252");
  try {
    d = new TextDecoder(label);
  } catch {
    d = new TextDecoder("windows-1252");
  }
  decoders.set(cp, d);
  return d;
}

/** Baytları kod sayfasıyla çözer; simge yazı tipinde (cp -1) baytlar özel kullanım alanına değil doğrudan karaktere çevrilir. */
export function decodeBytes(bytes: number[], cp: number): string {
  if (!bytes.length) return "";
  if (cp === -1) return String.fromCharCode(...bytes);
  return decoderFor(cp).decode(new Uint8Array(bytes));
}

const encoders = new Map<number, Map<string, number>>();

/** Karakterin tek baytlık ANSI karşılığı (yoksa null; \uN? yedeği için). */
export function ansiByte(ch: string, cp = 1254): number | null {
  const code = ch.charCodeAt(0);
  if (code < 0x80) return code;
  let map = encoders.get(cp);
  if (!map) {
    map = new Map();
    const dec = decoderFor(cp);
    for (let b = 0x80; b < 0x100; b++) {
      const s = dec.decode(new Uint8Array([b]));
      if (s.length === 1 && s !== "�" && !map.has(s)) map.set(s, b);
    }
    encoders.set(cp, map);
  }
  return map.get(ch) ?? null;
}
