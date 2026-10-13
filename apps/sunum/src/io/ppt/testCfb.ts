/**
 * Yalnız testler ve örnek üretimi için en küçük CFB yazıcısı (MS-CFB sürüm 3,
 * 512 baytlık kesim). Kökte düz akışlar; 4096 bayttan küçükler küçük akışa
 * (mini stream) yazılır. Uygulama .ppt yazmaz.
 */
const END = 0xfffffffe;
const FREE = 0xffffffff;
const FATSECT = 0xfffffffd;
const NOSTREAM = 0xffffffff;
const SEC = 512;
const MINI = 64;
const CUTOFF = 4096;

function cmpName(a: string, b: string): number {
  if (a.length !== b.length) return a.length - b.length;
  const x = a.toUpperCase();
  const y = b.toUpperCase();
  return x < y ? -1 : x > y ? 1 : 0;
}

export function writeCfb(streams: { name: string; data: Uint8Array }[]): Uint8Array {
  const list = [...streams].sort((a, b) => cmpName(a.name, b.name));
  // Küçük akış verisi.
  const miniParts: { start: number; count: number }[] = [];
  let miniLen = 0;
  const miniFat: number[] = [];
  for (const s of list) {
    if (s.data.length >= CUTOFF || s.data.length === 0) {
      miniParts.push({ start: END, count: 0 });
      continue;
    }
    const n = Math.ceil(s.data.length / MINI);
    const start = miniLen / MINI;
    for (let i = 0; i < n; i++) miniFat.push(i === n - 1 ? END : start + i + 1);
    miniParts.push({ start, count: n });
    miniLen += n * MINI;
  }
  const miniStream = new Uint8Array(miniLen);
  list.forEach((s, i) => {
    if (miniParts[i].count) miniStream.set(s.data, miniParts[i].start * MINI);
  });

  // Kesim yerleşimi: büyük akışlar, küçük akış, mini FAT, dizin, FAT.
  const fat: number[] = [];
  const sectors: Uint8Array[] = [];
  const place = (data: Uint8Array): number => {
    if (!data.length) return END;
    const n = Math.ceil(data.length / SEC);
    const start = sectors.length;
    for (let i = 0; i < n; i++) {
      const sec = new Uint8Array(SEC);
      sec.set(data.subarray(i * SEC, Math.min(data.length, (i + 1) * SEC)));
      sectors.push(sec);
      fat.push(i === n - 1 ? END : start + i + 1);
    }
    return start;
  };
  const bigStart = list.map((s) => (s.data.length >= CUTOFF ? place(s.data) : END));
  const miniStart = place(miniStream);
  const mfBytes = new Uint8Array(Math.ceil((miniFat.length * 4) / SEC) * SEC).fill(0xff);
  const mfv = new DataView(mfBytes.buffer);
  miniFat.forEach((v, i) => mfv.setUint32(i * 4, v, true));
  const miniFatStart = miniFat.length ? place(mfBytes) : END;

  // Dizin: kök + akışlar; çocuklar dengeli ikili ağaç.
  const count = list.length + 1;
  const dir = new Uint8Array(Math.ceil((count * 128) / SEC) * SEC);
  const dv = new DataView(dir.buffer);
  const tree = (lo: number, hi: number): number => {
    if (lo > hi) return NOSTREAM;
    const mid = (lo + hi) >> 1;
    const entry = mid + 1;
    dv.setUint32(entry * 128 + 0x44, tree(lo, mid - 1), true);
    dv.setUint32(entry * 128 + 0x48, tree(mid + 1, hi), true);
    return entry;
  };
  const writeEntry = (i: number, name: string, type: number, start: number, size: number) => {
    const o = i * 128;
    for (let k = 0; k < name.length && k < 31; k++) dv.setUint16(o + k * 2, name.charCodeAt(k), true);
    dv.setUint16(o + 0x40, (Math.min(name.length, 31) + 1) * 2, true);
    dv.setUint8(o + 0x42, type);
    dv.setUint8(o + 0x43, 1);
    dv.setUint32(o + 0x44, NOSTREAM, true);
    dv.setUint32(o + 0x48, NOSTREAM, true);
    dv.setUint32(o + 0x4c, NOSTREAM, true);
    dv.setUint32(o + 0x74, start, true);
    dv.setUint32(o + 0x78, size, true);
  };
  for (let i = count; i < dir.length / 128; i++) {
    dv.setUint32(i * 128 + 0x44, NOSTREAM, true);
    dv.setUint32(i * 128 + 0x48, NOSTREAM, true);
    dv.setUint32(i * 128 + 0x4c, NOSTREAM, true);
  }
  writeEntry(0, "Root Entry", 5, miniStart, miniLen);
  list.forEach((s, i) => writeEntry(i + 1, s.name, 2, s.data.length >= CUTOFF ? bigStart[i] : miniParts[i].start, s.data.length));
  dv.setUint32(0x4c, tree(0, list.length - 1), true);
  const dirStart = place(dir);

  // FAT kesimleri (kendileri de FAT'te işaretlenir).
  let nFat = 1;
  while (Math.ceil((sectors.length + nFat) / (SEC / 4)) > nFat) nFat++;
  if (nFat > 109) throw new Error("Örnek dosya çok büyük");
  const fatStart = sectors.length;
  for (let i = 0; i < nFat; i++) fat.push(FATSECT);
  const fatBytes = new Uint8Array(nFat * SEC).fill(0xff);
  const fv = new DataView(fatBytes.buffer);
  fat.forEach((v, i) => fv.setUint32(i * 4, v, true));
  for (let i = 0; i < nFat; i++) sectors.push(fatBytes.subarray(i * SEC, (i + 1) * SEC));

  const head = new Uint8Array(SEC);
  const hv = new DataView(head.buffer);
  head.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  hv.setUint16(0x18, 0x3e, true);
  hv.setUint16(0x1a, 3, true);
  hv.setUint16(0x1c, 0xfffe, true);
  hv.setUint16(0x1e, 9, true);
  hv.setUint16(0x20, 6, true);
  hv.setUint32(0x2c, nFat, true);
  hv.setUint32(0x30, dirStart, true);
  hv.setUint32(0x38, CUTOFF, true);
  hv.setUint32(0x3c, miniFatStart, true);
  hv.setUint32(0x40, miniFat.length ? mfBytes.length / SEC : 0, true);
  hv.setUint32(0x44, END, true);
  hv.setUint32(0x48, 0, true);
  for (let i = 0; i < 109; i++) hv.setUint32(0x4c + i * 4, i < nFat ? fatStart + i : FREE, true);

  const out = new Uint8Array(SEC * (sectors.length + 1));
  out.set(head, 0);
  sectors.forEach((s, i) => out.set(s, (i + 1) * SEC));
  return out;
}
