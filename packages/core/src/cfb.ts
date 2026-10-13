// Bileşik Dosya İkili Biçimi (Compound File Binary, MS-CFB) okuyucusu.
// Eski Office dosyaları (.doc, .xls, .ppt) bu kapsayıcının içindeki akışlarda saklanır:
// "WordDocument" + "1Table", "Workbook", "PowerPoint Document" + "Current User" gibi.
// Yalnız okur; bozuk ya da kötü niyetli dosyalarda döngü ve taşmalara karşı korunur.

const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;
const NOSTREAM = 0xffffffff;

export interface CfbEntry {
  /** Girdinin adı (ör. "WordDocument"). */
  name: string;
  /** Kökten "/" ile ayrılmış yol (ör. "ObjectPool/_1234/Ole"). Kök girdisi yer almaz. */
  path: string;
  type: "storage" | "stream";
  size: number;
}

export interface CfbFile {
  entries: CfbEntry[];
  /** Yoldaki akışın baytları; yol büyük/küçük harfe duyarsızdır. Yoksa null. */
  read(path: string): Uint8Array | null;
  has(path: string): boolean;
}

function bytesOf(data: ArrayBuffer | Uint8Array): Uint8Array {
  return data instanceof Uint8Array ? data : new Uint8Array(data);
}

/** İçerik bir CFB kapsayıcısı mı (imza: D0 CF 11 E0 A1 B1 1A E1)? */
export function isCfb(data: ArrayBuffer | Uint8Array): boolean {
  const b = bytesOf(data);
  if (b.length < 512) return false;
  return SIGNATURE.every((v, i) => b[i] === v);
}

interface DirEntry {
  name: string;
  type: number;
  left: number;
  right: number;
  child: number;
  start: number;
  size: number;
}

/** CFB kapsayıcısını ayrıştırır. Geçersizse Error fırlatır. */
export function readCfb(data: ArrayBuffer | Uint8Array): CfbFile {
  const b = bytesOf(data);
  if (!isCfb(b)) throw new Error("Bileşik dosya imzası yok");
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  const u16 = (o: number) => dv.getUint16(o, true);
  const u32 = (o: number) => dv.getUint32(o, true);

  const sectorShift = u16(0x1e);
  const miniShift = u16(0x20);
  if (sectorShift !== 9 && sectorShift !== 12) throw new Error("Geçersiz kesim boyutu");
  if (miniShift !== 6) throw new Error("Geçersiz küçük kesim boyutu");
  const sectorSize = 1 << sectorShift;
  const miniSize = 1 << miniShift;
  const sectorCount = Math.floor((b.length - sectorSize) / sectorSize) + (b.length % sectorSize ? 1 : 0);
  const sectorOffset = (n: number) => (n + 1) * sectorSize;
  const perSector = sectorSize / 4;

  const numFat = u32(0x2c);
  const firstDir = u32(0x30);
  const miniCutoff = u32(0x38) || 4096;
  const firstMiniFat = u32(0x3c);
  const numMiniFat = u32(0x40);
  let difat = u32(0x44);
  let numDifat = u32(0x48);

  // FAT kesimlerinin listesi: başlıktaki 109 girdi + DIFAT zinciri.
  const fatSectors: number[] = [];
  for (let i = 0; i < 109 && fatSectors.length < numFat; i++) {
    const s = u32(0x4c + i * 4);
    if (s !== FREESECT) fatSectors.push(s);
  }
  const seenDifat = new Set<number>();
  while (numDifat-- > 0 && difat !== ENDOFCHAIN && difat !== FREESECT && fatSectors.length < numFat) {
    if (difat >= sectorCount || seenDifat.has(difat)) throw new Error("Bozuk DIFAT zinciri");
    seenDifat.add(difat);
    const off = sectorOffset(difat);
    for (let i = 0; i < perSector - 1 && fatSectors.length < numFat; i++) {
      const s = u32(off + i * 4);
      if (s !== FREESECT) fatSectors.push(s);
    }
    difat = u32(off + (perSector - 1) * 4);
  }

  const fat = new Uint32Array(fatSectors.length * perSector);
  fatSectors.forEach((s, k) => {
    if (s >= sectorCount) throw new Error("FAT kesimi dosya dışında");
    const off = sectorOffset(s);
    for (let i = 0; i < perSector; i++) fat[k * perSector + i] = u32(off + i * 4);
  });

  /** Bir zincirin kesim numaraları; döngü ya da sınır aşımında hata. */
  const chain = (start: number, table: Uint32Array, limit: number): number[] => {
    const out: number[] = [];
    const seen = new Set<number>();
    for (let s = start; s !== ENDOFCHAIN && s !== FREESECT; s = table[s]) {
      if (s >= limit || s >= table.length || seen.has(s)) throw new Error("Bozuk kesim zinciri");
      seen.add(s);
      out.push(s);
    }
    return out;
  };

  const readChain = (start: number, size?: number): Uint8Array => {
    const secs = chain(start, fat, sectorCount);
    const total = secs.length * sectorSize;
    const out = new Uint8Array(size === undefined ? total : Math.min(size, total));
    secs.forEach((s, i) => {
      const from = sectorOffset(s);
      const at = i * sectorSize;
      if (at >= out.length) return;
      out.set(b.subarray(from, Math.min(from + sectorSize, b.length, from + out.length - at)), at);
    });
    return out;
  };

  // Dizin girdileri.
  const dirBytes = readChain(firstDir);
  const dirs: DirEntry[] = [];
  const ddv = new DataView(dirBytes.buffer, dirBytes.byteOffset, dirBytes.byteLength);
  for (let off = 0; off + 128 <= dirBytes.length; off += 128) {
    const nameLen = Math.min(64, ddv.getUint16(off + 0x40, true));
    let name = "";
    for (let i = 0; i + 1 < nameLen - 1; i += 2) name += String.fromCharCode(ddv.getUint16(off + i, true));
    const lo = ddv.getUint32(off + 0x78, true);
    const hi = sectorShift === 9 ? 0 : ddv.getUint32(off + 0x7c, true);
    dirs.push({
      name,
      type: ddv.getUint8(off + 0x42),
      left: ddv.getUint32(off + 0x44, true),
      right: ddv.getUint32(off + 0x48, true),
      child: ddv.getUint32(off + 0x4c, true),
      start: ddv.getUint32(off + 0x74, true),
      size: hi * 0x100000000 + lo,
    });
  }
  const root = dirs[0];
  if (!root || root.type !== 5) throw new Error("Kök girdi yok");

  // Küçük akışlar: kökün akışı (mini stream) + mini FAT.
  const miniStream = root.start === ENDOFCHAIN ? new Uint8Array(0) : readChain(root.start, root.size);
  const miniFatBytes = numMiniFat && firstMiniFat !== ENDOFCHAIN ? readChain(firstMiniFat) : new Uint8Array(0);
  const miniFat = new Uint32Array(miniFatBytes.length / 4);
  const mdv = new DataView(miniFatBytes.buffer, miniFatBytes.byteOffset, miniFatBytes.byteLength);
  for (let i = 0; i < miniFat.length; i++) miniFat[i] = mdv.getUint32(i * 4, true);
  const miniCount = Math.floor(miniStream.length / miniSize);

  const readMini = (start: number, size: number): Uint8Array => {
    const secs = chain(start, miniFat, miniCount);
    const out = new Uint8Array(Math.min(size, secs.length * miniSize));
    secs.forEach((s, i) => {
      const at = i * miniSize;
      if (at >= out.length) return;
      out.set(miniStream.subarray(s * miniSize, s * miniSize + Math.min(miniSize, out.length - at)), at);
    });
    return out;
  };

  // Ağacı dolaş: her depolamanın çocukları kırmızı-siyah ağaçta (sol/sağ kardeşler).
  const entries: CfbEntry[] = [];
  const byPath = new Map<string, DirEntry>();
  const visited = new Set<number>();
  const walk = (id: number, prefix: string) => {
    const stack = [id];
    while (stack.length) {
      const i = stack.pop()!;
      if (i === NOSTREAM || i >= dirs.length || visited.has(i)) continue;
      visited.add(i);
      const d = dirs[i];
      stack.push(d.left, d.right);
      if (d.type !== 1 && d.type !== 2) continue;
      const path = prefix ? `${prefix}/${d.name}` : d.name;
      entries.push({ name: d.name, path, type: d.type === 1 ? "storage" : "stream", size: d.type === 2 ? d.size : 0 });
      byPath.set(path.toLowerCase(), d);
      if (d.type === 1) walk(d.child, path);
    }
  };
  visited.add(0);
  walk(root.child, "");

  const read = (path: string): Uint8Array | null => {
    const d = byPath.get(path.replace(/^\/+/, "").toLowerCase());
    if (!d || d.type !== 2) return null;
    if (d.size === 0) return new Uint8Array(0);
    return d.size < miniCutoff ? readMini(d.start, d.size) : readChain(d.start, d.size);
  };

  return { entries, read, has: (p) => byPath.has(p.replace(/^\/+/, "").toLowerCase()) };
}
