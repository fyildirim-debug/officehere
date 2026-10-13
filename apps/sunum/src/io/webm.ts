/**
 * En küçük WebM (Matroska alt kümesi) yazıcısı: tek video izi (VP8/VP9),
 * kodlanmış kareler SimpleBlock olarak kümelere yazılır; süre, arama başlığı
 * (SeekHead) ve işaretler (Cues) eklenir. Dosya bellekte bir kez kurulur,
 * bu yüzden tüm boyutlar kesindir (bilinmeyen boyut kullanılmaz).
 * Biçim: EBML (RFC 8794) ve Matroska (RFC 9559) öğe kimlikleri.
 */

const ID = {
  EBML: 0x1a45dfa3,
  EBMLVersion: 0x4286,
  EBMLReadVersion: 0x42f7,
  EBMLMaxIDLength: 0x42f2,
  EBMLMaxSizeLength: 0x42f3,
  DocType: 0x4282,
  DocTypeVersion: 0x4287,
  DocTypeReadVersion: 0x4285,
  Segment: 0x18538067,
  SeekHead: 0x114d9b74,
  Seek: 0x4dbb,
  SeekID: 0x53ab,
  SeekPosition: 0x53ac,
  Info: 0x1549a966,
  TimecodeScale: 0x2ad7b1,
  MuxingApp: 0x4d80,
  WritingApp: 0x5741,
  Duration: 0x4489,
  Tracks: 0x1654ae6b,
  TrackEntry: 0xae,
  TrackNumber: 0xd7,
  TrackUID: 0x73c5,
  TrackType: 0x83,
  FlagLacing: 0x9c,
  CodecID: 0x86,
  Video: 0xe0,
  PixelWidth: 0xb0,
  PixelHeight: 0xba,
  Cluster: 0x1f43b675,
  Timecode: 0xe7,
  SimpleBlock: 0xa3,
  Cues: 0x1c53bb6b,
  CuePoint: 0xbb,
  CueTime: 0xb3,
  CueTrackPositions: 0xb7,
  CueTrack: 0xf7,
  CueClusterPosition: 0xf1,
} as const;

type Bytes = Uint8Array;

function concat(parts: Bytes[]): Bytes {
  let n = 0;
  for (const p of parts) n += p.length;
  const out = new Uint8Array(n);
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}

/** Öğe kimliği (işaret bitleri kimliğin içindedir) büyük uçlu baytlar. */
function idBytes(id: number): Bytes {
  const out: number[] = [];
  let v = id;
  while (v > 0) {
    out.unshift(v % 256);
    v = Math.floor(v / 256);
  }
  return Uint8Array.from(out);
}

/** EBML değişken uzunluklu boyut (en kısa biçim; len verilirse o uzunlukta). */
export function vint(n: number, len?: number): Bytes {
  for (let l = len ?? 1; l <= 8; l++) {
    if (n < 2 ** (7 * l) - 1) {
      const b = new Uint8Array(l);
      let v = n;
      for (let i = l - 1; i >= 0; i--) {
        b[i] = v % 256;
        v = Math.floor(v / 256);
      }
      b[0] |= 1 << (8 - l);
      return b;
    }
  }
  throw new Error("EBML boyutu çok büyük");
}

function uintBytes(v: number, len?: number): Bytes {
  const out: number[] = [];
  let x = Math.floor(v);
  do {
    out.unshift(x % 256);
    x = Math.floor(x / 256);
  } while (x > 0);
  while (len && out.length < len) out.unshift(0);
  return Uint8Array.from(out);
}

const el = (id: number, data: Bytes): Bytes => concat([idBytes(id), vint(data.length), data]);
const elList = (id: number, kids: Bytes[]): Bytes => el(id, concat(kids));
const u = (id: number, v: number, len?: number): Bytes => el(id, uintBytes(v, len));
const str = (id: number, s: string): Bytes => el(id, new TextEncoder().encode(s));
function f64(id: number, v: number): Bytes {
  const b = new Uint8Array(8);
  new DataView(b.buffer).setFloat64(0, v);
  return el(id, b);
}

interface Frame {
  data: Bytes;
  /** ms */
  time: number;
  key: boolean;
}

export class WebmMuxer {
  private frames: Frame[] = [];

  constructor(
    private width: number,
    private height: number,
    /** "V_VP8" ya da "V_VP9" */
    private codecId: string,
  ) {}

  /** Kodlanmış kare (zaman ms, artan sırada). */
  add(data: Bytes, timeMs: number, key: boolean) {
    this.frames.push({ data, time: Math.max(0, Math.round(timeMs)), key });
  }

  get frameCount() {
    return this.frames.length;
  }

  /** WebM dosyası; durationMs görüntülenme süresidir. */
  finish(durationMs: number): Bytes {
    const frames = this.frames.slice().sort((a, b) => a.time - b.time);
    const header = elList(ID.EBML, [u(ID.EBMLVersion, 1), u(ID.EBMLReadVersion, 1), u(ID.EBMLMaxIDLength, 4), u(ID.EBMLMaxSizeLength, 8), str(ID.DocType, "webm"), u(ID.DocTypeVersion, 2), u(ID.DocTypeReadVersion, 2)]);
    const info = elList(ID.Info, [u(ID.TimecodeScale, 1_000_000), str(ID.MuxingApp, "OfficeHere"), str(ID.WritingApp, "OfficeHere Sunum"), f64(ID.Duration, Math.max(1, durationMs))]);
    const tracks = elList(ID.Tracks, [
      elList(ID.TrackEntry, [u(ID.TrackNumber, 1), u(ID.TrackUID, 1), u(ID.TrackType, 1), u(ID.FlagLacing, 0), str(ID.CodecID, this.codecId), elList(ID.Video, [u(ID.PixelWidth, this.width), u(ID.PixelHeight, this.height)])]),
    ]);

    // Kümeler: her anahtar karede (ve göreli zaman 16 bite sığmazsa) yeni küme.
    const clusters: { time: number; bytes: Bytes }[] = [];
    let cur: { time: number; blocks: Bytes[] } | null = null;
    const flush = () => {
      if (cur) clusters.push({ time: cur.time, bytes: elList(ID.Cluster, [u(ID.Timecode, cur.time), ...cur.blocks]) });
      cur = null;
    };
    for (const f of frames) {
      if (!cur || (f.key && f.time > cur.time) || f.time - cur.time > 30000) {
        flush();
        cur = { time: f.time, blocks: [] };
      }
      const rel = f.time - cur.time;
      const head = new Uint8Array(4);
      head[0] = 0x81; // iz 1
      head[1] = (rel >> 8) & 0xff;
      head[2] = rel & 0xff;
      head[3] = f.key ? 0x80 : 0;
      cur.blocks.push(el(ID.SimpleBlock, concat([head, f.data])));
    }
    flush();

    // Yerleşim: SeekHead, Info, Tracks, Cluster…, Cues (konumlar segment verisinin başına göre).
    const seekEntry = (id: number, pos: number) => elList(ID.Seek, [el(ID.SeekID, idBytes(id)), u(ID.SeekPosition, pos, 8)]);
    const seekSize = elList(ID.SeekHead, [seekEntry(ID.Info, 0), seekEntry(ID.Tracks, 0), seekEntry(ID.Cues, 0)]).length;
    const infoPos = seekSize;
    const tracksPos = infoPos + info.length;
    let pos = tracksPos + tracks.length;
    const cuePoints: Bytes[] = [];
    for (const c of clusters) {
      cuePoints.push(elList(ID.CuePoint, [u(ID.CueTime, c.time), elList(ID.CueTrackPositions, [u(ID.CueTrack, 1), u(ID.CueClusterPosition, pos)])]));
      pos += c.bytes.length;
    }
    const cues = elList(ID.Cues, cuePoints.length ? cuePoints : [elList(ID.CuePoint, [u(ID.CueTime, 0), elList(ID.CueTrackPositions, [u(ID.CueTrack, 1), u(ID.CueClusterPosition, tracksPos + tracks.length)])])]);
    const seekHead = elList(ID.SeekHead, [seekEntry(ID.Info, infoPos), seekEntry(ID.Tracks, tracksPos), seekEntry(ID.Cues, pos)]);
    const segment = elList(ID.Segment, [seekHead, info, tracks, ...clusters.map((c) => c.bytes), cues]);
    return concat([header, segment]);
  }
}

/* ------------------------- Okuma (doğrulama/testler için) ------------------------- */

export interface EbmlNode {
  id: number;
  start: number;
  size: number;
  data: Bytes;
}

function readVint(b: Bytes, p: number, keepMarker: boolean): { value: number; len: number } {
  const first = b[p];
  let len = 1;
  while (len <= 8 && !(first & (1 << (8 - len)))) len++;
  let v = keepMarker ? first : first & ((1 << (8 - len)) - 1);
  for (let i = 1; i < len; i++) v = v * 256 + b[p + i];
  return { value: v, len };
}

/** Bir düzeydeki EBML öğeleri. */
export function readEbml(b: Bytes, start = 0, end = b.length): EbmlNode[] {
  const out: EbmlNode[] = [];
  let p = start;
  while (p < end) {
    const id = readVint(b, p, true);
    const size = readVint(b, p + id.len, false);
    const ds = p + id.len + size.len;
    out.push({ id: id.value, start: p, size: size.value, data: b.subarray(ds, ds + size.value) });
    p = ds + size.value;
  }
  return out;
}

export const WEBM_ID = ID;
