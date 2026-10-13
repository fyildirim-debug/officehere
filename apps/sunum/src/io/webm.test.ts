/** WebM yazıcısı: EBML yapısı, süre, kümeler, işaretler ve arama başlığı. */
import { describe, expect, it } from "vitest";
import { readEbml, vint, WEBM_ID, WebmMuxer } from "./webm";

const find = (nodes: ReturnType<typeof readEbml>, id: number) => nodes.find((n) => n.id === id)!;
const uintOf = (d: Uint8Array) => d.reduce((v, b) => v * 256 + b, 0);

describe("WebM yazıcısı", () => {
  it("EBML değişken uzunluklu boyut en kısa biçimde yazılır", () => {
    expect([...vint(0)]).toEqual([0x80]);
    expect([...vint(126)]).toEqual([0xfe]);
    expect([...vint(127)]).toEqual([0x40, 0x7f]);
    expect([...vint(1000)]).toEqual([0x43, 0xe8]);
  });

  it("izler, süre, anahtar karede başlayan kümeler ve küme konumlarını gösteren işaretler", () => {
    const m = new WebmMuxer(854, 480, "V_VP9");
    m.add(new Uint8Array([1, 2, 3]), 0, true);
    m.add(new Uint8Array([4]), 33, false);
    m.add(new Uint8Array([5, 6]), 1000, true);
    m.add(new Uint8Array([7]), 2500, false);
    const b = m.finish(3000);
    const top = readEbml(b);
    expect(top.map((n) => n.id)).toEqual([WEBM_ID.EBML, WEBM_ID.Segment]);
    const header = readEbml(top[0].data);
    expect(new TextDecoder().decode(find(header, WEBM_ID.DocType).data)).toBe("webm");
    const seg = top[1];
    const kids = readEbml(seg.data);
    expect(kids.map((n) => n.id)).toEqual([WEBM_ID.SeekHead, WEBM_ID.Info, WEBM_ID.Tracks, WEBM_ID.Cluster, WEBM_ID.Cluster, WEBM_ID.Cues]);
    const info = readEbml(find(kids, WEBM_ID.Info).data);
    expect(new DataView(find(info, WEBM_ID.Duration).data.buffer, find(info, WEBM_ID.Duration).data.byteOffset).getFloat64(0)).toBe(3000);
    const entry = readEbml(find(readEbml(find(kids, WEBM_ID.Tracks).data), WEBM_ID.TrackEntry).data);
    expect(new TextDecoder().decode(find(entry, WEBM_ID.CodecID).data)).toBe("V_VP9");
    const video = readEbml(find(entry, WEBM_ID.Video).data);
    expect(uintOf(find(video, WEBM_ID.PixelWidth).data)).toBe(854);
    // İkinci küme 1000 ms'de başlar; içindeki ikinci blok göreli 1500 ms, anahtar değil.
    const clusters = kids.filter((n) => n.id === WEBM_ID.Cluster);
    const c2 = readEbml(clusters[1].data);
    expect(uintOf(find(c2, WEBM_ID.Timecode).data)).toBe(1000);
    const blocks = c2.filter((n) => n.id === WEBM_ID.SimpleBlock);
    expect(blocks).toHaveLength(2);
    expect((blocks[1].data[1] << 8) | blocks[1].data[2]).toBe(1500);
    expect(blocks[0].data[3] & 0x80).toBe(0x80);
    expect(blocks[1].data[3] & 0x80).toBe(0);
    // İşaretler kümelerin segment verisi içindeki konumlarını gösterir (readEbml konumları segment verisine göredir).
    const cues = readEbml(find(kids, WEBM_ID.Cues).data).map((cp) => readEbml(find(readEbml(cp.data), WEBM_ID.CueTrackPositions).data));
    const positions = cues.map((tp) => uintOf(find(tp, WEBM_ID.CueClusterPosition).data));
    expect(positions).toEqual(clusters.map((c) => c.start));
    // Arama başlığındaki Cues konumu gerçek konumdur.
    const seeks = readEbml(find(kids, WEBM_ID.SeekHead).data).map((s) => readEbml(s.data));
    const cuesSeek = seeks.find((s) => uintOf(find(s, WEBM_ID.SeekID).data) === WEBM_ID.Cues)!;
    expect(uintOf(find(cuesSeek, WEBM_ID.SeekPosition).data)).toBe(find(kids, WEBM_ID.Cues).start);
  });
});
