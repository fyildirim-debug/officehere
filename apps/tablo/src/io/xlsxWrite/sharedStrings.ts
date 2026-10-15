/** Paylaşılan metinler (sharedStrings.xml): ilk görülme sırasıyla dizinlenir, sonda parça parça yazılır. */
import { ChunkedText, DeflateSink, type PackedEntry } from "./zip";
import { esc, XML_DECL } from "./xml";

export class SharedStrings {
  private index = new Map<string, number>();
  private values: string[] = [];
  /** Toplam başvuru sayısı (count özniteliği). */
  count = 0;

  add(s: string): number {
    this.count++;
    let i = this.index.get(s);
    if (i === undefined) {
      i = this.values.length;
      this.values.push(s);
      this.index.set(s, i);
    }
    return i;
  }

  get size(): number {
    return this.values.length;
  }

  async pack(): Promise<PackedEntry> {
    const sink = new DeflateSink();
    const out = new ChunkedText(sink);
    out.put(`${XML_DECL}<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${this.count}" uniqueCount="${this.values.length}">`);
    for (const v of this.values) {
      // Baştaki/sondaki boşluk ve satır sonu korunur
      out.put(/^\s|\n|\s$/.test(v) ? `<si><t xml:space="preserve">${esc(v)}</t></si>` : `<si><t>${esc(v)}</t></si>`);
      if (out.pending >= 32768) await out.drain();
    }
    out.put("</sst>");
    await out.drain(true);
    return sink.finish();
  }
}
