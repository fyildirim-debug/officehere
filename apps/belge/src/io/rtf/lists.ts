// RTF listeleri: \listtable düzey tanımları, \listoverridetable eşlemeleri ve eski \pn numaralandırması.
import type { ListInfo } from "../common";
import { NFC_STYLE, type Par } from "./state";

interface Level {
  nfc: number;
  start: number;
}

export class RtfLists {
  private lists = new Map<number, Level[]>();
  private cur: Level[] | null = null;
  private level: Level | null = null;
  private overrides = new Map<number, number>();
  private override: { list: number; ls: number } | null = null;

  /** \listtable içindeki sözcükler. */
  word(w: string, param: number | null) {
    switch (w) {
      case "list":
        this.cur = [];
        break;
      case "listid":
        if (this.cur) this.lists.set(param ?? 0, this.cur);
        break;
      case "listlevel":
        this.level = { nfc: 0, start: 1 };
        this.cur?.push(this.level);
        break;
      case "levelnfc":
      case "levelnfcn":
        if (this.level) this.level.nfc = param ?? 0;
        break;
      case "levelstartat":
        if (this.level) this.level.start = param ?? 1;
        break;
    }
  }

  startOverride() {
    this.override = { list: 0, ls: 0 };
  }

  overrideWord(w: string, param: number | null) {
    if (!this.override) return;
    if (w === "listid") this.override.list = param ?? 0;
    else if (w === "ls") this.override.ls = param ?? 0;
  }

  endOverride() {
    if (this.override) this.overrides.set(this.override.ls, this.override.list);
    this.override = null;
  }

  /** Paragrafın liste bilgisi (\ls + \ilvl ya da \pn). */
  info(pa: Par): ListInfo | undefined {
    if (pa.ls !== null && pa.ls > 0) {
      const lvl = this.lists.get(this.overrides.get(pa.ls) ?? -1)?.[pa.ilvl];
      const nfc = lvl ? (NFC_STYLE[lvl.nfc] ?? { ordered: lvl.nfc !== 255, style: null }) : { ordered: false, style: null };
      return { key: `ls${pa.ls}`, level: pa.ilvl, ordered: nfc.ordered, listStyle: nfc.style, start: lvl?.start ?? 1 };
    }
    if (pa.pn) return { key: `pn${pa.pn.ordered}`, level: pa.pn.level, ordered: pa.pn.ordered };
    return undefined;
  }
}
