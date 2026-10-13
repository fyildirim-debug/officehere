// RTF okuyucu (RTF 1.9.1): grup yığını, yazı tipi/renk/stil tabloları, kod sayfaları, \uN, karakter ve
// paragraf biçimi, listeler, tablolar, resimler (\pngblip/\jpegblip), alanlar, dipnotlar, üst/alt bilgi.
import type { JSONContent } from "@tiptap/core";
import { defaultSettings, normalizeDoc, type BelgeDocument } from "../../model/document";
import { noteText } from "../../editor/notes";
import { fieldPicture } from "../../editor/nodes";
import { pictureOfInstr } from "../../editor/dateFormat";
import { charMarks, foldLists, noteNode, paragraphNode, stripStyleProps, styleChar, stylesOf, type CharProps, type FlowItem } from "../common";
import { charsetCodepage, decodeBytes } from "./codepage";
import { RtfTables } from "./tables";
import { RtfPict } from "./pict";
import { RtfLists } from "./lists";
import { HL_COLORS, LINK_COLOR, SKIP_DESTS, charWord, newChr, newPar, paraWord, headingLevelOf, styleIdByName, applyInfo, type Chr, type Par } from "./state";

/** Sayfa ve tarih/saat alanları (sonuç metni yerine alan düğümü okunur). */
const PAGE_FIELD = /^\s*(PAGE|NUMPAGES|DATE|TIME)\b/i;

/** Alan düğümü; tarih ve saatte \@ biçim resmi korunur (varsayılansa yazılmaz). */
function pageFieldNode(field: string, instr: string): JSONContent {
  if (field !== "DATE" && field !== "TIME") return { type: "pageField", attrs: { field } };
  const pic = pictureOfInstr(instr);
  return { type: "pageField", attrs: { field, format: pic && pic !== fieldPicture(field) ? pic : null } };
}

type Dest =
  | "body"
  | "skip"
  | "fonttbl"
  | "colortbl"
  | "stylesheet"
  | "info"
  | "infoText"
  | "pict"
  | "fldinst"
  | "listtable"
  | "listoverride"
  | "bkmk"
  | "pn"
  | "shp"
  | "sn"
  | "shpPib";

interface Group {
  chr: Chr;
  par: Par;
  dest: Dest;
  /** Grup kapanınca çalışan işlev. */
  onEnd?: () => void;
  ignorable: boolean;
  first: boolean;
}

/** Bir metin akışı (gövde, dipnot, üst/alt bilgi). */
class Story {
  note = false;
  items: FlowItem[] = [];
  inline: JSONContent[] = [];
  tables = new RtfTables();
}

const SPECIAL: Record<string, string> = { tab: "\t", emdash: "—", endash: "–", bullet: "•", lquote: "‘", rquote: "’", ldblquote: "“", rdblquote: "”", emspace: " ", enspace: " ", qmspace: " " };

type PageKey = "width" | "height" | "left" | "right" | "top" | "bottom" | "header" | "footer";
const PAGE_WORDS: Record<string, PageKey> = {
  paperw: "width",
  pgwsxn: "width",
  paperh: "height",
  pghsxn: "height",
  margl: "left",
  marglsxn: "left",
  margr: "right",
  margrsxn: "right",
  margt: "top",
  margtsxn: "top",
  margb: "bottom",
  margbsxn: "bottom",
  headery: "header",
  footery: "footer",
};

export class RtfReader {
  private lists = new RtfLists();
  private pos = 0;
  private stack: Group[] = [];
  private g!: Group;
  private ansi = 1252;
  private deff = 0;
  private fonts = new Map<number, { name: string; charset?: number }>();
  private curFont: { id: number; name: string; charset?: number } | null = null;
  private colors: (string | null)[] = [];
  private colorCur: { r: number; g: number; b: number; set: boolean } = { r: 0, g: 0, b: 0, set: false };
  private styleNames = new Map<number, string>();
  private styleCur: { id: number; name: string } | null = null;
  private pending: number[] = [];
  private skipUc = 0;
  private story = new Story();
  private main = this.story;
  private infoKey = "";
  private infoText = "";
  private pict: RtfPict | null = null;
  /** Geçerli şekil (\shpinst): kutu ve resim özelliği. */
  private shape: { shpleft: number; shptop: number; shpright: number; shpbottom: number; prop: string; pict: boolean } | null = null;
  private fields: { instr: string; result: boolean }[] = [];
  private bkmk = "";
  private noteSeq = 0;
  readonly doc: BelgeDocument = { content: { type: "doc" }, settings: defaultSettings() };
  private page = { ...this.doc.settings.page, margin: { ...this.doc.settings.page.margin } };
  private pageTouched = false;

  constructor(private readonly b: Uint8Array) {}

  read(): BelgeDocument {
    // Baştaki BOM ve boşluklar atlanır.
    while (this.pos < Math.min(64, this.b.length) && this.b[this.pos] !== 0x7b) this.pos++;
    const head = new TextDecoder("latin1").decode(this.b.subarray(this.pos, this.pos + 6));
    if (!head.startsWith("{\\rtf")) throw new Error("Geçerli bir RTF dosyası değil.");
    this.g = { chr: newChr(), par: newPar(), dest: "body", ignorable: false, first: false };
    let guard = 0;
    while (this.pos < this.b.length) {
      if (++guard > 50_000_000) throw new Error("RTF dosyası çok büyük ya da bozuk.");
      const c = this.b[this.pos];
      if (c === 0x7b) {
        this.flushText();
        if (this.stack.length > 2000) throw new Error("RTF grupları çok derin iç içe.");
        this.stack.push(this.g);
        this.g = { chr: { ...this.g.chr }, par: { ...this.g.par, pn: this.g.par.pn }, dest: this.g.dest, ignorable: false, first: true };
        this.skipUc = 0;
        this.pos++;
      } else if (c === 0x7d) {
        this.flushText();
        this.g.onEnd?.();
        const prev = this.stack.pop();
        if (!prev) break;
        this.g = prev;
        this.skipUc = 0;
        this.pos++;
      } else if (c === 0x5c) {
        this.control();
      } else if (c === 0x0d || c === 0x0a) {
        this.pos++;
      } else {
        this.pos++;
        if (this.skipUc > 0) {
          this.skipUc--;
          continue;
        }
        this.g.first = false;
        this.pending.push(c);
      }
    }
    this.flushText();
    this.endParagraph(true);
    this.story.tables.close(this.story.items);
    const blocks = foldLists(this.main.items);
    const s = this.doc.settings;
    if (this.pageTouched) s.page = this.page;
    this.doc.content = normalizeDoc({ type: "doc", content: blocks.length ? blocks : [{ type: "paragraph" }] });
    return this.doc;
  }

  /* -------------------------- denetim sözcükleri -------------------------- */

  private control() {
    const b = this.b;
    let p = this.pos + 1;
    const c = b[p];
    if (c === undefined) {
      this.pos = p;
      return;
    }
    const isAlpha = (x: number) => (x >= 0x41 && x <= 0x5a) || (x >= 0x61 && x <= 0x7a);
    if (!isAlpha(c)) {
      // Denetim simgesi
      this.pos = p + 1;
      if (c === 0x27) {
        const hex = String.fromCharCode(b[p + 1] ?? 0, b[p + 2] ?? 0);
        this.pos = p + 3;
        if (this.skipUc > 0) {
          this.skipUc--;
          return;
        }
        const v = parseInt(hex, 16);
        if (Number.isFinite(v)) this.pending.push(v);
        return;
      }
      if (c === 0x2a) {
        this.g.ignorable = true;
        return;
      }
      this.flushText();
      if (this.skipUc > 0) {
        this.skipUc--;
        return;
      }
      const sym: Record<number, string> = { 0x5c: "\\", 0x7b: "{", 0x7d: "}", 0x7e: " ", 0x2d: "­", 0x5f: "‑", 0x09: "\t" };
      if (sym[c] !== undefined) this.text(sym[c]);
      else if (c === 0x0a || c === 0x0d) this.word("par", null);
      return;
    }
    let w = "";
    while (p < b.length && isAlpha(b[p])) w += String.fromCharCode(b[p++]);
    let param: number | null = null;
    let num = "";
    if (b[p] === 0x2d && b[p + 1] >= 0x30 && b[p + 1] <= 0x39) num = String.fromCharCode(b[p++]);
    while (p < b.length && b[p] >= 0x30 && b[p] <= 0x39 && num.length < 11) num += String.fromCharCode(b[p++]);
    if (num) param = parseInt(num, 10);
    if (b[p] === 0x20) p++;
    this.pos = p;
    if (w === "bin") {
      const n = Math.max(0, Math.min(param ?? 0, b.length - p));
      if (this.g.dest === "pict" && this.pict) this.pict.bin = b.slice(p, p + n);
      this.pos = p + n;
      return;
    }
    if (w === "u") {
      this.flushText();
      if (this.skipUc > 0) {
        this.skipUc--;
        return;
      }
      let cp = param ?? 0;
      if (cp < 0) cp += 65536;
      this.text(String.fromCharCode(cp));
      this.skipUc = this.g.chr.uc;
      return;
    }
    if (this.skipUc > 0) {
      this.skipUc--;
      return;
    }
    this.flushText();
    this.word(w, param);
  }

  /** Bekleyen ANSI baytlarını geçerli yazı tipinin kod sayfasıyla çözer. */
  private flushText() {
    if (!this.pending.length) return;
    const font = this.fonts.get(this.g.chr.f ?? this.deff);
    const cp = this.g.dest === "fonttbl" ? charsetCodepage(this.curFont?.charset, this.ansi) : charsetCodepage(font?.charset, this.ansi);
    const s = decodeBytes(this.pending, cp);
    this.pending = [];
    this.text(s);
  }

  private text(s: string) {
    switch (this.g.dest) {
      case "body":
        this.bodyText(s);
        break;
      case "fonttbl":
        if (this.curFont) {
          this.curFont.name += s;
          if (s.includes(";")) this.saveFont();
        }
        break;
      case "colortbl":
        if (s.includes(";")) {
          for (let k = 0; k < (s.match(/;/g)?.length ?? 0); k++) {
            const c = this.colorCur;
            this.colors.push(c.set ? `#${[c.r, c.g, c.b].map((x) => Math.max(0, Math.min(255, x)).toString(16).padStart(2, "0")).join("").toUpperCase()}` : null);
            this.colorCur = { r: 0, g: 0, b: 0, set: false };
          }
        }
        break;
      case "stylesheet":
        if (this.styleCur) this.styleCur.name += s;
        break;
      case "infoText":
        this.infoText += s;
        break;
      case "pict":
        this.pict?.text(s);
        break;
      case "fldinst":
        if (this.fields.length) this.fields[this.fields.length - 1].instr += s;
        break;
      case "listtable":
        break;
      case "bkmk":
        this.bkmk += s;
        break;
      case "sn":
        if (this.shape) this.shape.prop += s;
        break;
    }
  }

  private bodyText(s: string) {
    const f = this.fields[this.fields.length - 1];
    if (f && !f.result) return;
    if (f && f.result && PAGE_FIELD.test(f.instr)) return;
    const c = this.g.chr;
    const font = this.fonts.get(c.f ?? this.deff);
    const props: CharProps = {
      bold: c.b,
      italic: c.i,
      underline: c.ul,
      strike: c.strike,
      dstrike: c.dstrike,
      sup: c.sup,
      sub: c.sub,
      font: font?.name ?? null,
      size: c.fs !== null ? c.fs / 2 : 12,
      color: this.colors[c.cf] ?? null,
      highlight: c.hl ? (this.colors[c.hl] ?? HL_COLORS[c.hl] ?? null) : null,
      bg: c.cb ? (this.colors[c.cb] ?? null) : null,
      caps: c.caps,
      smallCaps: c.scaps,
      hidden: c.v,
      spacing: c.expnd ? c.expnd / 20 : null,
      link: c.link,
    };
    this.story.inline.push({ type: "text", text: s, attrs: { _props: props } } as JSONContent);
  }

  private node(n: JSONContent) {
    if (this.g.dest !== "body") return;
    this.story.inline.push(n);
  }

  private word(w: string, param: number | null) {
    const g = this.g;
    const c = g.chr;
    const pa = g.par;
    const first = g.first;
    g.first = false;
    if (g.dest === "skip") return;
    // Hedef (destination) sözcükleri: yalnız grubun ilk sözcüğü olarak anlamlıdır.
    if (first || g.ignorable) {
      if (this.destination(w)) return;
      if (g.ignorable && first) {
        g.dest = "skip";
        return;
      }
    }
    switch (g.dest) {
      case "fonttbl":
        if (w === "f") {
          this.saveFont();
          this.curFont = { id: param ?? 0, name: "" };
          const prev = g.onEnd;
          g.onEnd = () => {
            prev?.();
            this.saveFont();
          };
        } else if (w === "fcharset" && this.curFont) this.curFont.charset = param ?? 0;
        return;
      case "colortbl":
        if (w === "red" || w === "green" || w === "blue") {
          this.colorCur[w === "red" ? "r" : w === "green" ? "g" : "b"] = param ?? 0;
          this.colorCur.set = true;
        }
        return;
      case "stylesheet":
        if (w === "s" || w === "cs" || w === "ds" || w === "ts") {
          const cur = { id: w === "s" ? (param ?? 0) : -1 - (param ?? 0), name: "" };
          this.styleCur = cur;
          const prev = g.onEnd;
          g.onEnd = () => {
            prev?.();
            this.styleNames.set(cur.id, cur.name.replace(/;[\s\S]*$/, "").trim());
            this.styleCur = null;
          };
        }
        return;
      case "listtable":
        this.lists.word(w, param);
        return;
      case "listoverride":
        this.lists.overrideWord(w, param);
        return;
      case "pict":
        this.pict?.word(w, param);
        return;
      case "shp":
        if (this.shape && (w === "shpleft" || w === "shptop" || w === "shpright" || w === "shpbottom")) this.shape[w] = param ?? 0;
        return;
      case "sn":
      case "shpPib":
        return;
      case "pn":
        if (w === "pnlvlblt") pa.pn = { ordered: false, level: pa.pn?.level ?? 0 };
        else if (w === "pnlvlbody") pa.pn = { ordered: true, level: pa.pn?.level ?? 0 };
        else if (w === "pnlvl") pa.pn = { ordered: true, level: Math.max(0, (param ?? 1) - 1) };
        return;
      case "info":
      case "infoText":
      case "fldinst":
      case "bkmk":
        if (w === "yr" || w === "mo" || w === "dy" || w === "hr" || w === "min") this.infoText += `${w}${param} `;
        return;
    }
    if (charWord(c, w, param) || paraWord(pa, w, param)) return;
    if (this.story.tables.word(w, param, (i) => this.colors[i] ?? null)) return;
    const pageKey = PAGE_WORDS[w];
    if (pageKey) {
      if (param !== null && param !== 0 && (pageKey === "width" || pageKey === "height")) this.setPage(pageKey, param);
      else if (param !== null && pageKey !== "width" && pageKey !== "height") this.setMargin(pageKey, param);
      return;
    }
    if (SPECIAL[w]) return this.bodyText(SPECIAL[w]);
    switch (w) {
      case "ansicpg":
        this.ansi = param ?? 1252;
        break;
      case "deff":
        this.deff = param ?? 0;
        break;
      case "landscape":
      case "lndscpsxn":
        if (this.story === this.main && !this.main.items.length) {
          this.page.orientation = "landscape";
          this.pageTouched = true;
        }
        break;
      case "titlepg":
        this.page.titlePage = true;
        break;
      case "sbknone":
        this.sectBreak = false;
        break;
      case "par":
        this.endParagraph(false);
        break;
      case "line":
        this.node({ type: "hardBreak" });
        break;
      case "page":
      case "column":
        this.endParagraph(true);
        if (this.g.dest === "body") this.story.items.push({ node: { type: w === "page" ? "pageBreak" : "columnBreak" } });
        break;
      case "sect":
        this.endParagraph(true);
        this.sectBreak = true;
        break;
      case "ftnalt":
        this.noteKind = "endnote";
        break;
      case "chpgn":
        this.node({ type: "pageField", attrs: { field: "PAGE" } });
        break;
      case "cell":
      case "nestcell":
        this.endParagraph(false, true);
        this.story.tables.endCell();
        break;
    }
  }

  /** Yazı tipi tablosundaki geçerli girdiyi kaydeder. */
  private saveFont() {
    const f = this.curFont;
    if (!f) return;
    const name = f.name.split(";")[0].trim();
    if (name || !this.fonts.has(f.id)) this.fonts.set(f.id, { name, charset: f.charset });
    this.curFont = null;
  }

  private setPage(k: "width" | "height", v: number) {
    if (this.story !== this.main || this.main.items.length) return;
    this.page[k] = Math.max(1440, Math.min(31680, v));
    this.pageTouched = true;
  }

  private setMargin(k: keyof BelgeDocument["settings"]["page"]["margin"], v: number) {
    if (this.story !== this.main || this.main.items.length) return;
    this.page.margin[k] = Math.max(0, Math.min(15840, Math.abs(v)));
    this.pageTouched = true;
  }

  /** Hedef sözcükleri; işlendiyse true. */
  private destination(w: string): boolean {
    const g = this.g;
    const end = (fn: () => void) => {
      const prev = g.onEnd;
      g.onEnd = () => {
        prev?.();
        fn();
      };
    };
    switch (w) {
      case "fonttbl":
        g.dest = "fonttbl";
        return true;
      case "colortbl":
        g.dest = "colortbl";
        return true;
      case "stylesheet":
        g.dest = "stylesheet";
        return true;
      case "info":
        g.dest = "info";
        return true;
      case "title":
      case "author":
      case "subject":
      case "keywords":
      case "creatim":
      case "revtim":
        if (g.dest !== "info") return false;
        g.dest = "infoText";
        this.infoKey = w;
        this.infoText = "";
        end(() => applyInfo(this.doc.settings.properties, this.infoKey, this.infoText.trim()));
        return true;
      case "pict": {
        if (g.dest !== "body" && g.dest !== "shpPib") return false;
        const shape = g.dest === "shpPib" ? this.shape : null;
        g.dest = "pict";
        this.pict = new RtfPict();
        end(() => {
          // Şekil içindeki resimde boyut şekil kutusundan alınır.
          const img = this.pict?.node(shape ? (shape.shpright - shape.shpleft) / 15 : 0, shape ? (shape.shpbottom - shape.shptop) / 15 : 0);
          this.pict = null;
          if (img) {
            this.story.inline.push(img);
            if (shape) shape.pict = true;
          }
        });
        return true;
      }
      case "shpinst":
        g.dest = "shp";
        this.shape = { shpleft: 0, shptop: 0, shpright: 0, shpbottom: 0, prop: "", pict: false };
        return true;
      case "sp":
        return g.dest === "shp";
      case "sn":
      case "sv":
        if (g.dest !== "shp" || !this.shape) return false;
        if (w === "sn") this.shape.prop = "";
        g.dest = w === "sn" ? "sn" : this.shape.prop.trim() === "pib" ? "shpPib" : "skip";
        return true;
      case "shprslt":
        // Şeklin resmi okunduysa yedek görünüm atlanır.
        if (this.shape?.pict) g.dest = "skip";
        return true;
      case "shppict":
      case "result":
      case "object":
      case "shp":
      case "fldrslt":
        if (w === "fldrslt") {
          const f = this.fields[this.fields.length - 1];
          if (f) {
            f.result = true;
            const href = /^\s*HYPERLINK\s+(?:"([^"]*)"|(\S+))(?:.*?\\l\s+"([^"]*)")?/i.exec(f.instr);
            if (href) g.chr.link = (href[1] ?? href[2] ?? "") + (href[3] ? `#${href[3]}` : "");
            const kw = PAGE_FIELD.exec(f.instr);
            if (kw) this.node(pageFieldNode(kw[1].toUpperCase(), f.instr));
          }
        }
        return true;
      case "field":
        this.fields.push({ instr: "", result: false });
        end(() => {
          const f = this.fields.pop();
          // Sonucu olmayan sayfa alanı
          const kw = f && !f.result ? PAGE_FIELD.exec(f.instr) : null;
          if (kw && f) this.node(pageFieldNode(kw[1].toUpperCase(), f.instr));
        });
        return true;
      case "fldinst":
        g.dest = "fldinst";
        return true;
      case "listtable":
        g.dest = "listtable";
        return true;
      case "listoverridetable":
        g.dest = "listoverride";
        return true;
      case "listoverride":
        if (g.dest !== "listoverride") return false;
        this.lists.startOverride();
        end(() => this.lists.endOverride());
        return true;
      case "pn":
        g.dest = "pn";
        end(() => {
          const parent = this.stack[this.stack.length - 1];
          if (parent && g.par.pn) parent.par.pn = g.par.pn;
        });
        return true;
      case "bkmkstart":
        g.dest = "bkmk";
        this.bkmk = "";
        end(() => {
          const name = this.bkmk.trim();
          if (name && name !== "_GoBack") this.story.inline.push({ type: "bookmark", attrs: { name } });
        });
        return true;
      case "footnote":
        this.startStory(null);
        return true;
      case "header":
      case "headerr":
      case "footer":
      case "footerr":
      case "headerf":
      case "footerf":
      case "headerl":
      case "footerl":
        this.startStory((w.startsWith("header") ? "header" : "footer") + (w.endsWith("f") ? "First" : w.endsWith("l") ? "Even" : ""));
        return true;
      default:
        if (SKIP_DESTS.has(w)) {
          g.dest = "skip";
          return true;
        }
        return false;
    }
  }

  private noteKind: "footnote" | "endnote" = "footnote";
  /** \sect sonrası bekleyen sayfa sonu (\sbknone ile iptal). */
  private sectBreak = false;

  /** Ayrı metin akışı (dipnot ya da üst/alt bilgi): grup bitince sonuç yerine konur. */
  private startStory(hfKey: string | null) {
    const outer = this.story;
    const story = new Story();
    story.note = !hfKey;
    this.story = story;
    this.noteKind = "footnote";
    const g = this.g;
    g.par = newPar();
    g.chr = { ...newChr(), uc: g.chr.uc };
    const prev = g.onEnd;
    g.onEnd = () => {
      prev?.();
      this.endParagraph(true);
      story.tables.close(story.items);
      this.story = outer;
      const blocks = foldLists(story.items);
      if (!hfKey) outer.inline.push(noteNode(blocks, this.noteKind, `nr${++this.noteSeq}`, noteText));
      else if (JSON.stringify(blocks).match(/"text":"[^"]*\S|pageField|"image"/)) (this.doc.settings as unknown as Record<string, unknown>)[hfKey] = { type: "doc", content: blocks };
    };
  }

  /* -------------------------- paragraf ve liste -------------------------- */

  /** Paragraf sonu: biriken satır içi içerik paragraf düğümüne dönüşür. */
  private endParagraph(soft: boolean, cell = false) {
    if (this.g.dest !== "body" && !soft && !cell) return;
    const st = this.story;
    const pa = this.g.par;
    if ((soft || (cell && st.tables.hasBlocks())) && !st.inline.length) return;
    const level = headingLevelOf(this.styleNames.get(pa.s), pa.outline);
    const styleId = level ? null : styleIdByName(this.styleNames.get(pa.s));
    const ours = level ? `Heading${level}` : (styleId ?? "Normal");
    const base = styleChar(stylesOf(), ours).base;
    const content: JSONContent[] = st.inline.map((n) => {
      const props = (n.attrs as { _props?: CharProps } | undefined)?._props;
      if (n.type !== "text" || !props) return n;
      // Yazıcının bağlantılara verdiği renk biçim sayılmaz.
      const own: CharProps = { ...props };
      // Bağlantı stilinin rengi/alt çizgisi ve not metninin boyutu biçim sayılmaz.
      if (own.link && own.color && (own.color === LINK_COLOR || /^#(0000FF|000080)$/i.test(own.color))) own.color = null;
      if (own.link && own.underline === "single") own.underline = null;
      if (st.note) own.size = own.font = null;
      const marks = charMarks(stripStyleProps(own, base));
      return { type: "text", text: n.text, ...(marks.length ? { marks } : {}) };
    });
    st.inline = [];
    const attrs: Record<string, unknown> = {};
    if (level) attrs.level = level;
    else if (styleId) attrs.styleId = styleId;
    if (pa.align && pa.align !== "left") attrs.textAlign = pa.align;
    const list = this.lists.info(pa);
    const tw = (v: number | null) => (v === null ? null : Math.round((v / 20) * 100) / 100);
    if (!list) {
      if (pa.li) attrs.indentLeft = tw(pa.li);
      if (pa.fi) attrs.indentFirst = tw(pa.fi);
    }
    if (pa.ri) attrs.indentRight = tw(pa.ri);
    // Stilin zaten verdiği paragraf özellikleri doğrudan biçim sayılmaz.
    const near = (a: number | null, b: number | undefined) => a !== null && Math.abs(a - (b ?? 0)) < 0.06;
    if (pa.sb !== null && !near(tw(pa.sb), base.spaceBefore)) attrs.spaceBefore = tw(pa.sb);
    if (pa.sa !== null && !near(tw(pa.sa), base.spaceAfter)) attrs.spaceAfter = tw(pa.sa);
    if (pa.sl !== null && pa.sl !== 0) {
      if (pa.slmult) {
        const lh = Math.round((Math.abs(pa.sl) / 240) * 100) / 100;
        if (!near(lh, base.line ?? 1)) attrs.lineHeight = lh;
      } else {
        attrs.lineHeight = tw(Math.abs(pa.sl));
        attrs.lineRule = pa.sl < 0 ? "exact" : "atLeast";
      }
    }
    if (attrs.textAlign && attrs.textAlign === base.align) delete attrs.textAlign;
    if (pa.keepn && !base.keepNext) attrs.keepNext = true;
    if (pa.keep) attrs.keepLines = true;
    if (pa.pagebb) attrs.pageBreakBefore = true;
    if (pa.shade) attrs.shading = this.colors[pa.shade] ?? null;
    const node = !content.length && pa.brdrb && !level && !cell && !pa.intbl ? { type: "horizontalRule" } : paragraphNode(attrs, content, level ? "heading" : "paragraph");
    if (pa.intbl || cell) {
      st.tables.addBlock(node);
      return;
    }
    st.tables.close(st.items);
    if (this.sectBreak && st === this.main) {
      if (st.items.length) st.items.push({ node: { type: "pageBreak" } });
      this.sectBreak = false;
    }
    st.items.push({ node, list: level ? undefined : list });
  }
}

/** RTF baytları → belge. */
export function readRtf(data: ArrayBuffer | Uint8Array): BelgeDocument {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  return new RtfReader(bytes).read();
}
