// RTF okuyucu (RTF 1.9.1): grup yığını, yazı tipi/renk/stil tabloları, kod sayfaları, \uN, karakter ve
// paragraf biçimi, listeler, tablolar, resimler (\pngblip/\jpegblip), alanlar, dipnotlar, üst/alt bilgi.
import type { JSONContent } from "@tiptap/core";
import { defaultSettings, normalizeDoc, type BelgeDocument } from "../../model/document";
import { noteText } from "../../editor/notes";
import { fieldPicture } from "../../editor/nodes";
import { pictureOfInstr } from "../../editor/dateFormat";
import { charMarks, extendCommentToWord, foldLists, noteNode, paragraphNode, stripStyleProps, styleChar, stylesOf, type CharProps, type FlowItem } from "../common";
import { charsetCodepage, decodeBytes } from "./codepage";
import { RtfTables } from "./tables";
import { RtfPict } from "./pict";
import { RtfLists } from "./lists";
import { HL_COLORS, LINK_COLOR, SKIP_DESTS, charWord, dttmIso, newChr, newPar, paraWord, headingLevelOf, styleIdByName, applyInfo, type Chr, type Par } from "./state";
import { LoadWarnings } from "../warnings";
import type { StyleDef } from "../../model/styles";

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
  | "shpPib"
  | "revtbl"
  | "atnauthor"
  | "atnid"
  | "annotation"
  | "atnval";

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
  private styleCur: { id: number; name: string; chr: Chr; par: Par } | null = null;
  /** Kaynağın Normal (\s0) stil tanımı (Normal stile taşınır). */
  private normalDef: { chr: Chr; par: Par } | null = null;
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
  /** Okunamayan (kaydedince kaybolacak) içerik uyarıları. */
  warnings = new LoadWarnings();
  /** Değişiklik izleme yazarları (\revtbl) ve açıklamalar (\annotation). */
  private revText = "";
  private revAuthors: string[] = [];
  private atn = { author: "", initials: "", text: "", date: "", ref: null as string | null };
  private pendingAtn: string[] = [];
  /** Açıklama aralıkları (\atrfstart N … \atrfend N) ve \atnref N → açıklama kimliği. */
  private openAtr = new Set<string>();
  private atrSeen = new Set<string>();
  private atrMap = new Map<string, string>();
  private atnVal = "";
  /** Son \par'dan sonra \pard görüldü mü (son paragraf \par olmadan da vardır). */
  private pardAfterPar = false;
  private styleTable: Record<string, StyleDef> | null = null;
  /** Açılan gruplar kapananlardan çoksa (bozuk dosya) hoşgörülü okuma. */
  private lenient = false;
  /** Derinlik sınırını aşan (açılmayan) grup sayısı. */
  private overflow = 0;

  /** Kaçışlı olmayan { ve } sayıları dengesiz mi (açılan > kapanan). */
  private unbalanced(): boolean {
    const b = this.b;
    let depth = 0;
    for (let i = 0; i < b.length; i++) {
      const c = b[i];
      if (c === 0x5c) i++;
      else if (c === 0x7b) depth++;
      else if (c === 0x7d) depth--;
    }
    return depth > 1;
  }

  constructor(private readonly b: Uint8Array) {}

  read(): BelgeDocument {
    // Baştaki BOM ve boşluklar atlanır.
    while (this.pos < Math.min(64, this.b.length) && this.b[this.pos] !== 0x7b) this.pos++;
    const head = new TextDecoder("latin1").decode(this.b.subarray(this.pos, this.pos + 6));
    if (!head.startsWith("{\\rtf")) throw new Error("Geçerli bir RTF dosyası değil.");
    this.lenient = this.unbalanced();
    this.g = { chr: newChr(), par: newPar(), dest: "body", ignorable: false, first: false };
    let guard = 0;
    while (this.pos < this.b.length) {
      if (++guard > 50_000_000) throw new Error("RTF dosyası çok büyük ya da bozuk.");
      const c = this.b[this.pos];
      if (c === 0x7b) {
        this.flushText();
        // Aşırı derin gruplar düzleştirilir (ayrı grup açılmaz; eşleşen kapanışlar da yok sayılır).
        if (this.stack.length > 2000) {
          this.overflow++;
          this.warnings.add("deep", "RTF grup derinliği");
          this.pos++;
          continue;
        }
        this.stack.push(this.g);
        this.g = { chr: { ...this.g.chr }, par: { ...this.g.par, pn: this.g.par.pn }, dest: this.g.dest, ignorable: false, first: true };
        this.skipUc = 0;
        this.pos++;
      } else if (c === 0x7d) {
        this.flushText();
        if (this.overflow > 0) {
          this.overflow--;
          this.pos++;
          continue;
        }
        this.g.onEnd?.();
        // Kök grup kapanıyor: son paragraf kendi (grubun içindeki) biçimiyle biter.
        if (this.stack.length === 1 && this.g.dest === "body" && this.story === this.main) this.endDocument();
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
    if (this.stack.length) this.warnings.add("unclosed", `kapanmamış grup: ${this.stack.length}`);
    this.endDocument();
    this.story.tables.close(this.story.items);
    const blocks = foldLists(this.main.items);
    if (this.atrSeen.size) this.resolveRanges(blocks);
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
      // Dosyanın kalanından uzun ikili veri bozukluktur: atlanmaz (içerik okunmaya devam eder).
      if ((param ?? 0) > b.length - p) {
        this.warnings.add("corrupt", "\\bin uzunluğu dosyayı aşıyor");
        return;
      }
      const n = Math.max(0, param ?? 0);
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
      case "revtbl":
        this.revText += s;
        break;
      case "atnauthor":
        this.atn.author += s;
        break;
      case "atnid":
        this.atn.initials += s;
        break;
      case "annotation":
        this.atn.text += s;
        break;
      case "atnval":
        this.atnVal += s;
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
    const rev = c.rev ? { type: c.rev === "ins" ? "insertion" : "deletion", attrs: { author: this.revAuthors[c.rauth] ?? "", date: dttmIso(c.rdate) } } : null;
    this.story.inline.push({ type: "text", text: s, attrs: { _props: props, _rev: rev, _atr: this.openAtr.size ? [...this.openAtr] : null } } as JSONContent);
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
    // Grupları kapanmamış (dengesiz) bozuk dosya: paragraf başlangıcı gövdeye dönüldüğünü gösterir; aksi hâlde
    // belgenin tamamı ilk kapanmamış tablo grubunda kalır ve boş açılırdı.
    if (this.lenient && (w === "pard" || w === "sectd") && g.dest !== "body" && g.dest !== "pict") {
      g.dest = "body";
      g.ignorable = false;
      // Kapanmamış alanlar da biter (yoksa sonraki metin alan yönergesi sayılırdı).
      this.fields = [];
      g.chr.link = null;
    }
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
          const cur = { id: w === "s" ? (param ?? 0) : -1 - (param ?? 0), name: "", chr: newChr(), par: newPar() };
          this.styleCur = cur;
          const prev = g.onEnd;
          g.onEnd = () => {
            prev?.();
            this.styleNames.set(cur.id, cur.name.replace(/;[\s\S]*$/, "").trim());
            if (cur.id === 0) this.normalDef = { chr: cur.chr, par: cur.par };
            this.styleCur = null;
          };
        } else if (this.styleCur) charWord(this.styleCur.chr, w, param) || paraWord(this.styleCur.par, w, param);
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
      case "annotation":
        if (w === "par" && this.atn.text) this.atn.text += "\n";
        return;
      case "revtbl":
      case "atnauthor":
      case "atnid":
      case "atnval":
        return;
    }
    if (w === "pard" && g.dest === "body" && this.story === this.main) this.pardAfterPar = true;
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
      case "shpinst": {
        g.dest = "shp";
        const shape = { shpleft: 0, shptop: 0, shpright: 0, shpbottom: 0, prop: "", pict: false };
        this.shape = shape;
        // Resmi olmayan şekil (çizim, metin kutusu) okunamaz.
        end(() => {
          if (!shape.pict) this.warnings.add("shapes", "\\shp");
        });
        return true;
      }
      case "revtbl":
        g.dest = "revtbl";
        this.revText = "";
        end(() => {
          this.revAuthors = this.revText.split(";").map((x) => x.trim());
        });
        return true;
      case "atnauthor":
      case "atnid":
        g.dest = w;
        if (w === "atnauthor") this.atn.author = "";
        else this.atn.initials = "";
        return true;
      case "annotation":
        g.dest = "annotation";
        this.atn.text = "";
        end(() => this.addAnnotation());
        return true;
      case "atnref":
      case "atndate":
      case "atrfstart":
      case "atrfend": {
        // Küçük değer grupları: açıklama başvuru numarası, tarihi ve aralık sınırları.
        g.dest = "atnval";
        this.atnVal = "";
        end(() => {
          const v = this.atnVal.trim();
          if (w === "atnref") this.atn.ref = v;
          else if (w === "atndate") this.atn.date = dttmIso(Number(v) >>> 0);
          else if (w === "atrfstart") {
            this.openAtr.add(v);
            this.atrSeen.add(v);
          } else this.openAtr.delete(v);
        });
        return true;
      }
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
      case "object":
        // Gömülü nesne: yalnız önizleme resmi (\result) okunur.
        this.warnings.add("objects", "\\object");
        return true;
      case "shppict":
      case "result":
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
          const name = f?.instr.trim().split(/\s+/)[0]?.toUpperCase();
          if (name && !/^(HYPERLINK|SHAPE)$/.test(name) && !PAGE_FIELD.test(name)) this.warnings.add("fields", `alan: ${name}`);
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
        // Dengesiz dosyada grubun sonu bilinemez: not/üst-alt bilgi metni gövdeye akar (içerik kaybolmaz).
        if (!this.lenient) this.startStory(null);
        return true;
      case "header":
      case "headerr":
      case "footer":
      case "footerr":
      case "headerf":
      case "footerf":
      case "headerl":
      case "footerl":
        if (!this.lenient) this.startStory((w.startsWith("header") ? "header" : "footer") + (w.endsWith("f") ? "First" : w.endsWith("l") ? "Even" : ""));
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

  /** Belgenin stil tablosu: kaynağın varsayılanları (varsayılan yazı tipi, 12 nk, aralıksız tek satır) Normal stile taşınır. */
  private styles(): Record<string, StyleDef> {
    if (this.styleTable) return this.styleTable;
    const s = this.doc.settings;
    const ours = stylesOf().Normal;
    // Normal (\s0) stil tanımı varsa onun değerleri; yoksa RTF varsayılanları.
    const n = this.normalDef;
    const font = this.fonts.get(n?.chr.f ?? this.deff)?.name || ours.props.font;
    const size = n?.chr.fs ? n.chr.fs / 2 : 12;
    const pt = (v: number | null | undefined) => (v ? Math.round((v / 20) * 100) / 100 : 0);
    const line = n?.par.sl && n.par.slmult ? Math.round((Math.abs(n.par.sl) / 240) * 100) / 100 : 1;
    const props = { ...ours.props, font, size, spaceBefore: pt(n?.par.sb), spaceAfter: pt(n?.par.sa), line };
    const same = (Object.keys(props) as (keyof typeof props)[]).every((k) => props[k] === ours.props[k] || (typeof props[k] === "number" && Math.abs((props[k] as number) - ((ours.props[k] as number | undefined) ?? NaN)) < 0.01));
    if (!same) s.customStyles = { ...s.customStyles, Normal: { ...ours, props } };
    this.styleTable = stylesOf(s);
    return this.styleTable;
  }

  /** Belge sonu: açık paragraf biter; son \par'dan sonra yeni paragraf başladıysa (boş da olsa) korunur. */
  private endDocument() {
    this.endParagraph(true);
    if (this.pardAfterPar && this.story === this.main && !this.story.inline.length && this.g.dest === "body") this.endParagraph(false);
    this.pardAfterPar = false;
    if (this.pendingAtn.length) this.warnings.add("comments", "bağlanamayan açıklama");
  }

  /** Aralıklı açıklamaların yer tutucu kimlikleri gerçek kimliklerle değiştirilir (eşlenmeyenler atılır). */
  private resolveRanges(blocks: JSONContent[]) {
    const walk = (n: JSONContent) => {
      if (n.marks?.some((m) => m.type === "comment" && String(m.attrs?.id).startsWith("__atr:"))) {
        n.marks = n.marks.flatMap((m) => {
          const id = String(m.attrs?.id);
          if (m.type !== "comment" || !id.startsWith("__atr:")) return [m];
          const real = this.atrMap.get(id.slice(6));
          return real ? [{ type: "comment", attrs: { id: real } }] : [];
        });
        if (!n.marks.length) delete n.marks;
      }
      n.content?.forEach(walk);
    };
    blocks.forEach(walk);
  }

  /** \annotation grubu bitti: açıklama kaydedilir ve konumuna yer tutucu bırakılır. */
  private addAnnotation() {
    const comments = this.doc.settings.comments;
    let k = Object.keys(comments).length;
    while (comments[`rc${k}`]) k++;
    const id = `rc${k}`;
    const author = this.atn.author.trim();
    const initials = this.atn.initials.trim() || author.split(/\s+/).filter(Boolean).map((w) => w[0]!.toLocaleUpperCase("tr-TR")).join("").slice(0, 3) || "K";
    comments[id] = { id, author, initials, date: this.atn.date || new Date().toISOString(), text: this.atn.text.replace(/\s+$/, "").trim() };
    const ref = this.atn.ref;
    this.atn = { author: "", initials: "", text: "", date: "", ref: null };
    // Aralıklı açıklama: aralıktaki metinler yer tutucu kimlikle işaretlendi; belge sonunda eşlenir.
    if (ref !== null && this.atrSeen.has(ref)) this.atrMap.set(ref, id);
    else if (this.story === this.main && this.stack.some((x) => x.dest === "body")) this.story.inline.push({ type: "__atn", attrs: { id } });
    else this.warnings.add("comments", "üst/alt bilgi ya da not içinde açıklama");
  }

  private takeAtn(): { type: string; attrs: { id: string } }[] {
    const out = this.pendingAtn.map((id) => ({ type: "comment", attrs: { id } }));
    for (const id of this.pendingAtn) this.pointIds.add(id);
    this.pendingAtn = [];
    return out;
  }

  /** Bu paragrafta noktaya bağlanan açıklamalar (sözcüğün tamamına yayılır). */
  private pointIds = new Set<string>();

  /* -------------------------- paragraf ve liste -------------------------- */

  /** Paragraf sonu: biriken satır içi içerik paragraf düğümüne dönüşür. */
  private endParagraph(soft: boolean, cell = false) {
    if (this.g.dest !== "body" && !soft && !cell) return;
    const st = this.story;
    const pa = this.g.par;
    if ((soft || (cell && st.tables.hasBlocks())) && !st.inline.length) return;
    if (st === this.main) this.pardAfterPar = false;
    const level = headingLevelOf(this.styleNames.get(pa.s), pa.outline);
    const styleId = level ? null : styleIdByName(this.styleNames.get(pa.s));
    const ours = level ? `Heading${level}` : (styleId ?? "Normal");
    const base = styleChar(this.styles(), ours).base;
    const converted: JSONContent[] = st.inline.map((n) => {
      const props = (n.attrs as { _props?: CharProps } | undefined)?._props;
      if (n.type !== "text" || !props) return n;
      const rev = (n.attrs as { _rev?: { type: string; attrs: Record<string, unknown> } | null })._rev;
      const atr = (n.attrs as { _atr?: string[] | null })._atr ?? [];
      // Yazıcının bağlantılara verdiği renk biçim sayılmaz.
      const own: CharProps = { ...props };
      // Bağlantı stilinin rengi/alt çizgisi ve not metninin boyutu biçim sayılmaz.
      if (own.link && own.color && (own.color === LINK_COLOR || /^#(0000FF|000080)$/i.test(own.color))) own.color = null;
      if (own.link && own.underline === "single") own.underline = null;
      if (st.note) own.size = own.font = null;
      const marks: JSONContent["marks"] = [...charMarks(stripStyleProps(own, base)), ...(rev ? [rev] : []), ...atr.map((k) => ({ type: "comment", attrs: { id: `__atr:${k}` } }))];
      return { type: "text", text: n.text, ...(marks.length ? { marks } : {}) };
    });
    st.inline = [];
    // Açıklamalar (\annotation): ardından gelen ilk sözcüğe, yoksa paragrafın son metnine bağlanır.
    const content: JSONContent[] = [];
    for (const n of converted) {
      if (n.type === "__atn") {
        this.pendingAtn.push(String(n.attrs?.id));
        continue;
      }
      if (n.type === "text" && n.text && this.pendingAtn.length) {
        const m = /^\s*\S+/.exec(n.text);
        const head = m ? m[0] : n.text;
        content.push({ ...n, text: head, marks: [...(n.marks ?? []), ...this.takeAtn()] });
        if (head.length < n.text.length) content.push({ ...n, text: n.text.slice(head.length) });
        continue;
      }
      content.push(n);
    }
    if (this.pendingAtn.length) {
      const last = [...content].reverse().find((n) => n.type === "text" && n.text);
      if (last) last.marks = [...(last.marks ?? []), ...this.takeAtn()];
    }
    if (this.pointIds.size) {
      extendCommentToWord(content, this.pointIds);
      this.pointIds.clear();
    }
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
export function readRtf(data: ArrayBuffer | Uint8Array, warnings?: LoadWarnings): BelgeDocument {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const r = new RtfReader(bytes);
  if (warnings) r.warnings = warnings;
  return r.read();
}
