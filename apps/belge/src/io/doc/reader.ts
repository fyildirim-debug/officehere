// Eski ikili .doc (97-2003) okuyucu: CFB → WordDocument/Table akışları → FIB → parça tablosu → metin;
// CHPX/PAPX FKP'lerinden karakter ve paragraf biçimi, stil sayfasından başlıklar, listeler, tablolar,
// alanlar (köprü, sayfa numarası), dipnot/son notlar, üst/alt bilgi, satır içi resimler, bölümler.
import type { JSONContent } from "@tiptap/core";
import { isCfb, readCfb } from "@officehere/core";
import { defaultSettings, normalizeDoc, samePage, type BelgeDocument, type CommentData, type PageSetup } from "../../model/document";
import { noteText } from "../../editor/notes";
import type { StyleProps } from "../../model/styles";
import { bulletStyleAt, charMarks, extendCommentToWord, foldLists, imageNode, noteNode, paragraphNode, stripStyleProps, styleChar, stylesOf, type CharProps, type FlowItem, type ListInfo } from "../common";
import { RtfTables } from "../rtf/tables";
import { Bytes, FC, parseFib, plc, type Fib } from "./fib";
import { applyChpx, applyPapx, applySepx, newSep, parseTDef, parseTShd, type Chp, type Pap, type SectionProps } from "./sprm";
import { Lists, StyleSheet, parseBStore, type PictureStore, cpRuns, parseFkps, parseFonts, parsePieces, pieceText, readPicture, runAt, type CpRun, type Piece } from "./structs";
import { dttmIso } from "../rtf/state";
import { LoadWarnings } from "../warnings";

const KNOWN: Record<string, string> = { title: "Title", subtitle: "Subtitle", quote: "Quote", "intense quote": "IntenseQuote", "no spacing": "NoSpacing", "list paragraph": "ListParagraph", caption: "Caption" };
const NFC: Record<number, string | null> = { 0: null, 1: "upper-roman", 2: "lower-roman", 3: "upper-alpha", 4: "lower-alpha" };
const BULLETS: Record<string, string> = { "o": "circle", "○": "circle", "§": "square", "■": "square", "▪": "square", "–": "dash", "-": "dash", "ü": "check", "✓": "check", "Ø": "arrow", "➢": "arrow", "v": "diamond", "❖": "diamond" };

/** Bir metin akışının (gövde, not, üst/alt bilgi) kurucusu. */
class Builder {
  items: FlowItem[] = [];
  inline: (JSONContent | "page" | "column")[] = [];
  tables = new RtfTables();
}

interface Field {
  instr: string;
  result: boolean;
  href: string | null;
  suppress: boolean;
}

class DocReader {
  pieces: Piece[];
  text: string;
  chp: CpRun[];
  pap: CpRun[];
  stsh: StyleSheet;
  fonts: string[];
  lists: Lists;
  bstore: PictureStore;
  sections: { end: number; sep: SectionProps }[] = [];
  foot: { refs: Map<number, number>; txt: number[]; base: number };
  end: { refs: Map<number, number>; txt: number[]; base: number };
  private counts = new Map<string, number>();
  private noteSeq = 0;
  /** Bozuk yapı nedeniyle atlanan karakter sayısı (uyarı için). */
  skipped = 0;
  /** Değişiklik izleme yazarları (SttbfRMark). */
  private revAuthors: string[] = [];
  /** Açıklamalar: başvuru CP → sıra; açıklama metni aralıkları; yazarlar (GrpXstAtnOwners). */
  private atn: { refs: Map<number, { k: number; ibst: number; initials: string }>; txt: number[]; base: number; owners: string[] } = { refs: new Map(), txt: [], base: 0, owners: [] };
  comments: Record<string, CommentData> = {};
  private pendingComments: string[] = [];
  /** Bu paragrafta noktaya bağlanan açıklamalar (sözcüğün tamamına yayılır). */
  private pointIds = new Set<string>();
  private chpCache = new Map<string, CharProps>();

  constructor(
    readonly wd: Bytes,
    readonly t: Bytes,
    readonly data: Bytes | null,
    readonly fib: Fib,
    /** Okunamayan (kaydedince kaybolacak) içerik uyarıları. */
    readonly warnings: LoadWarnings = new LoadWarnings(),
  ) {
    this.pieces = parsePieces(t, fib);
    const total = fib.ccpText + fib.ccpFtn + fib.ccpHdd + fib.ccpMcr + fib.ccpAtn + fib.ccpEdn + fib.ccpTxbx + fib.ccpHdrTxbx + 1;
    this.text = pieceText(wd, this.pieces, total);
    this.chp = cpRuns(this.pieces, parseFkps(wd, t, fib, false));
    this.pap = cpRuns(this.pieces, parseFkps(wd, t, fib, true));
    this.stsh = new StyleSheet(t, fib);
    this.fonts = parseFonts(t, fib);
    this.lists = new Lists(t, fib);
    this.bstore = parseBStore(t, fib, wd);
    const notes = (refIdx: number, txtIdx: number, base: number) => {
      const r = fib.fcLcb(refIdx);
      const x = fib.fcLcb(txtIdx);
      const refs = new Map<number, number>();
      plc(t, r.fc, r.lcb, 2).cps.slice(0, -1).forEach((cp, i) => refs.set(cp, i));
      const txt = x.lcb ? Array.from({ length: Math.floor(x.lcb / 4) }, (_, i) => t.i32(x.fc + i * 4)) : [];
      return { refs, txt, base };
    };
    this.foot = notes(FC.PlcffndRef, FC.PlcffndTxt, fib.ccpText);
    try {
      this.revAuthors = sttb(t, fib.fcLcb(FC.SttbfRMark));
    } catch {
      this.revAuthors = [];
    }
    try {
      this.atn = this.readAtnTables();
    } catch {
      if (fib.ccpAtn) this.warnings.add("comments", "açıklama tablosu bozuk");
    }
    this.end = notes(FC.PlcfendRef, FC.PlcfendTxt, fib.ccpText + fib.ccpFtn + fib.ccpHdd + fib.ccpMcr + fib.ccpAtn);
    const sed = fib.fcLcb(FC.PlcfSed);
    const s = plc(t, sed.fc, sed.lcb, 12);
    s.data.forEach((at, i) => {
      const sep = newSep();
      const fcSepx = t.i32(at + 2);
      if (fcSepx >= 0 && fcSepx + 2 <= wd.length) {
        const cb = wd.i16(fcSepx);
        if (cb > 0 && fcSepx + 2 + cb <= wd.length) applySepx(sep, wd.slice(fcSepx + 2, cb));
      }
      this.sections.push({ end: s.cps[i + 1], sep });
    });
  }

  /** CP konumundaki paragraf özellikleri (paragraf işaretinin PAPX'i). */
  papAt(cp: number): Pap {
    const r = runAt(this.pap, cp);
    const p: Pap = { ...this.stsh.pap(r?.run.istd ?? 0) };
    if (r) {
      p.istd = r.run.istd;
      applyPapx(p, r.run.grpprl);
      if (r.piece.prm) applyPapx(p, r.piece.prm);
    }
    return p;
  }

  /** Koşunun karakter özellikleri (paragraf stili → karakter stili → doğrudan biçim). */
  charAt(cp: number, istd: number): { props: CharProps; chp: Chp; end: number } {
    const r = runAt(this.chp, cp);
    const base = this.stsh.chp(istd);
    const c: Chp = { ...base };
    if (r) {
      const probe = { ...c };
      applyChpx(probe, r.run.grpprl, base);
      const charIstd = probe.istd;
      const cs = charIstd !== null ? this.stsh.styles[charIstd] : null;
      if (cs?.chpx) applyChpx(c, cs.chpx, base);
      applyChpx(c, r.run.grpprl, base);
      if (r.piece.prm) applyChpx(c, r.piece.prm, base);
    }
    const key = `${r?.cpStart ?? -1}|${istd}`;
    let props = this.chpCache.get(key);
    if (!props) {
      props = { bold: c.bold, italic: c.italic, underline: c.underline, strike: c.strike, dstrike: c.dstrike, sup: c.iss === 1, sub: c.iss === 2, font: c.ftc !== null ? (this.fonts[c.ftc] ?? null) : null, size: c.hps / 2, color: c.color, highlight: c.highlight, bg: c.bg, caps: c.caps, smallCaps: c.smallCaps, hidden: c.hidden, spacing: c.spacing ? c.spacing / 20 : null, shadow: c.shadow, outline: c.outline, emboss: c.emboss, imprint: c.imprint };
      this.chpCache.set(key, props);
    }
    return { props, chp: c, end: r ? r.cpEnd : cp + 1 };
  }

  /** Paragraf stili → bizim stilimiz. */
  styleOf(istd: number, outline: number | null): { id: string; level: number | null } {
    const s = this.stsh.styles[istd];
    if (s && s.sti >= 1 && s.sti <= 9) return { id: `Heading${Math.min(6, s.sti)}`, level: Math.min(6, s.sti) };
    const name = (s?.name ?? "").toLowerCase().trim();
    const h = /^(?:heading|başlık)\s*(\d)$/.exec(name);
    if (h) return { id: `Heading${Math.min(6, Number(h[1]))}`, level: Math.min(6, Number(h[1])) };
    if (s?.sti === 62) return { id: "Title", level: null };
    if (s?.sti === 74) return { id: "Subtitle", level: null };
    if (KNOWN[name]) return { id: KNOWN[name], level: null };
    if (outline !== null && outline < 6 && name && name !== "normal") return { id: `Heading${outline + 1}`, level: outline + 1 };
    return { id: "Normal", level: null };
  }

  private nextEnd(i: number, end: number): number {
    for (let j = i; j < end; j++) {
      const c = this.text.charCodeAt(j);
      if (c === 0x0d || c === 0x07 || (c === 0x0c && this.isSectionEnd(j + 1))) return j;
    }
    return end - 1;
  }

  private isSectionEnd(cp: number): boolean {
    return this.sections.some((s) => s.end === cp && cp < this.fib.ccpText);
  }

  /** Metin aralığını (bir hikâye) bloklara çevirir. */
  /** Kapsayıcı hikâyelerde (not, üst/alt bilgi) paragraf stilinin kendi biçimi temel alınır. */
  private docBase(istd: number): StyleProps {
    const c = this.stsh.chp(istd);
    const p = this.stsh.pap(istd);
    const pt = (tw: number | null) => (tw === null ? undefined : Math.round((tw / 20) * 100) / 100);
    const line = p.line && p.line.mult ? Math.round((p.line.dya / 240) * 100) / 100 : 1;
    return { font: c.ftc !== null ? this.fonts[c.ftc] : undefined, size: c.hps / 2, bold: c.bold, italic: c.italic, color: c.color ?? undefined, caps: c.caps, spaceBefore: pt(p.before) ?? 0, spaceAfter: pt(p.after) ?? 0, line, indentLeft: pt(p.left), indentRight: pt(p.right), keepNext: p.keepNext, align: p.jc === 1 ? "center" : p.jc === 2 ? "right" : p.jc === 3 ? "justify" : "left" };
  }

  story(start: number, end: number, main: boolean, container = false): FlowItem[] {
    const b = new Builder();
    const fields: Field[] = [];
    let i = start;
    let guard = 0;
    while (i < end && guard++ < 5_000_000) {
      const paraEnd = this.nextEnd(i, end);
      const pap = this.papAt(paraEnd);
      const map = this.styleOf(pap.istd, pap.outline);
      const base = container ? this.docBase(pap.istd) : styleChar(stylesOf(), map.id).base;
      const firstBase = container ? this.stsh.pap(pap.istd).first : null;
      for (; i <= paraEnd && i < end; ) {
        const code = this.text.charCodeAt(i);
        const f = fields[fields.length - 1];
        if (code === 0x13) {
          fields.push({ instr: "", result: false, href: null, suppress: false });
          i++;
          continue;
        }
        if (code === 0x14 && f) {
          f.result = true;
          const href = /^\s*HYPERLINK\s+(?:"([^"]*)"|(\S+))(?:.*?\\l\s+"([^"]*)")?/i.exec(f.instr);
          if (href) f.href = (href[1] ?? href[2] ?? "") + (href[3] ? `#${href[3]}` : "");
          const kw = /^\s*(PAGE|NUMPAGES)\b/i.exec(f.instr);
          if (kw) {
            b.inline.push({ type: "pageField", attrs: { field: kw[1].toUpperCase() } });
            f.suppress = true;
          }
          i++;
          continue;
        }
        if (code === 0x15 && f) {
          fields.pop();
          const kw = !f.result ? /^\s*(PAGE|NUMPAGES)\b/i.exec(f.instr) : null;
          if (kw) b.inline.push({ type: "pageField", attrs: { field: kw[1].toUpperCase() } });
          const name = f.instr.trim().split(/\s+/)[0]?.toUpperCase();
          if (name && !/^(HYPERLINK|PAGE|NUMPAGES|SHAPE)$/.test(name)) this.warnings.add("fields", `alan: ${name}`);
          i++;
          continue;
        }
        if (f && !f.result && i < paraEnd) {
          f.instr += this.text[i];
          i++;
          continue;
        }
        if (i === paraEnd) {
          this.endParagraph(b, pap, map, base, code, main && i + 1 < end, firstBase);
          i++;
          break;
        }
        const ch = this.charAt(i, pap.istd);
        // Bozuk biçim koşusu (bitişi bu konumda ya da gerisinde) döngüyü durdurmasın: en az bir karakter ilerlenir.
        const stop = Math.max(i + 1, Math.min(ch.end, paraEnd));
        if (fields.some((x) => x.suppress && x.result)) {
          i++;
          continue;
        }
        const link = [...fields].reverse().find((x) => x.href)?.href ?? null;
        let run = "";
        const flush = () => {
          if (!run) return;
          const own = stripStyleProps(ch.props, base);
          // Köprü karakter stilinin mavi rengi ve alt çizgisi biçim sayılmaz.
          if (link && own.color && /^#(0563C1|0000FF|000080)$/i.test(own.color)) delete own.color;
          if (link && own.underline === "single") delete own.underline;
          const marks: JSONContent["marks"] = charMarks({ ...own, link });
          if (ch.chp.rmDel) marks.push({ type: "deletion", attrs: { author: this.revAuthors[ch.chp.ibstRMarkDel] ?? this.revAuthors[ch.chp.ibstRMark] ?? "", date: dttmIso(ch.chp.dttmRMarkDel || ch.chp.dttmRMark) } });
          else if (ch.chp.rmIns) marks.push({ type: "insertion", attrs: { author: this.revAuthors[ch.chp.ibstRMark] ?? "", date: dttmIso(ch.chp.dttmRMark) } });
          if (this.pendingComments.length && main) {
            // Başvurusu metinden önce gelen açıklama ardından gelen ilk sözcüğe bağlanır.
            const m = /^\s*\S+/.exec(run);
            const head = m ? m[0] : run;
            b.inline.push({ type: "text", text: head, marks: [...marks, ...this.pendingComments.map((id) => ({ type: "comment", attrs: { id } }))] });
            for (const id of this.pendingComments) this.pointIds.add(id);
            this.pendingComments = [];
            run = run.slice(head.length);
            if (!run) return;
          }
          b.inline.push({ type: "text", text: run, ...(marks.length ? { marks } : {}) });
          run = "";
        };
        const from = i;
        for (; i < stop; i++) {
          const c = this.text.charCodeAt(i);
          // Alan karakteri bir sonraki turda işlenir; eşi olmayan (açılmamış alanın) ayırıcı/sonu atlanır.
          if (c === 0x13 || c === 0x14 || c === 0x15) {
            if (i === from) {
              this.skipped++;
              i++;
            }
            break;
          }
          if (c >= 0x20 && c !== 0x1e && c !== 0x1f) {
            run += c >= 0xf020 && c <= 0xf0ff ? String.fromCharCode(c - 0xf000) : this.text[i];
            continue;
          }
          flush();
          if (c === 0x09) run += "\t";
          else if (c === 0x0b) b.inline.push({ type: "hardBreak" });
          else if (c === 0x0c) b.inline.push("page");
          else if (c === 0x0e) b.inline.push("column");
          else if (c === 0x1e) run += "‑";
          else if (c === 0x1f) run += "­";
          else if (c === 0x01 && ch.chp.spec && ch.chp.picLoc !== null && !ch.chp.ole) {
            const pic = readPicture(this.data, ch.chp.picLoc, this.bstore);
            const img = pic ? imageNode(pic.bytes, { width: pic.w > 0 ? pic.w : null, height: pic.h > 0 ? pic.h : null }) : null;
            if (img) b.inline.push(img);
          } else if (c === 0x02 && ch.chp.spec && main) {
            const note = this.note(i);
            if (note) b.inline.push(note);
          } else if (c === 0x05 && ch.chp.spec && main) {
            const id = this.annotation(i);
            if (id) this.attachComment(b, id);
          } else if (c === 0x01 && ch.chp.spec && ch.chp.ole) this.warnings.add("objects", "OLE");
          // Çizim nesnesi (kayan şekil/resim); SHAPE alanının sonucundaki satır içi resmin eşi sayılmaz.
          else if (c === 0x08 && ch.chp.spec && !fields.some((x) => x.result && /^\s*SHAPE\b/i.test(x.instr))) this.warnings.add("shapes", "çizim nesnesi");
        }
        flush();
      }
    }
    // Açık kalan içerik
    if (b.inline.length) this.endParagraph(b, this.papAt(end - 1), { id: "Normal", level: null }, styleChar(stylesOf(), "Normal").base, 0x0d, false);
    b.tables.close(b.items);
    return b.items;
  }

  private endParagraph(b: Builder, pap: Pap, map: { id: string; level: number | null }, base: StyleProps, mark: number, more: boolean, firstBase: number | null = null) {
    // İç içe tablonun satır sonu paragrafı içerik taşımaz.
    if (pap.ttp && mark === 0x07) {
      const def = parseTDef(pap.tdef ?? new Uint8Array(0));
      const left = def?.edges[0] ?? 0;
      const shd = parseTShd(pap.tshd);
      b.inline = [];
      b.tables.endRow({ cellx: def ? def.edges.slice(1).map((e) => e - left) : [], merge: def?.merge ?? [], vmerge: def?.vmerge ?? [], bg: shd });
      return;
    }
    if (pap.ttp && pap.itap > 1) {
      b.inline = [];
      return;
    }
    const attrs: Record<string, unknown> = {};
    if (map.level) attrs.level = map.level;
    else if (map.id !== "Normal") attrs.styleId = map.id;
    const jc = pap.jc === 1 ? "center" : pap.jc === 2 ? "right" : pap.jc === 3 || pap.jc === 4 ? "justify" : "left";
    if (jc !== (base.align ?? "left")) attrs.textAlign = jc;
    const pt = (tw: number) => Math.round((tw / 20) * 100) / 100;
    const near = (v: number, b2: number | undefined) => Math.abs(v - (b2 ?? 0)) < 0.06;
    const list = this.listInfo(pap);
    if (!list) {
      if (pap.left && !near(pt(pap.left), base.indentLeft)) attrs.indentLeft = pt(pap.left);
      if (pap.first && pap.first !== firstBase) attrs.indentFirst = pt(pap.first);
    }
    if (pap.right && !near(pt(pap.right), base.indentRight)) attrs.indentRight = pt(pap.right);
    if (!near(pt(pap.before ?? 0), base.spaceBefore)) attrs.spaceBefore = pt(pap.before ?? 0);
    if (!near(pt(pap.after ?? 0), base.spaceAfter)) attrs.spaceAfter = pt(pap.after ?? 0);
    const line = pap.line ?? { dya: 240, mult: true };
    if (line.mult) {
      const lh = Math.round((line.dya / 240) * 100) / 100;
      if (Math.abs(lh - (base.line ?? 1)) > 0.015 && lh > 0) attrs.lineHeight = lh;
    } else if (line.dya) {
      attrs.lineHeight = pt(Math.abs(line.dya));
      attrs.lineRule = line.dya < 0 ? "exact" : "atLeast";
    }
    if (pap.keepNext && !base.keepNext) attrs.keepNext = true;
    if (pap.keep) attrs.keepLines = true;
    if (pap.pageBreakBefore) attrs.pageBreakBefore = true;
    if (pap.shading) attrs.shading = pap.shading;
    // Sayfa/sütun sonları paragrafı parçalara böler.
    const segs: JSONContent[][] = [[]];
    const kinds: string[] = [];
    for (const n of b.inline) {
      if (n === "page" || n === "column") {
        kinds.push(n === "page" ? "pageBreak" : "columnBreak");
        segs.push([]);
      } else segs[segs.length - 1].push(n);
    }
    b.inline = [];
    if (this.pointIds.size) {
      for (const seg of segs) extendCommentToWord(seg, this.pointIds);
      this.pointIds.clear();
    }
    const nodes: { node: JSONContent; first: boolean }[] = [];
    segs.forEach((content, k) => {
      if (k > 0) nodes.push({ node: { type: kinds[k - 1] }, first: false });
      if (!content.length && segs.length > 1 && (k === 0 || k === segs.length - 1)) return;
      nodes.push({ node: paragraphNode({ ...attrs }, content, map.level ? "heading" : "paragraph"), first: k === 0 });
    });
    for (const { node, first } of nodes) {
      if (pap.inTable && node.type !== "pageBreak" && node.type !== "columnBreak") {
        // Tablonun ilk paragrafındaki "önce sayfa sonu" tablodan önce sayfa sonu olur.
        if (node.attrs?.pageBreakBefore && b.tables.isEmpty()) {
          delete node.attrs.pageBreakBefore;
          b.tables.close(b.items);
          b.items.push({ node: { type: "pageBreak" } });
        }
        b.tables.addBlock(node);
      }
      else if (!pap.inTable) {
        b.tables.close(b.items);
        b.items.push({ node, list: first && !map.level && node.type === "paragraph" ? list : undefined });
      }
    }
    if (pap.inTable && mark === 0x07 && !pap.ttp) b.tables.endCell();
    if (mark === 0x0c && more) {
      b.tables.close(b.items);
      b.items.push({ node: { type: "__section" } });
    }
  }

  private listInfo(pap: Pap): ListInfo | undefined {
    if (!pap.ilfo || pap.ilfo === 0xf801) return undefined;
    const lv = this.lists.level(pap.ilfo, pap.ilvl);
    const ordered = !!lv && lv.nfc !== 23 && lv.nfc !== 255;
    const key = `lfo${pap.ilfo}`;
    const ck = `${key}:${pap.ilvl}`;
    const n = this.counts.get(ck) ?? (lv?.start ?? 1) - 1;
    this.counts.set(ck, n + 1);
    // Alt düzeyler üst düzeyde yeni öğe açılınca yeniden başlar.
    for (let k = pap.ilvl + 1; k < 9; k++) this.counts.delete(`${key}:${k}`);
    const listStyle = !lv ? null : ordered ? (NFC[lv.nfc] ?? null) : bulletStyleAt(BULLETS[lv.text[0] ?? ""] ?? BULLETS[String.fromCharCode((lv.text.charCodeAt(0) || 0xf000) - 0xf000)] ?? null, pap.ilvl);
    return { key, level: pap.ilvl, ordered, listStyle, start: n + 1 };
  }

  /** Açıklama tabloları: başvurular (PlcfandRef + ATRD), metin aralıkları (PlcfandTxt), yazarlar. */
  private readAtnTables() {
    const out = { refs: new Map<number, { k: number; ibst: number; initials: string }>(), txt: [] as number[], base: 0, owners: [] as string[] };
    if (!this.fib.ccpAtn) return out;
    const r = this.fib.fcLcb(FC.PlcfandRef);
    const x = this.fib.fcLcb(FC.PlcfandTxt);
    const p = plc(this.t, r.fc, r.lcb, 30);
    p.cps.slice(0, -1).forEach((cp, k) => {
      const at = p.data[k];
      const cch = Math.min(9, this.t.u16(at));
      out.refs.set(cp, { k, ibst: this.t.u16(at + 20), initials: this.t.utf16(at + 2, cch) });
    });
    out.txt = x.lcb ? Array.from({ length: Math.floor(x.lcb / 4) }, (_, i) => this.t.i32(x.fc + i * 4)) : [];
    out.base = this.fib.ccpText + this.fib.ccpFtn + this.fib.ccpHdd + this.fib.ccpMcr;
    const own = this.fib.fcLcb(FC.GrpXstAtnOwners);
    for (let o = own.fc, end = own.fc + own.lcb, guard = 0; own.lcb && o + 2 <= end && guard < 10000; guard++) {
      const n = this.t.u16(o);
      out.owners.push(this.t.utf16(o + 2, Math.min(n, 255)));
      o += 2 + n * 2;
    }
    return out;
  }

  /** Açıklama başvurusu (CP) → açıklama kimliği (metni açıklama hikâyesinden). */
  private annotation(cp: number): string | null {
    const ref = this.atn.refs.get(cp);
    if (!ref || ref.k + 1 >= this.atn.txt.length) return null;
    const blocks = foldLists(this.story(this.atn.base + this.atn.txt[ref.k], this.atn.base + this.atn.txt[ref.k + 1], false, true)).filter((x) => !x.type?.startsWith("__"));
    const text = blocks
      .map((bl) => (bl.content ?? []).map((n) => (n.type === "text" ? (n.text ?? "") : "")).join(""))
      .join("\n")
      .trim();
    const id = `dc${ref.k}`;
    const author = this.atn.owners[ref.ibst] ?? "";
    this.comments[id] = { id, author, initials: ref.initials.trim() || author.split(/\s+/).filter(Boolean).map((w) => w[0]!.toLocaleUpperCase("tr-TR")).join("").slice(0, 3) || "K", date: new Date().toISOString(), text };
    return id;
  }

  /** Açıklama başvurusu açıklanan aralığın sonundadır: önceki sözcüğe, yoksa sonraki sözcüğe bağlanır. */
  private attachComment(b: Builder, id: string) {
    for (let k = b.inline.length - 1; k >= 0; k--) {
      const n = b.inline[k];
      if (typeof n === "string") break;
      if (n.type !== "text" || !n.text) continue;
      const m = /\S+\s*$/.exec(n.text);
      const tail = m ? m[0] : n.text;
      const mark = { type: "comment", attrs: { id } };
      const rest = n.text.slice(0, n.text.length - tail.length);
      const parts: JSONContent[] = [];
      if (rest) parts.push({ ...n, text: rest });
      parts.push({ ...n, text: tail, marks: [...(n.marks ?? []), mark] });
      b.inline.splice(k, 1, ...parts);
      this.pointIds.add(id);
      return;
    }
    this.pendingComments.push(id);
  }

  /** Dipnot/son not başvurusu (CP) → not düğümü. */
  private note(cp: number): JSONContent | null {
    for (const [kind, set] of [["footnote", this.foot], ["endnote", this.end]] as const) {
      const k = set.refs.get(cp);
      if (k === undefined || k + 1 >= set.txt.length) continue;
      const blocks = foldLists(this.story(set.base + set.txt[k], set.base + set.txt[k + 1], false, true)).filter((x) => !x.type?.startsWith("__"));
      return noteNode(blocks, kind, `nd${++this.noteSeq}`, noteText);
    }
    return null;
  }

  private pageOf(sep: SectionProps): PageSetup {
    const p = defaultSettings().page;
    p.width = Math.max(1440, Math.min(31680, sep.width));
    p.height = Math.max(1440, Math.min(31680, sep.height));
    p.orientation = sep.landscape || p.width > p.height ? "landscape" : "portrait";
    p.margin = { top: sep.top, bottom: sep.bottom, left: sep.left, right: sep.right, header: sep.header, footer: sep.footer, gutter: 0 };
    p.columns = sep.columns;
    p.columnSpace = sep.colSpace;
    if (sep.titlePage) p.titlePage = true;
    return p;
  }

  read(): BelgeDocument {
    const settings = defaultSettings();
    const items = this.story(0, this.fib.ccpText, true);
    if (this.fib.ccpTxbx || this.fib.ccpHdrTxbx) this.warnings.add("shapes", "metin kutusu");
    if (this.skipped) this.warnings.add("corrupt", `atlanan karakter: ${this.skipped}`);
    if (this.fib.ccpAtn && !Object.keys(this.comments).length) this.warnings.add("comments");
    if (this.pendingComments.length) this.warnings.add("comments", "bağlanamayan açıklama");
    settings.comments = this.comments;
    const blocks = foldLists(items);
    const seps = this.sections.length ? this.sections.map((s) => s.sep) : [newSep()];
    const last = this.pageOf(seps[seps.length - 1]);
    settings.page = last;
    const sections: JSONContent[] = [];
    let k = 0;
    const open = () => {
      const sep = seps[Math.min(k, seps.length - 1)];
      const page = this.pageOf(sep);
      sections.push({ type: "section", attrs: { kind: sep.bkc === 0 ? "continuous" : sep.bkc === 3 ? "evenPage" : sep.bkc === 4 ? "oddPage" : "nextPage", page: samePage(page, last) ? null : page }, content: [] });
    };
    open();
    for (const n of blocks) {
      if (n.type === "__section") {
        k++;
        open();
        continue;
      }
      sections[sections.length - 1].content!.push(n);
    }
    this.headers(settings);
    return { content: normalizeDoc({ type: "doc", content: sections }), settings };
  }

  /** İlk bölümün üst/alt bilgileri (PlcfHdd). */
  private headers(settings: BelgeDocument["settings"]) {
    const { fc, lcb } = this.fib.fcLcb(FC.PlcfHdd);
    if (!lcb || !this.fib.ccpHdd) return;
    const cps = Array.from({ length: Math.floor(lcb / 4) }, (_, i) => this.t.i32(fc + i * 4));
    const base = this.fib.ccpText + this.fib.ccpFtn;
    const read = (idx: number): JSONContent | null => {
      if (idx + 1 >= cps.length || cps[idx + 1] <= cps[idx]) return null;
      const blocks = foldLists(this.story(base + cps[idx], base + cps[idx + 1], false, true)).filter((x) => !x.type?.startsWith("__"));
      while (blocks.length && blocks[blocks.length - 1].type === "paragraph" && !blocks[blocks.length - 1].content?.length) blocks.pop();
      return JSON.stringify(blocks).match(/"text":"[^"]*\S|pageField|"image"/) ? { type: "doc", content: blocks } : null;
    };
    settings.headerEven = read(6);
    settings.header = read(7);
    settings.footerEven = read(8);
    settings.footer = read(9);
    settings.headerFirst = read(10);
    settings.footerFirst = read(11);
  }
}

/** STTB (genişletilmiş, UTF-16) dizgi tablosu. */
function sttb(t: Bytes, { fc, lcb }: { fc: number; lcb: number }): string[] {
  if (!lcb) return [];
  const ext = t.u16(fc) === 0xffff;
  const n = Math.min(ext ? t.u16(fc + 2) : t.u16(fc), 4096);
  const cbExtra = t.u16(fc + (ext ? 4 : 2));
  let o = fc + (ext ? 6 : 4);
  const out: string[] = [];
  for (let i = 0; i < n && o + 2 <= fc + lcb; i++) {
    const cch = t.u16(o);
    out.push(ext ? t.utf16(o + 2, Math.min(cch, 255)) : "");
    o += 2 + (ext ? cch * 2 : cch) + cbExtra;
  }
  return out;
}

/** .doc baytları → belge. Şifreli, 97 öncesi ya da bozuk dosyada anlaşılır Türkçe hata. */
export function readDoc(data: ArrayBuffer | Uint8Array, warnings?: LoadWarnings): BelgeDocument {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  if (!isCfb(bytes)) throw new Error("Geçerli bir .doc dosyası değil.");
  let cfb: ReturnType<typeof readCfb>;
  try {
    cfb = readCfb(bytes);
  } catch (err) {
    throw new Error(`Bozuk .doc dosyası (${(err as Error).message}).`);
  }
  if (cfb.has("EncryptionInfo") || cfb.has("EncryptedPackage")) throw new Error("Bu belge parolayla şifrelenmiş; şifreli belgeler açılamıyor.");
  const wdRaw = cfb.read("WordDocument");
  if (!wdRaw) throw new Error("Bu dosya bir metin belgesi değil (.doc ana akışı bulunamadı).");
  const wd = new Bytes(wdRaw);
  const fib = parseFib(wd);
  const tRaw = cfb.read(fib.table1 ? "1Table" : "0Table");
  if (!tRaw) throw new Error("Bozuk .doc dosyası (biçim tablosu akışı yok).");
  const dataRaw = cfb.read("Data");
  return new DocReader(wd, new Bytes(tRaw), dataRaw ? new Bytes(dataRaw) : null, fib, warnings).read();
}
