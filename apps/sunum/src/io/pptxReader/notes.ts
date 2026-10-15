/**
 * Konuşmacı notları (notesSlide) ve not asıl sayfasının üst/alt bilgisi
 * (notesMaster p:hf ve yer tutucu metinleri).
 * Geçiş katmanı: not sayfası olan slayt notsuz da not sayfasıyla yazılır (`SlideRaw.notesPage`);
 * not sayfasındaki gövde dışı şekiller ve not asıl sayfasının görünümü okunmuyor.
 */
import { DEFAULT_NOTES_HF } from "../../model/presentation";
import type { NotesHeaderFooter } from "../../model/types";
import { readTextBody } from "../readDrawing";
import { attr, boolAttr, descendants, kid, kidPath, kids, parseRels, parseXml, relId, relsPathFor, REL, type Rel } from "../xml";
import type { MkCtx } from "./common";
import type { PptxPackage } from "./package";
import type { PresPart } from "./presentation";

/** Not asıl sayfasındaki p:hf ve üst/alt bilgi metinleri; p:hf yoksa tanımsız (varsayılan). */
export function readNotesHf(xml: string): NotesHeaderFooter | undefined {
  let root: Element;
  try {
    root = parseXml(xml).documentElement;
  } catch {
    return undefined;
  }
  const hf = kid(root, "hf");
  if (!hf) return undefined;
  const on = (k: string) => boolAttr(hf, k) !== false;
  const phText = (type: string) => {
    const sp = descendants(root, "sp").find((e) => attr(descendants(e, "ph")[0], "type") === type);
    return sp ? descendants(sp, "p").map((p) => descendants(p, "t").map((t) => t.textContent ?? "").join("")).join(" ").trim() : "";
  };
  const dtSp = descendants(root, "sp").find((e) => attr(descendants(e, "ph")[0], "type") === "dt");
  const dateAuto = !dtSp || descendants(dtSp, "fld").some((f) => (attr(f, "type") ?? "").startsWith("datetime"));
  return {
    ...DEFAULT_NOTES_HF,
    pageNumber: on("sldNum"),
    header: on("hdr"),
    headerText: phText("hdr"),
    footer: on("ftr"),
    footerText: phText("ftr"),
    dateTime: on("dt"),
    dateAuto,
    dateFixed: dateAuto ? "" : phText("dt"),
  };
}

/** Not sayfaları/dinleyici notları üst-alt bilgisi (sunudaki not asıl sayfasından). */
export async function readNotesMasterHf(pres: PresPart): Promise<NotesHeaderFooter | undefined> {
  const nmRid = relId(kid(kid(pres.doc, "notesMasterIdLst"), "notesMasterId"));
  const nmRel = nmRid && pres.rels.get(nmRid)?.target;
  const nmPath = nmRel && pres.pk.zip.has(nmRel) ? nmRel : pres.partOf(REL.notesMaster, /^ppt\/notesMasters\/[^/]+\.xml$/i);
  const nmXml = nmPath ? await pres.pk.text(nmPath) : undefined;
  return nmXml ? readNotesHf(nmXml) : undefined;
}

/** Slaydın konuşmacı notları (düz metin; `n`: 1 tabanlı slayt numarası); yoksa "". */
export async function readSlideNotes(pk: PptxPackage, mkCtx: MkCtx, rels: Map<string, Rel>, n: number): Promise<string> {
  const nrel = [...rels.values()].find((r) => r.type === REL.notesSlide);
  if (!nrel) return "";
  const nPart = await pk.part(nrel.target, "notes");
  if (nPart.problem) pk.W.slide("notes", n, nPart.detail);
  if (!nPart.root) return "";
  const nctx = mkCtx(parseRels(await pk.text(relsPathFor(nrel.target)), nrel.target));
  const nsp = notesBody(kids(kidPath(nPart.root, "cSld", "spTree"), "sp"));
  if (!nsp) return "";
  const body = readTextBody(kid(nsp, "txBody"), nctx);
  return body ? body.paragraphs.map((p) => p.runs.map((r) => r.text.replace(/\v/g, "\n")).join("")).join("\n").replace(/\s+$/, "") : "";
}

/** Not gövdesi yer tutucusu; yoksa (bazı yazıcılar) yer tutucu olmayan ilk metin kutusu. */
function notesBody(nsps: Element[]): Element | undefined {
  const phType = (nsp: Element) => {
    const ph = kidPath(nsp, "nvSpPr", "nvPr", "ph");
    return ph ? (attr(ph, "type") ?? "body") : undefined;
  };
  return nsps.find((nn) => phType(nn) === "body") ?? nsps.find((nn) => phType(nn) === undefined && !!kid(nn, "txBody"));
}
