/** p:transition ve p:timing okuyucusu. */
import { uid } from "../model/factory";
import type { AnimClass, Animation, Transition, TransitionType } from "../model/types";
import { ANIMATIONS, TRANSITIONS } from "../show/catalog";
import { attr, descendants, kid, kids, numAttr, OH_EXT_URI } from "./xml";

function findTransition(sld: Element): Element | undefined {
  const direct = kid(sld, "transition");
  if (direct) return direct;
  for (const ac of kids(sld, "AlternateContent")) {
    const choice = kid(ac, "Choice");
    const fb = kid(ac, "Fallback");
    const t = kid(fb, "transition") ?? kid(choice, "transition");
    if (t) {
      // Choice içindeki süre bilgisi (p14:dur) daha kesindir.
      const ct = kid(choice, "transition");
      const dur = ct ? numAttr(ct, "dur") : undefined;
      if (dur !== undefined && !t.hasAttribute("dur")) t.setAttribute("data-dur", String(dur));
      return t;
    }
  }
  return undefined;
}

function extEl(el: Element | undefined, local: string): Element | undefined {
  const lst = kid(el, "extLst");
  for (const e of kids(lst, "ext")) {
    if ((attr(e, "uri") ?? "").toUpperCase() !== OH_EXT_URI.toUpperCase()) continue;
    const x = kid(e, local);
    if (x) return x;
  }
  return undefined;
}

export function readTransition(sld: Element): Transition | undefined {
  const t = findTransition(sld);
  if (!t) return undefined;
  const spd = attr(t, "spd");
  let duration = spd === "fast" ? 0.5 : spd === "med" ? 0.75 : 1;
  const dur = numAttr(t, "dur") ?? numAttr(t, "data-dur");
  if (dur !== undefined) duration = dur / 1000;
  const advClick = attr(t, "advClick");
  const advTm = numAttr(t, "advTm");
  const base: Transition = {
    type: "none",
    duration,
    advanceOnClick: advClick !== "0" && advClick !== "false",
    advanceAfter: advTm !== undefined ? advTm / 1000 : undefined,
  };
  const mine = extEl(t, "transition");
  if (mine) {
    const type = attr(mine, "type") as TransitionType | undefined;
    if (type && TRANSITIONS.some((x) => x.type === type)) {
      base.type = type;
      base.option = attr(mine, "option") || undefined;
      base.sound = attr(mine, "snd") || undefined;
      const d = numAttr(mine, "dur");
      if (d !== undefined) base.duration = d / 1000;
      return base;
    }
  }
  if (kid(kid(t, "sndAc"), "endSnd")) base.sound = "stop";
  const child = kids(t).find((k) => k.localName !== "sndAc" && k.localName !== "extLst");
  if (!child) return base;
  const dir = attr(child, "dir");
  switch (child.localName) {
    case "fade":
      base.type = "fade";
      if (attr(child, "thruBlk") === "1" || attr(child, "thruBlk") === "true") base.option = "black";
      break;
    case "push":
      base.type = "push";
      base.option = dir ?? "u";
      break;
    case "wipe":
      base.type = "wipe";
      base.option = dir ?? "l";
      break;
    case "split": {
      base.type = "split";
      const orient = attr(child, "orient") === "horz" ? "horz" : "vert";
      base.option = `${orient}${dir === "in" ? "In" : "Out"}`;
      break;
    }
    case "cut":
      base.type = "cut";
      break;
    case "randomBar":
      base.type = "randomBar";
      base.option = dir === "horz" ? "horz" : "vert";
      break;
    case "circle":
      base.type = "shape";
      base.option = "circle";
      break;
    case "diamond":
      base.type = "shape";
      base.option = "diamond";
      break;
    case "plus":
      base.type = "shape";
      base.option = "plus";
      break;
    case "pull":
      base.type = "uncover";
      base.option = dir && dir.length === 1 ? dir : "l";
      break;
    case "cover":
      base.type = "cover";
      base.option = dir && dir.length === 1 ? dir : "l";
      break;
    case "zoom":
      base.type = "zoom";
      base.option = dir === "out" ? "out" : "in";
      break;
    case "dissolve":
      base.type = "dissolve";
      break;
    case "checker":
      base.type = "checker";
      base.option = dir === "vert" ? "vert" : "horz";
      break;
    case "blinds":
      base.type = "blinds";
      base.option = dir === "horz" ? "horz" : "vert";
      break;
    case "wheel":
      base.type = "wheel";
      break;
    case "wedge":
      base.type = "wedge";
      break;
    case "newsflash":
      base.type = "newsflash";
      break;
    case "strips":
    case "comb":
      base.type = "wipe";
      base.option = "l";
      break;
    default:
      base.type = "fade";
      break;
  }
  return base;
}

const SUB_DIR: Record<number, string> = { 1: "t", 2: "r", 4: "b", 8: "l", 3: "tr", 6: "br", 12: "bl", 9: "tl" };

function mapPreset(cls: AnimClass, presetId: number, sub: number): { effect: string; option?: string } {
  const cand = ANIMATIONS.filter((a) => a.cls === cls && a.presetId === presetId);
  const def = cand[0];
  if (!def) return { effect: cls === "entr" ? "fade" : cls === "exit" ? "fadeOut" : "pulse" };
  let option: string | undefined;
  if (def.id === "flyIn" || def.id === "flyOut") option = SUB_DIR[sub] ?? "b";
  else if (def.id === "wipe" || def.id === "wipeOut") option = SUB_DIR[sub] && SUB_DIR[sub].length === 1 ? SUB_DIR[sub] : "b";
  else option = def.defaultOption;
  return { effect: def.id, option };
}

function firstDuration(el: Element): number {
  for (const c of descendants(el, "cTn")) {
    const d = attr(c, "dur");
    if (d && d !== "indefinite") {
      const n = Number(d);
      if (Number.isFinite(n)) return n / 1000;
    }
  }
  return 0.5;
}

export function readAnimations(sld: Element, idMap: Map<number, string>): Animation[] {
  const mine = extEl(sld, "anims");
  if (mine) {
    const out: Animation[] = [];
    for (const a of kids(mine, "anim")) {
      const spid = numAttr(a, "spid");
      const shapeId = spid !== undefined ? idMap.get(spid) : undefined;
      if (!shapeId) continue;
      out.push({
        id: uid("a"),
        shapeId,
        cls: (attr(a, "cls") as AnimClass) ?? "entr",
        effect: attr(a, "effect") ?? "fade",
        option: attr(a, "option") || undefined,
        trigger: (attr(a, "trigger") as Animation["trigger"]) ?? "click",
        duration: (numAttr(a, "dur") ?? 500) / 1000,
        delay: (numAttr(a, "delay") ?? 0) / 1000,
      });
    }
    return out;
  }
  const timing = kid(sld, "timing");
  if (!timing) return [];
  const out: Animation[] = [];
  for (const ctn of descendants(timing, "cTn")) {
    const cls = attr(ctn, "presetClass");
    if (cls !== "entr" && cls !== "exit" && cls !== "emph") continue;
    const spidEl = descendants(ctn, "spTgt")[0];
    const spid = numAttr(spidEl, "spid");
    const shapeId = spid !== undefined ? idMap.get(spid) : undefined;
    if (!shapeId) continue;
    const nodeType = attr(ctn, "nodeType");
    const trigger: Animation["trigger"] = nodeType === "withEffect" ? "with" : nodeType === "afterEffect" ? "after" : "click";
    const { effect, option } = mapPreset(cls, numAttr(ctn, "presetID") ?? 10, numAttr(ctn, "presetSubtype") ?? 0);
    const delayAttr = attr(kid(kid(ctn, "stCondLst"), "cond"), "delay");
    const delay = delayAttr && delayAttr !== "indefinite" ? Number(delayAttr) / 1000 : 0;
    out.push({ id: uid("a"), shapeId, cls, effect, option, trigger, duration: firstDuration(ctn), delay: Number.isFinite(delay) ? delay : 0 });
  }
  return out;
}
