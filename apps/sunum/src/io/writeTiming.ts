/** Slayt geçişi (p:transition) ve animasyon zamanlaması (p:timing) yazıcısı. */
import type { Animation, Slide, Transition } from "../model/types";
import { animDef } from "../show/catalog";
import { buildSteps } from "../show/effects";
import { attrs, esc, NS, OH_EXT_URI } from "./xml";

function speed(d: number): string {
  if (d <= 0.5) return "fast";
  if (d <= 0.75) return "med";
  return "slow";
}

function transitionChild(t: Transition): string {
  const o = t.option;
  switch (t.type) {
    case "fade":
      return o === "black" ? '<p:fade thruBlk="1"/>' : "<p:fade/>";
    case "reveal":
      return '<p:fade thruBlk="1"/>';
    case "push":
      return `<p:push dir="${o ?? "u"}"/>`;
    case "wipe":
      return `<p:wipe dir="${o ?? "l"}"/>`;
    case "split": {
      const v = o ?? "vertOut";
      return `<p:split orient="${v.startsWith("horz") ? "horz" : "vert"}" dir="${v.endsWith("In") ? "in" : "out"}"/>`;
    }
    case "cut":
      return "<p:cut/>";
    case "randomBar":
      return `<p:randomBar dir="${o === "horz" ? "horz" : "vert"}"/>`;
    case "shape":
      return o === "diamond" ? "<p:diamond/>" : o === "plus" ? "<p:plus/>" : "<p:circle/>";
    case "uncover":
      return `<p:pull dir="${o ?? "l"}"/>`;
    case "cover":
      return `<p:cover dir="${o ?? "l"}"/>`;
    case "zoom":
      return `<p:zoom dir="${o === "out" ? "out" : "in"}"/>`;
    case "dissolve":
      return "<p:dissolve/>";
    case "checker":
      return `<p:checker dir="${o === "vert" ? "vert" : "horz"}"/>`;
    case "blinds":
      return `<p:blinds dir="${o === "horz" ? "horz" : "vert"}"/>`;
    case "wheel":
      return '<p:wheel spokes="1"/>';
    case "wedge":
      return "<p:wedge/>";
    case "newsflash":
      return "<p:newsflash/>";
    default:
      return "";
  }
}

export function transitionXml(t: Transition | undefined): string {
  if (!t) return "";
  if (t.type === "none" && t.advanceOnClick && t.advanceAfter === undefined) return "";
  const a = attrs({
    spd: t.type === "none" ? undefined : speed(t.duration),
    advClick: t.advanceOnClick ? undefined : "0",
    advTm: t.advanceAfter !== undefined ? Math.round(t.advanceAfter * 1000) : undefined,
  });
  const ext = `<p:extLst><p:ext uri="${OH_EXT_URI}"><oh:transition xmlns:oh="${NS.oh}"${attrs({ type: t.type, option: t.option, dur: Math.round(t.duration * 1000) })}/></p:ext></p:extLst>`;
  return `<p:transition${a}>${transitionChild(t)}${ext}</p:transition>`;
}

const FLY_SUB: Record<string, number> = { t: 1, r: 2, b: 4, l: 8, tr: 3, br: 6, bl: 12, tl: 9 };

interface Ctx {
  id: number;
}

function cTnId(c: Ctx) {
  c.id += 1;
  return c.id;
}

function setVisibility(c: Ctx, spid: number, visible: boolean, delay = 0): string {
  return `<p:set><p:cBhvr><p:cTn id="${cTnId(c)}" dur="1" fill="hold"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="${visible ? "visible" : "hidden"}"/></p:to></p:set>`;
}

function animEffect(c: Ctx, spid: number, filter: string, dir: "in" | "out", dur: number): string {
  return `<p:animEffect transition="${dir}" filter="${filter}"><p:cBhvr><p:cTn id="${cTnId(c)}" dur="${dur}"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:animEffect>`;
}

function animProp(c: Ctx, spid: number, prop: string, from: string, to: string, dur: number): string {
  return `<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="${cTnId(c)}" dur="${dur}" fill="hold"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>${prop}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="${esc(from)}"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="${esc(to)}"/></p:val></p:tav></p:tavLst></p:anim>`;
}

function behaviors(c: Ctx, a: Animation, spid: number): string {
  const dur = Math.max(1, Math.round(a.duration * 1000));
  const o = a.option ?? "";
  const entr = a.cls === "entr";
  const exit = a.cls === "exit";
  const vis = entr ? setVisibility(c, spid, true) : "";
  const hide = exit ? setVisibility(c, spid, false, Math.max(0, dur - 1)) : "";
  switch (a.effect) {
    case "appear":
      return vis;
    case "disappear":
      return setVisibility(c, spid, false);
    case "fade":
    case "fadeOut":
      return vis + animEffect(c, spid, "fade", entr ? "in" : "out", dur) + hide;
    case "wipe":
    case "wipeOut": {
      const d = o === "l" ? "right" : o === "r" ? "left" : o === "t" ? "down" : "up";
      return vis + animEffect(c, spid, `wipe(${d})`, entr ? "in" : "out", dur) + hide;
    }
    case "split":
    case "splitOut": {
      const v = o || "vertOut";
      const barn = `${v.endsWith("Out") ? "out" : "in"}${v.startsWith("horz") ? "Horizontal" : "Vertical"}`;
      return vis + animEffect(c, spid, `barn(${barn})`, entr ? "in" : "out", dur) + hide;
    }
    case "shape":
      return vis + animEffect(c, spid, o === "diamond" ? "diamond(in)" : o === "plus" ? "plus(in)" : "circle(in)", "in", dur);
    case "wheel":
      return vis + animEffect(c, spid, "wheel(1)", "in", dur);
    case "randomBars":
      return vis + animEffect(c, spid, `randombar(${o === "vert" ? "vertical" : "horizontal"})`, "in", dur);
    case "flyIn":
    case "flyOut": {
      const dir = o || "b";
      const fx = dir.includes("l") ? "0-#ppt_w/2" : dir.includes("r") ? "1+#ppt_w/2" : "#ppt_x";
      const fy = dir.includes("t") ? "0-#ppt_h/2" : dir.includes("b") ? "1+#ppt_h/2" : "#ppt_y";
      if (entr) return vis + animProp(c, spid, "ppt_x", fx, "#ppt_x", dur) + animProp(c, spid, "ppt_y", fy, "#ppt_y", dur);
      return animProp(c, spid, "ppt_x", "#ppt_x", fx, dur) + animProp(c, spid, "ppt_y", "#ppt_y", fy, dur) + hide;
    }
    case "zoom":
    case "growTurn":
    case "swivel":
    case "bounce":
    case "floatIn":
      return vis + animEffect(c, spid, "fade", "in", dur) + animProp(c, spid, "ppt_w", "0", "#ppt_w", dur) + animProp(c, spid, "ppt_h", "0", "#ppt_h", dur);
    case "zoomOut":
    case "shrinkTurn":
    case "floatOut":
      return animEffect(c, spid, "fade", "out", dur) + animProp(c, spid, "ppt_w", "#ppt_w", "0", dur) + animProp(c, spid, "ppt_h", "#ppt_h", "0", dur) + hide;
    case "spin":
      return `<p:animRot by="21600000"><p:cBhvr><p:cTn id="${cTnId(c)}" dur="${dur}" fill="hold"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>`;
    case "growShrink":
    case "pulse":
      return `<p:animScale><p:cBhvr><p:cTn id="${cTnId(c)}" dur="${dur}" fill="hold"${a.effect === "pulse" ? ' autoRev="1"' : ""}/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr><p:by x="${a.effect === "pulse" ? 105000 : 150000}" y="${a.effect === "pulse" ? 105000 : 150000}"/></p:animScale>`;
    case "teeter":
      return `<p:animRot by="300000"><p:cBhvr><p:cTn id="${cTnId(c)}" dur="${dur}" fill="hold" autoRev="1"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>`;
    case "transparency":
    case "colorPulse":
      return `<p:set><p:cBhvr><p:cTn id="${cTnId(c)}" dur="${dur}" fill="hold"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="0.5"/></p:to></p:set>`;
    default:
      return vis + animEffect(c, spid, "fade", entr ? "in" : "out", dur) + hide;
  }
}

/** Animasyon zamanlama ağacı; şekil kimlikleri cNvPr kimliklerine eşlenir. */
export function timingXml(slide: Slide, idMap: Map<string, number>): string {
  const anims = slide.animations.filter((a) => idMap.has(a.shapeId));
  if (!anims.length) return "";
  const c: Ctx = { id: 2 };
  const steps = buildSteps({ ...slide, animations: anims });
  const groups: { delayIndefinite: boolean; effects: ReturnType<typeof buildSteps>["auto"] }[] = [];
  if (steps.auto.length) groups.push({ delayIndefinite: false, effects: steps.auto });
  for (const g of steps.clicks) groups.push({ delayIndefinite: true, effects: g });
  const grpIds = new Map<string, number>();
  const bld: string[] = [];
  const pars = groups
    .map((g) => {
      // Aynı başlangıç zamanındaki efektler bir alt grupta toplanır.
      const byStart = new Map<number, typeof g.effects>();
      for (const e of g.effects) {
        const k = Math.round(e.start * 1000);
        byStart.set(k, [...(byStart.get(k) ?? []), e]);
      }
      const subs = [...byStart.entries()]
        .sort((a, b) => a[0] - b[0])
        .map(([start, list]) => {
          const inner = list
            .map((e) => {
              const a = e.anim;
              const spid = idMap.get(a.shapeId)!;
              const def = animDef(a.effect);
              const sub = a.option && FLY_SUB[a.option] !== undefined ? FLY_SUB[a.option] : 0;
              const grpId = grpIds.get(a.shapeId) ?? grpIds.size;
              grpIds.set(a.shapeId, grpId);
              if (!bld.some((b) => b.includes(`spid="${spid}"`))) bld.push(`<p:bldP spid="${spid}" grpId="${grpId}" animBg="1"/>`);
              const nodeType = a.trigger === "click" ? "clickEffect" : a.trigger === "with" ? "withEffect" : "afterEffect";
              const id = cTnId(c);
              return `<p:par><p:cTn${attrs({ id, presetID: def?.presetId ?? 10, presetClass: a.cls, presetSubtype: sub, fill: "hold", grpId, nodeType })}><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${behaviors(c, a, spid)}</p:childTnLst></p:cTn></p:par>`;
            })
            .join("");
          return `<p:par><p:cTn id="${cTnId(c)}" fill="hold"><p:stCondLst><p:cond delay="${start}"/></p:stCondLst><p:childTnLst>${inner}</p:childTnLst></p:cTn></p:par>`;
        })
        .join("");
      const cond = g.delayIndefinite ? '<p:cond delay="indefinite"/>' : '<p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>';
      return `<p:par><p:cTn id="${cTnId(c)}" fill="hold"><p:stCondLst>${cond}</p:stCondLst><p:childTnLst>${subs}</p:childTnLst></p:cTn></p:par>`;
    })
    .join("");
  const seq =
    `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${pars}</p:childTnLst></p:cTn>` +
    '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>';
  return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>${seq}</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst>${bld.join("")}</p:bldLst></p:timing>`;
}

/** Uygulamanın kendi animasyon listesi (kayıpsız gidiş-dönüş için slayt uzantısı). */
export function slideExtXml(slide: Slide, idMap: Map<string, number>): string {
  const anims = slide.animations.filter((a) => idMap.has(a.shapeId));
  if (!anims.length) return "";
  const items = anims
    .map((a) => `<oh:anim${attrs({ spid: idMap.get(a.shapeId), cls: a.cls, effect: a.effect, option: a.option, trigger: a.trigger, dur: Math.round(a.duration * 1000), delay: Math.round(a.delay * 1000) })}/>`)
    .join("");
  return `<p:extLst><p:ext uri="${OH_EXT_URI}"><oh:anims xmlns:oh="${NS.oh}">${items}</oh:anims></p:ext></p:extLst>`;
}
