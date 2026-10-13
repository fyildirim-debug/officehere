/** ODF slayt geçişleri (smil:type/subtype/direction, hız, otomatik ilerleme) ↔ model. */
import type { Transition } from "../../model/types";
import type { Props } from "./styles";

const FROM: Record<string, string> = { fromBottom: "u", fromTop: "d", fromRight: "l", fromLeft: "r" };
const TO_FROM: Record<string, string> = { u: "fromBottom", d: "fromTop", l: "fromRight", r: "fromLeft" };

/** "PT5S", "PT00H00M05S", "5s" → saniye. */
function seconds(v: string | undefined): number | undefined {
  if (!v) return undefined;
  const m = /^P?T?(?:(\d+(?:\.\d+)?)H)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)S)?$/i.exec(v.trim());
  if (m && (m[1] || m[2] || m[3])) return Number(m[1] ?? 0) * 3600 + Number(m[2] ?? 0) * 60 + Number(m[3] ?? 0);
  const s = /^(\d+(?:\.\d+)?)\s*s$/i.exec(v.trim());
  return s ? Number(s[1]) : undefined;
}

/** Sayfa stili (drawing-page-properties) → geçiş. `dur`: zamanlama kökündeki transitionFilter süresi. */
export function readOdpTransition(p: Props, dur?: string): Transition | undefined {
  const type = p["type"];
  const sub = p["subtype"] ?? "";
  const rev = p["direction"] === "reverse";
  const auto = p["transition-type"] === "automatic";
  const after = auto ? seconds(p["duration"]) : undefined;
  const legacy = p["transition-style"];
  if (!type && after === undefined && (!legacy || legacy === "none")) return undefined;
  const speed = p["transition-speed"];
  const t: Transition = { type: "none", duration: seconds(dur) ?? (speed === "fast" ? 0.5 : speed === "slow" ? 1 : 0.75), advanceOnClick: true };
  if (after !== undefined) t.advanceAfter = Math.min(86400, after);
  switch (type) {
    case undefined:
      if (legacy && legacy !== "none") t.type = "fade";
      break;
    case "fade":
      t.type = "fade";
      t.option = /color/i.test(sub) ? "black" : "smooth";
      break;
    case "pushWipe":
      t.type = "push";
      t.option = FROM[sub] ?? "u";
      break;
    case "slideWipe":
      t.type = rev ? "uncover" : "cover";
      t.option = FROM[sub] ?? "l";
      break;
    case "barWipe":
      t.type = "wipe";
      t.option = sub === "topToBottom" ? (rev ? "u" : "d") : rev ? "l" : "r";
      break;
    case "barnDoorWipe": {
      t.type = "split";
      t.option = `${sub === "horizontal" ? "horz" : "vert"}${rev ? "In" : "Out"}`;
      break;
    }
    case "randomBarWipe":
      t.type = "randomBar";
      t.option = sub === "horizontal" ? "horz" : "vert";
      break;
    case "ellipseWipe":
      t.type = "shape";
      t.option = "circle";
      break;
    case "irisWipe":
      t.type = "shape";
      t.option = sub === "diamond" ? "diamond" : "plus";
      break;
    case "checkerBoardWipe":
      t.type = "checker";
      t.option = sub === "down" ? "vert" : "horz";
      break;
    case "blindsWipe":
      t.type = "blinds";
      t.option = sub === "horizontal" ? "horz" : "vert";
      break;
    case "dissolve":
      t.type = "dissolve";
      break;
    case "clockWipe":
      t.type = "wheel";
      break;
    case "fanWipe":
      t.type = "wedge";
      break;
    default:
      t.type = "fade";
  }
  return t;
}

/** Geçiş → sayfa stili öznitelikleri. */
export function odpTransitionAttrs(t: Transition | undefined): Record<string, string | undefined> {
  if (!t) return {};
  const a: Record<string, string | undefined> = {};
  const o = t.option ?? "";
  const set = (type: string, subtype: string, reverse = false) => {
    a["smil:type"] = type;
    a["smil:subtype"] = subtype;
    if (reverse) a["smil:direction"] = "reverse";
  };
  switch (t.type) {
    case "none":
    case "morph":
      break;
    case "fade":
      set("fade", o === "black" ? "fadeOverColor" : "crossfade");
      break;
    case "push":
      set("pushWipe", TO_FROM[o] ?? "fromBottom");
      break;
    case "cover":
    case "uncover":
      set("slideWipe", TO_FROM[o] ?? "fromRight", t.type === "uncover");
      break;
    case "wipe":
      set("barWipe", o === "u" || o === "d" ? "topToBottom" : "leftToRight", o === "u" || o === "l");
      break;
    case "split":
      set("barnDoorWipe", o.startsWith("horz") ? "horizontal" : "vertical", o.endsWith("In"));
      break;
    case "randomBar":
      set("randomBarWipe", o === "horz" ? "horizontal" : "vertical");
      break;
    case "shape":
      if (o === "circle") set("ellipseWipe", "circle");
      else set("irisWipe", o === "diamond" ? "diamond" : "rectangle");
      break;
    case "checker":
      set("checkerBoardWipe", o === "vert" ? "down" : "across");
      break;
    case "blinds":
      set("blindsWipe", o === "horz" ? "horizontal" : "vertical");
      break;
    case "dissolve":
      set("dissolve", "default");
      break;
    case "wheel":
      set("clockWipe", "clockwiseTwelve");
      break;
    case "wedge":
      set("fanWipe", "centerTop");
      break;
    default:
      set("fade", "crossfade");
  }
  if (a["smil:type"]) a["presentation:transition-speed"] = t.duration >= 0.9 ? "slow" : t.duration >= 0.6 ? "medium" : "fast";
  if (t.advanceAfter !== undefined) {
    a["presentation:transition-type"] = "automatic";
    a["presentation:duration"] = `PT${Math.round(t.advanceAfter * 1000) / 1000}S`;
  }
  return a;
}

/** Zamanlama kökü: geçiş süresi için anim:transitionFilter (sayfa kimliğiyle başlar). */
export function transitionTimingXml(t: Transition | undefined, pageId: string): string {
  const a = odpTransitionAttrs(t);
  if (!t || !a["smil:type"]) return "";
  const attrs = [`smil:dur="${Math.round(t.duration * 1000) / 1000}s"`, `smil:type="${a["smil:type"]}"`, `smil:subtype="${a["smil:subtype"]}"`, a["smil:direction"] ? `smil:direction="${a["smil:direction"]}"` : ""].filter(Boolean).join(" ");
  return `<anim:par presentation:node-type="timing-root"><anim:par smil:begin="${pageId}.begin"><anim:transitionFilter ${attrs}/></anim:par></anim:par>`;
}

/** Sayfadaki slayt geçişi süresi (zamanlama kökündeki transitionFilter). */
export function transitionDur(page: Element): string | undefined {
  const walk = (el: Element, depth: number): string | undefined => {
    if (depth > 6) return undefined;
    for (let c = el.firstElementChild; c; c = c.nextElementSibling) {
      if (c.localName === "transitionFilter" && /\.begin$/.test(el.getAttribute("smil:begin") ?? Array.from(el.attributes).find((x) => x.localName === "begin")?.value ?? "")) {
        return Array.from(c.attributes).find((x) => x.localName === "dur")?.value;
      }
      if (c.localName === "par") {
        const r = walk(c, depth + 1);
        if (r) return r;
      }
    }
    return undefined;
  };
  for (let c = page.firstElementChild; c; c = c.nextElementSibling) if (c.localName === "par") return walk(c, 0);
  return undefined;
}
