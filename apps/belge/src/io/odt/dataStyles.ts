// ODF tarih/saat veri stilleri (number:date-style, number:time-style) ↔ biçim resmi ("d MMMM yyyy").
import { parsePicture } from "../../editor/dateFormat";
import { attr, esc, kids } from "../docx/xml";

/** Biçim resminden veri stili öğeleri. */
export function pictureToOdf(name: string, pic: string, time: boolean): string {
  const long = (on: boolean) => (on ? ' number:style="long"' : "");
  const inner = parsePicture(pic)
    .map((t) => {
      switch (t.kind) {
        case "lit":
          return `<number:text>${esc(t.text)}</number:text>`;
        case "day":
          return t.len >= 3 ? `<number:day-of-week${long(t.len >= 4)}/>` : `<number:day${long(t.len === 2)}/>`;
        case "month":
          return `<number:month${t.len >= 3 ? ' number:textual="true"' : ""}${long(t.len === 2 || t.len >= 4)}/>`;
        case "year":
          return `<number:year${long(t.len > 2)}/>`;
        case "hour24":
        case "hour12":
          return `<number:hours${long(t.len >= 2)}/>`;
        case "minute":
          return `<number:minutes${long(t.len >= 2)}/>`;
        case "second":
          return `<number:seconds${long(t.len >= 2)}/>`;
        case "ampm":
          return "<number:am-pm/>";
      }
    })
    .join("");
  const tag = time ? "number:time-style" : "number:date-style";
  return `<${tag} style:name="${name}">${inner}</${tag}>`;
}

/** Veri stili öğesinden biçim resmi (desteklenmeyen öğeler atlanır). */
export function odfToPicture(el: Element): string {
  let pic = "";
  const isLong = (e: Element) => attr(e, "style") === "long";
  const hasAmPm = kids(el).some((e) => e.localName === "am-pm");
  for (const e of kids(el)) {
    switch (e.localName) {
      case "text": {
        const t = e.textContent ?? "";
        // Belirteç harfleri tırnak içinde yazılır (ör. "de" gibi sözcükler).
        pic += /[dDMyYHhmsS']|am\/pm/i.test(t) ? `'${t.replace(/'/g, "")}'` : t;
        break;
      }
      case "day":
        pic += isLong(e) ? "dd" : "d";
        break;
      case "day-of-week":
        pic += isLong(e) ? "dddd" : "ddd";
        break;
      case "month":
        pic += attr(e, "textual") === "true" ? (isLong(e) ? "MMMM" : "MMM") : isLong(e) ? "MM" : "M";
        break;
      case "year":
        pic += isLong(e) ? "yyyy" : "yy";
        break;
      case "hours":
        pic += hasAmPm ? (isLong(e) ? "hh" : "h") : isLong(e) ? "HH" : "H";
        break;
      case "minutes":
        pic += isLong(e) ? "mm" : "m";
        break;
      case "seconds":
        pic += isLong(e) ? "ss" : "s";
        break;
      case "am-pm":
        pic += "AM/PM";
        break;
    }
  }
  return pic;
}
