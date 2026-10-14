// Okuyucuların ihtiyaç duyduğu DOM API'leri (DOMParser, XMLSerializer, document…) jsdom ile sağlanır;
// vitest'in jsdom ortamına denktir. İşçi iş parçacığında, uygulama paketleri yüklenmeden önce içe aktarılır.
import { createRequire } from "node:module";
import { join } from "node:path";
import { KOK } from "./ortak.mjs";

const { JSDOM } = createRequire(join(KOK, "package.json"))("jsdom");
const dom = new JSDOM("<!doctype html><html><head></head><body></body></html>", { url: "http://localhost/", pretendToBeVisual: true });
const w = dom.window;
const ADLAR = [
  "window", "document", "DOMParser", "XMLSerializer", "Node", "Element", "HTMLElement", "Document", "DocumentFragment", "Text", "Comment",
  "NodeFilter", "XPathResult", "getComputedStyle", "HTMLImageElement", "HTMLCanvasElement", "Image", "MutationObserver", "Range", "Selection",
  "CharacterData", "HTMLAnchorElement", "HTMLTableElement", "CSSStyleDeclaration", "DOMRect", "requestAnimationFrame", "cancelAnimationFrame",
  "localStorage", "sessionStorage", "Blob", "FileReader",
];
for (const k of ADLAR) {
  // Node'un kendi sağladıkları (Blob, structuredClone…) korunur.
  if (k in globalThis && k !== "window" && k !== "document" && globalThis[k]) continue;
  const v = k === "window" ? w : typeof w[k] === "function" && !/^[A-Z]/.test(k) ? w[k].bind(w) : w[k];
  try {
    Object.defineProperty(globalThis, k, { value: v, configurable: true, writable: true });
  } catch {
    /* salt okunur genel ad: atla */
  }
}
if (typeof globalThis.self === "undefined") globalThis.self = globalThis;
export { dom };
