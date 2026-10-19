// Worker'da DOM: DOMParser ve XMLSerializer yoktur. Okuyucuların kullandığı DOM, @xmldom/xmldom ile kurulur;
// ayrıştırma docx/xml.ts'teki yerleşik (tarayıcı kadar katı) ayrıştırıcıyla yapılır. xmldom 0.8'de olmayan
// üyeler (firstElementChild, nextElementSibling, parentElement, contains) eklenir; getAttribute/getAttributeNS
// tarayıcıdaki gibi yoksa null döner. Yalnız io.worker.ts (ve Worker'ı taklit eden testler) yükler.
import { DOMImplementation, XMLSerializer as XmldomSerializer } from "@xmldom/xmldom";
import { setBuiltinXml } from "./docx/xml";

type AnyNode = { nodeType: number; parentNode: AnyNode | null; firstChild: AnyNode | null; lastChild: AnyNode | null; nextSibling: AnyNode | null; previousSibling: AnyNode | null };

let installed = false;
const impl = new DOMImplementation();

function patch() {
  const probe = impl.createDocument(null, "x", null);
  const el = probe.documentElement as unknown as object;
  const elementProto = Object.getPrototypeOf(el) as Record<string, unknown>;
  let nodeProto = elementProto;
  while (nodeProto && !Object.prototype.hasOwnProperty.call(nodeProto, "textContent")) nodeProto = Object.getPrototypeOf(nodeProto) as Record<string, unknown>;
  const getter = (proto: object, name: string, get: (this: AnyNode) => unknown) => {
    if (!(name in proto)) Object.defineProperty(proto, name, { get, configurable: true });
  };
  const element = (n: AnyNode | null, dir: "nextSibling" | "previousSibling"): AnyNode | null => {
    while (n && n.nodeType !== 1) n = n[dir];
    return n;
  };
  getter(nodeProto, "parentElement", function () {
    const p = this.parentNode;
    return p && p.nodeType === 1 ? p : null;
  });
  getter(nodeProto, "firstElementChild", function () {
    return element(this.firstChild, "nextSibling");
  });
  getter(nodeProto, "lastElementChild", function () {
    return element(this.lastChild, "previousSibling");
  });
  getter(nodeProto, "nextElementSibling", function () {
    return element(this.nextSibling, "nextSibling");
  });
  getter(nodeProto, "previousElementSibling", function () {
    return element(this.previousSibling, "previousSibling");
  });
  getter(nodeProto, "children", function () {
    const out: AnyNode[] = [];
    for (let c = element(this.firstChild, "nextSibling"); c; c = element(c.nextSibling, "nextSibling")) out.push(c);
    return out;
  });
  getter(nodeProto, "childElementCount", function () {
    let n = 0;
    for (let c = element(this.firstChild, "nextSibling"); c; c = element(c.nextSibling, "nextSibling")) n++;
    return n;
  });
  if (!("contains" in nodeProto)) {
    Object.defineProperty(nodeProto, "contains", {
      value(this: AnyNode, other: AnyNode | null) {
        for (let n = other; n; n = n.parentNode) if (n === this) return true;
        return false;
      },
      configurable: true,
      writable: true,
    });
  }
  // Tarayıcıda yok olan öznitelik null'dur (xmldom boş dize döner).
  type Attrs = { getAttributeNode(name: string): { value: string } | null; getAttributeNodeNS(ns: string | null, name: string): { value: string } | null };
  Object.defineProperty(elementProto, "getAttribute", {
    value(this: Attrs, name: string) {
      const a = this.getAttributeNode(name);
      return a ? a.value : null;
    },
    configurable: true,
    writable: true,
  });
  Object.defineProperty(elementProto, "getAttributeNS", {
    value(this: Attrs, ns: string | null, name: string) {
      const a = this.getAttributeNodeNS(ns || null, name);
      return a ? a.value : null;
    },
    configurable: true,
    writable: true,
  });
}

/**
 * Yerleşik XML ayrıştırıcısını ve xmldom DOM'unu etkinleştirir; genel XMLSerializer yoksa (Worker) xmldom'unki
 * kurulur. `force`: DOMParser olan ortamda da yerleşik ayrıştırıcı (karşılaştırma testleri).
 */
export function installXmlDom(opts: { force?: boolean } = {}): void {
  if (!installed) {
    patch();
    installed = true;
  }
  setBuiltinXml(() => impl.createDocument(null, "", null) as unknown as Document, opts);
  const g = globalThis as { XMLSerializer?: unknown };
  if (typeof g.XMLSerializer === "undefined" || opts.force) g.XMLSerializer = XmldomSerializer;
}
