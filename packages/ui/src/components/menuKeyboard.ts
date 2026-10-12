/**
 * Açılır katmanlarda (menü, renk paleti, galeri menüsü, daraltılmış grup) klavye gezintisi
 * ve odak yönetimi. Popover bunu kullanır; uygulamalardaki menüler ek kod istemeden yararlanır.
 */
import type { KeyboardEvent as ReactKeyboardEvent } from "react";
import { createTypeahead, firstEnabled, lastEnabled, matchByPrefix, scrollToReveal, spatialNext, type Box } from "./menuNav";

/* ------------------------------------------------------------------ */
/* Girdi kipi: katman klavyeyle mi açıldı?                               */
/* ------------------------------------------------------------------ */

let keyboardModality = false;
// Sağ tık odaklanamayan bir yüzeyde (ızgara tuvali) odağı boşa düşürür; menü kapanınca odak buraya döner.
let focusBeforePointer: HTMLElement | null = null;
const MODIFIERS = new Set(["Shift", "Control", "Alt", "Meta", "AltGraph", "CapsLock"]);
if (typeof window !== "undefined") {
  const onPointer = () => {
    keyboardModality = false;
    const a = document.activeElement;
    focusBeforePointer = a instanceof HTMLElement && a !== document.body ? a : null;
  };
  window.addEventListener("keydown", (e) => !MODIFIERS.has(e.key) && (keyboardModality = true), true);
  window.addEventListener("pointerdown", onPointer, true);
  window.addEventListener("mousedown", () => (keyboardModality = false), true);
}

/** Son etkileşim klavyeyle miydi (fareyle açılan menüde ilk öğe vurgulanmaz). */
export function lastInputWasKeyboard(): boolean {
  return keyboardModality;
}

/** Son fare basışından hemen önce odaktaki öğe (hâlâ sayfadaysa). */
export function focusBeforeLastPointer(): HTMLElement | null {
  return !keyboardModality && focusBeforePointer?.isConnected ? focusBeforePointer : null;
}

/* ------------------------------------------------------------------ */
/* Öğeler                                                               */
/* ------------------------------------------------------------------ */

const ITEM = 'button, a[href], [role="menuitem"], [role="menuitemcheckbox"], [role="menuitemradio"]';
const FIELD = 'input:not([type="hidden"]):not([type="color"]):not([type="file"]):not(.oh-hidden-input), select, textarea, [contenteditable="true"]';
const TABBABLE = 'button, a[href], input:not([type="hidden"]), select, textarea, [tabindex], [contenteditable="true"]';

const visible = (el: HTMLElement) => el.getClientRects().length > 0 && !el.closest("[aria-hidden='true'], [inert]");
const enabled = (el: HTMLElement) => !el.matches(":disabled, [aria-disabled='true']");

export function isField(el: Element | null): boolean {
  return !!el && el.matches(FIELD);
}

/** Ok tuşlarıyla gezilen öğeler (yazı alanları hariç; onlara Tab ile geçilir). */
export function navItems(root: HTMLElement): HTMLElement[] {
  return Array.from(root.querySelectorAll<HTMLElement>(ITEM)).filter((el) => !el.matches(FIELD) && enabled(el) && visible(el));
}

/** Tab sırasındaki öğeler. */
export function tabbables(root: HTMLElement): HTMLElement[] {
  return Array.from(root.querySelectorAll<HTMLElement>(TABBABLE)).filter((el) => el.tabIndex >= 0 && enabled(el) && visible(el) && !el.classList.contains("oh-hidden-input"));
}

/** Harfle atlama etiketi: öğenin adı (kısayol ve açıklama hariç). */
function itemLabel(el: HTMLElement): string {
  const label = el.querySelector(".oh-menu-label");
  const own = label?.firstChild?.textContent ?? label?.textContent;
  return el.getAttribute("aria-label") ?? own ?? el.textContent ?? "";
}

/** Odağı sayfayı kaydırmadan taşır; öğe kaydırılan bir menüdeyse yalnız o menüyü kaydırır. */
export function focusEl(el: HTMLElement | null | undefined, stop?: HTMLElement): void {
  if (!el) return;
  el.focus({ preventScroll: true });
  for (let p = el.parentElement; p; p = p.parentElement) {
    if (p.scrollHeight > p.clientHeight + 1 && /auto|scroll/.test(getComputedStyle(p).overflowY)) {
      const r = el.getBoundingClientRect();
      const pr = p.getBoundingClientRect();
      p.scrollTop = scrollToReveal(p.scrollTop, p.clientHeight, r.top - pr.top + p.scrollTop, r.height);
    }
    if (p === stop) break;
  }
}

/* ------------------------------------------------------------------ */
/* Katman türü                                                          */
/* ------------------------------------------------------------------ */

export interface LayerInfo {
  /** "menu": ok tuşlarıyla gezilen menü; "panel": yazı alanı da içeren menü benzeri katman. */
  kind: "menu" | "panel" | null;
  /** Fareyle açılınca odaklanan kap (role="menu"). */
  container: HTMLElement;
}

/** Katmanın içeriği menü mü? role="menu" ya da .oh-menu kabı (listbox/dialog rollü olanlar değil). */
export function layerInfo(root: HTMLElement): LayerInfo {
  const box = root.matches('[role="menu"]') ? root : root.querySelector<HTMLElement>('[role="menu"], .oh-menu');
  if (!box) return { kind: null, container: root };
  const role = box.getAttribute("role");
  if (role && role !== "menu") return { kind: null, container: root };
  const pure = !box.querySelector(FIELD);
  return { kind: pure ? "menu" : "panel", container: box };
}

/** Uygulamaların düz .oh-menu kapları da menü rolü ve odaklanabilirlik kazanır. */
export function prepareLayer(info: LayerInfo): void {
  const box = info.container;
  if (info.kind === "menu" && !box.hasAttribute("role")) box.setAttribute("role", "menu");
  if (info.kind && !box.hasAttribute("tabindex")) box.setAttribute("tabindex", "-1");
}

/** Açılışta odak: klavyeyle açıldıysa ilk öğe, fareyle açıldıysa kap (vurgu yok). */
export function focusOnOpen(root: HTMLElement, info: LayerInfo, mode: "first" | "container"): boolean {
  const box = info.container;
  if (!box.hasAttribute("tabindex")) box.setAttribute("tabindex", "-1");
  if (mode === "first") {
    const first = info.kind === "menu" ? navItems(root)[0] : (tabbables(root)[0] ?? navItems(root)[0]);
    if (first) {
      setKeyboardMode(root, true);
      focusEl(first, root);
      return true;
    }
  }
  box.focus({ preventScroll: true });
  return document.activeElement === box;
}

/**
 * Klavyeyle gezilirken (data-oh-kbd) imlecin altında kalan öğenin üzerine gelme vurgusu bastırılır;
 * tek vurgu odaklı öğede kalır. Fare gerçekten kıpırdayınca kalkar.
 */
export function setKeyboardMode(el: HTMLElement, on: boolean): void {
  if (on) el.setAttribute("data-oh-kbd", "");
  else el.removeAttribute("data-oh-kbd");
}

/* ------------------------------------------------------------------ */
/* Tuşlar                                                               */
/* ------------------------------------------------------------------ */

export interface MenuNavOptions {
  /** Alt menüyse ← kapatır. */
  submenu: boolean;
  /** Katmanı kapatır (← alt menüde, Tab saf menüde). */
  close: (reason: "arrow" | "tab") => void;
  /** Tab ile kapanınca odağın döneceği öğe ve bu öğe menüyü açan düğme mi. */
  tabTarget: () => { el: HTMLElement | null; opener: boolean };
  /** Menü olmayan ama odak alan katman (daraltılmış grup): yalnız Tab içeride döner. */
  tabLoop?: boolean;
}

const CONSUMED = new Set(["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight", "Home", "End", "PageUp", "PageDown", "Enter", " ", "Backspace", "Delete"]);

export interface MenuNav {
  onKeyDown: (e: ReactKeyboardEvent<HTMLElement>, root: HTMLElement, opts: MenuNavOptions) => void;
  onMouseMove: (e: { target: EventTarget; clientX: number; clientY: number }, root: HTMLElement) => void;
}

export function createMenuNav(): MenuNav {
  const typeahead = createTypeahead();
  let hovered: HTMLElement | null = null;
  let pointer: { x: number; y: number } | null = null;

  const onKeyDown: MenuNav["onKeyDown"] = (e, root, opts) => {
    const target = e.target as HTMLElement;
    // İç içe alt menüler React ağacında üst katmana da kabarır; yalnız kendi DOM'umuzdaki olaylar.
    if (e.defaultPrevented || !root.contains(target) || e.nativeEvent.isComposing) return;
    const info = layerInfo(root);
    if (!info.kind) {
      if (e.key === "Tab" && opts.tabLoop) onTab(e, root, { ...info, kind: "panel" }, opts);
      return;
    }
    if (e.key === "Tab") return onTab(e, root, info, opts);
    if (isField(target) || e.ctrlKey || e.metaKey || e.altKey) return;

    const items = navItems(root);
    const cur = items.indexOf(target);
    const go = (i: number) => {
      if (i < 0 || !items[i]) return;
      hovered = null;
      focusEl(items[i], root);
    };
    const boxes = (): Box[] => items.map((el) => el.getBoundingClientRect());
    const from = cur >= 0 ? cur : hovered ? items.indexOf(hovered) : -1;

    let handled = true;
    switch (e.key) {
      case "ArrowDown":
      case "ArrowUp": {
        const down = e.key === "ArrowDown";
        if (from < 0) go(down ? firstEnabled(items.length) : lastEnabled(items.length));
        else if (cur < 0) go(from);
        else go(spatialNext(boxes(), cur, down ? "down" : "up"));
        break;
      }
      case "ArrowRight":
      case "ArrowLeft": {
        const right = e.key === "ArrowRight";
        const side = cur >= 0 ? spatialNext(boxes(), cur, right ? "right" : "left") : -1;
        if (side >= 0) go(side);
        else if (right && cur >= 0 && target.getAttribute("aria-haspopup") && target.getAttribute("aria-expanded") !== "true") target.click();
        else if (!right && opts.submenu) opts.close("arrow");
        break;
      }
      case "Home":
      case "PageUp":
        go(firstEnabled(items.length));
        break;
      case "End":
      case "PageDown":
        go(lastEnabled(items.length));
        break;
      case "Enter":
      case " ":
        if (cur >= 0) target.click();
        break;
      default:
        if (e.key.length === 1 || /^\p{L}$/u.test(e.key)) {
          const q = typeahead(e.key);
          go(matchByPrefix(items.map(itemLabel), from, q));
        } else handled = CONSUMED.has(e.key);
    }
    if (handled) {
      e.preventDefault();
      e.stopPropagation();
      setKeyboardMode(root, true);
    }
  };

  const onTab = (e: ReactKeyboardEvent<HTMLElement>, root: HTMLElement, info: LayerInfo, opts: MenuNavOptions) => {
    e.stopPropagation();
    if (info.kind === "menu") {
      // Saf menü: Tab menüyü kapatır. Klavyeyle açıldıysa odak düğmeye döner ve tarayıcı
      // Tab'ı oradan sürdürür (sonraki denetime geçer); fareyle açıldıysa odak eski yerine döner.
      const { el, opener } = opts.tabTarget();
      if (el?.isConnected) el.focus({ preventScroll: true });
      if (!opener || !el) e.preventDefault();
      opts.close("tab");
      return;
    }
    // Yazı alanı içeren katman: Tab içeride döner.
    const list = tabbables(root);
    if (!list.length) return e.preventDefault();
    const i = list.indexOf(e.target as HTMLElement);
    const next = e.shiftKey ? (i <= 0 ? list.length - 1 : -1) : i < 0 || i === list.length - 1 ? 0 : -1;
    if (next >= 0) {
      e.preventDefault();
      focusEl(list[next], root);
    }
  };

  const onMouseMove: MenuNav["onMouseMove"] = (e, root) => {
    // Tarayıcı, içerik imlecin altında kayınca yerinde duran fare için de olay üretir; yalnız gerçek hareket sayılır.
    const moved = pointer !== null && (pointer.x !== e.clientX || pointer.y !== e.clientY);
    pointer = { x: e.clientX, y: e.clientY };
    if (!moved) return;
    setKeyboardMode(root, false);
    const item = (e.target as Element | null)?.closest?.<HTMLElement>(ITEM);
    if (!item || !root.contains(item) || item.matches(FIELD) || !enabled(item)) {
      hovered = null;
      return;
    }
    if (item === hovered) return;
    hovered = item;
    // Klavyeyle bir öğedeyken fare başka öğeye gelirse tek vurgu kalsın: odak da oraya geçer.
    const a = document.activeElement as HTMLElement | null;
    if (a && a !== item && root.contains(a) && a.matches(ITEM) && !a.matches(FIELD)) item.focus({ preventScroll: true });
  };

  return { onKeyDown, onMouseMove };
}
