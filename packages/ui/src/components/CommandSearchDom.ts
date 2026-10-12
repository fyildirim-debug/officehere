/**
 * Komut aramasının şeritle konuşan yardımcıları: gizli kapta çizilen sekmelerden
 * dizin çıkarma ve canlı şeritte komutun öğesini bulma.
 */
import type { RibbonTabDef } from "./Ribbon";
import type { CommandEntry } from "./commandIndex";

/** `[data-cmd-tab]` kaplarındaki ipuçlu öğelerden (data-tip) komut dizini çıkarır. */
export function scrapeCommands(root: HTMLElement, tabs: RibbonTabDef[]): CommandEntry[] {
  const out: CommandEntry[] = [];
  root.querySelectorAll<HTMLElement>("[data-cmd-tab]").forEach((pane) => {
    const tabId = pane.dataset.cmdTab!;
    const tabLabel = tabs.find((t) => t.id === tabId)?.label ?? "";
    const counts = new Map<string, number>();
    pane.querySelectorAll<HTMLElement>("[data-tip]").forEach((el) => {
      const title = el.dataset.tip!;
      const group = el.closest<HTMLElement>(".oh-ribbon-group")?.getAttribute("aria-label") ?? "";
      if (el.classList.contains("oh-ribbon-launcher")) return;
      const key = `${group}\u0000${title}`;
      const nth = counts.get(key) ?? 0;
      counts.set(key, nth + 1);
      const icon = el.querySelector(".oh-rbtn-icon > svg")?.outerHTML;
      out.push({ id: `${tabId}:${key}:${nth}`, title, desc: el.dataset.tipDesc, shortcut: el.dataset.tipKey, tabId, tabLabel, group, nth, icon });
    });
  });
  // Aynı adlı komutlardan yalnız ilki (ör. birden çok sekmedeki "Yazı Tipi").
  const seen = new Set<string>();
  return out.filter((c) => {
    const k = `${c.title}|${c.desc ?? ""}`;
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  });
}

/** Şerit komutlarını ek komutlarla birleştirir: aynı sekme + addaki ek komut yinelenmez, açıklaması şerit girdisine katılır. */
export function mergeCommands(scraped: CommandEntry[], extra: CommandEntry[] = []): CommandEntry[] {
  const byName = new Map(extra.map((c) => [`${c.tabLabel}|${c.title}`, c]));
  for (const c of scraped) {
    const x = byName.get(`${c.tabLabel}|${c.title}`);
    if (x?.desc && !c.desc) c.desc = x.desc;
  }
  const names = new Set(scraped.map((c) => `${c.tabLabel}|${c.title}`));
  return [...scraped, ...extra.filter((c) => !names.has(`${c.tabLabel}|${c.title}`))];
}

/** Şeritte bu sekmenin içeriği şu an görünüyor mu (sekme düğmeleri görünür sekmelerle aynı sırada çizilir). */
export function isShowing(tabs: RibbonTabDef[], tabId: string, collapsed: boolean): boolean {
  if (collapsed || !document.querySelector(".oh-ribbon-body")) return false;
  const buttons = Array.from(document.querySelectorAll(".oh-ribbon-tabs [role=tab]"));
  const i = buttons.findIndex((b) => b.getAttribute("aria-selected") === "true");
  return i >= 0 && tabs.filter((t) => !t.hidden)[i]?.id === tabId;
}

export const nextFrame = () => new Promise<void>((r) => requestAnimationFrame(() => requestAnimationFrame(() => r())));

/** Canlı şeritte komutun öğesini bulur (gerekirse daraltılmış grubu açar). */
export async function locateCommand(c: CommandEntry): Promise<HTMLElement | null> {
  const body = document.querySelector(".oh-ribbon-body");
  if (!body) return null;
  const sel = `[data-tip="${CSS.escape(c.title)}"]`;
  const group = Array.from(body.querySelectorAll<HTMLElement>(".oh-ribbon-group[role=group]")).find((g) => g.getAttribute("aria-label") === c.group);
  if (group) return group.querySelectorAll<HTMLElement>(sel)[c.nth] ?? null;
  const collapsed = Array.from(body.querySelectorAll<HTMLElement>(".oh-ribbon-group.is-collapsed")).find((g) => g.textContent?.trim().startsWith(c.group));
  const btn = collapsed?.querySelector("button");
  if (btn) {
    btn.click();
    await nextFrame();
    const pop = Array.from(document.querySelectorAll<HTMLElement>(".oh-ribbon-group-popup")).pop();
    return pop?.querySelectorAll<HTMLElement>(sel)[c.nth] ?? null;
  }
  return body.querySelector<HTMLElement>(sel);
}

/** Bulunan öğeyi çalıştırır; düğme yoksa ya da devre dışıysa false döner. */
export function activateCommand(el: HTMLElement): boolean {
  if (el instanceof HTMLInputElement || el instanceof HTMLSelectElement) {
    el.focus();
    return true;
  }
  const target = el.matches("button") ? el : (el.querySelector<HTMLElement>("button, input") ?? el);
  if ((target as HTMLButtonElement).disabled) return false;
  target.click();
  return true;
}
