/**
 * ODP gösteri ayarları ve özel gösteriler (ODF 1.3 §9.6 presentation:settings,
 * presentation:show). Slayt aralığında ODF yalnız başlangıç sayfasını saklar.
 */
import { uid } from "../../model/factory";
import type { CustomShow, Presentation } from "../../model/types";
import { DEFAULT_SHOW_SETTINGS, type ShowSettings } from "../../show/settings";
import { attr, kids } from "../xml";
import { odfAttrs } from "./common";

/** office:presentation içinde sayfalardan sonra yazılan ayar öğesi; ayar ve özel gösteri yoksa boş. */
export function odpSettingsXml(pres: Presentation, pageNames: string[]): string {
  const s = pres.showSettings;
  const shows = pres.customShows ?? [];
  if (!s && !shows.length) return "";
  const nameOf = new Map(pres.slides.map((sl, i) => [sl.id, pageNames[i]]));
  const custom = s?.range.kind === "custom" ? shows.find((c) => c.id === (s.range as { id: string }).id) : undefined;
  const a = odfAttrs({
    "presentation:show": custom?.name,
    "presentation:start-page": s?.range.kind === "range" ? pageNames[s.range.from - 1] : undefined,
    "presentation:endless": s && (s.loop || s.type === "kiosk") ? "true" : undefined,
    "presentation:pause": s?.type === "kiosk" ? "PT0S" : undefined,
    "presentation:force-manual": s?.advance === "manual" ? "true" : undefined,
    "presentation:full-screen": s?.type === "browsed" ? "false" : undefined,
    "presentation:animations": s?.noAnimations ? "disabled" : undefined,
  });
  const list = shows.map((c) => `<presentation:show${odfAttrs({ "presentation:name": c.name, "presentation:pages": c.slideIds.map((id) => nameOf.get(id)).filter(Boolean).join(",") })}/>`).join("");
  return `<presentation:settings${a}>${list}</presentation:settings>`;
}

/** presentation:settings okuması; sayfa adı → slayt kimliği eşlemesiyle. */
export function readOdpSettings(body: Element | undefined, byName: Map<string, string>, slideIds: string[]): { customShows?: CustomShow[]; showSettings?: ShowSettings } {
  const el = kids(body, "settings")[0];
  if (!el) return {};
  const customShows: CustomShow[] = kids(el, "show").map((sh, i) => ({
    id: uid("cs"),
    name: attr(sh, "name") ?? `Özel Gösteri ${i + 1}`,
    slideIds: (attr(sh, "pages") ?? "")
      .split(",")
      .map((n) => byName.get(n.trim()))
      .filter((x): x is string => !!x),
  }));
  const on = (n: string) => attr(el, n) === "true";
  const has = ["show", "start-page", "endless", "pause", "force-manual", "full-screen", "animations"].some((n) => attr(el, n) !== undefined);
  let showSettings: ShowSettings | undefined;
  if (has) {
    const s: ShowSettings = { ...DEFAULT_SHOW_SETTINGS, range: { kind: "all" } };
    s.loop = on("endless");
    s.type = on("endless") && attr(el, "pause") !== undefined ? "kiosk" : attr(el, "full-screen") === "false" ? "browsed" : "speaker";
    s.advance = on("force-manual") ? "manual" : "timings";
    s.noAnimations = attr(el, "animations") === "disabled";
    const custom = customShows.find((c) => c.name === attr(el, "show"));
    const start = attr(el, "start-page") ? byName.get(attr(el, "start-page")!) : undefined;
    if (custom) s.range = { kind: "custom", id: custom.id };
    else if (start) s.range = { kind: "range", from: slideIds.indexOf(start) + 1, to: slideIds.length };
    showSettings = s;
  }
  return { ...(customShows.length ? { customShows } : {}), ...(showSettings ? { showSettings } : {}) };
}
