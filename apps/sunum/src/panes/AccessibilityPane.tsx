/** Erişilebilirlik bölmesi: denetim sonuçları (Hatalar / Uyarılar), nesneye gitme ve alternatif metin ekleme. */
import { useMemo, useState } from "react";
import { checkAccessibility, ISSUE_HELP, ISSUE_TITLES, type A11yIssue, type IssueKind, type IssueLevel } from "../model/accessibility";
import type { Actions } from "../state/actions";
import type { AppState } from "../state/store";
import { PaneShell } from "./paneControls";

const LEVELS: { level: IssueLevel; label: string }[] = [
  { level: "error", label: "Hatalar" },
  { level: "warning", label: "Uyarılar" },
  { level: "tip", label: "İpuçları" },
];

export function AccessibilityPane({ state, actions: a, onClose }: { state: AppState; actions: Actions; onClose: () => void }) {
  const issues = useMemo(() => checkAccessibility(state.pres), [state.pres]);
  const [active, setActive] = useState<A11yIssue | null>(null);
  const [alt, setAlt] = useState("");
  const pick = (it: A11yIssue) => {
    setActive(it);
    setAlt("");
    a.goToSlide(it.slideIndex);
    if (it.shapeId) a.store.set({ sel: [it.shapeId], editing: null, focus: "canvas" });
  };
  const saveAlt = () => {
    if (!active?.shapeId || !alt.trim()) return;
    a.updateShapes((s) => ({ ...s, descr: alt.trim() }), [active.shapeId]);
    setActive(null);
  };
  const groups = LEVELS.map((l) => ({ ...l, kinds: [...new Set(issues.filter((i) => i.level === l.level).map((i) => i.kind))] })).filter((g) => g.kinds.length);
  return (
    <PaneShell title="Erişilebilirlik" onClose={onClose}>
      <div className="a11y-pane">
        <div className="a11y-head">Denetim Sonuçları</div>
        {!issues.length && <p className="a11y-ok">Erişilebilirlik sorunu bulunamadı. Engelli kişiler bu sunuyu okumakta zorlanmamalıdır.</p>}
        {groups.map((g) => (
          <details key={g.level} open className="a11y-group">
            <summary>{g.label}</summary>
            {g.kinds.map((k: IssueKind) => {
              const list = issues.filter((i) => i.level === g.level && i.kind === k);
              return (
                <details key={k} open className="a11y-kind">
                  <summary>
                    {ISSUE_TITLES[k]} ({list.length})
                  </summary>
                  {list.map((it, n) => (
                    <div key={n} role="button" tabIndex={0} className={`list-row${active === it ? " is-selected" : ""}`} onClick={() => pick(it)} onKeyDown={(e) => e.key === "Enter" && pick(it)}>
                      <span className="grow">{it.label}</span>
                      <span className="a11y-slide">Slayt {it.slideIndex + 1}</span>
                    </div>
                  ))}
                </details>
              );
            })}
          </details>
        ))}
        {active && (
          <div className="a11y-info">
            <div className="a11y-head">Ek Bilgi</div>
            <div style={{ fontWeight: 600 }}>{ISSUE_TITLES[active.kind]}</div>
            <p>{ISSUE_HELP[active.kind]}</p>
            {active.kind === "altText" && (
              <div className="a11y-alt">
                <label htmlFor="a11y-alt">Alternatif metin:</label>
                <textarea id="a11y-alt" rows={3} value={alt} placeholder="Nesneyi kısaca tanımlayın" onChange={(e) => setAlt(e.target.value)} onKeyDown={(e) => e.stopPropagation()} />
                <button type="button" className="oh-btn" disabled={!alt.trim()} onClick={saveAlt}>
                  Alternatif Metin Ekle
                </button>
              </div>
            )}
          </div>
        )}
      </div>
    </PaneShell>
  );
}
