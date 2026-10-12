/** Seçim Bölmesi: sayfadaki grafik, resim ve şekilleri listeler; seçme, gizleme/gösterme. */
import { Dismiss16Regular, Eye16Regular, EyeOff16Regular } from "@fluentui/react-icons";
import type { Controller } from "../controller";
import { objectList, setObjectHidden } from "../actions/objectClip";

export function SelectionPane({ ctl }: { ctl: Controller }) {
  const items = objectList(ctl).reverse();
  const close = () => {
    ctl.formatPane = null;
    ctl.emit();
    ctl.focusGrid();
  };
  const all = (hidden: boolean) => {
    for (const o of items) if (o.hidden !== hidden) setObjectHidden(ctl, o.kind, o.id, hidden);
  };
  return (
    <aside className="tb-pane tb-sel-pane" aria-label="Seçim">
      <div className="tb-pane-head">
        <span>Seçim</span>
        <button type="button" className="tb-pane-close" aria-label="Kapat" onClick={close}>
          <Dismiss16Regular />
        </button>
      </div>
      <div className="tb-sel-actions">
        <button type="button" className="oh-btn" onClick={() => all(false)} disabled={!items.length}>
          Tümünü Göster
        </button>
        <button type="button" className="oh-btn" onClick={() => all(true)} disabled={!items.length}>
          Tümünü Gizle
        </button>
      </div>
      <div className="tb-sel-list" role="listbox" aria-label="Nesneler">
        {!items.length && <div className="tb-sel-empty">Bu sayfada nesne yok.</div>}
        {items.map((o) => (
          <div
            key={o.id}
            role="option"
            aria-selected={ctl.selObj?.id === o.id}
            className={`tb-sel-item${ctl.selObj?.id === o.id ? " is-active" : ""}${o.hidden ? " is-hidden" : ""}`}
            onClick={() => {
              if (o.hidden) return;
              ctl.selObj = { kind: o.kind, id: o.id };
              ctl.emit();
            }}
          >
            <span className="tb-sel-name">{o.name}</span>
            <button type="button" className="tb-pane-close" aria-label={o.hidden ? "Göster" : "Gizle"} onClick={(e) => (e.stopPropagation(), setObjectHidden(ctl, o.kind, o.id, !o.hidden))}>
              {o.hidden ? <EyeOff16Regular /> : <Eye16Regular />}
            </button>
          </div>
        ))}
      </div>
    </aside>
  );
}
