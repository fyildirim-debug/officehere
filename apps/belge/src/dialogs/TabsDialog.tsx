// Sekmeler iletişim kutusu: sekme durak konumu, hizalama (sol/orta/sağ/ondalık/çubuk), öncü karakter;
// Ayarla / Temizle / Tümünü Temizle. Duraklar imleçteki (seçimdeki) paragraflara uygulanır.
import { useState } from "react";
import { Dialog, DialogButton, FieldSet } from "@officehere/ui";
import { useBelge } from "../state/context";
import { currentPara, setParaAttrs } from "../editor/commands";
import { DEFAULT_TAB, TAB_ALIGN_LABEL, normalizeTabs, type TabAlign, type TabLeader, type TabStop } from "../editor/tabStops";

const PT_PER_CM = 72 / 2.54;
const fmtCm = (pt: number) => `${Number((pt / PT_PER_CM).toFixed(2)).toLocaleString("tr-TR")} cm`;
const parseCm = (s: string) => {
  const n = parseFloat(s.replace(",", ".").replace(/[^\d.-]/g, ""));
  return Number.isFinite(n) ? n * PT_PER_CM : null;
};

const LEADERS: { id: TabLeader; label: string }[] = [
  { id: "none", label: "1 Yok" },
  { id: "dot", label: "2 ......." },
  { id: "hyphen", label: "3 -------" },
  { id: "underscore", label: "4 ______" },
];

export function TabsDialog({ onClose }: { onClose: () => void }) {
  const { editor } = useBelge();
  const p = currentPara(editor);
  const [tabs, setTabs] = useState<TabStop[]>(() => ((p?.node.attrs.tabs as TabStop[] | null) ?? []).slice());
  const [text, setText] = useState(tabs[0] ? fmtCm(tabs[0].pos) : "");
  const [align, setAlign] = useState<TabAlign>(tabs[0]?.align ?? "left");
  const [leader, setLeader] = useState<TabLeader>(tabs[0]?.leader ?? "none");
  const [cleared, setCleared] = useState<string[]>([]);
  const pos = parseCm(text);

  const set = () => {
    if (pos === null || pos < 0) return;
    setTabs((t) => normalizeTabs([...t.filter((x) => Math.abs(x.pos - pos) > 0.5), { pos, align, leader }]));
  };
  const clear = () => {
    if (pos === null) return;
    const hit = tabs.find((x) => Math.abs(x.pos - pos) < 0.5);
    if (!hit) return;
    setTabs((t) => t.filter((x) => x !== hit));
    setCleared((c) => [...c, fmtCm(hit.pos)]);
  };
  const ok = () => {
    let list = tabs;
    // Yazılmış ama "Ayarla"ya basılmamış durak da eklenir (Word'deki gibi).
    if (pos !== null && pos >= 0 && !tabs.some((x) => Math.abs(x.pos - pos) < 0.5) && !cleared.includes(fmtCm(pos))) list = normalizeTabs([...tabs, { pos, align, leader }]);
    setParaAttrs(editor, { tabs: list.length ? list : null });
    onClose();
  };

  return (
    <Dialog
      open
      title="Sekmeler"
      width={460}
      onClose={onClose}
      onSubmit={ok}
      footer={
        <>
          <DialogButton onClick={set} disabled={pos === null}>
            Ayarla
          </DialogButton>
          <DialogButton onClick={clear} disabled={pos === null || !tabs.some((x) => Math.abs(x.pos - (pos ?? -1)) < 0.5)}>
            Temizle
          </DialogButton>
          <DialogButton
            onClick={() => {
              setCleared((c) => [...c, ...tabs.map((x) => fmtCm(x.pos))]);
              setTabs([]);
            }}
            disabled={!tabs.length}
          >
            Tümünü Temizle
          </DialogButton>
          <DialogButton primary onClick={ok}>
            Tamam
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <div className="bl-tabs-dlg">
        <div className="bl-tabs-col">
          <label className="bl-tabs-field">
            Sekme durak konumu:
            <input type="text" aria-label="Sekme durak konumu" value={text} onChange={(e) => setText(e.target.value)} />
          </label>
          <div className="bl-tabs-list" role="listbox" aria-label="Sekme durakları">
            {tabs.map((t) => (
              <div
                key={t.pos}
                role="option"
                aria-selected={pos !== null && Math.abs(t.pos - pos) < 0.5}
                className={pos !== null && Math.abs(t.pos - pos) < 0.5 ? "is-selected" : undefined}
                onClick={() => {
                  setText(fmtCm(t.pos));
                  setAlign(t.align);
                  setLeader(t.leader);
                }}
              >
                {fmtCm(t.pos)} <span>{TAB_ALIGN_LABEL[t.align]}</span>
              </div>
            ))}
          </div>
        </div>
        <div className="bl-tabs-col">
          <label className="bl-tabs-field">
            Varsayılan sekme durakları:
            <input type="text" aria-label="Varsayılan sekme durakları" value={fmtCm(DEFAULT_TAB)} readOnly />
          </label>
          {cleared.length > 0 && <div className="bl-tabs-cleared">Temizlenecek sekme durakları: {cleared.join("; ")}</div>}
        </div>
      </div>
      <FieldSet legend="Hizalama">
        <div className="bl-tabs-radios">
          {(Object.keys(TAB_ALIGN_LABEL) as TabAlign[]).map((a) => (
            <label key={a} className="bl-radio">
              <input type="radio" name="tab-align" checked={align === a} onChange={() => setAlign(a)} /> {TAB_ALIGN_LABEL[a]}
            </label>
          ))}
        </div>
      </FieldSet>
      <FieldSet legend="Öncü karakter">
        <div className="bl-tabs-radios">
          {LEADERS.map((l) => (
            <label key={l.id} className="bl-radio">
              <input type="radio" name="tab-leader" checked={leader === l.id} onChange={() => setLeader(l.id)} /> {l.label}
            </label>
          ))}
        </div>
      </FieldSet>
    </Dialog>
  );
}
