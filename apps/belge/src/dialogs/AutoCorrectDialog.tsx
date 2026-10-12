// Otomatik Düzeltme seçenekleri: Otomatik Düzelt (değiştirme listesi, büyük harf kuralları) ve
// Yazarken Otomatik Biçimlendir (akıllı tırnak, kesir, tire, sekme girintisi).
import { useMemo, useState } from "react";
import { Dialog, DialogButton, DialogTabs, FieldSet } from "@officehere/ui";
import { defaultAutoCorrect, getAutoCorrect, setAutoCorrect, type AutoCorrectOptions } from "../editor/autocorrectRules";
import { Check } from "./FontDialog";

export function AutoCorrectDialog({ onClose }: { onClose: () => void }) {
  const [tab, setTab] = useState("autocorrect");
  const [o, setO] = useState<AutoCorrectOptions>(() => ({ ...getAutoCorrect(), entries: getAutoCorrect().entries.slice() }));
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");
  const [selected, setSelected] = useState<string | null>(null);
  const set = <K extends keyof AutoCorrectOptions>(k: K) => (v: AutoCorrectOptions[K]) => setO((x) => ({ ...x, [k]: v }));
  const sorted = useMemo(() => o.entries.slice().sort((a, b) => a.from.localeCompare(b.from, "tr")), [o.entries]);
  const existing = o.entries.find((e) => e.from === from);

  const add = () => {
    if (!from.trim() || !to) return;
    setO((x) => ({ ...x, entries: [...x.entries.filter((e) => e.from !== from), { from, to }] }));
    setSelected(from);
  };
  const remove = () => {
    const key = selected ?? from;
    setO((x) => ({ ...x, entries: x.entries.filter((e) => e.from !== key) }));
    setSelected(null);
    setFrom("");
    setTo("");
  };
  const ok = () => {
    setAutoCorrect(o);
    onClose();
  };

  return (
    <Dialog
      open
      title="Otomatik Düzelt"
      width={520}
      onClose={onClose}
      onSubmit={ok}
      footer={
        <>
          <DialogButton onClick={() => setO(defaultAutoCorrect())}>Varsayılanlara Dön</DialogButton>
          <DialogButton primary onClick={ok}>
            Tamam
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <DialogTabs
        tabs={[
          { id: "autocorrect", label: "Otomatik Düzelt" },
          { id: "format", label: "Yazarken Otomatik Biçimlendir" },
        ]}
        active={tab}
        onChange={setTab}
      />
      {tab === "autocorrect" ? (
        <div className="bl-ac">
          <Check label="İKi BÜyük harfle başlayan sözcükleri düzelt" v={o.twoInitialCaps} set={set("twoInitialCaps")} />
          <Check label="Cümlelerin ilk harfini büyük yap" v={o.capFirstSentence} set={set("capFirstSentence")} />
          <Check label="Tablo hücrelerinin ilk harfini büyük yap" v={o.capFirstCell} set={set("capFirstCell")} />
          <Check label="Yazarken metni değiştir" v={o.replaceText} set={set("replaceText")} />
          <div className="bl-ac-edit">
            <label>
              Değiştir:
              <input type="text" aria-label="Değiştir" value={from} onChange={(e) => setFrom(e.target.value)} />
            </label>
            <label>
              Yerine:
              <input type="text" aria-label="Yerine" value={to} onChange={(e) => setTo(e.target.value)} />
            </label>
          </div>
          <div className="bl-ac-list" role="listbox" aria-label="Değiştirme listesi">
            {sorted.map((e) => (
              <div
                key={e.from}
                role="option"
                aria-selected={selected === e.from}
                className={selected === e.from ? "is-selected" : undefined}
                onClick={() => {
                  setSelected(e.from);
                  setFrom(e.from);
                  setTo(e.to);
                }}
              >
                <span>{e.from}</span>
                <span>{e.to}</span>
              </div>
            ))}
          </div>
          <div className="bl-ac-buttons">
            <DialogButton onClick={add} disabled={!from.trim() || !to}>
              {existing ? "Değiştir" : "Ekle"}
            </DialogButton>
            <DialogButton onClick={remove} disabled={!existing && !selected}>
              Sil
            </DialogButton>
          </div>
        </div>
      ) : (
        <div className="bl-ac">
          <FieldSet legend="Yazarken değiştir">
            <Check label={"“Düz tırnakları” “akıllı tırnaklarla”"} v={o.smartQuotes} set={set("smartQuotes")} />
            <Check label="Kesirleri (1/2) kesir karakteriyle (½)" v={o.fractions} set={set("fractions")} />
            <Check label="Kısa çizgileri (--) uzun çizgiyle (—)" v={o.dashes} set={set("dashes")} />
          </FieldSet>
          <FieldSet legend="Yazarken otomatik olarak">
            <Check label="Sol ve ilk satır girintisini sekme ve geri silme tuşlarıyla ayarla" v={o.tabIndent} set={set("tabIndent")} />
          </FieldSet>
        </div>
      )}
    </Dialog>
  );
}
