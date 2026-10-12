// Sayfa Yapısı iletişim kutusu ve Sütunlar iletişim kutusu.
import { useState } from "react";
import { Dialog, DialogButton, DialogTabs, Field, FieldSet } from "@officehere/ui";
import { useBelge } from "../state/context";
import { AutoFocus } from "./AutoFocus";
import { PAGE_SIZE_LIST, cmToTw, pageSizeId, twToCm, type PageSetup } from "../model/document";

const f = (tw: number) => Number(twToCm(tw).toFixed(2)).toLocaleString("tr-TR");
const p = (s: string) => {
  const n = parseFloat(s.replace(",", ".").replace(/[^\d.-]/g, ""));
  return Number.isFinite(n) ? n : 0;
};

export function PageSetupDialog({ initialTab, onClose }: { initialTab?: string; onClose: () => void }) {
  const { settings, updateSettings } = useBelge();
  const pg = settings.page;
  const [tab, setTab] = useState(initialTab === "paper" ? "paper" : initialTab === "layout" ? "layout" : "margins");
  const [top, setTop] = useState(f(pg.margin.top));
  const [bottom, setBottom] = useState(f(pg.margin.bottom));
  const [left, setLeft] = useState(f(pg.margin.left));
  const [right, setRight] = useState(f(pg.margin.right));
  const [gutter, setGutter] = useState(f(pg.margin.gutter));
  const [orient, setOrient] = useState(pg.orientation);
  const [size, setSize] = useState(pageSizeId(pg));
  const [w, setW] = useState(f(pg.width));
  const [h, setH] = useState(f(pg.height));
  const [hd, setHd] = useState(f(pg.margin.header));
  const [fd, setFd] = useState(f(pg.margin.footer));

  const pickSize = (id: string) => {
    setSize(id);
    const s = PAGE_SIZE_LIST.find((x) => x.id === id);
    if (!s) return;
    setW(f(orient === "landscape" ? s.h : s.w));
    setH(f(orient === "landscape" ? s.w : s.h));
  };

  const pickOrient = (o: "portrait" | "landscape") => {
    if (o === orient) return;
    setOrient(o);
    setW(h);
    setH(w);
  };

  const apply = () => {
    const next: PageSetup = {
      ...pg,
      orientation: orient,
      width: cmToTw(p(w)) || pg.width,
      height: cmToTw(p(h)) || pg.height,
      margin: { top: cmToTw(p(top)), bottom: cmToTw(p(bottom)), left: cmToTw(p(left)), right: cmToTw(p(right)), gutter: cmToTw(p(gutter)), header: cmToTw(p(hd)), footer: cmToTw(p(fd)) },
    };
    updateSettings((s) => ({ ...s, page: next }));
    onClose();
  };

  const num = (label: string, v: string, set: (s: string) => void) => (
    <Field label={label}>
      <input type="text" value={v} onChange={(e) => set(e.target.value)} />
    </Field>
  );

  return (
    <Dialog
      open
      title="Sayfa Yapısı"
      width={460}
      onClose={onClose}
      onSubmit={apply}
      footer={
        <>
          <DialogButton primary onClick={apply}>
            Tamam
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <DialogTabs
        tabs={[
          { id: "margins", label: "Kenar Boşlukları" },
          { id: "paper", label: "Kağıt" },
          { id: "layout", label: "Düzen" },
        ]}
        active={tab}
        onChange={setTab}
      />
      <AutoFocus deps={[tab]} />
      {tab === "margins" && (
        <>
          <FieldSet legend="Kenar Boşlukları (cm)">
            <div className="bl-dlg-grid">
              {num("Üst:", top, setTop)}
              {num("Alt:", bottom, setBottom)}
              {num("Sol:", left, setLeft)}
              {num("Sağ:", right, setRight)}
              {num("Cilt payı:", gutter, setGutter)}
            </div>
          </FieldSet>
          <FieldSet legend="Yönlendirme">
            <div className="bl-orient">
              <button type="button" className={orient === "portrait" ? "is-selected" : ""} onClick={() => pickOrient("portrait")}>
                <span className="bl-orient-ico is-portrait" />
                Dikey
              </button>
              <button type="button" className={orient === "landscape" ? "is-selected" : ""} onClick={() => pickOrient("landscape")}>
                <span className="bl-orient-ico is-landscape" />
                Yatay
              </button>
            </div>
          </FieldSet>
        </>
      )}
      {tab === "paper" && (
        <FieldSet legend="Kağıt boyutu">
          <Field label="Boyut:">
            <select value={size} onChange={(e) => pickSize(e.target.value)}>
              {PAGE_SIZE_LIST.map((s) => (
                <option key={s.id} value={s.id}>
                  {s.label}
                </option>
              ))}
              <option value="custom">Özel boyut</option>
            </select>
          </Field>
          <div className="bl-dlg-grid">
            {num("Genişlik (cm):", w, (v) => (setW(v), setSize("custom")))}
            {num("Yükseklik (cm):", h, (v) => (setH(v), setSize("custom")))}
          </div>
        </FieldSet>
      )}
      {tab === "layout" && (
        <FieldSet legend="Üst bilgi ve alt bilgi — kenardan (cm)">
          <div className="bl-dlg-grid">
            {num("Üst bilgi:", hd, setHd)}
            {num("Alt bilgi:", fd, setFd)}
          </div>
        </FieldSet>
      )}
    </Dialog>
  );
}

export function ColumnsDialog({ onClose }: { onClose: () => void }) {
  const { settings, updateSettings } = useBelge();
  const [n, setN] = useState(settings.page.columns);
  const [gap, setGap] = useState(f(settings.page.columnSpace));
  const apply = () => {
    updateSettings((s) => ({ ...s, page: { ...s.page, columns: Math.max(1, Math.min(3, n)), columnSpace: cmToTw(p(gap)) || 708 } }));
    onClose();
  };
  return (
    <Dialog
      open
      title="Sütunlar"
      width={380}
      onClose={onClose}
      onSubmit={apply}
      footer={
        <>
          <DialogButton primary onClick={apply}>
            Tamam
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <FieldSet legend="Önayarlar">
        <div className="bl-orient">
          {[1, 2, 3].map((k) => (
            <button key={k} type="button" className={n === k ? "is-selected" : ""} onClick={() => setN(k)}>
              <span className={`bl-cols-ico is-${k}`}>
                {Array.from({ length: k }, (_, i) => (
                  <i key={i} />
                ))}
              </span>
              {["Bir", "İki", "Üç"][k - 1]}
            </button>
          ))}
        </div>
      </FieldSet>
      <Field label="Aralık (cm):" inline>
        <input type="text" value={gap} onChange={(e) => setGap(e.target.value)} />
      </Field>
      <p className="bl-dlg-note">Sütunlar tüm belgeye uygulanır. Paragraflar sütunlar arasında bölünmez.</p>
    </Dialog>
  );
}
