/** Tablo Ekle, Üst Bilgi ve Alt Bilgi, Simge, Tarih ve Saat, Simgeler iletişim kutuları. */
import { useMemo, useRef, useState } from "react";
import {
  Home24Filled,
  Person24Filled,
  People24Filled,
  Mail24Filled,
  Call24Filled,
  Chat24Filled,
  Calendar24Filled,
  Clock24Filled,
  Globe24Filled,
  Location24Filled,
  Map24Filled,
  Airplane24Filled,
  VehicleCar24Filled,
  Building24Filled,
  BuildingFactory24Filled,
  Lightbulb24Filled,
  Rocket24Filled,
  Trophy24Filled,
  Star24Filled,
  Heart24Filled,
  ThumbLike24Filled,
  Checkmark24Filled,
  Dismiss24Filled,
  Warning24Filled,
  Info24Filled,
  Question24Filled,
  Shield24Filled,
  LockClosed24Filled,
  Key24Filled,
  Settings24Filled,
  Wrench24Filled,
  DataBarVertical24Filled,
  DataPie24Filled,
  DataTrending24Filled,
  ChartMultiple24Filled,
  Money24Filled,
  Wallet24Filled,
  Cart24Filled,
  Gift24Filled,
  Briefcase24Filled,
  Document24Filled,
  Folder24Filled,
  Book24Filled,
  HatGraduation24Filled,
  Beaker24Filled,
  WeatherSunny24Filled,
  Cloud24Filled,
  Flash24Filled,
  Fire24Filled,
  Drop24Filled,
  Camera24Filled,
  Video24Filled,
  MusicNote224Filled,
  Headphones24Filled,
  Desktop24Filled,
  Phone24Filled,
  Laptop24Filled,
  Code24Filled,
  Bug24Filled,
  Target24Filled,
  Flag24Filled,
  Diamond24Filled,
} from "@fluentui/react-icons";
import type { FluentIcon } from "@fluentui/react-icons";
import { Dialog, DialogButton, DialogTabs, Field } from "@officehere/ui";
import { makeSp, NO_LINE } from "../model/factory";
import type { CustomPath, HeaderFooter } from "../model/types";
import { svgPathToCmds } from "../render/svgPath";
import type { Actions } from "../state/actions";

export function InsertTableDialog({ open, onClose, actions: a }: { open: boolean; onClose: () => void; actions: Actions }) {
  const [cols, setCols] = useState(5);
  const [rows, setRows] = useState(2);
  const ok = () => {
    a.insertTable(Math.max(1, Math.min(75, rows)), Math.max(1, Math.min(75, cols)));
    onClose();
  };
  return (
    <Dialog open={open} title="Tablo Ekle" onClose={onClose} onSubmit={ok} width={300} footer={<><DialogButton primary onClick={ok}>Tamam</DialogButton><DialogButton onClick={onClose}>İptal</DialogButton></>}>
      <Field label="Sütun sayısı:" inline>
        <input type="number" min={1} max={75} value={cols} onChange={(e) => setCols(Number(e.target.value))} style={{ width: 80 }} />
      </Field>
      <Field label="Satır sayısı:" inline>
        <input type="number" min={1} max={75} value={rows} onChange={(e) => setRows(Number(e.target.value))} style={{ width: 80 }} />
      </Field>
    </Dialog>
  );
}

const DATE_FORMATS: ((d: Date) => string)[] = [
  (d) => new Intl.DateTimeFormat("tr-TR", { day: "2-digit", month: "2-digit", year: "numeric" }).format(d),
  (d) => new Intl.DateTimeFormat("tr-TR", { day: "numeric", month: "long", year: "numeric" }).format(d),
  (d) => new Intl.DateTimeFormat("tr-TR", { weekday: "long", day: "numeric", month: "long", year: "numeric" }).format(d),
  (d) => new Intl.DateTimeFormat("tr-TR", { day: "2-digit", month: "2-digit", year: "2-digit" }).format(d),
  (d) => new Intl.DateTimeFormat("tr-TR", { month: "long", year: "numeric" }).format(d),
  (d) => new Intl.DateTimeFormat("tr-TR", { hour: "2-digit", minute: "2-digit" }).format(d),
  (d) => new Intl.DateTimeFormat("tr-TR", { hour: "2-digit", minute: "2-digit", second: "2-digit" }).format(d),
  (d) => new Intl.DateTimeFormat("tr-TR", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" }).format(d),
];

export function HeaderFooterDialog({ open, onClose, actions: a }: { open: boolean; onClose: () => void; actions: Actions }) {
  const [hf, setHf] = useState<HeaderFooter>(a.pres.headerFooter);
  const [lastOpen, setLastOpen] = useState(false);
  const [fmt, setFmt] = useState(0);
  if (open !== lastOpen) {
    setLastOpen(open);
    if (open) setHf(a.pres.headerFooter);
  }
  const now = new Date();
  const apply = (all: boolean) => {
    a.applyHeaderFooter({ ...hf, dateFixed: hf.dateFixed || DATE_FORMATS[fmt](now) }, all);
    onClose();
  };
  return (
    <Dialog
      open={open}
      title="Üst Bilgi ve Alt Bilgi"
      onClose={onClose}
      width={460}
      footer={
        <>
          <DialogButton onClick={() => apply(false)}>Uygula</DialogButton>
          <DialogButton primary onClick={() => apply(true)}>
            Tümüne Uygula
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <DialogTabs tabs={[{ id: "slide", label: "Slayt" }]} active="slide" onChange={() => {}} />
      <fieldset className="oh-fieldset">
        <legend>Slayta ekle</legend>
        <label className="fp-radio" style={{ marginLeft: 0 }}>
          <input type="checkbox" checked={hf.dateTime} onChange={(e) => setHf({ ...hf, dateTime: e.target.checked })} /> Tarih ve saat
        </label>
        <div style={{ marginLeft: 22, opacity: hf.dateTime ? 1 : 0.5 }}>
          <label className="fp-radio" style={{ marginLeft: 0 }}>
            <input type="radio" disabled={!hf.dateTime} checked={hf.dateAuto} onChange={() => setHf({ ...hf, dateAuto: true })} /> Otomatik olarak güncelleştir
          </label>
          <select disabled={!hf.dateTime || !hf.dateAuto} value={fmt} onChange={(e) => setFmt(Number(e.target.value))} style={{ marginLeft: 22, height: 24 }}>
            {DATE_FORMATS.map((f, i) => (
              <option key={i} value={i}>
                {f(now)}
              </option>
            ))}
          </select>
          <label className="fp-radio" style={{ marginLeft: 0 }}>
            <input type="radio" disabled={!hf.dateTime} checked={!hf.dateAuto} onChange={() => setHf({ ...hf, dateAuto: false, dateFixed: hf.dateFixed || DATE_FORMATS[fmt](now) })} /> Sabit
          </label>
          <input type="text" disabled={!hf.dateTime || hf.dateAuto} value={hf.dateFixed} onChange={(e) => setHf({ ...hf, dateFixed: e.target.value })} style={{ marginLeft: 22, height: 24, width: 260 }} />
        </div>
        <label className="fp-radio" style={{ marginLeft: 0 }}>
          <input type="checkbox" checked={hf.slideNumber} onChange={(e) => setHf({ ...hf, slideNumber: e.target.checked })} /> Slayt numarası
        </label>
        <label className="fp-radio" style={{ marginLeft: 0 }}>
          <input type="checkbox" checked={hf.footer} onChange={(e) => setHf({ ...hf, footer: e.target.checked })} /> Alt Bilgi
        </label>
        <input type="text" disabled={!hf.footer} value={hf.footerText} onChange={(e) => setHf({ ...hf, footerText: e.target.value })} style={{ marginLeft: 22, height: 24, width: 300 }} />
      </fieldset>
      <label className="fp-radio" style={{ marginLeft: 0 }}>
        <input type="checkbox" checked={hf.hideOnTitle} onChange={(e) => setHf({ ...hf, hideOnTitle: e.target.checked })} /> Başlık slaytında gösterme
      </label>
    </Dialog>
  );
}

export function DateTimeDialog({ open, onClose, onInsert }: { open: boolean; onClose: () => void; onInsert: (text: string) => void }) {
  const [sel, setSel] = useState(0);
  const now = new Date();
  const ok = () => {
    onInsert(DATE_FORMATS[sel](now));
    onClose();
  };
  return (
    <Dialog open={open} title="Tarih ve Saat" onClose={onClose} onSubmit={ok} width={360} footer={<><DialogButton primary onClick={ok}>Tamam</DialogButton><DialogButton onClick={onClose}>İptal</DialogButton></>}>
      <div className="list-pane" style={{ border: "1px solid var(--oh-border)", padding: 4, maxHeight: 240, overflow: "auto" }}>
        {DATE_FORMATS.map((f, i) => (
          <div key={i} className={`list-row${i === sel ? " is-selected" : ""}`} onClick={() => setSel(i)} onDoubleClick={ok}>
            {f(now)}
          </div>
        ))}
      </div>
      <p style={{ color: "var(--oh-text-muted)" }}>Dil: Türkçe (Türkiye)</p>
    </Dialog>
  );
}

const SYMBOL_SETS: { name: string; from: number; to: number }[] = [
  { name: "Latin-1 Ek", from: 0xa1, to: 0xff },
  { name: "Genişletilmiş Latin-A", from: 0x100, to: 0x17f },
  { name: "Yunanca", from: 0x391, to: 0x3c9 },
  { name: "Genel Noktalama", from: 0x2010, to: 0x2044 },
  { name: "Para Birimi Simgeleri", from: 0x20a0, to: 0x20bf },
  { name: "Harf Benzeri Simgeler", from: 0x2100, to: 0x214f },
  { name: "Oklar", from: 0x2190, to: 0x21ff },
  { name: "Matematiksel İşleçler", from: 0x2200, to: 0x22ff },
  { name: "Geometrik Şekiller", from: 0x25a0, to: 0x25ff },
  { name: "Çeşitli Simgeler", from: 0x2600, to: 0x26ff },
  { name: "Dingbatlar", from: 0x2700, to: 0x27bf },
];

export function SymbolDialog({ open, onClose, onInsert }: { open: boolean; onClose: () => void; onInsert: (ch: string) => void }) {
  const [set, setSet] = useState(0);
  const [sel, setSel] = useState<string>("©");
  const chars = useMemo(() => {
    const s = SYMBOL_SETS[set];
    const out: string[] = [];
    for (let c = s.from; c <= s.to; c++) out.push(String.fromCodePoint(c));
    return out;
  }, [set]);
  return (
    <Dialog
      open={open}
      title="Simge"
      onClose={onClose}
      width={540}
      footer={
        <>
          <DialogButton primary onClick={() => onInsert(sel)}>
            Ekle
          </DialogButton>
          <DialogButton onClick={onClose}>Kapat</DialogButton>
        </>
      }
    >
      <div className="dlg-row">
        <label>
          Alt küme:
          <select value={set} onChange={(e) => setSet(Number(e.target.value))}>
            {SYMBOL_SETS.map((s, i) => (
              <option key={s.name} value={i}>
                {s.name}
              </option>
            ))}
          </select>
        </label>
        <span style={{ marginLeft: "auto" }}>
          Karakter kodu: <strong>{sel.codePointAt(0)!.toString(16).toUpperCase().padStart(4, "0")}</strong> (Unicode onaltılık)
        </span>
      </div>
      <div className="symbol-grid">
        {chars.map((c) => (
          <button key={c} type="button" className={`symbol-cell${c === sel ? " is-selected" : ""}`} onClick={() => setSel(c)} onDoubleClick={() => onInsert(c)}>
            {c}
          </button>
        ))}
      </div>
    </Dialog>
  );
}

const ICON_LIB: Record<string, FluentIcon> = {
  Home: Home24Filled,
  Person: Person24Filled,
  People: People24Filled,
  Mail: Mail24Filled,
  Call: Call24Filled,
  Chat: Chat24Filled,
  Calendar: Calendar24Filled,
  Clock: Clock24Filled,
  Globe: Globe24Filled,
  Location: Location24Filled,
  Map: Map24Filled,
  Airplane: Airplane24Filled,
  VehicleCar: VehicleCar24Filled,
  Building: Building24Filled,
  BuildingFactory: BuildingFactory24Filled,
  Lightbulb: Lightbulb24Filled,
  Rocket: Rocket24Filled,
  Trophy: Trophy24Filled,
  Star: Star24Filled,
  Heart: Heart24Filled,
  ThumbLike: ThumbLike24Filled,
  Checkmark: Checkmark24Filled,
  Dismiss: Dismiss24Filled,
  Warning: Warning24Filled,
  Info: Info24Filled,
  Question: Question24Filled,
  Shield: Shield24Filled,
  LockClosed: LockClosed24Filled,
  Key: Key24Filled,
  Settings: Settings24Filled,
  Wrench: Wrench24Filled,
  DataBarVertical: DataBarVertical24Filled,
  DataPie: DataPie24Filled,
  DataTrending: DataTrending24Filled,
  ChartMultiple: ChartMultiple24Filled,
  Money: Money24Filled,
  Wallet: Wallet24Filled,
  Cart: Cart24Filled,
  Gift: Gift24Filled,
  Briefcase: Briefcase24Filled,
  Document: Document24Filled,
  Folder: Folder24Filled,
  Book: Book24Filled,
  HatGraduation: HatGraduation24Filled,
  Beaker: Beaker24Filled,
  WeatherSunny: WeatherSunny24Filled,
  Cloud: Cloud24Filled,
  Flash: Flash24Filled,
  Fire: Fire24Filled,
  Drop: Drop24Filled,
  Camera: Camera24Filled,
  Video: Video24Filled,
  MusicNote2: MusicNote224Filled,
  Headphones: Headphones24Filled,
  Desktop: Desktop24Filled,
  Phone: Phone24Filled,
  Laptop: Laptop24Filled,
  Code: Code24Filled,
  Bug: Bug24Filled,
  Target: Target24Filled,
  Flag: Flag24Filled,
  Diamond: Diamond24Filled,
};

export function IconsDialog({ open, onClose, actions: a }: { open: boolean; onClose: () => void; actions: Actions }) {
  const [sel, setSel] = useState<Set<string>>(new Set());
  const [q, setQ] = useState("");
  const gridRef = useRef<HTMLDivElement>(null);
  const names = Object.keys(ICON_LIB).filter((n) => n.toLowerCase().includes(q.toLowerCase()));
  const insert = () => {
    const grid = gridRef.current;
    if (!grid) return;
    let i = 0;
    for (const n of sel) {
      const svg = grid.querySelector(`[data-icon="${n}"] svg`);
      if (!svg) continue;
      const vb = (svg.getAttribute("viewBox") ?? "0 0 24 24").split(/\s+/).map(Number);
      const paths: CustomPath[] = Array.from(svg.querySelectorAll("path")).map((p) => ({ w: vb[2] || 24, h: vb[3] || 24, cmds: svgPathToCmds(p.getAttribute("d") ?? "") }));
      if (!paths.length) continue;
      const size = 120;
      const sh = makeSp("custom", (a.pres.width - size) / 2 + i * 30, (a.pres.height - size) / 2 + i * 30, size, size, {
        name: `Simge ${n}`,
        paths,
        fill: { type: "solid", color: { scheme: "dk1" } },
        line: NO_LINE,
      });
      i += 1;
      a.store.commit((p) => {
        const slides = p.slides.slice();
        const s = slides[a.s.slideIdx];
        if (!s) return p;
        slides[a.s.slideIdx] = { ...s, shapes: [...s.shapes, sh] };
        return { ...p, slides };
      }, { sel: [sh.id] });
    }
    setSel(new Set());
    onClose();
  };
  return (
    <Dialog
      open={open}
      title="Simgeler Ekle"
      onClose={onClose}
      width={560}
      footer={
        <>
          <DialogButton primary disabled={!sel.size} onClick={insert}>
            Ekle ({sel.size})
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <input type="text" placeholder="Simgeleri ara (İngilizce ad)" value={q} onChange={(e) => setQ(e.target.value)} style={{ width: "100%", height: 28, marginBottom: 8 }} onKeyDown={(e) => e.key !== "Escape" && e.stopPropagation()} />
      <div ref={gridRef} style={{ display: "grid", gridTemplateColumns: "repeat(10, 46px)", gap: 4, maxHeight: 320, overflowY: "auto" }}>
        {names.map((n) => {
          const C = ICON_LIB[n];
          const on = sel.has(n);
          return (
            <button
              key={n}
              type="button"
              data-icon={n}
              title={n}
              className={`symbol-cell${on ? " is-selected" : ""}`}
              style={{ width: 46, height: 46 }}
              onClick={() => {
                const s = new Set(sel);
                if (on) s.delete(n);
                else s.add(n);
                setSel(s);
              }}
            >
              <C style={{ width: 28, height: 28 }} />
            </button>
          );
        })}
      </div>
    </Dialog>
  );
}
