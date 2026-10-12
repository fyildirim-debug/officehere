// Sayfa Yapısı iletişim kutusu ve Sütunlar iletişim kutusu (bölüm başına; "Uygulama yeri").
import { useState } from "react";
import { Dialog, DialogButton, DialogTabs, Field, FieldSet } from "@officehere/ui";
import { useBelge } from "../state/context";
import { AutoFocus } from "./AutoFocus";
import { Check } from "./FontDialog";
import { PAGE_SIZE_LIST, SECTION_KIND_LABEL, cmToTw, pageSizeId, twToCm, type PageSetup, type SectionAttrs, type SectionKind } from "../model/document";
import { applyPageSetup, sectionInfo, type PageScope } from "../state/pageSetup";
import { applyColumns } from "../ribbon/LayoutTab";

const f = (tw: number) => Number(twToCm(tw).toFixed(2)).toLocaleString("tr-TR");
const p = (s: string) => {
  const n = parseFloat(s.replace(",", ".").replace(/[^\d.-]/g, ""));
  return Number.isFinite(n) ? n : 0;
};

/** "Uygulama yeri" seçimi (tek bölümlü belgede yalnız Tüm belge ve Bu noktadan sonra). */
function ScopeField({ scope, setScope, multi }: { scope: PageScope; setScope: (s: PageScope) => void; multi: boolean }) {
  return (
    <Field label="Uygulama yeri:" inline>
      <select value={scope} onChange={(e) => setScope(e.target.value as PageScope)}>
        {multi && <option value="section">Bu bölüm</option>}
        <option value="forward">Bu noktadan sonra</option>
        <option value="all">Tüm belge</option>
      </select>
    </Field>
  );
}

export function PageSetupDialog({ initialTab, onClose }: { initialTab?: string; onClose: () => void }) {
  const { settings, updateSettings, mainEditor: editor } = useBelge();
  const sec = sectionInfo(editor, settings.page);
  const pg = sec.page;
  const multi = sec.count > 1;
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
  const [titlePage, setTitlePage] = useState(!!pg.titlePage);
  const [oddEven, setOddEven] = useState(settings.oddEven);
  const [kind, setKind] = useState<SectionKind>(((sec.node?.attrs.kind as SectionKind | undefined) ?? "nextPage") as SectionKind);
  const [scope, setScope] = useState<PageScope>(multi ? "section" : "all");

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
    const fn = (cur: PageSetup): PageSetup => ({
      ...cur,
      orientation: orient,
      width: cmToTw(p(w)) || cur.width,
      height: cmToTw(p(h)) || cur.height,
      margin: { top: cmToTw(p(top)), bottom: cmToTw(p(bottom)), left: cmToTw(p(left)), right: cmToTw(p(right)), gutter: cmToTw(p(gutter)), header: cmToTw(p(hd)), footer: cmToTw(p(fd)) },
      titlePage: titlePage || undefined,
    });
    applyPageSetup(editor, settings, updateSettings, fn, scope);
    if (oddEven !== settings.oddEven) updateSettings((s) => ({ ...s, oddEven }));
    // Bölüm başlangıcı (yalnız ilk bölüm dışındakiler için anlamlı)
    const now = sectionInfo(editor, settings.page);
    if (now.node && now.index > 0 && now.node.attrs.kind !== kind) editor.view.dispatch(editor.state.tr.setNodeMarkup(now.pos, undefined, { ...(now.node.attrs as SectionAttrs), kind }));
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
      width={470}
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
      <div className="bl-pagesetup">
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
          <>
            <FieldSet legend="Bölüm">
              <Field label="Bölüm başlangıcı:" inline>
                <select value={kind} disabled={sec.index === 0} onChange={(e) => setKind(e.target.value as SectionKind)}>
                  {(Object.keys(SECTION_KIND_LABEL) as SectionKind[]).map((k) => (
                    <option key={k} value={k}>
                      {k === "nextPage" ? "Yeni sayfa" : k === "continuous" ? "Sürekli" : k === "evenPage" ? "Çift sayfa" : "Tek sayfa"}
                    </option>
                  ))}
                </select>
              </Field>
            </FieldSet>
            <FieldSet legend="Üst bilgiler ve alt bilgiler">
              <div className="bl-dlg-checks is-single">
                <Check label="Tek ve çift sayfalarda farklı" v={oddEven} set={setOddEven} />
                <Check label="İlk sayfada farklı" v={titlePage} set={setTitlePage} />
              </div>
              <div className="bl-dlg-grid">
                {num("Üst bilgi kenardan (cm):", hd, setHd)}
                {num("Alt bilgi kenardan (cm):", fd, setFd)}
              </div>
            </FieldSet>
          </>
        )}
        <ScopeField scope={scope} setScope={setScope} multi={multi} />
      </div>
    </Dialog>
  );
}

export function ColumnsDialog({ onClose }: { onClose: () => void }) {
  const { settings, updateSettings, mainEditor: editor } = useBelge();
  const sec = sectionInfo(editor, settings.page);
  const multi = sec.count > 1;
  const hasSel = !editor.state.selection.empty;
  const [n, setN] = useState(sec.page.columns);
  const [gap, setGap] = useState(f(sec.page.columnSpace));
  const [line, setLine] = useState(!!sec.page.columnLine);
  const [scope, setScope] = useState<PageScope>(hasSel ? "section" : multi ? "section" : "all");
  const apply = () => {
    const fn = (pg: PageSetup): PageSetup => ({ ...pg, columns: Math.max(1, Math.min(13, n)), columnSpace: cmToTw(p(gap)) || 708, columnLine: n > 1 && line ? true : undefined });
    if (hasSel && scope === "section") applyColumns(editor, settings, updateSettings, fn);
    else applyPageSetup(editor, settings, updateSettings, fn, scope);
    onClose();
  };
  return (
    <Dialog
      open
      title="Sütunlar"
      width={400}
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
      <div className="bl-dlg-grid">
        <Field label="Sütun sayısı:">
          <input type="number" min={1} max={13} value={n} onChange={(e) => setN(Math.max(1, Math.min(13, Number(e.target.value) || 1)))} />
        </Field>
        <Field label="Aralık (cm):">
          <input type="text" value={gap} onChange={(e) => setGap(e.target.value)} />
        </Field>
      </div>
      <div className="bl-dlg-checks is-single">
        <Check label="Arasına çizgi" v={line} set={setLine} disabled={n < 2} />
      </div>
      <Field label="Uygulama yeri:" inline>
        <select value={scope} onChange={(e) => setScope(e.target.value as PageScope)}>
          {(multi || hasSel) && <option value="section">{hasSel ? "Seçili metin" : "Bu bölüm"}</option>}
          <option value="forward">Bu noktadan sonra</option>
          <option value="all">Tüm belge</option>
        </select>
      </Field>
    </Dialog>
  );
}
