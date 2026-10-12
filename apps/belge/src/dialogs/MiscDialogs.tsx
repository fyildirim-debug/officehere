// Diğer iletişim kutuları: Yakınlaştır, Sözcük Sayısı, Klavye Kısayolları, Filigran, Sayfa Kenarlıkları, Tablo Özellikleri,
// Hücreleri Böl, Sırala, Alternatif Metin, Resim Boyutu.
import { useMemo, useState } from "react";
import { NodeSelection } from "@tiptap/pm/state";
import { FONT_LIST } from "@officehere/core";
import { Dialog, DialogButton, Field, FieldSet } from "@officehere/ui";
import { useBelge } from "../state/context";
import { docStats, countWords, sortParagraphs } from "../editor/commands";
import type { PageBorder } from "../model/document";
import { tableInfo } from "../ribbon/tableUtil";
import { splitCellGrid, splitDefaults } from "../editor/tableSplit";
import { Check, ColorField } from "./FontDialog";

const okCancel = (ok: () => void, close: () => void) => (
  <>
    <DialogButton primary onClick={ok}>
      Tamam
    </DialogButton>
    <DialogButton onClick={close}>İptal</DialogButton>
  </>
);

export type ZoomTarget = { pct: number } | { fit: "width" | "text" | "page" | "multi" };

export function ZoomDialog({ current, onApply, onClose }: { current: number; onApply: (z: ZoomTarget) => void; onClose: () => void }) {
  const [choice, setChoice] = useState<string>(String(current));
  const [pct, setPct] = useState(current);
  const pick = (v: string) => {
    setChoice(v);
    if (/^\d+$/.test(v)) setPct(Number(v));
  };
  const apply = () => {
    if (choice === "width" || choice === "text" || choice === "page" || choice === "multi") onApply({ fit: choice });
    else onApply({ pct: Math.max(10, Math.min(500, pct)) });
    onClose();
  };
  const radio = (v: string, label: string) => (
    <label className="bl-radio">
      <input type="radio" name="zoom" checked={choice === v} onChange={() => pick(v)} /> {label}
    </label>
  );
  return (
    <Dialog open title="Yakınlaştır" width={460} onClose={onClose} onSubmit={apply} footer={okCancel(apply, onClose)}>
      <FieldSet legend="Yakınlaştırma oranı">
        <div className="bl-dlg-grid">
          <div>
            {radio("200", "%200")}
            {radio("100", "%100")}
            {radio("75", "%75")}
          </div>
          <div>
            {radio("width", "Sayfa genişliği")}
            {radio("text", "Metin genişliği")}
            {radio("page", "Tam sayfa")}
            {radio("multi", "Birden çok sayfa")}
          </div>
        </div>
        <Field label="Yüzde:" inline>
          <input
            type="number"
            min={10}
            max={500}
            value={pct}
            onChange={(e) => {
              setPct(Number(e.target.value));
              setChoice("custom");
            }}
          />
        </Field>
      </FieldSet>
      <div className="bl-zoom-preview">
        <span style={{ fontSize: Math.max(8, Math.min(28, (11 * pct) / 100)) }}>AaBbÇçĞğIıİi</span>
      </div>
    </Dialog>
  );
}

export function WordCountDialog({ pages, onClose }: { pages: number; onClose: () => void }) {
  const { editor } = useBelge();
  const st = useMemo(() => docStats(editor.state.doc), [editor]);
  const { from, to, empty } = editor.state.selection;
  const selWords = empty ? null : countWords(editor.state.doc.textBetween(from, to, "\n", " "));
  const lines = useMemo(() => {
    const dom = editor.view.dom as HTMLElement;
    let n = 0;
    dom.querySelectorAll("p, h1, h2, h3, h4, h5, h6").forEach((el) => {
      const lh = parseFloat(getComputedStyle(el).lineHeight) || 18;
      const text = el.textContent?.trim();
      if (text) n += Math.max(1, Math.round((el as HTMLElement).getBoundingClientRect().height / (lh * (dom.getBoundingClientRect().width / dom.offsetWidth || 1))));
    });
    return n;
  }, [editor]);
  const rows: [string, number][] = [
    ["Sayfa", pages],
    ["Sözcük", selWords ?? st.words],
    ["Karakter (boşluksuz)", st.charsNoSpaces],
    ["Karakter (boşluklu)", st.chars],
    ["Paragraf", st.paragraphs],
    ["Satır", lines],
  ];
  return (
    <Dialog
      open
      title="Sözcük Sayısı"
      width={340}
      onClose={onClose}
      onSubmit={onClose}
      footer={
        <DialogButton primary onClick={onClose}>
          Kapat
        </DialogButton>
      }
    >
      <p className="bl-dlg-note">İstatistikler{selWords !== null ? " (sözcük sayısı seçim için)" : ""}:</p>
      <table className="bl-stats">
        <tbody>
          {rows.map(([k, v]) => (
            <tr key={k}>
              <td>{k}</td>
              <td>{v.toLocaleString("tr-TR")}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </Dialog>
  );
}

export const SHORTCUTS: [string, string][] = [
  ["Ctrl+N", "Yeni belge"],
  ["Ctrl+O", "Aç"],
  ["Ctrl+S", "Kaydet"],
  ["F12", "Farklı Kaydet"],
  ["Ctrl+P", "Yazdır"],
  ["Ctrl+Z / Ctrl+Y", "Geri Al / Yinele"],
  ["Ctrl+X / Ctrl+C / Ctrl+V", "Kes / Kopyala / Yapıştır"],
  ["Ctrl+Shift+C / Ctrl+Shift+V", "Biçimi kopyala / Biçimi yapıştır"],
  ["Ctrl+A", "Tümünü Seç"],
  ["Ctrl+B / Ctrl+I / Ctrl+U", "Kalın / İtalik / Altı Çizili"],
  ["Ctrl+Shift+D", "Çift altı çizili"],
  ["Ctrl+Shift+A", "Tümü büyük harf"],
  ["Ctrl+Shift+K", "Küçük büyük harfler"],
  ["Ctrl+= / Ctrl+Shift+=", "Alt simge / Üst simge"],
  ["Ctrl+Shift+> / Ctrl+Shift+<", "Yazı tipini büyült / küçült"],
  ["Ctrl+] / Ctrl+[", "Yazı tipi boyutunu 1 nk artır / azalt"],
  ["Shift+F3", "Büyük/küçük harf değiştir"],
  ["Ctrl+Space", "Karakter biçimini temizle"],
  ["Ctrl+D", "Yazı Tipi iletişim kutusu"],
  ["Ctrl+L / Ctrl+E / Ctrl+R / Ctrl+J", "Sola hizala / Ortala / Sağa hizala / İki yana yasla"],
  ["Ctrl+1 / Ctrl+2 / Ctrl+5", "Satır aralığı 1 / 2 / 1,5"],
  ["Ctrl+0", "Paragraftan önce 12 nk boşluk ekle/kaldır"],
  ["Ctrl+M / Ctrl+Shift+M", "Girintiyi artır / azalt"],
  ["Ctrl+T", "Asılı girinti"],
  ["Ctrl+Alt+1 / 2 / 3", "Başlık 1 / 2 / 3"],
  ["Ctrl+Shift+N", "Normal stil"],
  ["Ctrl+Shift+L", "Madde işaretli liste"],
  ["Tab / Shift+Tab", "Liste düzeyini artır / azalt"],
  ["Ctrl+Enter", "Sayfa sonu"],
  ["Shift+Enter", "Satır sonu"],
  ["Ctrl+K", "Köprü ekle"],
  ["Ctrl+Alt+M", "Yeni açıklama"],
  ["Ctrl+Shift+E", "Değişiklikleri İzle"],
  ["Ctrl+F", "Bul (Gezinti bölmesi)"],
  ["Ctrl+H", "Değiştir"],
  ["Ctrl+G", "Git"],
  ["Ctrl+Shift+8", "Tümünü Göster (¶)"],
  ["Ctrl+Shift+G", "Sözcük Sayısı"],
  ["Ctrl+fare tekerleği", "Yakınlaştır / Uzaklaştır"],
  ["Alt+Q", "Arama kutusu"],
];

export function ShortcutsDialog({ onClose }: { onClose: () => void }) {
  return (
    <Dialog
      open
      title="Klavye Kısayolları"
      width={560}
      onClose={onClose}
      onSubmit={onClose}
      footer={
        <DialogButton primary onClick={onClose}>
          Kapat
        </DialogButton>
      }
    >
      <div className="bl-shortcuts oh-scroll">
        <table>
          <tbody>
            {SHORTCUTS.map(([k, d]) => (
              <tr key={k}>
                <td>
                  <kbd>{k}</kbd>
                </td>
                <td>{d}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Dialog>
  );
}

export function WatermarkDialog({ onClose }: { onClose: () => void }) {
  const { settings, updateSettings } = useBelge();
  const wm = settings.watermark;
  const [kind, setKind] = useState<"none" | "text">(wm ? "text" : "none");
  const [text, setText] = useState(wm?.text ?? "GİZLİ");
  const [font, setFont] = useState(wm?.font ?? "Aptos");
  const [color, setColor] = useState<string | null>(wm?.color ?? "#C0C0C0");
  const [semi, setSemi] = useState(wm?.semitransparent ?? true);
  const [layout, setLayout] = useState<"diagonal" | "horizontal">(wm?.layout ?? "diagonal");
  const apply = () => {
    updateSettings((s) => ({ ...s, watermark: kind === "none" || !text.trim() ? null : { text: text.trim(), font, color: color ?? "#C0C0C0", semitransparent: semi, layout } }));
    onClose();
  };
  return (
    <Dialog open title="Basılı Filigran" width={420} onClose={onClose} onSubmit={apply} footer={okCancel(apply, onClose)}>
      <label className="bl-radio">
        <input type="radio" checked={kind === "none"} onChange={() => setKind("none")} /> Filigran yok
      </label>
      <label className="bl-radio">
        <input type="radio" checked={kind === "text"} onChange={() => setKind("text")} /> Metin filigranı
      </label>
      <fieldset className="oh-fieldset" disabled={kind !== "text"}>
        <Field label="Metin:" inline>
          <input type="text" list="bl-wm-list" value={text} onChange={(e) => setText(e.target.value)} />
        </Field>
        <datalist id="bl-wm-list">
          {["GİZLİ", "TASLAK", "ÖRNEK", "ACİL", "KOPYALAMAYIN", "ÇOK ACİL", "ÖZEL"].map((t) => (
            <option key={t} value={t} />
          ))}
        </datalist>
        <Field label="Yazı tipi:" inline>
          <select value={font} onChange={(e) => setFont(e.target.value)}>
            {FONT_LIST.map((f) => (
              <option key={f}>{f}</option>
            ))}
          </select>
        </Field>
        <Field label="Renk:" inline>
          <ColorField value={color} onChange={setColor} noneLabel="Otomatik" />
        </Field>
        <Check label="Yarı saydam" v={semi} set={setSemi} />
        <div className="bl-dlg-row">
          <label className="bl-radio">
            <input type="radio" checked={layout === "diagonal"} onChange={() => setLayout("diagonal")} /> Çapraz
          </label>
          <label className="bl-radio">
            <input type="radio" checked={layout === "horizontal"} onChange={() => setLayout("horizontal")} /> Yatay
          </label>
        </div>
      </fieldset>
    </Dialog>
  );
}

export function PageBordersDialog({ onClose }: { onClose: () => void }) {
  const { settings, updateSettings } = useBelge();
  const pb = settings.pageBorder;
  const [setting, setSetting] = useState<"none" | "box">(pb ? "box" : "none");
  const [style, setStyle] = useState<PageBorder["style"]>(pb?.style ?? "single");
  const [color, setColor] = useState<string | null>(pb?.color ?? "#000000");
  const [width, setWidth] = useState(pb?.width ?? 1);
  const apply = () => {
    updateSettings((s) => ({ ...s, pageBorder: setting === "none" ? null : { style, color: color ?? "#000000", width } }));
    onClose();
  };
  return (
    <Dialog open title="Kenarlıklar ve Gölgelendirme — Sayfa Kenarlığı" width={460} onClose={onClose} onSubmit={apply} footer={okCancel(apply, onClose)}>
      <div className="bl-dlg-grid">
        <FieldSet legend="Ayar">
          <label className="bl-radio">
            <input type="radio" checked={setting === "none"} onChange={() => setSetting("none")} /> Yok
          </label>
          <label className="bl-radio">
            <input type="radio" checked={setting === "box"} onChange={() => setSetting("box")} /> Kutu
          </label>
        </FieldSet>
        <div>
          <Field label="Stil:">
            <select value={style} onChange={(e) => setStyle(e.target.value as PageBorder["style"])}>
              <option value="single">Düz çizgi</option>
              <option value="double">Çift çizgi</option>
              <option value="dotted">Noktalı</option>
              <option value="dashed">Kesik çizgili</option>
              <option value="thick">Kalın</option>
            </select>
          </Field>
          <Field label="Renk:">
            <ColorField value={color} onChange={setColor} noneLabel="Otomatik" />
          </Field>
          <Field label="Genişlik:">
            <select value={width} onChange={(e) => setWidth(Number(e.target.value))}>
              {[0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6].map((w) => (
                <option key={w} value={w}>
                  {w.toLocaleString("tr-TR")} nk
                </option>
              ))}
            </select>
          </Field>
        </div>
      </div>
      <div className="bl-pb-preview">
        <span style={{ border: setting === "none" ? "1px dashed #ccc" : `${Math.max(1, width)}px ${style === "thick" || style === "single" ? "solid" : style} ${color ?? "#000"}` }} />
      </div>
    </Dialog>
  );
}

export function SplitCellsDialog({ onClose }: { onClose: () => void }) {
  const { editor, notify } = useBelge();
  const def = useMemo(() => splitDefaults(editor.state), [editor]);
  const [cols, setCols] = useState(def.cols);
  const [rows, setRows] = useState(def.rows);
  const [merge, setMerge] = useState(true);
  const apply = () => {
    if (!splitCellGrid(editor.state, editor.view.dispatch, cols, rows, merge)) {
      notify(`Hücre ${rows} satıra bölünemez: satır sayısı hücrenin kapladığı satır sayısını tam bölmelidir.`, "warning");
      return;
    }
    editor.commands.focus();
    onClose();
  };
  const num = (label: string, v: number, set: (n: number) => void) => (
    <Field label={label}>
      <input type="number" min={1} max={63} value={v} autoFocus={label.startsWith("Sütun")} onChange={(e) => set(Math.max(1, Math.min(63, Number(e.target.value) || 1)))} />
    </Field>
  );
  return (
    <Dialog open title="Hücreleri Böl" width={300} onClose={onClose} onSubmit={apply} footer={okCancel(apply, onClose)}>
      <div className="bl-dlg-grid">
        {num("Sütun sayısı:", cols, setCols)}
        {num("Satır sayısı:", rows, setRows)}
      </div>
      {def.multi && <Check label="Bölmeden önce hücreleri birleştir" v={merge} set={setMerge} />}
    </Dialog>
  );
}

export function TablePropsDialog({ onClose }: { onClose: () => void }) {
  const { editor, openDialog } = useBelge();
  const info = tableInfo(editor);
  const a = info?.node.attrs ?? {};
  const [align, setAlign] = useState<string>((a.align as string | null) ?? "left");
  const [headerRow, setHeaderRow] = useState(!!a.headerRow);
  const [repeat, setRepeat] = useState(!!a.repeatHeader);
  const apply = () => {
    editor
      .chain()
      .focus()
      .updateAttributes("table", { align: align === "left" ? null : align, headerRow, repeatHeader: repeat })
      .run();
    onClose();
  };
  return (
    <Dialog open title="Tablo Özellikleri" width={400} onClose={onClose} onSubmit={apply} footer={okCancel(apply, onClose)}>
      <FieldSet legend="Hizalama">
        <div className="bl-orient">
          {[
            ["left", "Sol"],
            ["center", "Orta"],
            ["right", "Sağ"],
          ].map(([v, l]) => (
            <button key={v} type="button" className={align === v ? "is-selected" : ""} onClick={() => setAlign(v)}>
              {l}
            </button>
          ))}
        </div>
      </FieldSet>
      <FieldSet legend="Satır">
        <Check label="Üst bilgi satırı olarak biçimlendir" v={headerRow} set={setHeaderRow} />
        <Check label="Her sayfanın üst kısmında başlık satırı olarak yinele" v={repeat} set={setRepeat} />
      </FieldSet>
      <DialogButton
        onClick={() => {
          apply();
          openDialog("cellMargins");
        }}
      >
        Seçenekler...
      </DialogButton>
      {info && (
        <p className="bl-dlg-note">
          Boyut: {info.rows} satır × {info.cols} sütun
        </p>
      )}
    </Dialog>
  );
}

export function SortDialog({ onClose }: { onClose: () => void }) {
  const { editor } = useBelge();
  const [type, setType] = useState<"text" | "number">("text");
  const [order, setOrder] = useState<"asc" | "desc">("asc");
  const apply = () => {
    sortParagraphs(editor, order === "desc", type === "number");
    onClose();
  };
  return (
    <Dialog open title="Metni Sırala" width={380} onClose={onClose} onSubmit={apply} footer={okCancel(apply, onClose)}>
      <FieldSet legend="Önce sıralama ölçütü: Paragraflar">
        <Field label="Tür:" inline>
          <select value={type} onChange={(e) => setType(e.target.value as "text" | "number")}>
            <option value="text">Metin</option>
            <option value="number">Sayı</option>
          </select>
        </Field>
        <label className="bl-radio">
          <input type="radio" checked={order === "asc"} onChange={() => setOrder("asc")} /> Artan
        </label>
        <label className="bl-radio">
          <input type="radio" checked={order === "desc"} onChange={() => setOrder("desc")} /> Azalan
        </label>
      </FieldSet>
    </Dialog>
  );
}

export function AltTextDialog({ onClose }: { onClose: () => void }) {
  const { editor } = useBelge();
  const sel = editor.state.selection;
  const node = sel instanceof NodeSelection ? sel.node : null;
  const [alt, setAlt] = useState((node?.attrs.alt as string | null) ?? "");
  const apply = () => {
    if (node) editor.view.dispatch(editor.state.tr.setNodeMarkup(sel.from, undefined, { ...node.attrs, alt: alt || null }));
    onClose();
  };
  return (
    <Dialog open title="Alternatif Metin" width={400} onClose={onClose} onSubmit={apply} footer={okCancel(apply, onClose)}>
      <p className="bl-dlg-note">Görme engelli kişilere bu nesneyi ve bağlamını nasıl tanımlarsınız?</p>
      <textarea className="bl-fill-textarea" rows={4} value={alt} onChange={(e) => setAlt(e.target.value)} autoFocus />
    </Dialog>
  );
}

export function ImageSizeDialog({ onClose }: { onClose: () => void }) {
  const { editor } = useBelge();
  const sel = editor.state.selection;
  const node = sel instanceof NodeSelection ? sel.node : null;
  const CM = 96 / 2.54;
  const w0 = (node?.attrs.width as number | null) ?? 200;
  const h0 = (node?.attrs.height as number | null) ?? 150;
  const [w, setW] = useState((w0 / CM).toFixed(2).replace(".", ","));
  const [h, setH] = useState((h0 / CM).toFixed(2).replace(".", ","));
  const [lock, setLock] = useState(true);
  const num = (s: string) => parseFloat(s.replace(",", ".")) || 0;
  const apply = () => {
    if (node) editor.view.dispatch(editor.state.tr.setNodeMarkup(sel.from, undefined, { ...node.attrs, width: Math.round(num(w) * CM), height: Math.round(num(h) * CM) }));
    onClose();
  };
  return (
    <Dialog open title="Düzen — Boyut" width={360} onClose={onClose} onSubmit={apply} footer={okCancel(apply, onClose)}>
      <Field label="Yükseklik (cm):" inline>
        <input
          type="text"
          value={h}
          onChange={(e) => {
            setH(e.target.value);
            if (lock) setW(((num(e.target.value) * w0) / h0).toFixed(2).replace(".", ","));
          }}
        />
      </Field>
      <Field label="Genişlik (cm):" inline>
        <input
          type="text"
          value={w}
          onChange={(e) => {
            setW(e.target.value);
            if (lock) setH(((num(e.target.value) * h0) / w0).toFixed(2).replace(".", ","));
          }}
        />
      </Field>
      <Check label="En boy oranını kilitle" v={lock} set={setLock} />
    </Dialog>
  );
}
