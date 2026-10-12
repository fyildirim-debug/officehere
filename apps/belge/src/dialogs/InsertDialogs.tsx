// Ekleme iletişim kutuları: Bağlantı, Tablo Ekle, Simge, Tarih ve Saat, Yer İşareti.
import { useMemo, useState } from "react";
import { TextSelection } from "@tiptap/pm/state";
import { Dialog, DialogButton, Field, FieldSet } from "@officehere/ui";
import { useBelge } from "../state/context";
import { DATE_FORMATS, collectHeadings, insertText, wordAt } from "../editor/commands";
import { insertTable } from "../ribbon/tableUtil";
import { pageMetrics } from "../model/document";

const okCancel = (ok: () => void, close: () => void, okLabel = "Tamam") => (
  <>
    <DialogButton primary onClick={ok}>
      {okLabel}
    </DialogButton>
    <DialogButton onClick={close}>İptal</DialogButton>
  </>
);

export function LinkDialog({ onClose }: { onClose: () => void }) {
  const { editor } = useBelge();
  const { from, to, empty } = editor.state.selection;
  const existing = editor.getAttributes("link").href as string | undefined;
  const [kind, setKind] = useState<"web" | "place" | "mail">(existing?.startsWith("mailto:") ? "mail" : existing?.startsWith("#") ? "place" : "web");
  const [text, setText] = useState(() => {
    if (!empty) return editor.state.doc.textBetween(from, to, " ");
    if (existing) {
      const w = wordAt(editor.state.doc, from);
      return w ? editor.state.doc.textBetween(w.from, w.to) : "";
    }
    return "";
  });
  const [href, setHref] = useState(existing?.replace(/^mailto:/, "") ?? "");
  const headings = useMemo(() => collectHeadings(editor.state.doc), [editor]);
  const bookmarks = useMemo(() => {
    const out: string[] = [];
    editor.state.doc.descendants((n) => {
      if (n.type.name === "bookmark") out.push(n.attrs.name as string);
      return true;
    });
    return out;
  }, [editor]);

  const apply = () => {
    let url = href.trim();
    if (!url) return onClose();
    if (kind === "mail" && !url.startsWith("mailto:")) url = `mailto:${url}`;
    if (kind === "web" && !/^[a-z]+:/i.test(url) && !url.startsWith("#")) url = `https://${url}`;
    const { state } = editor;
    const tr = state.tr;
    // Başlığa bağlantı: başlığın başına gizli bir yer işareti eklenir (OOXML w:anchor).
    const hm = /^#h-(\d+)$/.exec(url);
    if (hm) {
      const h = headings[Number(hm[1])];
      if (h) {
        const name = `_Ref${Date.now().toString(36)}`;
        tr.insert(h.pos + 1, state.schema.nodes.bookmark.create({ name }));
        url = `#${name}`;
      }
    }
    const linkMark = state.schema.marks.link.create({ href: url });
    if (tr.selection.empty && !existing) {
      const label = text || url.replace(/^mailto:/, "");
      tr.insert(tr.selection.from, state.schema.text(label, [linkMark]));
    } else {
      let { from: a, to: b } = tr.selection;
      if (a === b) {
        const w = wordAt(tr.doc, a);
        if (w) ({ from: a, to: b } = w);
      }
      const cur = tr.doc.textBetween(a, b, " ");
      if (text && text !== cur) tr.replaceWith(a, b, state.schema.text(text, [linkMark]));
      else tr.addMark(a, b, linkMark);
    }
    editor.view.dispatch(tr);
    editor.commands.focus();
    onClose();
  };

  return (
    <Dialog
      open
      title={existing ? "Köprüyü Düzenle" : "Köprü Ekle"}
      width={520}
      onClose={onClose}
      onSubmit={apply}
      footer={
        <>
          {existing && (
            <DialogButton
              onClick={() => {
                editor.chain().focus().extendMarkRange("link").unsetLink().run();
                onClose();
              }}
            >
              Bağlantıyı Kaldır
            </DialogButton>
          )}
          <span className="bl-flex" />
          {okCancel(apply, onClose)}
        </>
      }
    >
      <div className="bl-link-dlg">
        <div className="bl-link-kinds">
          <button type="button" className={kind === "web" ? "is-selected" : ""} onClick={() => setKind("web")}>
            Var Olan Dosya veya Web Sayfası
          </button>
          <button type="button" className={kind === "place" ? "is-selected" : ""} onClick={() => setKind("place")}>
            Bu Belgeye Yerleştir
          </button>
          <button type="button" className={kind === "mail" ? "is-selected" : ""} onClick={() => setKind("mail")}>
            E-posta Adresi
          </button>
        </div>
        <div className="bl-link-body">
          <Field label="Görüntülenecek metin:">
            <input type="text" value={text} onChange={(e) => setText(e.target.value)} />
          </Field>
          {kind === "place" ? (
            <div className="bl-listfield">
              <span>Bu belgede bir yer seçin:</span>
              <select size={8} value={href} onChange={(e) => setHref(e.target.value)}>
                <option value="#_top">Belgenin Başı</option>
                {headings.length > 0 && (
                  <optgroup label="Başlıklar">
                    {headings.map((h, i) => (
                      <option key={`${h.pos}-${i}`} value={`#h-${i}`}>
                        {" ".repeat((h.level - 1) * 3)}
                        {h.text}
                      </option>
                    ))}
                  </optgroup>
                )}
                {bookmarks.length > 0 && (
                  <optgroup label="Yer İşaretleri">
                    {bookmarks.map((b) => (
                      <option key={b} value={`#${b}`}>
                        {b}
                      </option>
                    ))}
                  </optgroup>
                )}
              </select>
            </div>
          ) : (
            <Field label={kind === "mail" ? "E-posta adresi:" : "Adres:"}>
              <input type="text" value={href} onChange={(e) => setHref(e.target.value)} autoFocus={!!text} placeholder={kind === "mail" ? "ad@ornek.com" : "https://"} />
            </Field>
          )}
        </div>
      </div>
    </Dialog>
  );
}

export function InsertTableDialog({ onClose }: { onClose: () => void }) {
  const { editor, settings } = useBelge();
  const [cols, setCols] = useState(5);
  const [rows, setRows] = useState(2);
  const [fit, setFit] = useState<"fixed" | "content" | "window">("fixed");
  const [colW, setColW] = useState("Otomatik");
  const apply = () => {
    const m = pageMetrics(settings.page);
    const c = Math.max(1, Math.min(63, cols));
    const r = Math.max(1, Math.min(500, rows));
    const cm = parseFloat(colW.replace(",", "."));
    const w = fit === "fixed" ? (Number.isFinite(cm) ? (cm / 2.54) * 96 : m.contentW / c) : fit === "window" ? m.contentW / c : undefined;
    insertTable(editor, r, c, w);
    onClose();
  };
  return (
    <Dialog open title="Tablo Ekle" width={360} onClose={onClose} onSubmit={apply} footer={okCancel(apply, onClose)}>
      <FieldSet legend="Tablo boyutu">
        <Field label="Sütun sayısı:" inline>
          <input type="number" min={1} max={63} value={cols} onChange={(e) => setCols(Number(e.target.value))} />
        </Field>
        <Field label="Satır sayısı:" inline>
          <input type="number" min={1} max={500} value={rows} onChange={(e) => setRows(Number(e.target.value))} />
        </Field>
      </FieldSet>
      <FieldSet legend="Otomatik Sığdır davranışı">
        <label className="bl-radio">
          <input type="radio" checked={fit === "fixed"} onChange={() => setFit("fixed")} /> Sabit sütun genişliği:
          <input type="text" value={colW} disabled={fit !== "fixed"} onChange={(e) => setColW(e.target.value)} style={{ width: 90, marginLeft: 8 }} />
        </label>
        <label className="bl-radio">
          <input type="radio" checked={fit === "content"} onChange={() => setFit("content")} /> İçeriğe otomatik sığdır
        </label>
        <label className="bl-radio">
          <input type="radio" checked={fit === "window"} onChange={() => setFit("window")} /> Pencereye otomatik sığdır
        </label>
      </FieldSet>
    </Dialog>
  );
}

const SYMBOL_SETS: { id: string; label: string; chars: string }[] = [
  { id: "latin", label: "Latin-1 Ek", chars: "¡¢£¤¥¦§¨©ª«¬®¯°±²³´µ¶·¸¹º»¼½¾¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõö÷øùúûüýþÿ" },
  { id: "tr", label: "Türkçe ve Latin Genişletilmiş", chars: "ĞğİıŞşÇçÖöÜüÂâÎîÛûĀāĂăĄąĆćČčĎďĒēĘęĚěĹĺĽľŁłŃńŇňŐőŒœŔŕŘřŚśŠšŢţŤťŰűŸŹźŻżŽž" },
  { id: "greek", label: "Yunanca", chars: "ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩαβγδεζηθικλμνξοπρστυφχψω" },
  { id: "punct", label: "Genel Noktalama", chars: "–—‘’‚“”„†‡•…‰′″‹›‼⁄€₺₹₽¢£¥©®™℃℉№℗℠" },
  { id: "math", label: "Matematik İşleçleri", chars: "∀∂∃∅∆∇∈∉∋∏∑−∓∕∗∘√∛∜∝∞∟∠∧∨∩∪∫∬∮∴∵≈≠≡≤≥⊂⊃⊆⊇⊕⊗⊥⋅±×÷¼½¾" },
  { id: "arrows", label: "Oklar", chars: "←↑→↓↔↕↖↗↘↙⇐⇑⇒⇓⇔⇕➔➘➙➚➛➜➝➞➟➠➡➢➣➤" },
  { id: "shapes", label: "Geometrik Şekiller ve Simgeler", chars: "■□▢▣▤▥▦▧▨▩▪▫▬▭▮▯▰▱▲△▴▵▶▷▸▹►▻▼▽◀◁◆◇◈◉◊○◌◍◎●◐◑◒◓★☆☎☑☒☐☺☻♠♣♥♦♪♫✓✔✕✖✗✘❖" },
];

export function SymbolDialog({ onClose }: { onClose: () => void }) {
  const { editor } = useBelge();
  const [set, setSet] = useState("tr");
  const [sel, setSel] = useState<string | null>(null);
  const chars = Array.from(SYMBOL_SETS.find((s) => s.id === set)?.chars ?? "");
  const insert = () => {
    if (sel) insertText(editor, sel);
  };
  return (
    <Dialog
      open
      title="Simge"
      width={520}
      onClose={onClose}
      onSubmit={insert}
      footer={
        <>
          <DialogButton primary onClick={insert} disabled={!sel}>
            Ekle
          </DialogButton>
          <DialogButton onClick={onClose}>Kapat</DialogButton>
        </>
      }
    >
      <Field label="Alt küme:" inline>
        <select value={set} onChange={(e) => setSet(e.target.value)}>
          {SYMBOL_SETS.map((s) => (
            <option key={s.id} value={s.id}>
              {s.label}
            </option>
          ))}
        </select>
      </Field>
      <div className="bl-symbols">
        {chars.map((c) => (
          <button key={c} type="button" className={sel === c ? "is-selected" : ""} onClick={() => setSel(c)} onDoubleClick={() => insertText(editor, c)}>
            {c}
          </button>
        ))}
      </div>
      <div className="bl-dlg-msg">{sel ? `Karakter kodu: ${sel.codePointAt(0)!.toString(16).toUpperCase().padStart(4, "0")} (Unicode onaltılık)` : "Bir simge seçin."}</div>
    </Dialog>
  );
}

export function DateTimeDialog({ onClose }: { onClose: () => void }) {
  const { editor } = useBelge();
  const now = useMemo(() => new Date(), []);
  const [fmt, setFmt] = useState(DATE_FORMATS[0].id);
  const apply = () => {
    const f = DATE_FORMATS.find((d) => d.id === fmt) ?? DATE_FORMATS[0];
    insertText(editor, f.fmt(now));
    onClose();
  };
  return (
    <Dialog open title="Tarih ve Saat" width={360} onClose={onClose} onSubmit={apply} footer={okCancel(apply, onClose)}>
      <div className="bl-listfield">
        <span>Kullanılabilir biçimler:</span>
        <select size={9} value={fmt} onChange={(e) => setFmt(e.target.value)} onDoubleClick={apply}>
          {DATE_FORMATS.map((d) => (
            <option key={d.id} value={d.id}>
              {d.fmt(now)}
            </option>
          ))}
        </select>
      </div>
      <Field label="Dil:" inline>
        <select disabled value="tr">
          <option value="tr">Türkçe (Türkiye)</option>
        </select>
      </Field>
    </Dialog>
  );
}

export function BookmarkDialog({ onClose }: { onClose: () => void }) {
  const { editor } = useBelge();
  const list = useMemo(() => {
    const out: { name: string; pos: number }[] = [];
    editor.state.doc.descendants((n, pos) => {
      if (n.type.name === "bookmark") out.push({ name: n.attrs.name as string, pos });
      return true;
    });
    return out;
  }, [editor]);
  const [name, setName] = useState("");
  const valid = /^[\p{L}_][\p{L}\p{N}_]*$/u.test(name);
  const exists = list.find((b) => b.name === name);
  const add = () => {
    if (!valid) return;
    if (exists) editor.view.dispatch(editor.state.tr.delete(exists.pos, exists.pos + 1));
    editor.chain().focus().insertBookmark(name).run();
    onClose();
  };
  return (
    <Dialog
      open
      title="Yer İşareti"
      width={360}
      onClose={onClose}
      onSubmit={add}
      footer={
        <>
          <DialogButton primary onClick={add} disabled={!valid}>
            Ekle
          </DialogButton>
          <DialogButton
            disabled={!exists}
            onClick={() => {
              if (exists) editor.view.dispatch(editor.state.tr.delete(exists.pos, exists.pos + 1));
              onClose();
            }}
          >
            Sil
          </DialogButton>
          <DialogButton
            disabled={!exists}
            onClick={() => {
              if (exists) editor.view.dispatch(editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(exists.pos))).scrollIntoView());
              editor.commands.focus();
              onClose();
            }}
          >
            Git
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <Field label="Yer işareti adı:">
        <input type="text" value={name} onChange={(e) => setName(e.target.value.replace(/\s/g, "_"))} />
      </Field>
      <select size={6} className="bl-fill-select" value={name} onChange={(e) => setName(e.target.value)}>
        {list.map((b) => (
          <option key={b.name} value={b.name}>
            {b.name}
          </option>
        ))}
      </select>
    </Dialog>
  );
}
