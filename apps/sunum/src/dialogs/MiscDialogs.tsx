/** Slayt Boyutu, Bul/Değiştir, Yazı Tiplerini Değiştir, Yakınlaştır, Boyut ve Konum, Gösteri Ayarları, Kısayollar, Hakkında. */
import { useState } from "react";
import { DeveloperCredit, Dialog, DialogButton, Field } from "@officehere/ui";
import { FONT_LIST } from "@officehere/core";
import { resizePresentation } from "../model/presentation";
import type { Actions } from "../state/actions";
import { findAll, replaceAll, replaceFonts, usedFonts, type Match } from "../state/find";
import { cm, fromCm } from "../ribbon/common";

const SIZES: { label: string; w: number; h: number }[] = [
  { label: "Ekran Gösterisi (4:3)", w: 25.4, h: 19.05 },
  { label: "Geniş ekran", w: 33.867, h: 19.05 },
  { label: "Ekran Gösterisi (16:9)", w: 25.4, h: 14.288 },
  { label: "Ekran Gösterisi (16:10)", w: 25.4, h: 15.875 },
  { label: "A4 Kağıt (210x297 mm)", w: 27.517, h: 19.05 },
  { label: "Letter Kağıdı (8,5x11 inç)", w: 25.4, h: 19.05 },
  { label: "35 mm Slaytlar", w: 28.575, h: 19.05 },
  { label: "Başlık", w: 33.867, h: 4.233 },
];

export function SlideSizeDialog({ open, onClose, actions: a }: { open: boolean; onClose: () => void; actions: Actions }) {
  const [w, setW] = useState(cm(a.pres.width));
  const [h, setH] = useState(cm(a.pres.height));
  const [lastOpen, setLastOpen] = useState(false);
  if (open !== lastOpen) {
    setLastOpen(open);
    if (open) {
      setW(cm(a.pres.width));
      setH(cm(a.pres.height));
    }
  }
  const apply = (mode: "scale" | "fit") => {
    const W = Math.max(2.54, Math.min(142.24, w));
    const H = Math.max(2.54, Math.min(142.24, h));
    a.store.commit((p) => resizePresentation(p, fromCm(W), fromCm(H), mode));
    onClose();
  };
  const preset = SIZES.find((s) => Math.abs(s.w - w) < 0.02 && Math.abs(s.h - h) < 0.02);
  return (
    <Dialog
      open={open}
      title="Slayt Boyutu"
      onClose={onClose}
      width={420}
      footer={
        <>
          <DialogButton primary onClick={() => apply("scale")}>
            En Büyük Boyuta Getir
          </DialogButton>
          <DialogButton onClick={() => apply("fit")}>Sığdırmayı Sağla</DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <Field label="Slayt boyutu:">
        <select
          value={preset?.label ?? "custom"}
          onChange={(e) => {
            const s = SIZES.find((x) => x.label === e.target.value);
            if (s) {
              setW(s.w);
              setH(s.h);
            }
          }}
        >
          {SIZES.map((s) => (
            <option key={s.label} value={s.label}>
              {s.label}
            </option>
          ))}
          <option value="custom">Özel</option>
        </select>
      </Field>
      <div className="dlg-row">
        <label>
          Genişlik:
          <input type="number" step={0.01} value={w} onChange={(e) => setW(Number(e.target.value))} style={{ width: 90 }} /> cm
        </label>
        <label>
          Yükseklik:
          <input type="number" step={0.01} value={h} onChange={(e) => setH(Number(e.target.value))} style={{ width: 90 }} /> cm
        </label>
      </div>
      <p style={{ color: "var(--oh-text-muted)" }}>Yeni slayt boyutuna ölçeklendiriyorsunuz. İçeriğin boyutunu en üst düzeye çıkarmak mı yoksa yeni slayda sığdırılmasını sağlamak mı istersiniz?</p>
    </Dialog>
  );
}

export function FindDialog({ open, onClose, actions: a, replaceMode, initialQuery }: { open: boolean; onClose: () => void; actions: Actions; replaceMode: boolean; initialQuery?: string }) {
  const [q, setQ] = useState("");
  const [wasOpen, setWasOpen] = useState(false);
  // Komut aramasından "Bul" ile açılınca aranan metin hazır gelir.
  if (open !== wasOpen) {
    setWasOpen(open);
    if (open && initialQuery) setQ(initialQuery);
  }
  const [r, setR] = useState("");
  const [matchCase, setMatchCase] = useState(false);
  const [whole, setWhole] = useState(false);
  const [cursor, setCursor] = useState(-1);
  const [msg, setMsg] = useState("");
  const opts = { matchCase, wholeWord: whole };
  const go = (m: Match) => {
    a.goToSlide(m.slideIdx);
    a.startEditing({ shapeId: m.shapeId, cell: m.cell, range: [m.from, m.to] });
  };
  const findNext = () => {
    const all = findAll(a.pres, q, opts);
    if (!all.length) {
      setMsg("Aranan öğe bulunamadı.");
      return;
    }
    const s = a.s;
    let next = (cursor + 1) % all.length;
    if (cursor < 0) {
      const after = all.findIndex((m) => m.slideIdx >= s.slideIdx);
      next = after >= 0 ? after : 0;
    }
    setCursor(next);
    setMsg(`${all.length} sonuçtan ${next + 1}.`);
    go(all[next]);
  };
  const replaceOne = () => {
    const h = a.textHandle;
    if (h && a.s.editing && h.selectedText().toLocaleLowerCase("tr-TR") === q.toLocaleLowerCase("tr-TR") && q) {
      h.insert(r);
      setCursor((c) => c - 1);
    }
    findNext();
  };
  const doAll = () => {
    a.stopEditing();
    let count = 0;
    a.store.commit((p) => {
      const res = replaceAll(p, q, r, opts);
      count = res.count;
      return res.pres;
    });
    setMsg(count ? `${count} değiştirme yapıldı.` : "Aranan öğe bulunamadı.");
  };
  return (
    <Dialog
      open={open}
      title={replaceMode ? "Değiştir" : "Bul"}
      onClose={onClose}
      onSubmit={findNext}
      width={520}
      footer={
        <>
          <DialogButton primary disabled={!q} onClick={findNext}>
            Sonrakini Bul
          </DialogButton>
          {replaceMode && (
            <>
              <DialogButton disabled={!q} onClick={replaceOne}>
                Değiştir
              </DialogButton>
              <DialogButton disabled={!q} onClick={doAll}>
                Tümünü Değiştir
              </DialogButton>
            </>
          )}
          <DialogButton onClick={onClose}>Kapat</DialogButton>
        </>
      }
    >
      <div className="find-bar" onKeyDown={(e) => e.key !== "Escape" && e.stopPropagation()}>
        <Field label="Aranan:">
          <input
            type="text"
            value={q}
            onChange={(e) => {
              setQ(e.target.value);
              setCursor(-1);
              setMsg("");
            }}
          />
        </Field>
        {replaceMode && (
          <Field label="Yeni değer:">
            <input type="text" value={r} onChange={(e) => setR(e.target.value)} />
          </Field>
        )}
        <label>
          <input type="checkbox" checked={matchCase} onChange={(e) => setMatchCase(e.target.checked)} /> Büyük/küçük harf eşleştir
        </label>
        <label>
          <input type="checkbox" checked={whole} onChange={(e) => setWhole(e.target.checked)} /> Yalnızca sözcüğün tamamını bul
        </label>
        {msg && <span style={{ color: "var(--oh-text-muted)" }}>{msg}</span>}
      </div>
    </Dialog>
  );
}

export function ReplaceFontsDialog({ open, onClose, actions: a }: { open: boolean; onClose: () => void; actions: Actions }) {
  const fonts = usedFonts(a.pres);
  const [from, setFrom] = useState(fonts[0] ?? "");
  const [to, setTo] = useState("Arial");
  const ok = () => {
    if (from && to && from !== to) a.store.commit((p) => replaceFonts(p, from, to));
    onClose();
  };
  return (
    <Dialog open={open} title="Yazı Tiplerini Değiştir" onClose={onClose} onSubmit={ok} width={360} footer={<><DialogButton primary onClick={ok}>Değiştir</DialogButton><DialogButton onClick={onClose}>Kapat</DialogButton></>}>
      <Field label="Değiştir:">
        <select value={from} onChange={(e) => setFrom(e.target.value)}>
          {fonts.map((f) => (
            <option key={f}>{f}</option>
          ))}
        </select>
      </Field>
      <Field label="Yeni:">
        <select value={to} onChange={(e) => setTo(e.target.value)}>
          {FONT_LIST.map((f) => (
            <option key={f}>{f}</option>
          ))}
        </select>
      </Field>
    </Dialog>
  );
}

export function ZoomDialog({ open, onClose, zoom, onZoom }: { open: boolean; onClose: () => void; zoom: number; onZoom: (z: number, fit?: boolean) => void }) {
  const [v, setV] = useState(Math.round(zoom));
  const [fit, setFit] = useState(false);
  const ok = () => {
    onZoom(v, fit);
    onClose();
  };
  return (
    <Dialog open={open} title="Yakınlaştır" onClose={onClose} onSubmit={ok} width={300} footer={<><DialogButton primary onClick={ok}>Tamam</DialogButton><DialogButton onClick={onClose}>İptal</DialogButton></>}>
      <fieldset className="oh-fieldset">
        <legend>Yakınlaştır</legend>
        <label className="fp-radio" style={{ marginLeft: 0 }}>
          <input type="radio" checked={fit} onChange={() => setFit(true)} /> Sığdır
        </label>
        {[400, 300, 200, 100, 66, 50, 33].map((z) => (
          <label key={z} className="fp-radio" style={{ marginLeft: 0 }}>
            <input
              type="radio"
              checked={!fit && v === z}
              onChange={() => {
                setFit(false);
                setV(z);
              }}
            />
            %{z}
          </label>
        ))}
        <div className="dlg-row">
          Yüzde:
          <input type="number" min={10} max={400} value={v} onChange={(e) => (setFit(false), setV(Number(e.target.value)))} style={{ width: 70 }} /> %
        </div>
      </fieldset>
    </Dialog>
  );
}

export function SizePosDialog({ open, onClose, actions: a }: { open: boolean; onClose: () => void; actions: Actions }) {
  const s = a.selectedShapes()[0];
  const [v, setV] = useState({ w: 0, h: 0, rot: 0, x: 0, y: 0, lock: false });
  const [lastOpen, setLastOpen] = useState(false);
  if (open !== lastOpen) {
    setLastOpen(open);
    if (open && s) setV({ w: cm(s.w), h: cm(s.h), rot: s.rot, x: cm(s.x), y: cm(s.y), lock: s.kind === "pic" });
  }
  const ok = () => {
    if (s) {
      a.setGeometry(s.id, { x: fromCm(v.x), y: fromCm(v.y), w: fromCm(v.w), h: fromCm(v.h), rot: ((v.rot % 360) + 360) % 360 });
    }
    onClose();
  };
  const setW = (w: number) => setV((o) => ({ ...o, w, h: o.lock && o.w ? (o.h * w) / o.w : o.h }));
  const setH = (h: number) => setV((o) => ({ ...o, h, w: o.lock && o.h ? (o.w * h) / o.h : o.w }));
  return (
    <Dialog open={open} title="Boyut ve Konum" onClose={onClose} onSubmit={ok} width={380} footer={<><DialogButton primary onClick={ok}>Tamam</DialogButton><DialogButton onClick={onClose}>İptal</DialogButton></>}>
      {!s ? (
        <p>Önce bir nesne seçin.</p>
      ) : (
        <>
          <fieldset className="oh-fieldset">
            <legend>Boyut</legend>
            <div className="dlg-row">
              <label>
                Yükseklik: <input type="number" step={0.01} value={Number(v.h.toFixed(2))} onChange={(e) => setH(Number(e.target.value))} style={{ width: 80 }} /> cm
              </label>
              <label>
                Genişlik: <input type="number" step={0.01} value={Number(v.w.toFixed(2))} onChange={(e) => setW(Number(e.target.value))} style={{ width: 80 }} /> cm
              </label>
            </div>
            <div className="dlg-row">
              <label>
                Döndürme: <input type="number" step={1} value={v.rot} onChange={(e) => setV({ ...v, rot: Number(e.target.value) })} style={{ width: 70 }} /> °
              </label>
              <label>
                <input type="checkbox" checked={v.lock} onChange={(e) => setV({ ...v, lock: e.target.checked })} /> En boy oranını kilitle
              </label>
            </div>
          </fieldset>
          <fieldset className="oh-fieldset">
            <legend>Konum (sol üst köşeden)</legend>
            <div className="dlg-row">
              <label>
                Yatay: <input type="number" step={0.01} value={Number(v.x.toFixed(2))} onChange={(e) => setV({ ...v, x: Number(e.target.value) })} style={{ width: 80 }} /> cm
              </label>
              <label>
                Dikey: <input type="number" step={0.01} value={Number(v.y.toFixed(2))} onChange={(e) => setV({ ...v, y: Number(e.target.value) })} style={{ width: 80 }} /> cm
              </label>
            </div>
          </fieldset>
        </>
      )}
    </Dialog>
  );
}

const SHORTCUTS: [string, string][] = [
  ["Ctrl+N / Ctrl+O / Ctrl+S", "Yeni / Aç / Kaydet"],
  ["F12", "Farklı Kaydet"],
  ["Ctrl+P", "Yazdır"],
  ["Ctrl+M", "Yeni slayt"],
  ["Ctrl+D", "Seçili nesneyi veya slaydı çoğalt"],
  ["Delete", "Seçili nesneyi veya slaydı sil"],
  ["Ctrl+Z / Ctrl+Y", "Geri al / Yinele"],
  ["Ctrl+C / Ctrl+X / Ctrl+V", "Kopyala / Kes / Yapıştır"],
  ["Ctrl+A", "Tümünü seç"],
  ["Ctrl+G / Ctrl+Shift+G", "Grupla / Grubu çöz"],
  ["Tab / Shift+Tab", "Sonraki / önceki nesneyi seç; metinde liste düzeyini artır/azalt"],
  ["Ok tuşları (Ctrl ile ince)", "Seçili nesneyi kaydır"],
  ["Alt+Sol/Sağ ok", "Nesneyi 15° döndür"],
  ["Shift + sürükle", "Oranı koruyarak boyutlandır / yatay-dikey taşı"],
  ["Ctrl + sürükle", "Kopyasını taşı"],
  ["Alt + sürükle", "Kılavuzlara yapışmadan taşı"],
  ["F2 / Enter", "Seçili şeklin metnini düzenle"],
  ["Esc", "Metin düzenlemeden çık / seçimi kaldır"],
  ["Ctrl+B / Ctrl+I / Ctrl+U", "Kalın / İtalik / Altı çizili"],
  ["Ctrl+] / Ctrl+[", "Yazı tipi boyutunu büyüt / küçült"],
  ["Ctrl+L / Ctrl+E / Ctrl+R / Ctrl+J", "Sola / Ortaya / Sağa hizala / İki yana yasla"],
  ["Ctrl+Boşluk", "Karakter biçimlendirmesini temizle"],
  ["Ctrl+Shift+C / Ctrl+Shift+V", "Biçimi kopyala / yapıştır"],
  ["Ctrl+F / Ctrl+H", "Bul / Değiştir"],
  ["Alt+Q", "Komut ara (şerit komutlarını bul ve çalıştır)"],
  ["F7", "Yazım Denetimi bölmesi"],
  ["Shift + sağ tık (metinde)", "Tarayıcının yazım önerileri menüsü"],
  ["Ctrl+fare tekerleği", "Yakınlaştır / uzaklaştır"],
  ["F5 / Shift+F5", "Gösteriyi baştan / geçerli slayttan başlat"],
  ["Gösteride: N, →, ↓, Boşluk, Enter", "Sonraki animasyon veya slayt"],
  ["Gösteride: P, ←, ↑, Geri", "Önceki animasyon veya slayt"],
  ["Gösteride: sayı + Enter", "O slayta git"],
  ["Gösteride: B / W", "Siyah / beyaz ekran"],
  ["Gösteride: G ya da -", "Tüm slaytları gör"],
  ["Gösteride: Ctrl+P / Ctrl+L / E", "Kalem / lazer işaretçisi / mürekkebi sil"],
  ["Gösteride: Esc", "Gösteriyi bitir"],
];

export function ShortcutsDialog({ open, onClose }: { open: boolean; onClose: () => void }) {
  return (
    <Dialog open={open} title="Klavye Kısayolları" onClose={onClose} width={560} footer={<DialogButton primary onClick={onClose}>Tamam</DialogButton>}>
      <div style={{ maxHeight: "60vh", overflow: "auto" }}>
        <table className="kbd-table">
          <tbody>
            {SHORTCUTS.map(([k, d]) => (
              <tr key={k}>
                <td>{k}</td>
                <td>{d}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Dialog>
  );
}

export function AboutDialog({ open, onClose }: { open: boolean; onClose: () => void }) {
  return (
    <Dialog open={open} title="OfficeHere Sunum Hakkında" onClose={onClose} width={440} footer={<DialogButton primary onClick={onClose}>Tamam</DialogButton>}>
      <p>
        <strong>OfficeHere Sunum</strong> — OfficeHere açık kaynak ofis paketinin sunu uygulaması.
      </p>
      <p>MIT lisansıyla dağıtılır. Dosya biçimi desteği açık ECMA-376 / ISO/IEC 29500 (Office Open XML) standardına dayanır.</p>
      <DeveloperCredit />
      <p>
        <a href="./THIRD_PARTY_NOTICES.txt" target="_blank" rel="noopener noreferrer">
          Üçüncü taraf yazılım bildirimleri
        </a>
      </p>
      <p style={{ fontSize: 12, opacity: 0.8 }}>
        OfficeHere bağımsız bir projedir; Microsoft Corporation ile bağlantılı değildir, onun tarafından desteklenmez. Microsoft ve PowerPoint, Microsoft
        Corporation'ın ticari markalarıdır.
      </p>
    </Dialog>
  );
}
