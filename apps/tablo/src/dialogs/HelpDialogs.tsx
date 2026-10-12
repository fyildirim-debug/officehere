/** Yardım iletişim kutuları: klavye kısayolları ve çalışma kitabı istatistikleri. */
import { useMemo } from "react";
import { Dialog, DialogButton } from "@officehere/ui";
import { colToName } from "@officehere/formula";
import type { Controller } from "../controller";

type P = { ctl: Controller; onClose: () => void; arg?: unknown };

export const SHORTCUTS: { group: string; items: [string, string][] }[] = [
  {
    group: "Gezinme",
    items: [
      ["Ok tuşları", "Bir hücre taşı"],
      ["Ctrl+Ok", "Veri bölgesinin kenarına git"],
      ["Home / Ctrl+Home / Ctrl+End", "Satır başı / A1 / son hücre"],
      ["Page Up / Page Down", "Bir ekran yukarı / aşağı"],
      ["Ctrl+G veya F5", "Git"],
      ["Ctrl+Page Up / Page Down", "Önceki / sonraki sayfa"],
    ],
  },
  {
    group: "Seçim",
    items: [
      ["Shift+Ok", "Seçimi genişlet"],
      ["Ctrl+Shift+Ok", "Seçimi veri kenarına genişlet"],
      ["Ctrl+A", "Geçerli bölgeyi / tümünü seç"],
      ["Ctrl+Boşluk / Shift+Boşluk", "Sütunu / satırı seç"],
    ],
  },
  {
    group: "Düzenleme",
    items: [
      ["F2", "Hücreyi düzenle"],
      ["Enter / Shift+Enter", "Onayla ve aşağı / yukarı git"],
      ["Tab / Shift+Tab", "Onayla ve sağa / sola git"],
      ["Alt+Enter", "Hücrede yeni satır"],
      ["Ctrl+Enter", "Seçili tüm hücreleri doldur"],
      ["Esc", "Düzenlemeyi iptal et"],
      ["F4", "Başvuru türünü değiştir ($A$1, A$1, $A1, A1)"],
      ["Delete", "İçeriği temizle"],
      ["Ctrl+D / Ctrl+R", "Aşağı / sağa doldur"],
      ["Ctrl+; / Ctrl+Shift+;", "Bugünün tarihi / saat"],
      ["Alt+=", "Otomatik Toplam"],
      ["Shift+F3", "İşlev Ekle"],
    ],
  },
  {
    group: "Pano ve geri alma",
    items: [
      ["Ctrl+C / Ctrl+X / Ctrl+V", "Kopyala / Kes / Yapıştır"],
      ["Ctrl+Alt+V", "Özel Yapıştır"],
      ["Ctrl+Z / Ctrl+Y", "Geri Al / Yinele"],
    ],
  },
  {
    group: "Biçimlendirme",
    items: [
      ["Ctrl+B / Ctrl+I / Ctrl+U", "Kalın / İtalik / Altı çizili"],
      ["Ctrl+5", "Üstü çizili"],
      ["Ctrl+1", "Hücreleri Biçimlendir"],
      ["Ctrl+Shift+4 / 5 / 3", "Para birimi / yüzde / tarih biçimi"],
    ],
  },
  {
    group: "Diğer",
    items: [
      ["Ctrl+F / Ctrl+H", "Bul / Değiştir"],
      ["Ctrl+`", "Formülleri göster"],
      ["F9", "Şimdi hesapla"],
      ["Ctrl+S / F12", "Kaydet / Farklı Kaydet"],
      ["Ctrl+O / Ctrl+N / Ctrl+P", "Aç / Yeni / Yazdır"],
      ["Ctrl+F1", "Şeridi daralt / genişlet"],
    ],
  },
];

export function ShortcutsDialog({ onClose }: P) {
  return (
    <Dialog open title="Klavye Kısayolları" onClose={onClose} width={620} footer={<DialogButton primary onClick={onClose}>Kapat</DialogButton>}>
      <div className="tb-shortcuts">
        {SHORTCUTS.map((g) => (
          <section key={g.group}>
            <h4>{g.group}</h4>
            <table>
              <tbody>
                {g.items.map(([k, d]) => (
                  <tr key={k}>
                    <td>
                      <kbd>{k}</kbd>
                    </td>
                    <td>{d}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </section>
        ))}
      </div>
    </Dialog>
  );
}

export function WorkbookStatsDialog({ ctl, onClose }: P) {
  const stats = useMemo(() => {
    const wb = ctl.wb;
    const sheet = ctl.sheet;
    const count = (s: typeof sheet) => {
      let cells = 0;
      let formulas = 0;
      wb.engine.forEachCell(s.id, (_r, _c, rec) => {
        cells++;
        if (rec.formula) formulas++;
      });
      return { cells, formulas };
    };
    const cur = count(sheet);
    const all = wb.sheets.map(count).reduce((a, b) => ({ cells: a.cells + b.cells, formulas: a.formulas + b.formulas }), { cells: 0, formulas: 0 });
    const u = wb.usedRange(sheet);
    const sum = (f: (s: typeof sheet) => number) => wb.sheets.reduce((n, s) => n + f(s), 0);
    return {
      sheet: [
        ["Son hücre", u ? `${colToName(u.c2)}${u.r2 + 1}` : "A1"],
        ["Veri içeren hücreler", cur.cells],
        ["Formüller", cur.formulas],
        ["Tablolar", sheet.tables.length],
        ["Grafikler", sheet.charts.length],
        ["Resimler ve şekiller", sheet.images.length + sheet.shapes.length],
        ["Notlar", sheet.notes.size],
      ] as [string, string | number][],
      book: [
        ["Sayfalar", wb.sheets.length],
        ["Veri içeren hücreler", all.cells],
        ["Formüller", all.formulas],
        ["Tablolar", sum((s) => s.tables.length)],
        ["Grafikler", sum((s) => s.charts.length)],
        ["Tanımlı adlar", wb.engine.getNames().length],
      ] as [string, string | number][],
    };
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const fmt = (v: string | number) => (typeof v === "number" ? v.toLocaleString("tr-TR") : v);
  return (
    <Dialog open title="Çalışma Kitabı İstatistikleri" onClose={onClose} width={380} footer={<DialogButton primary onClick={onClose}>Tamam</DialogButton>}>
      <div className="tb-stats">
        <h4>Geçerli sayfa: {ctl.sheet.name}</h4>
        <table>
          <tbody>
            {stats.sheet.map(([k, v]) => (
              <tr key={k}>
                <td>{k}</td>
                <td>{fmt(v)}</td>
              </tr>
            ))}
          </tbody>
        </table>
        <h4>Çalışma kitabı</h4>
        <table>
          <tbody>
            {stats.book.map(([k, v]) => (
              <tr key={k}>
                <td>{k}</td>
                <td>{fmt(v)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Dialog>
  );
}
