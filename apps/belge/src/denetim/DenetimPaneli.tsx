// Gözden Geçir › Resmî Yazı Denetimi: Yönetmeliğe uygunluk denetiminin sonuçları. Kural başına durum, bulgular ve
// bulguya git ("Göster"); kenar boşlukları tek tıkla düzeltilebilir.
import { useMemo, useState } from "react";
import { CheckmarkCircle16Filled, DismissCircle16Filled, Info16Filled, Warning16Filled } from "@fluentui/react-icons";
import type { Node as PMNode } from "@tiptap/pm/model";
import { Dialog, DialogButton, Field } from "@officehere/ui";
import { useBelge } from "../state/context";
import { YONLER } from "../templates/resmi/alanlar";
import { KENAR_TW, LOGO_UST_CM } from "../templates/resmi/olcu";
import { twToCm } from "../model/document";
import { resmiYaziDenetle, type Bulgu, type DenetimYonu, type Durum, type KuralSonucu } from "./resmiYazi";

const SINIF: Record<Durum, string> = { uygun: "is-success", bilgi: "is-info", uyari: "is-warning", hata: "is-error" };
const DURUM_ADI: Record<Durum, string> = { uygun: "Uygun", bilgi: "Bilgi", uyari: "Uyarı", hata: "Hata" };

function Simge({ durum }: { durum: Durum }) {
  const s = { flex: "none", marginTop: 1 } as const;
  if (durum === "uygun") return <CheckmarkCircle16Filled style={{ ...s, color: "#107c10" }} aria-hidden="true" />;
  if (durum === "hata") return <DismissCircle16Filled style={{ ...s, color: "#c50f1f" }} aria-hidden="true" />;
  if (durum === "uyari") return <Warning16Filled style={{ ...s, color: "#bc4b09" }} aria-hidden="true" />;
  return <Info16Filled style={{ ...s, color: "#0f6cbd" }} aria-hidden="true" />;
}

/** Bölgedeki n. paragraf/başlık düğümünün konumu (denetimin paragraf sırasıyla aynı gezinti). */
function paragrafKonumu(doc: PMNode, indeks: number): { from: number; to: number } | null {
  let i = 0;
  let bulunan: { from: number; to: number } | null = null;
  doc.descendants((n, pos) => {
    if (bulunan) return false;
    if (n.type.name === "paragraph" || n.type.name === "heading") {
      if (i++ === indeks) bulunan = { from: pos + 1, to: pos + 1 + n.content.size };
      return false;
    }
    return true;
  });
  return bulunan;
}

export function DenetimPaneli({ onClose }: { onClose: () => void }) {
  const ctx = useBelge();
  const [yon, setYon] = useState<DenetimYonu>("otomatik");
  const [tur, setTur] = useState(0);
  const sonuclar = useMemo(() => resmiYaziDenetle({ content: ctx.mainEditor.getJSON(), settings: ctx.settings }, { yon }), [ctx.mainEditor, ctx.settings, yon, tur]); // eslint-disable-line react-hooks/exhaustive-deps
  const say = (d: Durum) => sonuclar.filter((s) => s.durum === d).length;

  const goster = (b: Bulgu) => {
    if (b.paragraf === undefined) return;
    if (b.bolge === "altBilgi" || b.bolge === "ustBilgi") {
      onClose();
      ctx.editHeaderFooter(b.bolge === "altBilgi" ? "footer" : "header");
      return;
    }
    const ed = ctx.mainEditor;
    const yer = paragrafKonumu(ed.state.doc, b.paragraf);
    if (!yer) return;
    onClose();
    ed.chain().focus().setTextSelection(yer).scrollIntoView().run();
  };

  const kenarDuzelt = () =>
    ctx.updateSettings((s) => {
      const m = s.page.margin;
      // Logo için 0,5 cm üst kenar korunur; başka her değer 1,5 cm olur.
      const top = Math.abs(twToCm(m.top) - LOGO_UST_CM) < 0.05 ? m.top : KENAR_TW;
      return { ...s, page: { ...s.page, margin: { ...m, top, left: KENAR_TW, right: KENAR_TW } } };
    });

  const satir = (k: KuralSonucu) => (
    <li
      key={k.kural}
      className={`oh-messagebar ${SINIF[k.durum]}`}
      style={{ alignItems: "flex-start", border: 0, borderRadius: 4, marginBottom: 4, flexDirection: "column", gap: 2 }}
      data-kural={k.kural}
      data-durum={k.durum}
    >
      <div style={{ display: "flex", gap: 8, width: "100%", alignItems: "flex-start" }}>
        <Simge durum={k.durum} />
        <div style={{ flex: 1 }}>
          <strong>{k.ad}</strong>{" "}
          <span style={{ color: "var(--oh-text-muted)" }}>
            ({k.madde}) · {DURUM_ADI[k.durum]}
          </span>
          {k.durum === "uygun" && <div>{k.ozet}</div>}
        </div>
        {k.kural === "kenar" && k.bulgular.some((b) => b.durum === "hata") && (
          <button type="button" className="oh-btn" onClick={kenarDuzelt}>
            Düzelt
          </button>
        )}
      </div>
      {k.bulgular.length > 0 && (
        <ul style={{ margin: "2px 0 0 24px", padding: 0, listStyle: "none", width: "calc(100% - 24px)" }}>
          {k.bulgular.map((b, i) => (
            <li key={i} style={{ display: "flex", gap: 8, alignItems: "baseline", padding: "2px 0" }}>
              <span style={{ flex: 1 }}>{b.ileti}</span>
              {b.paragraf !== undefined && (
                <button type="button" className="oh-btn" style={{ height: 22, minWidth: 0 }} onClick={() => goster(b)}>
                  Göster
                </button>
              )}
            </li>
          ))}
        </ul>
      )}
    </li>
  );

  return (
    <Dialog
      open
      title="Resmî Yazı Denetimi"
      width={620}
      onClose={onClose}
      footer={
        <>
          <DialogButton onClick={() => setTur((n) => n + 1)}>Yeniden Denetle</DialogButton>
          <DialogButton primary onClick={onClose}>
            Kapat
          </DialogButton>
        </>
      }
    >
      <p className="bl-dlg-note">Belge, 2020 tarihli Resmî Yazışmalarda Uygulanacak Usul ve Esaslar Hakkında Yönetmelik'in ölçülerine göre denetlenir.</p>
      <Field label="Muhatap düzeyi (arz / rica uyumu için)" inline>
        <select value={yon} onChange={(e) => setYon(e.target.value as DenetimYonu)}>
          <option value="otomatik">Muhataptan anla</option>
          {YONLER.map((y) => (
            <option key={y.id} value={y.id}>
              {y.etiket}
            </option>
          ))}
        </select>
      </Field>
      <p className="bl-resmi-ozet" role="status">
        {say("hata") || say("uyari")
          ? `${say("hata")} hata, ${say("uyari")} uyarı; ${say("uygun")} kural uygun.`
          : `Denetlenen ${sonuclar.length} kuraldan ${say("uygun")} tanesi uygun; hata ya da uyarı yok.`}
      </p>
      <ul style={{ listStyle: "none", margin: 0, padding: 0, maxHeight: "min(56vh, 520px)", overflowY: "auto" }} aria-label="Denetim sonuçları">
        {sonuclar.map(satir)}
      </ul>
    </Dialog>
  );
}
