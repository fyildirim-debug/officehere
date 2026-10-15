// Dosya › Yeni › "Resmî yazışma": önizlemeli şablon galerisi ve alan formu. "Oluştur" form değerlerini bekletir
// ve yeni belgeyi her zamanki yoldan (useFileOps.newDocument) açtırır.
import { useMemo, useState, type ReactNode } from "react";
import { Dialog, DialogButton, Field, RibbonGallery } from "@officehere/ui";
import { konuSorunlari } from "../../denetim/resmiYazi";
import { BASLIK_TURLERI, KAPANIS, YONLER, yonelmeEkiVar, type ResmiAlanlar } from "./alanlar";
import { YAZI_TIPLERI, ZARF } from "./olcu";
import { ONIZLEMELER } from "./onizleme";
import { RESMI_SABLONLAR, resmiAlanlar, resmiSablon } from "./sablonlar";

interface Props {
  onClose: () => void;
  onOlustur: (sablon: string, alanlar: Partial<ResmiAlanlar>) => void;
}

type MetinAlani = { [K in keyof ResmiAlanlar]: ResmiAlanlar[K] extends string ? K : never }[keyof ResmiAlanlar];

const izgara = { display: "grid", gridTemplateColumns: "1fr 1fr", columnGap: 16 } as const;
const tamSatir = { gridColumn: "1 / -1" } as const;
const uyariStili = { margin: "-4px 0 6px", color: "var(--oh-text-muted)", fontSize: 11 } as const;

function Grup({ baslik, children }: { baslik: string; children: ReactNode }) {
  return (
    <fieldset className="oh-fieldset" style={izgara}>
      <legend>{baslik}</legend>
      {children}
    </fieldset>
  );
}

export function ResmiSablonPenceresi({ onClose, onOlustur }: Props) {
  const [secili, setSecili] = useState(RESMI_SABLONLAR[0].id);
  // Yalnız kullanıcının değiştirdiği alanlar; şablon değişince o şablona özgü varsayılanlar yeniden geçerli olur.
  const [form, setForm] = useState<Partial<ResmiAlanlar>>({});
  const sablon = resmiSablon(secili) ?? RESMI_SABLONLAR[0];
  const a = useMemo(() => resmiAlanlar(sablon.id, form), [sablon.id, form]);
  const g = new Set(sablon.gruplar);
  const ata = (patch: Partial<ResmiAlanlar>) => setForm((f) => ({ ...f, ...patch }));

  const sec = (id: string) => {
    const yeni = resmiSablon(id);
    if (!yeni) return;
    setSecili(id);
    setForm((f) => Object.fromEntries(Object.entries(f).filter(([k]) => !(k in yeni.varsayilan))) as Partial<ResmiAlanlar>);
  };

  const metin = (ad: MetinAlani, etiket: string, ek: { satir?: number; tam?: boolean; ipucu?: string } = {}) => (
    <div style={ek.tam ? tamSatir : undefined} key={ad}>
      <Field label={etiket}>
        {ek.satir ? (
          <textarea rows={ek.satir} value={a[ad]} placeholder={ek.ipucu} onChange={(e) => ata({ [ad]: e.target.value })} />
        ) : (
          <input type="text" value={a[ad]} placeholder={ek.ipucu} onChange={(e) => ata({ [ad]: e.target.value })} />
        )}
      </Field>
    </div>
  );

  const konuUyari = g.has("sayi") && sablon.id !== "resmi-zarf" ? konuSorunlari(a.konu) : [];
  const muhatapUyari = g.has("muhatap") && !yonelmeEkiVar(a.muhatap);
  const olustur = () => onOlustur(sablon.id, form);

  return (
    <Dialog
      open
      title="Resmî Yazışma Şablonları"
      width={700}
      onClose={onClose}
      onSubmit={olustur}
      footer={
        <>
          <DialogButton primary onClick={olustur}>
            Oluştur
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <div className="bl-resmi-galeri">
        <RibbonGallery
          items={RESMI_SABLONLAR.map((s) => ({ id: s.id, label: s.etiket, preview: ONIZLEMELER[s.id] }))}
          selectedId={sablon.id}
          onSelect={sec}
          itemWidth={72}
          itemHeight={96}
          columns={8}
        />
      </div>
      <p className="bl-dlg-note" aria-live="polite">
        <strong>{sablon.etiket}.</strong> {sablon.aciklama} Ölçüler 2020 tarihli Resmî Yazışma Yönetmeliği'ne göredir.
      </p>
      <div style={{ maxHeight: "min(52vh, 480px)", overflowY: "auto", paddingRight: 4 }}>
        {g.has("baslik") && (
          <Grup baslik="Başlık">
            <Field label="Başlık türü">
              <select
                value={a.baslikTuru}
                onChange={(e) => {
                  const b = BASLIK_TURLERI.find((x) => x.id === e.target.value)!;
                  ata({ baslikTuru: b.id, idare: b.idare, araIdare: b.araIdare ?? "", birim: b.birim });
                }}
              >
                {BASLIK_TURLERI.map((b) => (
                  <option key={b.id} value={b.id}>
                    {b.etiket}
                  </option>
                ))}
              </select>
            </Field>
            <Field label="Yazı tipi">
              <select value={a.yaziTipi} onChange={(e) => ata({ yaziTipi: e.target.value as ResmiAlanlar["yaziTipi"] })}>
                {Object.entries(YAZI_TIPLERI).map(([id, y]) => (
                  <option key={id} value={id}>
                    {y.etiket}
                  </option>
                ))}
              </select>
            </Field>
            {metin(
              "idare",
              a.baslikTuru === "il" || a.baslikTuru === "ilMem" ? "Valilik (2. satır)" : a.baslikTuru === "ilceMem" || a.baslikTuru === "okul" ? "Kaymakamlık (2. satır)" : "İdare adı (2. satır)",
            )}
            {a.baslikTuru === "bagli" && metin("araIdare", "Bağlı kurum (3. satır)")}
            {metin("birim", a.baslikTuru === "okul" ? "Okul müdürlüğü (son satır)" : "Birim adı (son satır)")}
            {g.has("iletisim") && (
              <label className="bl-check" style={{ alignSelf: "end", marginBottom: 10 }}>
                <input type="checkbox" checked={a.logo} onChange={(e) => ata({ logo: e.target.checked })} /> Sol üstte logo yeri bırak
              </label>
            )}
          </Grup>
        )}
        {g.has("sayi") && (
          <Grup baslik="Sayı ve tarih">
            <Field label="Belge ortamı">
              <select value={a.ortam} onChange={(e) => ata({ ortam: e.target.value as ResmiAlanlar["ortam"] })}>
                <option value="E">Elektronik (E)</option>
                <option value="Z">Zorunlu hâl, ıslak imza (Z)</option>
                <option value="O">Olağanüstü durum (O)</option>
              </select>
            </Field>
            {metin("tarih", "Tarih", { ipucu: "GG.AA.YYYY" })}
            {metin("detsis", "DETSİS numarası")}
            {metin("sdp", "Dosya planı kodu")}
            {metin("kayitNo", "Kayıt numarası", { ipucu: "Belge yönetim sistemi verir" })}
            <Field label="İvedilik">
              <select value={a.ivedilik} onChange={(e) => ata({ ivedilik: e.target.value as ResmiAlanlar["ivedilik"] })}>
                <option value="">Yok</option>
                <option value="ACELE">ACELE</option>
                <option value="GÜNLÜDÜR">GÜNLÜDÜR</option>
                {sablon.id === "resmi-zarf" && <option value="KİŞİYE ÖZEL">KİŞİYE ÖZEL</option>}
              </select>
            </Field>
            {sablon.id !== "resmi-zarf" && metin("konu", "Konu", { tam: true })}
            {konuUyari.map((u) => (
              <p key={u} role="note" style={{ ...uyariStili, ...tamSatir }}>
                {u}
              </p>
            ))}
          </Grup>
        )}
        {g.has("dilekce") && (
          <Grup baslik="Dilekçe sahibi">
            {metin("dilekceAd", "Adı SOYADI")}
            {metin("tarih", "Tarih", { ipucu: "GG.AA.YYYY" })}
            {metin("tcKimlik", "T.C. kimlik numarası")}
            {metin("dilekceTelefon", "Telefon")}
            {metin("dilekceAdres", "Adres", { tam: true })}
            {metin("dilekceEposta", "e-Posta")}
          </Grup>
        )}
        {g.has("tutanak") && (
          <Grup baslik="Tutanak ya da rapor">
            <Field label="Belge türü">
              <select value={a.belgeTuru} onChange={(e) => ata({ belgeTuru: e.target.value as ResmiAlanlar["belgeTuru"], belgeBasligi: e.target.value === "rapor" ? "RAPOR" : "TUTANAK" })}>
                <option value="tutanak">Tutanak</option>
                <option value="rapor">Rapor</option>
              </select>
            </Field>
            {metin("belgeBasligi", "Belge başlığı")}
            {metin("tarih", "Tarih (metnin sonunda)", { ipucu: "GG.AA.YYYY" })}
            {metin("imzacilar", "İmzalayanlar", { satir: 3, tam: true, ipucu: "Her satıra: Ad SOYAD, Unvan" })}
          </Grup>
        )}
        {g.has("muhatap") && (
          <Grup baslik="Muhatap">
            {metin("muhatap", "Muhatap", { ipucu: "… MÜDÜRLÜĞÜNE ya da Sayın Ad SOYAD" })}
            {metin("muhatapBirim", "Muhatabın birimi (parantez içinde)")}
            {muhatapUyari && (
              <p role="note" style={{ ...uyariStili, ...tamSatir }}>
                Muhatap yönelme ekiyle yazılır: “… MÜDÜRLÜĞÜNE”, “… BAKANLIĞINA” (m.14/2).
              </p>
            )}
            {(g.has("zarf") || sablon.id === "resmi-ust-yazi") && metin("muhatapAdres", "Muhatabın adresi", { satir: 2, tam: true })}
          </Grup>
        )}
        {g.has("ilgi") && <Grup baslik="İlgi">{metin("ilgiler", "İlgi (her satıra bir belge)", { satir: 2, tam: true, ipucu: "… tarihli ve … sayılı yazı." })}</Grup>}
        {g.has("metin") && (
          <Grup baslik="Metin">
            {metin("metin", "Metin (her satır bir paragraf)", { satir: 4, tam: true })}
            {g.has("imza") && (
              <Field label="Muhatap düzeyi">
                <select value={a.yon} onChange={(e) => ata({ yon: e.target.value as ResmiAlanlar["yon"], kapanis: KAPANIS[e.target.value as ResmiAlanlar["yon"]] })}>
                  {YONLER.map((y) => (
                    <option key={y.id} value={y.id}>
                      {y.etiket}
                    </option>
                  ))}
                </select>
              </Field>
            )}
            {sablon.id !== "resmi-tutanak" && metin("kapanis", "Kapanış")}
          </Grup>
        )}
        {g.has("imza") && (
          <Grup baslik={g.has("olur") ? "Teklif eden" : "İmza"}>
            {metin("imzaAd", "Adı SOYADI")}
            {metin("imzaUnvan", "Unvanı")}
            {metin("imzaYetki", "Yetki devri ya da vekâlet", { ipucu: "Vali a. / Genel Müdür V." })}
          </Grup>
        )}
        {g.has("olur") && (
          <Grup baslik="Olur">
            <label className="bl-check" style={tamSatir}>
              <input type="checkbox" checked={a.araMakam} onChange={(e) => ata({ araMakam: e.target.checked })} /> Ara makam görüşü (Uygun görüşle arz ederim.)
            </label>
            {a.araMakam && metin("araMakamAd", "Ara makam: adı SOYADI")}
            {a.araMakam && metin("araMakamUnvan", "Ara makam: unvanı")}
            {metin("olurAd", "Oluru veren: adı SOYADI")}
            {metin("olurUnvan", "Oluru veren: unvanı")}
          </Grup>
        )}
        {g.has("ek") && <Grup baslik="Ekler">{metin("ekler", "Ekler (her satıra bir ek)", { satir: 2, tam: true, ipucu: "Liste (3 sayfa)" })}</Grup>}
        {g.has("dagitim") && (
          <Grup baslik="Dağıtım">
            {metin("geregi", "Gereği (her satıra bir yer)", { satir: 3 })}
            {metin("bilgi", "Bilgi (her satıra bir yer)", { satir: 3 })}
          </Grup>
        )}
        {g.has("zarf") && (
          <Grup baslik="Zarf">
            <Field label="Zarf boyu">
              <select value={a.zarfBoyu} onChange={(e) => ata({ zarfBoyu: e.target.value as ResmiAlanlar["zarfBoyu"] })}>
                {Object.entries(ZARF).map(([id, z]) => (
                  <option key={id} value={id}>
                    {z.etiket}
                  </option>
                ))}
              </select>
            </Field>
            {metin("gonderenAdres", "Gönderen adresi (isteğe bağlı)", { satir: 2 })}
          </Grup>
        )}
        {g.has("iletisim") && (
          <details>
            <summary style={{ cursor: "pointer", margin: "8px 0" }}>İletişim bilgileri (her sayfanın altında)</summary>
            <Grup baslik="İletişim">
              {metin("adres", "Adres ve posta kodu", { tam: true })}
              {metin("telefon", "Telefon")}
              {metin("faks", "Faks")}
              {metin("eposta", "e-Posta")}
              {metin("kep", "KEP adresi")}
              {metin("internet", "İnternet adresi")}
              {metin("bilgiKisi", "Bilgi için: adı SOYADI")}
              {metin("bilgiUnvan", "Bilgi için: unvanı")}
              {metin("bilgiTelefon", "Bilgi için: telefonu")}
            </Grup>
          </details>
        )}
      </div>
    </Dialog>
  );
}
