// Güncelleştirme bildirimi: başlık çubuğunda rozet ve Hakkında sayfasında denetim bölümü (kanal
// seçimiyle); ayrıca yazım denetimi sözlüklerinin indirme onayı (ilk açılış sorusu ve Seçenekler).
// Yalnız masaüstünde görünür; tarayıcı sürümünde köprü olmadığı için hiçbir şey çizilmez.
import { useEffect, useState } from "react";
import { getNative, type SpellcheckState, type UpdateChannel, type UpdateState } from "@officehere/core";
import { ArrowClockwise16Regular, ArrowDownload16Regular } from "@fluentui/react-icons";
import { tipProps } from "./Tooltip";

/** Masaüstündeki güncelleştirme durumu; tarayıcıda null. */
export function useUpdates(): UpdateState | null {
  const [state, setState] = useState<UpdateState | null>(null);
  useEffect(() => {
    const api = getNative()?.updates;
    if (!api) return;
    let live = true;
    void api.getState().then((s) => live && setState(s));
    const off = api.onChange((s) => live && setState(s));
    return () => {
      live = false;
      off();
    };
  }, []);
  return state;
}

const updatesApi = () => getNative()?.updates;

/** Başlık çubuğu rozeti: yeni sürüm, indirme ilerlemesi ya da kuruluma hazır. */
export function UpdatePill() {
  const s = useUpdates();
  if (!s || (s.phase !== "available" && s.phase !== "downloading" && s.phase !== "ready")) return null;
  const v = s.version ?? "";
  const pill =
    s.phase === "ready"
      ? { label: "Yeniden başlat ve güncelleştir", tip: `OfficeHere ${v} indirildi`, desc: "Kaydedilmemiş belgeler için önce sorulur. Uygulamayı kapattığınızda da kurulur." }
      : s.phase === "downloading"
        ? { label: `Güncelleştirme indiriliyor %${s.percent ?? 0}`, tip: `OfficeHere ${v} indiriliyor`, desc: "İndirme bitince yeniden başlatıp kurabilirsiniz." }
        : { label: `Yeni sürüm: ${v}`, tip: `OfficeHere ${v} yayımlandı`, desc: "İndirme sayfasını açmak için tıklayın." };
  return (
    <button
      type="button"
      className={`oh-update-pill is-${s.phase}`}
      style={s.phase === "downloading" ? ({ "--oh-update-p": `${s.percent ?? 0}%` } as React.CSSProperties) : undefined}
      aria-disabled={s.phase === "downloading" || undefined}
      onClick={() => {
        if (s.phase !== "downloading") void updatesApi()?.install();
      }}
      {...tipProps(pill.tip, undefined, pill.desc)}
    >
      {s.phase === "ready" ? <ArrowClockwise16Regular /> : <ArrowDownload16Regular />}
      <span>{pill.label}</span>
    </button>
  );
}

function statusText(s: UpdateState): string {
  const at = s.checkedAt ? ` Son denetim: ${new Date(s.checkedAt).toLocaleTimeString("tr-TR", { hour: "2-digit", minute: "2-digit" })}.` : "";
  switch (s.phase) {
    case "checking":
      return "Güncelleştirmeler denetleniyor…";
    case "latest":
      return `OfficeHere güncel.${at}`;
    case "available":
      return s.mode === "auto" && s.error ? `Sürüm ${s.version} indirilemedi; indirme sayfasından kurabilirsiniz.` : `Sürüm ${s.version} yayımlandı.`;
    case "downloading":
      return `Sürüm ${s.version} indiriliyor (%${s.percent ?? 0})…`;
    case "ready":
      return `Sürüm ${s.version} kurulmaya hazır.`;
    case "error":
      return `Denetlenemedi: ${s.error ?? "bilinmeyen hata"}`;
    default:
      return s.mode === "auto" ? "Güncelleştirmeler arka planda indirilir, yeniden başlatınca kurulur." : "Yeni sürüm çıktığında bildirilir ve indirme sayfası açılır.";
  }
}

/** Hakkında / Seçenekler: sürüm, durum, elle denetim ve otomatik denetim seçeneği. */
export function UpdateSection() {
  const s = useUpdates();
  const [busy, setBusy] = useState(false);
  // Seçenek tıklanınca hemen değişir; ana süreçten gelen durum onu sonra doğrular.
  const [autoCheck, setAutoCheck] = useState<boolean | null>(null);
  const [channel, setChannel] = useState<UpdateChannel | null>(null);
  useEffect(() => {
    if (s) setAutoCheck(s.autoCheck);
  }, [s?.autoCheck]);
  useEffect(() => {
    if (s) setChannel(s.channel ?? "kararli");
  }, [s?.channel]);
  if (!s) return null;
  const api = updatesApi();
  const shownChannel = channel ?? s.channel ?? "kararli";
  return (
    <section className="oh-update-section" aria-label="Güncelleştirmeler">
      <div className="oh-update-row">
        <div>
          <div className="oh-update-version">Sürüm {s.current}</div>
          <div className={`oh-update-status is-${s.phase}`} role="status">
            {s.mode !== "off"
              ? statusText(s)
              : // Mağaza paketinde ana süreç "store: true" gönderir (köprü tipinde yok; eski kabuklar göndermez).
                "store" in s && s.store === true
                ? "Bu sürüm uygulama mağazasından güncellenir; OfficeHere kendisi güncelleştirme denetlemez."
                : "Geliştirme sürümünde güncelleştirme denetimi kapalı."}
          </div>
        </div>
        <div className="oh-update-actions">
          {(s.phase === "ready" || s.phase === "available") && (
            <button type="button" className="oh-btn is-primary" onClick={() => void api?.install()}>
              {s.phase === "ready" ? "Yeniden başlat ve güncelleştir" : "İndirme sayfasını aç"}
            </button>
          )}
          <button
            type="button"
            className="oh-btn"
            disabled={s.mode === "off" || busy || s.phase === "checking" || s.phase === "downloading" || s.phase === "ready"}
            onClick={async () => {
              setBusy(true);
              try {
                await api?.check();
              } finally {
                setBusy(false);
              }
            }}
          >
            Güncelleştirmeleri denetle
          </button>
        </div>
      </div>
      {s.mode !== "off" && (
        <label className="oh-update-auto">
          <input
            type="checkbox"
            checked={autoCheck ?? s.autoCheck}
            onChange={(e) => {
              setAutoCheck(e.target.checked);
              void api?.setAutoCheck(e.target.checked);
            }}
          />
          <span>Güncelleştirmeleri otomatik denetle</span>
        </label>
      )}
      {s.mode !== "off" && api?.setChannel && (
        <>
          <label className="oh-field is-inline oh-update-channel">
            <span className="oh-field-label">Güncelleme kanalı</span>
            <select
              value={shownChannel}
              onChange={(e) => {
                const next = e.target.value === "beta" ? "beta" : "kararli";
                setChannel(next);
                void api.setChannel?.(next);
              }}
            >
              <option value="kararli">Kararlı</option>
              <option value="beta">Beta</option>
            </select>
          </label>
          <p className="oh-update-note">
            {shownChannel === "beta"
              ? "Beta sürümleri yenilikleri önce getirir ama hata içerebilir. Kararlı kanala dönünce kurulu sürüm geri alınmaz; bir sonraki kararlı sürüm beklenir."
              : "Yalnız kararlı sürümler alınır."}
          </p>
        </>
      )}
    </section>
  );
}

/** Masaüstündeki yazım sözlüğü durumu; tarayıcıda null. */
export function useSpellcheck(): [SpellcheckState | null, (on: boolean) => void] {
  const [state, setState] = useState<SpellcheckState | null>(null);
  useEffect(() => {
    const api = getNative()?.spellcheck;
    if (!api) return;
    let live = true;
    void api.getState().then((s) => live && setState(s));
    const off = api.onChange((s) => live && setState(s));
    return () => {
      live = false;
      off();
    };
  }, []);
  const choose = (on: boolean) => {
    void getNative()
      ?.spellcheck?.setDownload(on)
      .then((s) => setState(s));
  };
  return [state, choose];
}

const SPELL_STATUS: Record<SpellcheckState["status"], string> = {
  sistem: "macOS'un kendi yazım denetimi kullanılır; sözlük indirilmez.",
  yok: "Sözlükler indirilmedi; yazım denetimi kapalı.",
  indiriliyor: "Sözlükler indiriliyor…",
  hazir: "Türkçe ve İngilizce sözlükler yüklü; yazım denetimi ağa çıkmadan çalışır.",
  hata: "Sözlükler indirilemedi. İnternet bağlantınızı denetleyip yeniden deneyin.",
};

/** Seçenekler: yazım denetimi sözlüklerini indir / kaldır ve durum. */
export function SpellcheckSection() {
  const [s, choose] = useSpellcheck();
  if (!s) return null;
  const installed = s.status === "hazir" || s.status === "indiriliyor";
  return (
    <section className="oh-spell-section" aria-label="Yazım denetimi sözlükleri">
      <div className="oh-update-row">
        <div>
          <div className="oh-update-version">Yazım denetimi sözlükleri</div>
          <div className={`oh-update-status is-${s.status === "hata" ? "error" : s.status}`} role="status">
            {SPELL_STATUS[s.status]}
          </div>
        </div>
        {s.status !== "sistem" && (
          <div className="oh-update-actions">
            {installed ? (
              <button type="button" className="oh-btn" onClick={() => choose(false)}>
                Kaldır
              </button>
            ) : (
              <button type="button" className="oh-btn is-primary" onClick={() => choose(true)}>
                İndir
              </button>
            )}
          </div>
        )}
      </div>
      {s.status !== "sistem" && (
        <p className="oh-update-note">
          Türkçe ve İngilizce sözlükler (yaklaşık {s.sizeMb} MB) yalnız siz isteyince, bir kez Google sunucusundan indirilir ve bilgisayarınızda
          saklanır. Sonrasında yazım denetimi ağa çıkmaz.
        </p>
      )}
    </section>
  );
}

/** İlk açılış sorusu (başlangıç ekranı): yalnız hiç sorulmadıysa ve sözlük yerelde yoksa görünür. */
export function SpellcheckPrompt() {
  const [s, choose] = useSpellcheck();
  if (!s?.ask) return null;
  return (
    <section className="oh-spell-prompt" aria-label="Yazım denetimi sözlükleri">
      <p>Yazım denetimi için Türkçe ve İngilizce sözlükler (yaklaşık {s.sizeMb} MB) bir kez Google sunucusundan indirilir. İndirilsin mi?</p>
      <div className="oh-spell-prompt-actions">
        <button type="button" className="oh-btn is-primary" onClick={() => choose(true)}>
          İndir
        </button>
        <button type="button" className="oh-btn" onClick={() => choose(false)}>
          Şimdi değil
        </button>
      </div>
      <p className="oh-update-note">Seçiminizi Seçenekler'den istediğiniz zaman değiştirebilirsiniz.</p>
    </section>
  );
}
