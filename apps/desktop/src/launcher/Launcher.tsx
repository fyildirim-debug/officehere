import { useEffect, useState } from "react";
import { Bug20Regular, Bug24Regular, FolderOpen24Regular, Home20Regular, DocumentAdd20Regular, FolderOpen20Regular, Settings20Regular } from "@fluentui/react-icons";
import { getNative, getRecentFiles, openFile, type AppKind, type RecentFile } from "@officehere/core";
import { ALL_EXTENSIONS, FORMAT_GROUPS, appForFileName } from "../shared/formats";
import {
  AppLogo,
  AppShell,
  Backstage,
  BackstageAction,
  BackstageHeading,
  BackstageSubheading,
  BugReportDialog,
  DeveloperCredit,
  RecentFileList,
  SpellcheckPrompt,
  SpellcheckSection,
  TemplateRow,
  TemplateTile,
  TitleBar,
  UI_THEMES,
  UpdateSection,
  useUiTheme,
} from "@officehere/ui";

const APPS: { app: AppKind; label: string; name: string; shape: "portrait" | "landscape" | "square" }[] = [
  { app: "belge", label: "Boş belge", name: "Belge", shape: "portrait" },
  { app: "tablo", label: "Boş çalışma kitabı", name: "Tablo", shape: "square" },
  { app: "sunum", label: "Boş sunu", name: "Sunum", shape: "landscape" },
];


function greeting() {
  const h = new Date().getHours();
  if (h >= 5 && h < 12) return "Günaydın";
  if (h >= 12 && h < 18) return "İyi günler";
  if (h >= 18 && h < 22) return "İyi akşamlar";
  return "İyi geceler";
}

/** Masaüstünde yeni pencere açar; web'de ilgili uygulamaya gider. */
function launch(app: AppKind, path?: string) {
  const native = getNative();
  if (native) {
    void native.newWindow(app, path);
    return;
  }
  window.location.href = `./${app}/`;
}

function Preview({ app }: { app: AppKind }) {
  if (app === "belge") {
    return (
      <svg viewBox="0 0 130 168" className="ohl-preview">
        <rect width="130" height="168" fill="#fff" />
      </svg>
    );
  }
  if (app === "tablo") {
    return (
      <svg viewBox="0 0 150 120" className="ohl-preview">
        <rect width="150" height="120" fill="#fff" />
        {Array.from({ length: 7 }, (_, i) => (
          <line key={`h${i}`} x1="0" x2="150" y1={i * 18 + 12} y2={i * 18 + 12} stroke="#e1e1e1" />
        ))}
        {Array.from({ length: 5 }, (_, i) => (
          <line key={`v${i}`} y1="0" y2="120" x1={i * 34 + 14} x2={i * 34 + 14} stroke="#e1e1e1" />
        ))}
        <rect x="14" y="12" width="34" height="18" fill="none" stroke="#1d7a46" strokeWidth="2" />
      </svg>
    );
  }
  return (
    <svg viewBox="0 0 196 110" className="ohl-preview">
      <rect width="196" height="110" fill="#fff" />
      <rect x="38" y="36" width="120" height="16" fill="none" stroke="#c8c8c8" strokeDasharray="3 2" />
      <rect x="58" y="60" width="80" height="10" fill="none" stroke="#c8c8c8" strokeDasharray="3 2" />
    </svg>
  );
}

/** Seçenekler: açılışta son oturumu geri yükleme (yalnız masaüstü; ayar ana süreçte saklanır). */
function SessionOption() {
  const native = getNative();
  const [on, setOn] = useState<boolean | null>(null);
  useEffect(() => {
    native
      ?.getSettings?.()
      .then((s) => setOn(s.oturumuGeriYukle === true))
      .catch(() => setOn(false));
  }, [native]);
  if (!native?.getSettings || !native.setSetting) return null;
  return (
    <div className="ohl-option">
      <label className="ohl-check">
        <input
          type="checkbox"
          checked={on === true}
          disabled={on === null}
          onChange={(e) => {
            const value = e.target.checked;
            setOn(value);
            void native.setSetting?.("oturumuGeriYukle", value).then((s) => setOn(s.oturumuGeriYukle === true));
          }}
        />
        <span>Açılışta son oturumu geri yükle</span>
      </label>
      <p className="ohl-option-note">OfficeHere dosya seçilmeden başlatıldığında, son kapatılışta açık olan belgeler kendi pencerelerinde yeniden açılır.</p>
    </div>
  );
}

/** Ana süreç "Hata bildir…" isteyince başlangıç ekranı ?hata-bildir=1 ile açılır. */
function wantsBugReport(): boolean {
  try {
    return new URLSearchParams(window.location.search).has("hata-bildir");
  } catch {
    return false;
  }
}

export function Launcher() {
  const [page, setPage] = useState("home");
  const [recent, setRecent] = useState<RecentFile[]>([]);
  const [theme, setTheme] = useUiTheme();
  const [reporting, setReporting] = useState(wantsBugReport);
  const native = getNative();

  useEffect(() => {
    getRecentFiles().then(setRecent).catch(() => setRecent([]));
    const onFocus = () => getRecentFiles().then(setRecent).catch(() => {});
    window.addEventListener("focus", onFocus);
    return () => window.removeEventListener("focus", onFocus);
  }, []);

  const browse = async () => {
    const f = await openFile([{ name: "Tüm desteklenen dosyalar", extensions: ALL_EXTENSIONS }, ...FORMAT_GROUPS.map(({ name, extensions }) => ({ name, extensions }))]);
    if (!f) return;
    const app = appForFileName(f.name);
    if (app && f.path) launch(app, f.path);
    else if (app) launch(app);
  };

  const newRow = (
    <TemplateRow>
      {APPS.map((a) => (
        <TemplateTile key={a.app} label={a.label} shape={a.shape} preview={<Preview app={a.app} />} onClick={() => launch(a.app)} />
      ))}
    </TemplateRow>
  );

  let content;
  if (page === "home") {
    content = (
      <>
        <BackstageHeading>{greeting()}</BackstageHeading>
        <SpellcheckPrompt />
        <div className="ohl-apps">
          {APPS.map((a) => (
            <button key={a.app} type="button" className={`ohl-app is-${a.app}`} onClick={() => launch(a.app)}>
              <AppLogo app={a.app} size={40} />
              <span>{a.name}</span>
            </button>
          ))}
        </div>
        <BackstageSubheading>Yeni</BackstageSubheading>
        {newRow}
        <BackstageSubheading>Son Kullanılanlar</BackstageSubheading>
        <RecentFileList files={recent.slice(0, 15)} onOpen={(f) => launch(f.app, f.path)} emptyText={native ? "Henüz dosya açmadınız." : "Son kullanılan dosyalar masaüstü uygulamasında listelenir."} />
        <DeveloperCredit variant="footer" />
      </>
    );
  } else if (page === "new") {
    content = (
      <>
        <BackstageHeading>Yeni</BackstageHeading>
        {newRow}
      </>
    );
  } else if (page === "open") {
    content = (
      <>
        <BackstageHeading>Aç</BackstageHeading>
        <BackstageAction icon={<FolderOpen24Regular />} title="Gözat" description="Belge, çalışma kitabı ya da sunu açın" onClick={browse} />
        <BackstageSubheading>Son Kullanılanlar</BackstageSubheading>
        <RecentFileList files={recent} onOpen={(f) => launch(f.app, f.path)} />
      </>
    );
  } else {
    content = (
      <>
        <BackstageHeading>Seçenekler</BackstageHeading>
        <label className="oh-field is-inline" style={{ maxWidth: 420 }}>
          <span className="oh-field-label">Arayüz teması</span>
          <select value={theme} onChange={(e) => setTheme(e.target.value as typeof theme)}>
            {UI_THEMES.map((t) => (
              <option key={t.id} value={t.id}>
                {t.label}
              </option>
            ))}
          </select>
        </label>
        <SessionOption />
        <UpdateSection />
        <SpellcheckSection />
        <BackstageSubheading>Sorun bildirme</BackstageSubheading>
        <BackstageAction
          icon={<Bug24Regular />}
          title="Hata bildir"
          description="Sürüm, işletim sistemi ve son günlük satırlarını görüp düzenleyebileceğiniz bir bildirim hazırlar. Siz göndermeden hiçbir şey paylaşılmaz."
          onClick={() => setReporting(true)}
        />
        <DeveloperCredit />
        <p className="oh-bs-note">
          OfficeHere açık kaynaklı (MIT) bağımsız bir projedir; Microsoft Corporation ile bağlantılı değildir. Microsoft, Word, Excel ve PowerPoint, Microsoft
          Corporation'ın ticari markalarıdır.
        </p>
        <p className="oh-bs-note">
          <a href="./THIRD_PARTY_NOTICES.txt" target="_blank" rel="noopener noreferrer">
            Üçüncü taraf yazılım bildirimleri
          </a>
        </p>
      </>
    );
  }

  return (
    <AppShell
      app="baslangic"
      titleBar={<TitleBar app="baslangic" title="Başlangıç" />}
      backstage={
        <Backstage
          items={[
            { id: "home", label: "Giriş", icon: <Home20Regular /> },
            { id: "new", label: "Yeni", icon: <DocumentAdd20Regular /> },
            { id: "open", label: "Aç", icon: <FolderOpen20Regular /> },
            { id: "report", label: "Hata bildir", icon: <Bug20Regular />, bottom: true, onClick: () => setReporting(true) },
            { id: "options", label: "Seçenekler", icon: <Settings20Regular />, bottom: true },
          ]}
          activeId={page}
          onSelect={(id) => {
            if (id !== "report") setPage(id);
          }}
          onBack={() => setPage("home")}
        >
          {content}
        </Backstage>
      }
    >
      <div />
      <BugReportDialog open={reporting} onClose={() => setReporting(false)} app="baslangic" />
    </AppShell>
  );
}
