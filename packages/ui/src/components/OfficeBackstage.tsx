import { useEffect, useState, type ReactElement, type ReactNode } from "react";
import {
  ArrowDownload24Regular,
  Document20Regular,
  DocumentAdd20Regular,
  DocumentArrowRight20Regular,
  FolderOpen20Regular,
  FolderOpen24Regular,
  Home20Regular,
  Info20Regular,
  Print20Regular,
  Print24Regular,
  Save20Regular,
  SaveEdit20Regular,
  Settings20Regular,
  Dismiss20Regular,
  QuestionCircle20Regular,
} from "@fluentui/react-icons";
import { getRecentFiles, isDesktop, type AppKind, type RecentFile } from "@officehere/core";
import {
  Backstage,
  BackstageAction,
  BackstageHeading,
  BackstageSubheading,
  PropertyList,
  RecentFileList,
  TemplateRow,
  TemplateTile,
  type BackstageNavItem,
} from "./Backstage";
import { APP_NAMES } from "./AppShell";
import { UI_THEMES, useUiTheme } from "../hooks/useUiTheme";
import { UpdateSection } from "./Updates";
import { useAutoSavePref } from "../hooks/useAutoSave";

export interface BackstageTemplate {
  id: string;
  label: string;
  preview: ReactNode;
}

export interface BackstageExport {
  id: string;
  title: string;
  description?: string;
  icon?: ReactElement;
}

export type OfficeBackstagePage = "home" | "new" | "open" | "info" | "print" | "export" | "options" | "about";

export interface OfficeBackstageProps {
  app: AppKind;
  fileName: string;
  filePath?: string;
  templates: BackstageTemplate[];
  templateShape?: "portrait" | "landscape" | "square";
  onNew: (templateId: string) => void;
  onOpen: () => void;
  onOpenRecent: (file: RecentFile) => void;
  onSave: () => void;
  onSaveAs: () => void;
  onPrint: () => void;
  exports: BackstageExport[];
  onExport: (id: string) => void;
  onClose: () => void;
  onBack: () => void;
  /** Bilgi sayfasındaki özellikler (Boyut, Sözcük sayısı, Sayfa sayısı…). */
  info?: [string, ReactNode][];
  initialPage?: OfficeBackstagePage;
  /** Seçenekler sayfasına uygulamaya özel ek ayarlar. */
  extraOptions?: ReactNode;
  /** Yazdır sayfasında baskı önizlemesi. */
  printPreview?: ReactNode;
  /** Yazdır sayfasının tüm içeriği (verilirse varsayılan Yazdır sayfasının yerine geçer). */
  printPage?: ReactNode;
}

function greeting(): string {
  const h = new Date().getHours();
  if (h >= 5 && h < 12) return "Günaydın";
  if (h >= 12 && h < 18) return "İyi günler";
  if (h >= 18 && h < 22) return "İyi akşamlar";
  return "İyi geceler";
}

const NEW_LABEL: Record<AppKind, string> = { belge: "Boş belge", tablo: "Boş çalışma kitabı", sunum: "Boş sunu" };

/** Üç uygulamada ortak "Dosya" görünümü: Giriş, Yeni, Aç, Bilgi, Kaydet, Yazdır, Dışarı Aktar, Seçenekler. */
export function OfficeBackstage(props: OfficeBackstageProps) {
  const { app, fileName, filePath, templates, templateShape = "portrait", onNew, onOpen, onOpenRecent, onSave, onSaveAs, onPrint, exports, onExport, onClose, onBack, info, initialPage = "home", extraOptions, printPreview, printPage } = props;
  const [page, setPage] = useState<OfficeBackstagePage>(initialPage);
  const [recent, setRecent] = useState<RecentFile[]>([]);
  const [theme, setTheme] = useUiTheme();

  useEffect(() => {
    getRecentFiles(app).then(setRecent).catch(() => setRecent([]));
  }, [app]);

  const items: BackstageNavItem[] = [
    { id: "home", label: "Giriş", icon: <Home20Regular /> },
    { id: "new", label: "Yeni", icon: <DocumentAdd20Regular /> },
    { id: "open", label: "Aç", icon: <FolderOpen20Regular /> },
    { id: "info", label: "Bilgi", icon: <Info20Regular />, separator: true },
    { id: "save", label: "Kaydet", icon: <Save20Regular />, onClick: onSave },
    { id: "saveas", label: "Farklı Kaydet", icon: <SaveEdit20Regular />, onClick: onSaveAs },
    { id: "print", label: "Yazdır", icon: <Print20Regular /> },
    { id: "export", label: "Dışarı Aktar", icon: <DocumentArrowRight20Regular /> },
    { id: "close", label: "Kapat", icon: <Dismiss20Regular />, onClick: onClose },
    { id: "about", label: "Hakkında", icon: <QuestionCircle20Regular />, bottom: true },
    { id: "options", label: "Seçenekler", icon: <Settings20Regular />, bottom: true },
  ];

  const templateTiles = (list: BackstageTemplate[]) => (
    <TemplateRow>
      {list.map((t) => (
        <TemplateTile key={t.id} label={t.label} preview={t.preview} shape={templateShape} onClick={() => onNew(t.id)} />
      ))}
    </TemplateRow>
  );

  let content: ReactNode;
  switch (page) {
    case "home":
      content = (
        <>
          <BackstageHeading>{greeting()}</BackstageHeading>
          <BackstageSubheading>Yeni</BackstageSubheading>
          {templateTiles(templates.slice(0, 6))}
          <BackstageSubheading>Son Kullanılanlar</BackstageSubheading>
          <RecentFileList files={recent.slice(0, 12)} onOpen={onOpenRecent} />
        </>
      );
      break;
    case "new":
      content = (
        <>
          <BackstageHeading>Yeni</BackstageHeading>
          {templateTiles(templates)}
        </>
      );
      break;
    case "open":
      content = (
        <>
          <BackstageHeading>Aç</BackstageHeading>
          <BackstageAction icon={<FolderOpen24Regular />} title="Gözat" description={isDesktop() ? "Bu bilgisayardaki bir dosyayı açın" : "Cihazınızdaki bir dosyayı seçin"} onClick={onOpen} />
          <BackstageSubheading>Son Kullanılanlar</BackstageSubheading>
          <RecentFileList files={recent} onOpen={onOpenRecent} />
        </>
      );
      break;
    case "info":
      content = (
        <>
          <BackstageHeading>Bilgi</BackstageHeading>
          <p className="oh-bs-file">
            <Document20Regular /> <strong>{fileName}</strong>
            {filePath && <span className="oh-bs-filepath">{filePath}</span>}
          </p>
          <BackstageSubheading>Özellikler</BackstageSubheading>
          <PropertyList rows={info ?? []} />
        </>
      );
      break;
    case "print":
      content = printPage ?? (
        <>
          <BackstageHeading>Yazdır</BackstageHeading>
          <div className="oh-bs-print">
            <div className="oh-bs-print-side">
              <BackstageAction icon={<Print24Regular />} title="Yazdır" description="Yazıcı seçimi ve sayfa ayarları sistem iletişim kutusunda yapılır." onClick={onPrint} />
              <p className="oh-bs-note">PDF dosyası için yazıcı olarak “PDF olarak kaydet” seçeneğini kullanın.</p>
            </div>
            {printPreview && <div className="oh-bs-print-preview">{printPreview}</div>}
          </div>
        </>
      );
      break;
    case "export":
      content = (
        <>
          <BackstageHeading>Dışarı Aktar</BackstageHeading>
          {exports.map((e) => (
            <BackstageAction key={e.id} icon={e.icon ?? <ArrowDownload24Regular />} title={e.title} description={e.description} onClick={() => onExport(e.id)} />
          ))}
        </>
      );
      break;
    case "options":
      content = (
        <>
          <BackstageHeading>Seçenekler</BackstageHeading>
          <BackstageSubheading>Genel</BackstageSubheading>
          <div className="oh-bs-options">
            <label className="oh-field is-inline">
              <span className="oh-field-label">Arayüz teması</span>
              <select value={theme} onChange={(e) => setTheme(e.target.value as typeof theme)}>
                {UI_THEMES.map((t) => (
                  <option key={t.id} value={t.id}>
                    {t.label}
                  </option>
                ))}
              </select>
            </label>
            <UserNameField />
            {isDesktop() && <AutoSaveField />}
            {extraOptions}
          </div>
        </>
      );
      break;
    case "about":
      content = (
        <>
          <BackstageHeading>Hakkında</BackstageHeading>
          <p>
            <strong>{APP_NAMES[app]}</strong> — OfficeHere açık kaynak ofis paketinin parçasıdır.
          </p>
          <p>
            MIT lisansıyla dağıtılır. Kaynak kodu:{" "}
            <a href="https://github.com/fyildirim-debug/officehere" target="_blank" rel="noopener noreferrer">
              github.com/fyildirim-debug/officehere
            </a>
          </p>
          <p>
            Kullanılan açık kaynak bileşenlerin lisansları:{" "}
            <a href="./THIRD_PARTY_NOTICES.txt" target="_blank" rel="noopener noreferrer">
              Üçüncü taraf yazılım bildirimleri
            </a>
          </p>
          <UpdateSection />
          <p className="oh-bs-note">
            OfficeHere bağımsız bir projedir; Microsoft Corporation ile bağlantılı değildir, onun tarafından desteklenmez. Microsoft, Word, Excel ve PowerPoint,
            Microsoft Corporation'ın ticari markalarıdır. Dosya biçimi desteği açık ECMA-376 / ISO/IEC 29500 (Office Open XML) standardına dayanır.
          </p>
        </>
      );
      break;
  }

  return (
    <Backstage items={items} activeId={page} onSelect={(id) => setPage(id as OfficeBackstagePage)} onBack={onBack}>
      {content}
    </Backstage>
  );
}

const USER_KEY = "officehere.userName";

/** Açıklamalar ve belge özelliklerinde kullanılan kullanıcı adı. */
export function getUserName(): string {
  try {
    return localStorage.getItem(USER_KEY) || "Kullanıcı";
  } catch {
    return "Kullanıcı";
  }
}

/** Otomatik kaydet (yalnız masaüstü; tarayıcıda dosyaya doğrudan yazılamaz). */
function AutoSaveField() {
  const [on, setOn] = useAutoSavePref();
  return (
    <label className="oh-field is-inline oh-option-check">
      <span className="oh-option-text">
        <span className="oh-field-label">Otomatik kaydet</span>
        <span className="oh-option-desc" id="oh-autosave-desc">
          Bilgisayarınızdaki belgeler değiştikçe kendiliğinden kaydedilir.
        </span>
      </span>
      <input type="checkbox" aria-label="Otomatik kaydet" aria-describedby="oh-autosave-desc" checked={on} onChange={(e) => setOn(e.target.checked)} />
    </label>
  );
}

function UserNameField() {
  const [name, setName] = useState(getUserName);
  return (
    <label className="oh-field is-inline">
      <span className="oh-field-label">Kullanıcı adı</span>
      <input
        type="text"
        value={name}
        onChange={(e) => {
          setName(e.target.value);
          try {
            localStorage.setItem(USER_KEY, e.target.value);
          } catch {
            // yok say
          }
        }}
      />
    </label>
  );
}

export { NEW_LABEL as BLANK_TEMPLATE_LABEL };
