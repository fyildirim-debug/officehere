/** Formüller sekmesi: İşlev Kitaplığı, Tanımlı Adlar, Formül Denetimi, Hesaplama. */
import {
  Book24Regular,
  BookOpen20Regular,
  Calculator24Regular,
  CalculatorArrowClockwise20Regular,
  CalendarLtr24Regular,
  Money24Regular,
  MathFormula24Regular,
  Search24Regular,
  TextT24Regular,
  Branch24Regular,
  Tag20Regular,
  TagMultiple20Regular,
  DocumentBulletList24Regular,
  History24Regular,
  Settings20Regular,
  Eye20Regular,
} from "@fluentui/react-icons";
import { useMemo } from "react";
import { Menu, MenuHeader, MenuItem, MenuSeparator, RibbonButton, RibbonGroup, RibbonMenuButton, RibbonSplitButton, RibbonStack } from "@officehere/ui";
import { FUNCTIONS, type FnDef } from "@officehere/formula";
import type { Controller } from "../controller";
import * as E from "../actions/edit";
import { allFunctions } from "../dialogs/FunctionDialog";
import { FxGlyph, SigmaGlyph } from "./glyphs";

function FnMenu({ ctl, list }: { ctl: Controller; list: FnDef[] }) {
  return (
    <Menu width={200} className="tb-menu-scroll">
      {list.map((d) => (
        <MenuItem key={d.name} label={d.tr} onClick={() => ctl.openDialog("insertFunction", { name: d.name })} />
      ))}
      <MenuSeparator />
      <MenuItem label="İşlev Ekle..." onClick={() => ctl.openDialog("insertFunction")} />
    </Menu>
  );
}

/** Ad ya da fonksiyonu düzenlenen formüle ekler (düzenleme yoksa başlatır). */
function insertIntoFormula(ctl: Controller, text: string): void {
  if (!ctl.edit) {
    ctl.startEdit("=" + text);
    return;
  }
  const e = ctl.edit;
  const t = e.text.slice(0, e.caret) + text + e.text.slice(e.caret);
  ctl.updateEdit(t, e.caret + text.length);
}

function recentList(): FnDef[] {
  let names: string[] = ["SUM", "AVERAGE", "IF", "HYPERLINK", "COUNT", "MAX", "SIN", "SUMIF", "PMT", "STDEV"];
  try {
    const s = localStorage.getItem("tablo.recentFunctions");
    if (s) names = JSON.parse(s) as string[];
  } catch {
    /* depolama yok */
  }
  return names.map((n) => FUNCTIONS.get(n)).filter((d): d is FnDef => !!d);
}

export function FormulasTab({ ctl }: { ctl: Controller }) {
  const all = useMemo(allFunctions, []);
  const byCat = (c: string) => all.filter((d) => d.category === c);
  const sheet = ctl.sheet;
  const auto = ctl.wb.engine.autoCalc;
  const names = ctl.wb.engine.getNames();
  return (
    <>
      <RibbonGroup label="İşlev Kitaplığı" icon={<MathFormula24Regular />}>
        <RibbonButton size="large" icon={<FxGlyph />} label="İşlev Ekle" shortcut="Shift+F3" description="Geçerli hücredeki formülü düzenlemek için işlev seçin." onClick={() => ctl.openDialog("insertFunction")} />
        <RibbonSplitButton
          size="large"
          icon={<SigmaGlyph size={32} />}
          label="Otomatik Toplam"
          shortcut="Alt+="
          onClick={() => E.autoSum(ctl)}
          menu={
            <Menu>
              <MenuItem label="Topla" onClick={() => E.autoSum(ctl, "SUM")} />
              <MenuItem label="Ortalama" onClick={() => E.autoSum(ctl, "AVERAGE")} />
              <MenuItem label="Sayıları Say" onClick={() => E.autoSum(ctl, "COUNT")} />
              <MenuItem label="En Büyük" onClick={() => E.autoSum(ctl, "MAX")} />
              <MenuItem label="En Küçük" onClick={() => E.autoSum(ctl, "MIN")} />
              <MenuSeparator />
              <MenuItem label="Diğer İşlevler..." onClick={() => ctl.openDialog("insertFunction")} />
            </Menu>
          }
        />
        <RibbonMenuButton size="large" icon={<History24Regular />} label="Son Kullanılan" menu={() => <FnMenu ctl={ctl} list={recentList()} />} />
        <RibbonMenuButton size="large" icon={<Money24Regular />} label="Finansal" menu={<FnMenu ctl={ctl} list={byCat("financial")} />} />
        <RibbonMenuButton size="large" icon={<Branch24Regular />} label="Mantıksal" menu={<FnMenu ctl={ctl} list={byCat("logical")} />} />
        <RibbonMenuButton size="large" icon={<TextT24Regular />} label="Metin" menu={<FnMenu ctl={ctl} list={byCat("text")} />} />
        <RibbonMenuButton size="large" icon={<CalendarLtr24Regular />} label="Tarih ve Saat" menu={<FnMenu ctl={ctl} list={byCat("date")} />} />
        <RibbonMenuButton size="large" icon={<Search24Regular />} label="Arama ve Başvuru" menu={<FnMenu ctl={ctl} list={byCat("lookup")} />} />
        <RibbonMenuButton size="large" icon={<Calculator24Regular />} label="Matematik ve Trig" menu={<FnMenu ctl={ctl} list={byCat("math")} />} />
        <RibbonMenuButton
          size="large"
          icon={<Book24Regular />}
          label="Diğer İşlevler"
          menu={
            <Menu width={200}>
              {(
                [
                  ["stats", "İstatistik"],
                  ["info", "Bilgi"],
                  ["engineering", "Mühendislik"],
                  ["database", "Veritabanı"],
                ] as const
              )
                .filter(([c]) => byCat(c).length)
                .map(([c, label]) => (
                  <MenuItem key={c} label={label} submenu={<FnMenu ctl={ctl} list={byCat(c)} />} />
                ))}
            </Menu>
          }
        />
      </RibbonGroup>
      <RibbonGroup label="Tanımlı Adlar" icon={<TagMultiple20Regular />}>
        <RibbonButton size="large" icon={<DocumentBulletList24Regular />} label="Ad Yöneticisi" shortcut="Ctrl+F3" description="Çalışma kitabındaki tüm adları oluşturur, düzenler, siler ve bulur." onClick={() => ctl.openDialog("nameManager")} />
        <RibbonStack>
          <RibbonButton icon={<Tag20Regular />} label="Ad Tanımla" onClick={() => ctl.openDialog("defineName")} />
          <RibbonMenuButton
            icon={<BookOpen20Regular />}
            label="Formülde Kullan"
            menu={
              <Menu width={200}>
                {names.length === 0 && <MenuHeader>Tanımlı ad yok</MenuHeader>}
                {names.map((n) => (
                  <MenuItem key={`${n.sheet ?? ""}${n.name}`} label={n.name} onClick={() => insertIntoFormula(ctl, n.name)} />
                ))}
                {ctl.wb.sheets.flatMap((s) => s.tables).map((t) => (
                  <MenuItem key={t.id} label={t.name} onClick={() => insertIntoFormula(ctl, t.name)} />
                ))}
              </Menu>
            }
          />
          <RibbonButton icon={<TagMultiple20Regular />} label="Seçimden Oluştur" shortcut="Ctrl+Shift+F3" onClick={() => ctl.openDialog("createNames")} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Formül Denetimi" icon={<Eye20Regular />}>
        <RibbonStack>
          <RibbonButton
            icon={<Eye20Regular />}
            label="Formülleri Göster"
            shortcut="Ctrl+`"
            checked={sheet.showFormulas}
            description="Her hücrede sonuç yerine formülü gösterir."
            onClick={() => {
              sheet.showFormulas = !sheet.showFormulas;
              ctl.emit();
            }}
          />
          <RibbonButton icon={<Search24Regular />} label="Formülleri Seç" description="Sayfadaki formül içeren tüm hücreleri seçer." onClick={() => selectFormulas(ctl)} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Hesaplama" icon={<Calculator24Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<Settings20Regular />}
          label="Hesaplama Seçenekleri"
          menu={
            <Menu width={180}>
              <MenuItem label="Otomatik" checked={auto} onClick={() => setAuto(ctl, true)} />
              <MenuItem label="El ile" checked={!auto} onClick={() => setAuto(ctl, false)} />
            </Menu>
          }
        />
        <RibbonStack>
          <RibbonButton
            icon={<CalculatorArrowClockwise20Regular />}
            label="Şimdi Hesapla"
            shortcut="F9"
            onClick={() => {
              ctl.wb.engine.recalc(true);
              ctl.emit();
            }}
          />
          <RibbonButton
            icon={<Calculator24Regular />}
            label="Sayfayı Hesapla"
            shortcut="Shift+F9"
            onClick={() => {
              ctl.wb.engine.recalc();
              ctl.emit();
            }}
          />
        </RibbonStack>
      </RibbonGroup>
    </>
  );
}

function setAuto(ctl: Controller, on: boolean): void {
  ctl.wb.engine.autoCalc = on;
  if (on) ctl.wb.engine.recalc();
  ctl.emit();
}

function selectFormulas(ctl: Controller): void {
  const ranges: { r1: number; c1: number; r2: number; c2: number }[] = [];
  ctl.wb.engine.forEachCell(ctl.sheet.id, (row, col, rec) => {
    if (rec.formula) ranges.push({ r1: row, c1: col, r2: row, c2: col });
  });
  if (!ranges.length) return ctl.alert("Hiçbir hücre bulunamadı.");
  ctl.setSelection({ ranges, active: { row: ranges[0].r1, col: ranges[0].c1 }, anchor: { row: ranges[0].r1, col: ranges[0].c1 } });
}
