/** Veri sekmesi: Veri Al, Sırala ve Filtre Uygula, Veri Araçları. */
import {
  ArrowSync20Regular,
  ArrowSync24Regular,
  DocumentText24Regular,
  Filter24Regular,
  FilterDismiss20Regular,
  FilterAdd20Regular,
  TextSortAscending20Regular,
  TextSortDescending20Regular,
  ArrowSort24Regular,
  TextColumnTwo24Regular,
  TableDismiss20Regular,
  CheckboxChecked24Regular,
  Target24Regular,
  Group24Regular,
  GroupDismiss24Regular,
  TableCalculator20Regular,
  AddSquare20Regular,
  SubtractSquare20Regular,
  Flash20Regular,
  TableStackBelow20Regular,
} from "@fluentui/react-icons";
import { openFile } from "@officehere/core";
import { Menu, MenuItem, RibbonButton, RibbonGroup, RibbonMenuButton, RibbonSplitButton, RibbonStack } from "@officehere/ui";
import { clearOutline, groupSelection, showDetail, ungroupSelection } from "../actions/outline";
import { flashFill } from "../actions/whatIf";
import type { Controller } from "../controller";
import * as D from "../actions/data";
import { refreshAll } from "../pivot/pivotOps";
import { subtractRects } from "../model/rect";

export async function importText(ctl: Controller): Promise<void> {
  const f = await openFile([{ name: "Metin dosyaları", extensions: ["csv", "txt", "tsv"] }]);
  if (f) ctl.openDialog("textImport", { name: f.name, data: f.data });
}

export function DataTab({ ctl }: { ctl: Controller }) {
  const sheet = ctl.sheet;
  const filters = D.filtersOf(sheet);
  const filterOn = !!sheet.autoFilter || !!sheet.tableAt(ctl.sel.active.row, ctl.sel.active.col)?.filterButton;
  const anyCriteria = filters.some((f) => Object.keys(f.columns).length > 0);
  return (
    <>
      <RibbonGroup label="Veri Al ve Dönüştür" icon={<DocumentText24Regular />}>
        <RibbonButton size="large" icon={<DocumentText24Regular />} label="Metin/CSV'den" description="Metin, virgülle ya da noktalı virgülle ayrılmış değer dosyasından veri alır." onClick={() => void importText(ctl)} />
      </RibbonGroup>
      <RibbonGroup label="Sorgular ve Bağlantılar" icon={<ArrowSync24Regular />}>
        <RibbonButton size="large" icon={<ArrowSync24Regular />} label="Tümünü Yenile" shortcut="Ctrl+Alt+F5" description="Çalışma kitabındaki tüm özet tabloları kaynak verilerinden yeniden hesaplar." disabled={!ctl.wb.sheets.some((s) => s.pivots.length)} onClick={() => refreshAll(ctl)} />
      </RibbonGroup>
      <RibbonGroup label="Sırala ve Filtre Uygula" icon={<Filter24Regular />}>
        <RibbonStack>
          <RibbonButton size="small" icon={<TextSortAscending20Regular />} tip="A'dan Z'ye Sırala" description="En düşükten en yükseğe sıralar." onClick={() => D.quickSort(ctl, false)} />
          <RibbonButton size="small" icon={<TextSortDescending20Regular />} tip="Z'den A'ya Sırala" description="En yüksekten en düşüğe sıralar." onClick={() => D.quickSort(ctl, true)} />
        </RibbonStack>
        <RibbonButton size="large" icon={<ArrowSort24Regular />} label="Sırala" description="Verileri birden çok ölçüte göre sıralar." onClick={() => ctl.openDialog("sort")} />
        <RibbonButton size="large" icon={<Filter24Regular />} label="Filtre" shortcut="Ctrl+Shift+L" checked={filterOn} description="Seçili hücreler için filtrelemeyi açar." onClick={() => D.toggleAutoFilter(ctl)} />
        <RibbonStack>
          <RibbonButton icon={<FilterDismiss20Regular />} label="Temizle" disabled={!anyCriteria && !sheet.advFilter} onClick={() => D.clearAllFilters(ctl)} />
          <RibbonButton icon={<ArrowSync20Regular />} label="Yeniden Uygula" disabled={!filters.length && !sheet.advFilter} onClick={() => D.reapplyFilters(ctl)} />
          <RibbonButton icon={<FilterAdd20Regular />} label="Gelişmiş" onClick={() => ctl.openDialog("advancedFilter")} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Veri Araçları" icon={<CheckboxChecked24Regular />}>
        <RibbonButton size="large" icon={<TextColumnTwo24Regular />} label="Metni Sütunlara Dönüştür" className="tb-wide" description="Tek bir sütundaki metni ayırıcıya göre birden çok sütuna böler." onClick={() => ctl.openDialog("textToColumns")} />
        <RibbonStack>
          <RibbonButton icon={<Flash20Regular />} label="Akıllı Doldurma" shortcut="Ctrl+E" description="Yazdığınız örneklerdeki deseni bulur ve sütunun geri kalanını otomatik doldurur." onClick={() => flashFill(ctl)} />
          <RibbonButton icon={<TableDismiss20Regular />} label="Yinelenenleri Kaldır" description="Sayfadaki yinelenen satırları siler." onClick={() => ctl.openDialog("removeDuplicates")} />
          <RibbonButton icon={<TableStackBelow20Regular />} label="Birleştir" description="Birden çok aralığın (başka sayfalardakiler de) değerlerini tek bir tabloda toplar, ortalamasını alır ya da sayar." onClick={() => ctl.openDialog("consolidate")} />
        </RibbonStack>
        <RibbonSplitButton
          size="large"
          icon={<CheckboxChecked24Regular />}
          label="Veri Doğrulama"
          description="Hücreye girilebilecek veri türünü sınırlar (ör. açılır liste)."
          onClick={() => ctl.openDialog("dataValidation")}
          menu={
            <Menu width={220}>
              <MenuItem label="Veri Doğrulama..." onClick={() => ctl.openDialog("dataValidation")} />
              <MenuItem
                label="Doğrulamayı Temizle"
                onClick={() =>
                  ctl.ops.sheetOp(sheet, "Veri Doğrulama", () => {
                    const sel = ctl.sel.ranges;
                    sheet.validations = sheet.validations.map((v) => ({ ...v, ranges: subtractRects(v.ranges, sel) })).filter((v) => v.ranges.length);
                  })
                }
              />
            </Menu>
          }
        />
      </RibbonGroup>
      <RibbonGroup label="Tahmin" icon={<Target24Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<Target24Regular />}
          label="Durum Çözümlemesi"
          className="tb-wide"
          menu={
            <Menu width={200}>
              <MenuItem label="Senaryo Yöneticisi..." onClick={() => ctl.openDialog("scenarioManager")} />
              <MenuItem label="Hedef Ara..." onClick={() => ctl.openDialog("goalSeek")} />
              <MenuItem label="Veri Tablosu..." onClick={() => ctl.openDialog("dataTable")} />
            </Menu>
          }
        />
      </RibbonGroup>
      <RibbonGroup label="Anahat" icon={<Group24Regular />}>
        <RibbonSplitButton
          size="large"
          icon={<Group24Regular />}
          label="Gruplandır"
          shortcut="Alt+Shift+Sağ Ok"
          description="Satır ya da sütun aralığını, daraltılıp genişletilebilen bir grupta birleştirir."
          onClick={() => groupSelection(ctl)}
          menu={
            <Menu width={200}>
              <MenuItem label="Gruplandır..." onClick={() => groupSelection(ctl)} />
            </Menu>
          }
        />
        <RibbonSplitButton
          size="large"
          icon={<GroupDismiss24Regular />}
          label="Grubu Çöz"
          shortcut="Alt+Shift+Sol Ok"
          description="Gruplandırılmış satır ya da sütunları gruptan çıkarır."
          onClick={() => ungroupSelection(ctl)}
          menu={
            <Menu width={200}>
              <MenuItem label="Grubu Çöz..." onClick={() => ungroupSelection(ctl)} />
              <MenuItem label="Anahattı Temizle" onClick={() => clearOutline(ctl)} />
            </Menu>
          }
        />
        <RibbonButton size="large" icon={<TableCalculator20Regular />} label="Alt Toplam" description="İlgili veri satırlarına alt toplam ve genel toplam ekleyerek gruplandırır." onClick={() => ctl.openDialog("subtotal")} />
        <RibbonStack>
          <RibbonButton icon={<AddSquare20Regular />} label="Ayrıntıyı Göster" onClick={() => showDetail(ctl, true)} />
          <RibbonButton icon={<SubtractSquare20Regular />} label="Ayrıntıyı Gizle" onClick={() => showDetail(ctl, false)} />
        </RibbonStack>
      </RibbonGroup>
    </>
  );
}
