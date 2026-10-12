/** Dosya > Yeni şablonları: gerçek Türkçe içerik, formüller ve biçimlerle. */
import { WorkbookModel, type SheetModel } from "../model/workbook";
import type { CellStyle } from "../model/styles";
import { ACCOUNTING_FORMAT } from "../model/styles";
import type { CellInput } from "../model/workbook";

export interface TemplateDef {
  id: string;
  label: string;
  build: () => WorkbookModel;
}

type Row = (CellInput | undefined)[];

function writeRows(wb: WorkbookModel, s: SheetModel, r0: number, c0: number, rows: Row[]): void {
  rows.forEach((row, i) =>
    row.forEach((v, j) => {
      if (v === undefined || v === null) return;
      if (typeof v === "string" && v.startsWith("=")) wb.setInputRaw(s, r0 + i, c0 + j, { formula: v });
      else wb.setInputRaw(s, r0 + i, c0 + j, v);
    }),
  );
}

function style(wb: WorkbookModel, s: SheetModel, r1: number, c1: number, r2: number, c2: number, st: CellStyle): void {
  for (let r = r1; r <= r2; r++)
    for (let c = c1; c <= c2; c++) {
      const id = wb.styleIdAt(s, r, c);
      wb.setStyleRaw(s, r, c, wb.styles.patch(id, st));
    }
}

const TL = '#,##0.00 "₺"';
const thin = { style: "thin" as const, color: "#BFBFBF" };

function budget(): WorkbookModel {
  const wb = new WorkbookModel(true);
  const s = wb.addSheet("Aylık Bütçe");
  s.showGridlines = false;
  s.cols.set(0, 24);
  s.cols.set(1, 200);
  s.cols.set(2, 120);
  s.cols.set(3, 120);
  s.cols.set(4, 120);
  s.cols.set(5, 24);
  s.cols.set(6, 170);
  s.cols.set(7, 130);
  writeRows(wb, s, 1, 1, [["Aylık Bütçe"], ["Ekim 2026 dönemi için gelir ve gider takibi"]]);
  style(wb, s, 1, 1, 1, 1, { size: 24, bold: true, color: "#0E2841" });
  style(wb, s, 2, 1, 2, 1, { italic: true, color: "#595959" });
  s.rows.set(1, 44);
  // Özet kutusu
  writeRows(wb, s, 4, 6, [["Özet"], ["Toplam Gelir", "=SUM(C7:C10)"], ["Toplam Gider", "=SUM(D14:D24)"], ["Bakiye", "=H6-H7"], ["Tasarruf Oranı", "=IFERROR(H8/H6,0)"]]);
  style(wb, s, 4, 6, 4, 7, { bold: true, color: "#FFFFFF", fill: "#156082" });
  style(wb, s, 5, 7, 7, 7, { numFmt: TL });
  style(wb, s, 8, 7, 8, 7, { numFmt: "0.0%" });
  style(wb, s, 7, 6, 7, 7, { bold: true });
  style(wb, s, 5, 6, 8, 7, { bBottom: thin });
  // Gelirler
  writeRows(wb, s, 5, 1, [["Gelirler", "Tutar"], ["Maaş", 42000], ["Ek Gelir", 6500], ["Kira Geliri", 0], ["Diğer", 1200], ["Toplam Gelir", "=SUM(C7:C10)"]]);
  style(wb, s, 5, 1, 5, 2, { bold: true, fill: "#C0E4F5", color: "#0E2841", bBottom: { style: "medium", color: "#156082" } });
  style(wb, s, 6, 2, 10, 2, { numFmt: TL });
  style(wb, s, 10, 1, 10, 2, { bold: true, bTop: { style: "thin", color: "#156082" }, bBottom: { style: "double", color: "#156082" } });
  // Giderler
  writeRows(wb, s, 12, 1, [
    ["Gider Kalemi", "Planlanan", "Gerçekleşen", "Fark"],
    ["Kira", 15000, 15000, "=C14-D14"],
    ["Market", 7000, 7850, "=C15-D15"],
    ["Faturalar (elektrik, su, doğalgaz)", 3200, 2980, "=C16-D16"],
    ["İnternet ve telefon", 900, 900, "=C17-D17"],
    ["Ulaşım", 2500, 2210, "=C18-D18"],
    ["Sağlık", 1000, 450, "=C19-D19"],
    ["Eğitim", 1500, 1500, "=C20-D20"],
    ["Eğlence", 2000, 2650, "=C21-D21"],
    ["Giyim", 1500, 980, "=C22-D22"],
    ["Birikim", 5000, 5000, "=C23-D23"],
    ["Diğer", 800, 620, "=C24-D24"],
    ["Toplam Gider", "=SUM(C14:C24)", "=SUM(D14:D24)", "=SUM(E14:E24)"],
  ]);
  style(wb, s, 12, 1, 12, 4, { bold: true, fill: "#FBE2D5", color: "#0E2841", bBottom: { style: "medium", color: "#E97132" } });
  style(wb, s, 13, 2, 25, 4, { numFmt: TL });
  style(wb, s, 13, 1, 23, 4, { bBottom: thin });
  style(wb, s, 25, 1, 25, 4, { bold: true, bTop: { style: "thin", color: "#E97132" }, bBottom: { style: "double", color: "#E97132" } });
  s.condFormats.push({ id: wb.newId("cf"), type: "cellIs", operator: "lessThan", values: ["=0"], ranges: [{ r1: 13, c1: 4, r2: 23, c2: 4 }], style: { color: "#9C0006", fill: "#FFC7CE" } });
  s.condFormats.push({ id: wb.newId("cf"), type: "dataBar", color: "#E97132", ranges: [{ r1: 13, c1: 3, r2: 23, c2: 3 }] });
  s.charts.push({ id: wb.newId("chart"), type: "doughnut", range: "'Aylık Bütçe'!B13:B24", seriesInRows: false, title: "Gider Dağılımı", showTitle: true, showLegend: true, legendPos: "right", showDataLabels: true, showGridlines: false, palette: 0, style: 1, anchor: { row: 11, col: 6, dx: 0, dy: 0 }, width: 380, height: 260 });
  // Pasta için tutar sütunu: B ve D
  s.charts[0].range = "'Aylık Bütçe'!B13:D24";
  s.charts[0].type = "column";
  s.charts[0].title = "Planlanan ve Gerçekleşen";
  s.charts[0].legendPos = "bottom";
  s.charts[0].showDataLabels = false;
  s.charts[0].showGridlines = true;
  s.frozen = { rows: 0, cols: 0 };
  s.page.orientation = "landscape";
  s.page.fitToWidth = true;
  wb.props.title = "Aylık Bütçe";
  return wb;
}

function invoice(): WorkbookModel {
  const wb = new WorkbookModel(true);
  const s = wb.addSheet("Fatura");
  s.showGridlines = false;
  [24, 90, 230, 70, 110, 110].forEach((w, i) => s.cols.set(i, w));
  writeRows(wb, s, 1, 1, [["FATURA"]]);
  style(wb, s, 1, 1, 1, 1, { size: 26, bold: true, color: "#196B24" });
  s.rows.set(1, 46);
  writeRows(wb, s, 1, 4, [["Fatura No:", "2026-0147"], ["Tarih:", 46307], ["Vade:", "=F3+30"]]);
  style(wb, s, 1, 4, 3, 4, { bold: true, hAlign: "right" });
  style(wb, s, 2, 5, 3, 5, { numFmt: "dd.mm.yyyy", hAlign: "right" });
  style(wb, s, 1, 5, 1, 5, { hAlign: "right" });
  writeRows(wb, s, 5, 1, [["Satıcı", undefined, undefined, "Alıcı"], ["Örnek Yazılım Ltd. Şti.", undefined, undefined, "Deniz Gıda A.Ş."], ["Atatürk Cad. No: 12, Kadıköy / İstanbul", undefined, undefined, "Cumhuriyet Mah. 5. Sk. No: 8, Çankaya / Ankara"], ["Vergi No: 1234567890", undefined, undefined, "Vergi No: 9876543210"]]);
  style(wb, s, 5, 1, 5, 4, { bold: true, color: "#196B24" });
  writeRows(wb, s, 10, 1, [
    ["Kod", "Açıklama", "Miktar", "Birim Fiyat", "Tutar"],
    ["YZ-01", "Yazılım lisansı (yıllık)", 3, 4500, "=D12*E12"],
    ["DS-04", "Kurulum ve eğitim (saat)", 8, 750, "=D13*E13"],
    ["BK-02", "Bakım ve destek paketi", 1, 3200, "=D14*E14"],
    ["DN-07", "Veri taşıma hizmeti", 2, 1250, "=D15*E15"],
    [undefined, undefined, undefined, undefined, undefined],
  ]);
  style(wb, s, 10, 1, 10, 5, { bold: true, fill: "#196B24", color: "#FFFFFF" });
  style(wb, s, 11, 1, 15, 5, { bBottom: thin });
  style(wb, s, 11, 3, 15, 3, { hAlign: "center" });
  style(wb, s, 11, 4, 15, 5, { numFmt: TL });
  writeRows(wb, s, 17, 4, [["Ara Toplam", "=SUM(F12:F16)"], ["KDV (%20)", "=ROUND(F18*20%,2)"], ["Genel Toplam", "=F18+F19"]]);
  style(wb, s, 17, 4, 19, 4, { hAlign: "right" });
  style(wb, s, 17, 5, 19, 5, { numFmt: TL });
  style(wb, s, 19, 4, 19, 5, { bold: true, size: 13, fill: "#DAF2D0", bTop: { style: "medium", color: "#196B24" } });
  writeRows(wb, s, 22, 1, [["Ödeme Bilgileri"], ["IBAN: TR00 0000 0000 0000 0000 0000 00"], ["Lütfen ödeme açıklamasına fatura numarasını yazınız. Teşekkür ederiz."]]);
  style(wb, s, 22, 1, 22, 1, { bold: true });
  style(wb, s, 24, 1, 24, 1, { italic: true, color: "#595959" });
  s.page.margins = { top: 0.6, bottom: 0.6, left: 0.5, right: 0.5, header: 0.3, footer: 0.3 };
  wb.props.title = "Fatura";
  return wb;
}

function todo(): WorkbookModel {
  const wb = new WorkbookModel(true);
  const s = wb.addSheet("Yapılacaklar");
  [30, 300, 110, 110, 110, 90].forEach((w, i) => s.cols.set(i, w));
  writeRows(wb, s, 0, 1, [["Yapılacaklar Listesi"]]);
  style(wb, s, 0, 1, 0, 1, { size: 20, bold: true, color: "#A02B93" });
  s.rows.set(0, 36);
  writeRows(wb, s, 1, 1, [["Tamamlanan:", '=COUNTIF(E4:E30,"Tamamlandı")&" / "&COUNTA(B4:B30)']]);
  style(wb, s, 1, 1, 1, 2, { color: "#595959" });
  const today = 46307;
  writeRows(wb, s, 2, 1, [
    ["Görev", "Öncelik", "Son Tarih", "Durum", "Kalan Gün"],
    ["Proje raporunu hazırla", "Yüksek", today + 2, "Devam Ediyor", "=IF(E4=\"Tamamlandı\",\"\",D4-TODAY())"],
    ["Müşteri toplantısı", "Yüksek", today + 1, "Başlamadı", "=IF(E5=\"Tamamlandı\",\"\",D5-TODAY())"],
    ["Bütçe tablosunu güncelle", "Orta", today + 5, "Tamamlandı", "=IF(E6=\"Tamamlandı\",\"\",D6-TODAY())"],
    ["Ekip değerlendirmeleri", "Orta", today + 9, "Başlamadı", "=IF(E7=\"Tamamlandı\",\"\",D7-TODAY())"],
    ["Sunum taslağını gözden geçir", "Düşük", today + 12, "Devam Ediyor", "=IF(E8=\"Tamamlandı\",\"\",D8-TODAY())"],
    ["Fatura ödemelerini kontrol et", "Yüksek", today - 1, "Tamamlandı", "=IF(E9=\"Tamamlandı\",\"\",D9-TODAY())"],
    ["Yedekleme planını güncelle", "Düşük", today + 20, "Başlamadı", "=IF(E10=\"Tamamlandı\",\"\",D10-TODAY())"],
  ]);
  s.tables.push({ id: wb.newId("tbl"), name: "Görevler", ref: { r1: 2, c1: 1, r2: 9, c2: 5 }, style: "TableStyleMedium5", headerRow: true, totalRow: false, bandedRows: true, bandedCols: false, firstColumn: false, lastColumn: false, filterButton: true, columns: ["Görev", "Öncelik", "Son Tarih", "Durum", "Kalan Gün"], filter: { range: { r1: 2, c1: 1, r2: 9, c2: 5 }, columns: {} } });
  style(wb, s, 3, 3, 9, 3, { numFmt: "dd.mm.yyyy", hAlign: "center" });
  style(wb, s, 3, 2, 9, 2, { hAlign: "center" });
  style(wb, s, 3, 4, 9, 5, { hAlign: "center" });
  s.validations.push({ id: wb.newId("dv"), ranges: [{ r1: 3, c1: 2, r2: 40, c2: 2 }], rule: { type: "list", formula1: "Yüksek;Orta;Düşük" } });
  s.validations.push({ id: wb.newId("dv"), ranges: [{ r1: 3, c1: 4, r2: 40, c2: 4 }], rule: { type: "list", formula1: "Başlamadı;Devam Ediyor;Tamamlandı" } });
  s.condFormats.push({ id: wb.newId("cf"), type: "containsText", text: "Tamamlandı", ranges: [{ r1: 3, c1: 4, r2: 40, c2: 4 }], style: { fill: "#C6EFCE", color: "#006100" } });
  s.condFormats.push({ id: wb.newId("cf"), type: "containsText", text: "Yüksek", ranges: [{ r1: 3, c1: 2, r2: 40, c2: 2 }], style: { color: "#C00000", bold: true } });
  s.condFormats.push({ id: wb.newId("cf"), type: "cellIs", operator: "lessThan", values: ["=0"], ranges: [{ r1: 3, c1: 5, r2: 40, c2: 5 }], style: { fill: "#FFC7CE", color: "#9C0006" } });
  s.frozen = { rows: 3, cols: 0 };
  wb.props.title = "Yapılacaklar Listesi";
  return wb;
}

function sales(): WorkbookModel {
  const wb = new WorkbookModel(true);
  const s = wb.addSheet("Satış Raporu");
  [120, 95, 95, 95, 95, 110, 90].forEach((w, i) => s.cols.set(i, w));
  writeRows(wb, s, 0, 0, [["2026 Bölgesel Satış Raporu"]]);
  style(wb, s, 0, 0, 0, 0, { size: 18, bold: true, color: "#0E2841" });
  s.rows.set(0, 32);
  writeRows(wb, s, 2, 0, [
    ["Bölge", "1. Çeyrek", "2. Çeyrek", "3. Çeyrek", "4. Çeyrek", "Toplam", "Pay"],
    ["Marmara", 482000, 515000, 498000, 610000, "=SUM(B4:E4)", "=F4/$F$10"],
    ["Ege", 296000, 312000, 344000, 365000, "=SUM(B5:E5)", "=F5/$F$10"],
    ["İç Anadolu", 351000, 338000, 362000, 401000, "=SUM(B6:E6)", "=F6/$F$10"],
    ["Akdeniz", 214000, 262000, 305000, 278000, "=SUM(B7:E7)", "=F7/$F$10"],
    ["Karadeniz", 158000, 171000, 189000, 176000, "=SUM(B8:E8)", "=F8/$F$10"],
    ["Doğu ve Güneydoğu", 132000, 141000, 150000, 163000, "=SUM(B9:E9)", "=F9/$F$10"],
    ["Toplam", "=SUM(B4:B9)", "=SUM(C4:C9)", "=SUM(D4:D9)", "=SUM(E4:E9)", "=SUM(F4:F9)", "=SUM(G4:G9)"],
  ]);
  style(wb, s, 2, 0, 2, 6, { bold: true, fill: "#156082", color: "#FFFFFF", hAlign: "center" });
  style(wb, s, 3, 1, 9, 5, { numFmt: "#,##0 \"₺\"" });
  style(wb, s, 3, 6, 9, 6, { numFmt: "0.0%" });
  style(wb, s, 3, 0, 8, 6, { bBottom: thin });
  style(wb, s, 9, 0, 9, 6, { bold: true, fill: "#C0E4F5", bTop: { style: "medium", color: "#156082" } });
  s.condFormats.push({ id: wb.newId("cf"), type: "colorScale", colors: ["#F8696B", "#FFEB84", "#63BE7B"], ranges: [{ r1: 3, c1: 1, r2: 8, c2: 4 }] });
  writeRows(wb, s, 11, 0, [["En yüksek çeyrek satışı:", undefined, "=MAX(B4:E9)"], ["Ortalama bölge satışı:", undefined, "=AVERAGE(F4:F9)"], ["En iyi bölge:", undefined, "=INDEX(A4:A9,MATCH(MAX(F4:F9),F4:F9,0))"]]);
  style(wb, s, 11, 2, 12, 2, { numFmt: "#,##0 \"₺\"", bold: true });
  style(wb, s, 13, 2, 13, 2, { bold: true });
  s.charts.push({ id: wb.newId("chart"), type: "column", range: "'Satış Raporu'!A3:E9", seriesInRows: false, title: "Çeyreklik Satışlar", showTitle: true, showLegend: true, legendPos: "bottom", showDataLabels: false, showGridlines: true, palette: 0, style: 1, anchor: { row: 15, col: 0, dx: 4, dy: 0 }, width: 470, height: 270 });
  s.charts.push({ id: wb.newId("chart"), type: "pie", range: "'Satış Raporu'!A3:A9", seriesInRows: false, title: "Bölgelere Göre Pay", showTitle: true, showLegend: true, legendPos: "right", showDataLabels: true, showGridlines: false, palette: 0, style: 1, anchor: { row: 15, col: 4, dx: 40, dy: 0 }, width: 360, height: 270 });
  s.charts[1].range = "'Satış Raporu'!A3:A9";
  // Pasta: bölge adları + toplam sütunu için ayrı yardımcı aralık
  writeRows(wb, s, 2, 8, [["Bölge", "Toplam"]]);
  for (let i = 0; i < 6; i++) writeRows(wb, s, 3 + i, 8, [[`=A${4 + i}`, `=F${4 + i}`]]);
  s.cols.setHidden(8, true);
  s.cols.setHidden(9, true);
  s.charts[1].range = "'Satış Raporu'!I3:J9";
  s.frozen = { rows: 3, cols: 1 };
  s.page.orientation = "landscape";
  s.page.fitToWidth = true;
  s.page.footer = "Sayfa &[Sayfa] / &[Sayfalar]";
  wb.props.title = "Satış Raporu";
  return wb;
}

export const TEMPLATES: TemplateDef[] = [
  { id: "blank", label: "Boş çalışma kitabı", build: () => new WorkbookModel() },
  { id: "budget", label: "Aylık Bütçe", build: budget },
  { id: "invoice", label: "Fatura", build: invoice },
  { id: "todo", label: "Yapılacaklar Listesi", build: todo },
  { id: "sales", label: "Satış Raporu", build: sales },
];

export function buildTemplate(id: string): WorkbookModel {
  const t = TEMPLATES.find((x) => x.id === id) ?? TEMPLATES[0];
  const wb = t.build();
  wb.activeIndex = 0;
  wb.engine.recalc(true);
  void ACCOUNTING_FORMAT;
  return wb;
}
