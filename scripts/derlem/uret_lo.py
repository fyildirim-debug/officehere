#!/usr/bin/env python3
"""LibreOffice UNO ile gerçek dünya derlemi üretir (bizim içeriğimiz; üçüncü taraf belgesi yoktur).

Writer (docx/odt/rtf/doc/html), Calc (xlsx/ods/xls/csv), Impress (pptx/odp/ppt) belgeleri;
çok sayfa, tablo, resim, üst/alt bilgi, dipnot, liste, değişiklik izleme, formül, biçim, grafik,
birleştirilmiş hücre, şekil, not, geçiş içerir.

Ön koşul: LibreOffice UNO dinleyicisi çalışıyor olmalı (scripts/derlem/uret.mjs bunu kendisi başlatır):
  soffice --headless --accept="socket,host=127.0.0.1,port=<LO_PORT>;urp;"
Kullanım: LO_PORT=2002 python3 scripts/derlem/uret_lo.py <çıktı_dizini>
Çıktı: <dizin>/*.{docx,...} ve <dizin>/envanter.json (hangi öğe başarıyla eklendi).
Görseller: DERLEM_GORSEL (varsayılan test-derlem/gorsel; uret_gorsel.py üretir).
"""
import json
import os
import sys
import traceback

import uno
from com.sun.star.awt import Point, Rectangle, Size
from com.sun.star.beans import PropertyValue
from com.sun.star.lang import Locale
from com.sun.star.table import BorderLine2, CellAddress
from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK, LINE_BREAK

KOK = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(KOK, ".derlem", "uretilen", "lo")
IMG = os.environ.get("DERLEM_GORSEL", os.path.join(KOK, "test-derlem", "gorsel"))
PORT = int(os.environ.get("LO_PORT", "2002"))
os.makedirs(OUT, exist_ok=True)

local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
ctx = resolver.resolve(f"uno:socket,host=127.0.0.1,port={PORT};urp;StarOffice.ComponentContext")
smgr = ctx.ServiceManager
desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
gp = smgr.createInstanceWithContext("com.sun.star.graphic.GraphicProvider", ctx)

ENV = {}  # dosya -> {öğe: True/hata}


def P(name, value):
    pv = PropertyValue()
    pv.Name = name
    pv.Value = value
    return pv


def url(p):
    return uno.systemPathToFileUrl(p)


def enum(t, v):
    return uno.Enum(t, v)


def graphic(path):
    return gp.queryGraphic((P("URL", url(path)),))


def feat(doc_key, name, fn):
    try:
        fn()
        ENV.setdefault(doc_key, {})[name] = True
    except Exception as e:  # noqa
        ENV.setdefault(doc_key, {})[name] = f"HATA: {e}"
        traceback.print_exc()


def kapat(doc):
    # LibreOffice bazı Impress belgelerini kapatırken çökebiliyor (dosyalar zaten yazılmış olur).
    try:
        doc.close(True)
    except Exception as e:  # noqa: BLE001
        print("uyarı: belge kapatılırken LibreOffice hatası:", e)


def store(doc, base, variants):
    for v in variants:
        ext, flt = v[0], v[1]
        opts = v[2] if len(v) > 2 else None
        path = os.path.join(OUT, f"{base}.{ext}")
        try:
            args = [P("FilterName", flt), P("Overwrite", True)]
            if opts:
                args.append(P("FilterOptions", opts))
            doc.storeToURL(url(path), tuple(args))
            print("yazıldı", path, os.path.getsize(path))
        except Exception as e:
            print("YAZILAMADI", path, e)
            ENV.setdefault(base, {})[f"kaydet:{ext}"] = f"HATA: {e}"


LOREM = (
    "OfficeHere ölçüm derlemi için üretilmiş bu paragraf, Türkçe karakterleri (ığüşöç İĞÜŞÖÇ) ve "
    "uzun cümleleri içerir. Amaç; okuyucuların metni, biçimleri ve yapıyı kayıpsız aktarıp aktarmadığını "
    "ölçmektir. Sayılar 1.234,56 ve tarihler 15.03.2024 biçiminde yazılmıştır."
)

# ------------------------------------------------------------------ Writer


def writer_doc(base, variant):
    doc = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, (P("Hidden", True),))
    text = doc.Text
    cur = text.createTextCursor()
    key = base

    def para(s="", style="Standard", align=None, brk=False):
        cur.ParaStyleName = style
        if align is not None:
            cur.ParaAdjust = enum("com.sun.star.style.ParagraphAdjust", align)
        if brk:
            cur.BreakType = enum("com.sun.star.style.BreakType", "PAGE_BEFORE")
        if s:
            text.insertString(cur, s, False)
        text.insertControlCharacter(cur, PARAGRAPH_BREAK, False)
        cur.ParaStyleName = "Standard"
        cur.setPropertyToDefault("ParaAdjust")
        cur.BreakType = enum("com.sun.star.style.BreakType", "NONE")

    def run(s, **props):
        names = list(props.keys())
        for k, v in props.items():
            cur.setPropertyValue(k, v)
        text.insertString(cur, s, False)
        for k in names:
            cur.setPropertyToDefault(k)

    def header_footer():
        ps = doc.StyleFamilies.getByName("PageStyles").getByName("Standard")
        ps.HeaderIsOn = True
        ps.HeaderText.setString(f"OfficeHere Derlem — {base}")
        ps.FooterIsOn = True
        ft = ps.FooterText
        fc = ft.createTextCursor()
        ft.insertString(fc, "Sayfa ", False)
        pn = doc.createInstance("com.sun.star.text.TextField.PageNumber")
        pn.NumberingType = 4
        pn.SubType = enum("com.sun.star.text.PageNumberType", "CURRENT")
        ft.insertTextContent(fc, pn, False)
        ft.insertString(fc, " / ", False)
        pc = doc.createInstance("com.sun.star.text.TextField.PageCount")
        pc.NumberingType = 4
        ft.insertTextContent(fc, pc, False)

    def table(rows, cols, title, merge=True):
        t = doc.createInstance("com.sun.star.text.TextTable")
        t.initialize(rows, cols)
        text.insertTextContent(cur, t, False)
        t.RepeatHeadline = True
        names = t.getCellNames()
        for i, n in enumerate(names):
            c = t.getCellByName(n)
            r = int("".join(ch for ch in n if ch.isdigit()))
            if r == 1:
                cc = c.createTextCursor()
                cc.CharWeight = 150.0
                c.insertString(cc, f"{title} {n}", False)
                c.BackColor = 0x4472C4
            else:
                c.setString(f"Hücre {n} değer {i * 3}")
        if merge and rows >= 3 and cols >= 2:
            tc = t.createCursorByCellName(f"A{rows}")
            tc.gotoCellByName(f"B{rows}", True)
            tc.mergeRange()
        return t

    def image(name, w, h):
        g = doc.createInstance("com.sun.star.text.TextGraphicObject")
        g.Graphic = graphic(os.path.join(IMG, name))
        g.AnchorType = enum("com.sun.star.text.TextContentAnchorType", "AS_CHARACTER")
        g.Size = Size(w, h)
        text.insertTextContent(cur, g, False)

    def footnote(s, kind="Footnote"):
        fn = doc.createInstance(f"com.sun.star.text.{kind}")
        text.insertTextContent(cur, fn, False)
        fn.setString(s)

    def comment(s):
        a = doc.createInstance("com.sun.star.text.TextField.Annotation")
        a.Author = "Ayşe Yılmaz"
        a.Content = s
        text.insertTextContent(cur, a, False)

    def blist(items, style="List 1"):
        for lvl, s in items:
            cur.NumberingStyleName = style
            cur.NumberingLevel = lvl
            text.insertString(cur, s, False)
            text.insertControlCharacter(cur, PARAGRAPH_BREAK, False)
        cur.NumberingStyleName = ""
        cur.setPropertyToDefault("NumberingLevel")

    feat(key, "üst/alt bilgi + sayfa alanı", header_footer)
    para("Gerçek Dünya Derlem Belgesi", "Title")
    para(f"Varyant: {variant}", "Subtitle")

    if variant in ("rapor", "uzun"):
        # içindekiler tablosu (bölüm başlıklarından)
        try:
            idx = doc.createInstance("com.sun.star.text.ContentIndex")
            idx.CreateFromOutline = True
            text.insertTextContent(cur, idx, False)
            ENV.setdefault(key, {})["içindekiler"] = True
        except Exception as e:
            idx = None
            ENV.setdefault(key, {})["içindekiler"] = f"HATA: {e}"

    sections = 8 if variant == "rapor" else 30 if variant == "uzun" else 3
    for s in range(1, sections + 1):
        para(f"Bölüm {s}: Ölçüm Başlığı", "Heading 1", brk=s > 1)
        para(LOREM)
        # biçimli koşular + dipnot + bağlantı
        run("Kalın metin", CharWeight=150.0)
        run(", ")
        run("italik metin", CharPosture=enum("com.sun.star.awt.FontSlant", "ITALIC"))
        run(", ")
        run("altı çizili", CharUnderline=1)
        run(", ")
        run("üstü çizili", CharStrikeout=1)
        run(", ")
        run("kırmızı Arial 14", CharColor=0xC00000, CharFontName="Arial", CharHeight=14.0)
        run(", ")
        run("vurgulu", CharBackColor=0xFFFF00)
        run(", H")
        run("2", CharEscapement=-33, CharEscapementHeight=58)
        run("O ve x")
        run("2", CharEscapement=33, CharEscapementHeight=58)
        run(" ve ")
        run("bağlantı", HyperLinkURL="https://ornek.com.tr/olcum")
        if s % 2 == 1:
            feat(key, "dipnot", lambda: footnote(f"Bölüm {s} dipnotu: kaynak açıklaması."))
        text.insertControlCharacter(cur, PARAGRAPH_BREAK, False)
        para(LOREM + " Sağa hizalı.", align="RIGHT")
        para(LOREM + " Ortalı.", align="CENTER")
        para((LOREM + " ") * 2 + "İki yana yaslı.", align="BLOCK")
        para(f"Alt başlık {s}.1", "Heading 2")
        feat(key, "madde işaretli liste", lambda: blist([(0, "Birinci madde"), (1, "İç madde a"), (1, "İç madde b"), (0, "İkinci madde")]))
        feat(key, "numaralı liste", lambda: blist([(0, "Bir"), (0, "İki"), (1, "İki-a"), (0, "Üç")], "Numbering 123"))
        if s % 2 == 0 or variant == "uzun":
            feat(key, "tablo + birleştirme", lambda: table(5, 4, f"T{s}"))
            para("")
        if s in (3, 6) or (variant == "karma" and s == 2):
            feat(key, "resim", lambda: image("grafik.png", 8000, 5000))
            para(" Şekil: örnek grafik resmi.")
        if s == 4 or (variant == "karma" and s == 1):
            feat(key, "resim (jpg)", lambda: image("foto.jpg", 10000, 7500))
            para("")
        if s == 2 or (variant == "karma" and s == 1):
            feat(key, "yorum", lambda: comment(f"Bölüm {s} için gözden geçirme notu."))
            para("Bu paragrafta bir yorum var. Ayrıca silinecek kelime ve kalınlaşacak ifade burada.")
        para("Satır sonu öncesi" if s == 1 else LOREM)
        if s == 1:
            text.insertString(cur, "Satır", False)
            text.insertControlCharacter(cur, LINE_BREAK, False)
            text.insertString(cur, "sonu\tsekme sonrası", False)
            text.insertControlCharacter(cur, PARAGRAPH_BREAK, False)

    if variant == "karma":
        def frame():
            f = doc.createInstance("com.sun.star.text.TextFrame")
            f.Size = Size(6000, 3000)
            f.AnchorType = enum("com.sun.star.text.TextContentAnchorType", "AT_PARAGRAPH")
            text.insertTextContent(cur, f, False)
            f.getText().setString("Metin çerçevesi içindeki yazı.")
        feat(key, "metin çerçevesi", frame)
        para("Çerçeveden sonraki paragraf.")

        def shape():
            sh = doc.createInstance("com.sun.star.drawing.RectangleShape")
            sh.AnchorType = enum("com.sun.star.text.TextContentAnchorType", "AT_PARAGRAPH")
            sh.Size = Size(5000, 2500)
            text.insertTextContent(cur, sh, False)
            sh.FillColor = 0x70AD47
            sh.setString("Şekil metni")
        feat(key, "çizim şekli", shape)
        para("Şekilden sonraki paragraf.")

        def cols():
            start = text.createTextCursorByRange(cur.getStart())
            for i in range(4):
                para(f"İki sütunlu bölümdeki paragraf {i + 1}. " + LOREM)
            rng = text.createTextCursorByRange(start.getStart())
            rng.gotoRange(cur.getStart(), True)
            sec = doc.createInstance("com.sun.star.text.TextSection")
            tcols = doc.createInstance("com.sun.star.text.TextColumns")
            tcols.ColumnCount = 2
            sec.TextColumns = tcols
            text.insertTextContent(rng, sec, True)
        feat(key, "iki sütunlu bölüm", cols)

        def nested():
            t = table(3, 3, "Dış", merge=False)
            c = t.getCellByName("B2")
            inner = doc.createInstance("com.sun.star.text.TextTable")
            inner.initialize(2, 2)
            c.insertTextContent(c.createTextCursor(), inner, False)
            inner.getCellByName("A1").setString("iç tablo")
        feat(key, "iç içe tablo", nested)
        para("")

        def landscape():
            para("Yatay sayfadaki metin.")
            cur.PageDescName = "Landscape"
            para("Yatay sayfa başlangıcı. " + LOREM)
        feat(key, "yatay sayfa stili", landscape)
        feat(key, "sonnot", lambda: footnote("Sonnot metni.", "Endnote"))
        para("")

    # değişiklik izleme: ekleme + silme + biçim değişikliği
    def track():
        doc.RecordChanges = True
        para("Bu paragraf değişiklik izleme açıkken eklendi (ekleme).")
        sd = doc.createSearchDescriptor()
        sd.SearchString = "silinecek kelime"
        f = doc.findFirst(sd)
        if f is not None:
            f.setString("")
        sd.SearchString = "kalınlaşacak ifade"
        f = doc.findFirst(sd)
        if f is not None:
            f.CharWeight = 150.0
        doc.RecordChanges = False
    if variant in ("rapor", "karma"):
        feat(key, "değişiklik izleme", track)
    try:
        doc.DocumentProperties.Title = f"Derlem {base}"
        doc.DocumentProperties.Author = "Ölçüm Betiği"
        doc.DocumentProperties.Subject = "Kayıt-aç turu"
    except Exception:
        pass
    try:
        idxs = doc.getDocumentIndexes()
        for i in range(idxs.getCount()):
            idxs.getByIndex(i).update()
    except Exception:
        pass
    doc.refresh() if hasattr(doc, "refresh") else None
    store(doc, base, [("odt", "writer8"), ("docx", "MS Word 2007 XML"), ("rtf", "Rich Text Format"), ("doc", "MS Word 97"), ("html", "HTML (StarWriter)")])
    kapat(doc)


# ------------------------------------------------------------------ Calc


def calc_doc(base, variant):
    doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, (P("Hidden", True),))
    key = base
    sheets = doc.Sheets
    sh = sheets.getByIndex(0)
    sh.Name = "Satışlar"
    loc = Locale("en", "US", "")
    nfs = doc.NumberFormats

    def nf(code):
        k = nfs.queryKey(code, loc, False)
        if k == -1:
            k = nfs.addNew(code, loc)
        return k

    def border(rng):
        bl = BorderLine2()
        bl.OuterLineWidth = 18
        bl.LineWidth = 18
        bl.Color = 0x404040
        rng.TopBorder = bl
        rng.BottomBorder = bl
        rng.LeftBorder = bl
        rng.RightBorder = bl

    products = ["Çay", "Şeker", "Ğofret", "Kahve", "Süt", "Peynir", "Zeytin", "Ekmek", "Bal", "Un"]
    cities = ["İstanbul", "Ankara", "İzmir", "Bursa", "Çanakkale"]
    N = 50 if variant == "butce" else 3000
    # başlık (birleştirilmiş)
    sh.getCellByPosition(0, 0).setString("2024 Satış Raporu — İĞÜŞÖÇ")
    feat(key, "birleştirilmiş başlık", lambda: sh.getCellRangeByName("A1:H1").merge(True))
    t = sh.getCellByPosition(0, 0)
    t.CharWeight = 150.0
    t.CharHeight = 16.0
    t.HoriJustify = enum("com.sun.star.table.CellHoriJustify", "CENTER")
    heads = ["Ürün", "Şehir", "Adet", "Fiyat", "Tarih", "Tutar", "KDV", "Durum"]
    for c, h in enumerate(heads):
        cell = sh.getCellByPosition(c, 1)
        cell.setString(h)
        cell.CharWeight = 150.0
        cell.CellBackColor = 0x1F4E79
        cell.CharColor = 0xFFFFFF
    for r in range(N):
        row = r + 2
        sh.getCellByPosition(0, row).setString(products[r % len(products)])
        sh.getCellByPosition(1, row).setString(cities[r % len(cities)])
        sh.getCellByPosition(2, row).setValue((r * 7) % 50 + 1)
        sh.getCellByPosition(3, row).setValue(round(5 + (r % 13) * 2.75, 2))
        sh.getCellByPosition(4, row).setFormula(f"=DATE(2024;{r % 12 + 1};{r % 28 + 1})")
        sh.getCellByPosition(5, row).setFormula(f"=C{row + 1}*D{row + 1}")
        sh.getCellByPosition(6, row).setFormula(f"=ROUND(F{row + 1}*0.2;2)")
        sh.getCellByPosition(7, row).setFormula(f'=IF(F{row + 1}>500;"Yüksek";IF(F{row + 1}>100;"Orta";"Düşük"))')
    last = N + 2
    tot = last + 1
    sh.getCellByPosition(0, tot - 1).setString("Toplam")
    sh.getCellByPosition(2, tot - 1).setFormula(f"=SUM(C3:C{last})")
    sh.getCellByPosition(5, tot - 1).setFormula(f"=SUM(F3:F{last})")
    sh.getCellByPosition(6, tot - 1).setFormula(f"=SUBTOTAL(9;G3:G{last})")
    sh.getCellByPosition(0, tot).setString("Ortalama")
    sh.getCellByPosition(5, tot).setFormula(f"=AVERAGE(F3:F{last})")
    sh.getCellByPosition(0, tot + 1).setString("Çay adedi")
    sh.getCellByPosition(5, tot + 1).setFormula(f'=SUMIF(A3:A{last};"Çay";C3:C{last})')
    sh.getCellByPosition(0, tot + 2).setString("Yüksek sayısı")
    sh.getCellByPosition(5, tot + 2).setFormula(f'=COUNTIF(H3:H{last};"Yüksek")')
    sh.getCellByPosition(0, tot + 3).setString("En büyük")
    sh.getCellByPosition(5, tot + 3).setFormula(f"=MAX(F3:F{last})")
    # biçimler
    def fmts():
        sh.getCellRangeByName(f"D3:D{last}").NumberFormat = nf('#,##0.00 "TL"')
        sh.getCellRangeByName(f"E3:E{last}").NumberFormat = nf("DD.MM.YYYY")
        sh.getCellRangeByName(f"F3:G{tot + 3}").NumberFormat = nf("#,##0.00")
    feat(key, "sayı/tarih biçimleri", fmts)
    feat(key, "kenarlıklar", lambda: border(sh.getCellRangeByName(f"A2:H{last}")))
    def tot_style():
        r = sh.getCellRangeByName(f"A{tot}:H{tot}")
        r.CharWeight = 150.0
        r.CellBackColor = 0xDDEBF7
    feat(key, "toplam satırı biçimi", tot_style)
    def widths():
        for c, w in enumerate([3200, 3000, 1800, 2600, 2600, 2800, 2400, 2200]):
            sh.Columns.getByIndex(c).Width = w
        sh.Rows.getByIndex(0).Height = 900
    feat(key, "sütun genişlikleri", widths)
    feat(key, "yorum", lambda: sh.Annotations.insertNew(CellAddress(0, 0, 1), "Ürün adı (açıklama)"))
    def wrap_rotate():
        c = sh.getCellByPosition(9, 1)
        c.setString("Döndürülmüş ve kaydırılmış uzun başlık metni")
        c.IsTextWrapped = True
        c.RotateAngle = 4500
    feat(key, "kaydırma/döndürme", wrap_rotate)
    def cond():
        from com.sun.star.sheet.ConditionOperator import GREATER
        rng = sh.getCellRangeByName(f"F3:F{last}")
        cf = rng.ConditionalFormat
        cf.addNew((P("Operator", GREATER), P("Formula1", "500"), P("StyleName", "Accent 3")))
        rng.ConditionalFormat = cf
    feat(key, "koşullu biçim", cond)
    def valid():
        rng = sh.getCellRangeByName(f"I3:I{min(last, 30)}")
        v = rng.Validation
        v.Type = enum("com.sun.star.sheet.ValidationType", "LIST")
        v.Formula1 = '"Evet";"Hayır"'
        rng.Validation = v
    feat(key, "veri doğrulama", valid)
    feat(key, "gizli sütun", lambda: setattr(sh.Columns.getByIndex(10), "IsVisible", False))
    def freeze():
        doc.CurrentController.setActiveSheet(sh)
        doc.CurrentController.freezeAtPosition(1, 2)
    feat(key, "bölmeleri dondur", freeze)
    def autofilter():
        doc.DatabaseRanges.addNewByName("Veri", sh.getCellRangeByName(f"A2:H{last}").RangeAddress)
        doc.DatabaseRanges.getByName("Veri").AutoFilter = True
    feat(key, "otomatik filtre", autofilter)
    def named():
        doc.NamedRanges.addNewByName("Fiyatlar", f"$Satışlar.$D$3:$D${last}", CellAddress(0, 0, 0), 0)
    feat(key, "adlı aralık", named)
    def chart(sheet, name, rng, kind, x, y):
        rect = Rectangle()
        rect.X, rect.Y, rect.Width, rect.Height = x, y, 12000, 7500
        sheet.Charts.addNewByName(name, rect, (sheet.getCellRangeByName(rng).RangeAddress,), True, True)
        emb = sheet.Charts.getByName(name).EmbeddedObject
        emb.Diagram = emb.createInstance(f"com.sun.star.chart.{kind}")
        emb.HasMainTitle = True
        emb.Title.String = f"{name} başlığı"

    # Özet sayfası
    sheets.insertNewByName("Özet", 1)
    o = sheets.getByName("Özet")
    o.getCellByPosition(0, 0).setString("Genel toplam")
    o.getCellByPosition(1, 0).setFormula(f"=$Satışlar.F{tot}")
    o.getCellByPosition(0, 1).setString("KDV dahil")
    o.getCellByPosition(1, 1).setFormula("=B1*1.2")
    o.getCellByPosition(0, 2).setString("Ürün araması")
    o.getCellByPosition(1, 2).setFormula(f'=VLOOKUP("Kahve";$Satışlar.A3:F{last};6;0)')
    o.getCellByPosition(0, 3).setString("INDEX/MATCH")
    o.getCellByPosition(1, 3).setFormula(f'=INDEX($Satışlar.B3:B{last};MATCH("Süt";$Satışlar.A3:A{last};0))')
    o.getCellByPosition(0, 4).setString("Metin")
    o.getCellByPosition(1, 4).setFormula('=CONCATENATE("Toplam: ";TEXT(B1;"0.00"))')
    o.getCellByPosition(0, 5).setString("Büyük harf")
    o.getCellByPosition(1, 5).setFormula('=UPPER("istanbul")')
    o.getCellByPosition(0, 6).setString("Tarih farkı")
    o.getCellByPosition(1, 6).setFormula("=DATE(2024;12;31)-DATE(2024;1;1)")
    o.getCellByPosition(0, 7).setString("SUMPRODUCT")
    o.getCellByPosition(1, 7).setFormula(f"=SUMPRODUCT($Satışlar.C3:C{min(last, 22)};$Satışlar.D3:D{min(last, 22)})")
    o.getCellByPosition(0, 8).setString("IFERROR")
    o.getCellByPosition(1, 8).setFormula("=IFERROR(1/0;-1)")
    o.getCellByPosition(0, 9).setString("Ortalama (adlı)")
    o.getCellByPosition(1, 9).setFormula("=AVERAGE(Fiyatlar)")
    o.getCellByPosition(0, 10).setString("Yuvarla")
    o.getCellByPosition(1, 10).setFormula("=ROUNDUP(PI()*1000;2)")
    o.getCellByPosition(0, 11).setString("Mantık")
    o.getCellByPosition(1, 11).setFormula("=AND(B1>0;OR(B2<0;TRUE()))")
    o.getCellByPosition(0, 12).setString("Ay adı")
    o.getCellByPosition(1, 12).setFormula("=MONTH(DATE(2024;7;4))")
    o.getCellByPosition(0, 13).setString("Bölü sıfır")
    o.getCellByPosition(1, 13).setFormula("=1/0")
    o.getCellByPosition(0, 14).setString("Bulunamadı")
    o.getCellByPosition(1, 14).setFormula('=VLOOKUP("Yok";$Satışlar.A3:F10;2;0)')
    o.getCellByPosition(0, 15).setString("Std sapma")
    o.getCellByPosition(1, 15).setFormula(f"=STDEV($Satışlar.F3:F{last})")
    o.getCellByPosition(0, 16).setString("Medyan")
    o.getCellByPosition(1, 16).setFormula(f"=MEDIAN($Satışlar.C3:C{last})")
    o.getCellByPosition(0, 17).setString("Mantıksal değer")
    o.getCellByPosition(1, 17).setValue(1)
    o.getCellByPosition(1, 17).setFormula("=TRUE()")
    for r in range(18):
        o.getCellByPosition(1, r).NumberFormat = nf("General")
    feat(key, "sütun grafiği", lambda: chart(sh, "Grafik1", f"C2:D{min(last, 12)}", "BarDiagram", 20000, 1000))
    feat(key, "pasta grafiği", lambda: chart(o, "Pasta", "A1:B3", "PieDiagram", 8000, 1000))
    feat(key, "çizgi grafiği", lambda: chart(sh, "Cizgi", f"F2:G{min(last, 20)}", "LineDiagram", 20000, 10000))
    # Zincir sayfası: 200 hücrelik formül zinciri
    sheets.insertNewByName("Zincir", 2)
    z = sheets.getByName("Zincir")
    z.getCellByPosition(0, 0).setValue(1)
    for r in range(1, 200):
        z.getCellByPosition(0, r).setFormula(f"=A{r}*1.01+1")
    z.getCellByPosition(1, 0).setFormula("=A200")
    # birleştirilmiş bloklar
    def merges():
        z.getCellRangeByName("C2:E4").merge(True)
        z.getCellByPosition(2, 1).setString("3x3 birleşik")
        z.getCellRangeByName("C6:C9").merge(True)
        z.getCellByPosition(2, 5).setString("Dikey birleşik")
    feat(key, "çoklu birleştirme", merges)
    if variant == "karma":
        sheets.insertNewByName("Gizli", 3)
        g = sheets.getByName("Gizli")
        g.getCellByPosition(0, 0).setString("gizli sayfa verisi")
        feat(key, "gizli sayfa", lambda: setattr(g, "IsVisible", False))
    try:
        doc.DocumentProperties.Title = f"Derlem {base}"
    except Exception:
        pass
    doc.calculateAll()
    store(doc, base, [("ods", "calc8"), ("xlsx", "Calc MS Excel 2007 XML"), ("xls", "MS Excel 97")])
    # CSV: yalnız etkin sayfa (Satışlar); ; ayırıcı, " metin sınırlayıcı, UTF-8 (76)
    doc.CurrentController.setActiveSheet(sh)
    store(doc, base, [("csv", "Text - txt - csv (StarCalc)", "59,34,76,1")])
    kapat(doc)


# ------------------------------------------------------------------ Impress


def impress_doc(base, variant):
    doc = desktop.loadComponentFromURL("private:factory/simpress", "_blank", 0, (P("Hidden", True),))
    key = base
    pages = doc.DrawPages
    n = 8 if variant == "sunu" else 24

    def placeholders(page):
        out = {}
        for i in range(page.getCount()):
            s = page.getByIndex(i)
            out.setdefault(s.ShapeType, []).append(s)
        return out

    def notes(page, s):
        np_ = page.NotesPage
        for i in range(np_.getCount()):
            sh = np_.getByIndex(i)
            if sh.ShapeType == "com.sun.star.presentation.NotesShape":
                sh.setString(s)
                return
        raise RuntimeError("not şekli yok")

    def add(page, kind, x, y, w, h, **props):
        s = doc.createInstance(f"com.sun.star.drawing.{kind}")
        page.add(s)
        s.Position = Point(x, y)
        s.Size = Size(w, h)
        for k, v in props.items():
            if k == "text":
                continue
            s.setPropertyValue(k, v)
        if "text" in props:
            s.setString(props["text"])
        return s

    effects = ["FADE_FROM_LEFT", "FADE_FROM_TOP", "UNCOVER_TO_RIGHT", "DISSOLVE", "CLOSE_VERTICAL", "SPIRALIN_LEFT", "STRETCH_FROM_LEFT", "ROLL_FROM_TOP"]
    for i in range(n):
        page = pages.getByIndex(0) if i == 0 else pages.insertNewByIndex(i - 1)
        kind = i % 8
        page.Layout = 0 if i == 0 else (1 if kind in (1, 6) else 19 if kind in (2, 3, 4, 7) else 3 if kind == 5 else 1)
        ph = placeholders(page)
        titles = ph.get("com.sun.star.presentation.TitleTextShape", [])
        if titles:
            titles[0].setString(f"Slayt {i + 1}: {'Kapak' if i == 0 else 'İçerik başlığı ığüşöç'}")
        subs = ph.get("com.sun.star.presentation.SubtitleShape", [])
        if subs:
            subs[0].setString("Alt başlık — OfficeHere ölçüm derlemi")
        outl = ph.get("com.sun.star.presentation.OutlinerShape", [])
        for oi, ol in enumerate(outl):
            ol.setString("Birinci madde\nİç madde (düzey 2)\nİkinci madde\nÜçüncü madde: Türkçe İĞÜŞÖÇ")
            try:
                en = ol.getText().createEnumeration()
                j = 0
                while en.hasMoreElements():
                    p = en.nextElement()
                    if j == 1:
                        p.NumberingLevel = 1
                    j += 1
            except Exception:
                pass
        if kind == 2:
            feat(key, "şekiller (dikdörtgen/elips/metin)", lambda: (
                add(page, "RectangleShape", 1500, 4000, 6000, 3500, FillColor=0x4472C4, LineColor=0x203864, text="Dikdörtgen"),
                add(page, "EllipseShape", 9000, 4000, 5000, 3500, FillColor=0xED7D31, text="Elips"),
                add(page, "TextShape", 15500, 4000, 8000, 2000, text="Serbest metin kutusu"),
            ))
            def line():
                l = doc.createInstance("com.sun.star.drawing.LineShape")
                page.add(l)
                l.Position = Point(1500, 9500)
                l.Size = Size(12000, 2000)
                l.LineWidth = 80
                l.LineColor = 0xC00000
                l.LineEndName = "Arrow"
            feat(key, "ok çizgisi", line)
        if kind == 3:
            def table():
                t = doc.createInstance("com.sun.star.drawing.TableShape")
                page.add(t)
                t.Position = Point(1500, 4000)
                t.Size = Size(20000, 6000)
                m = t.Model
                m.Columns.insertByIndex(m.Columns.getCount(), 4 - m.Columns.getCount())
                m.Rows.insertByIndex(m.Rows.getCount(), 4 - m.Rows.getCount())
                for r in range(4):
                    for c in range(4):
                        m.getCellByPosition(c, r).setString("Başlık" if r == 0 else f"R{r}C{c}")
            feat(key, "tablo", table)
        if kind == 4:
            def pic():
                g = doc.createInstance("com.sun.star.drawing.GraphicObjectShape")
                page.add(g)
                g.Graphic = graphic(os.path.join(IMG, "foto.jpg" if i % 2 else "grafik.png"))
                g.Position = Point(2000, 4000)
                g.Size = Size(12000, 9000)
            feat(key, "resim", pic)
        if kind == 7:
            def grp():
                coll = smgr.createInstanceWithContext("com.sun.star.drawing.ShapeCollection", ctx)
                coll.add(add(page, "RectangleShape", 2000, 5000, 4000, 3000, FillColor=0x70AD47, text="A"))
                coll.add(add(page, "EllipseShape", 7000, 5000, 4000, 3000, FillColor=0xFFC000, text="B"))
                page.group(coll)
            feat(key, "grup", grp)
            def ole_chart():
                s = doc.createInstance("com.sun.star.drawing.OLE2Shape")
                page.add(s)
                s.Position = Point(13000, 4000)
                s.Size = Size(10000, 7000)
                s.CLSID = "12dcae26-281f-416f-a234-c3086127382e"
            feat(key, "grafik (OLE)", ole_chart)
        feat(key, "konuşmacı notları", lambda: notes(page, f"Slayt {i + 1} konuşmacı notu: anlatılacak noktalar, ığüşöç."))
        def trans():
            page.Effect = enum("com.sun.star.presentation.FadeEffect", effects[i % len(effects)])
            page.Speed = enum("com.sun.star.presentation.AnimationSpeed", "MEDIUM")
        if i > 0:
            feat(key, "geçiş", trans)
        if i == n - 1 and variant != "sunu":
            feat(key, "gizli slayt", lambda: setattr(page, "Visible", False))
        if i == 1:
            def bg():
                b = doc.createInstance("com.sun.star.drawing.Background")
                b.FillStyle = enum("com.sun.star.drawing.FillStyle", "SOLID")
                b.FillColor = 0xFFF2CC
                page.Background = b
            feat(key, "arka plan rengi", bg)
    try:
        doc.DocumentProperties.Title = f"Derlem {base}"
    except Exception:
        pass
    store(doc, base, [("odp", "impress8"), ("pptx", "Impress MS PowerPoint 2007 XML"), ("ppt", "MS PowerPoint 97")])
    kapat(doc)


for b, v in [("lo-rapor", "rapor"), ("lo-karma", "karma"), ("lo-uzun", "uzun")]:
    writer_doc(b, v)
for b, v in [("lo-butce", "butce"), ("lo-buyuk", "karma")]:
    calc_doc(b, v)
for b, v in [("lo-sunu", "sunu"), ("lo-sunu24", "uzun")]:
    impress_doc(b, v)

with open(os.path.join(OUT, "envanter.json"), "w") as f:
    json.dump(ENV, f, ensure_ascii=False, indent=1)
print(json.dumps(ENV, ensure_ascii=False, indent=1))
