#!/usr/bin/env python3
"""Şema doğrulama kapısı (pnpm derlem:xsd).

Derlem turunun (pnpm derlem) kaydettiği çıktılar biçimin resmî şemasıyla doğrulanır:
  - OOXML (docx/xlsx/pptx ve makrolu/şablon türleri): ECMA-376 5. baskı Bölüm 4 Transitional XSD'leri,
    lxml (libxml2). Önce Biçimlendirme Uyumluluğu (MCE, ECMA-376 Bölüm 3) uygulanır: mc:Ignorable ad
    alanlarındaki öğe ve öznitelikler atılır (mc:ProcessContent olanların içeriği korunur),
    mc:AlternateContent anlaşılan ilk Choice'a (yoksa Fallback'e) indirgenir. Böylece w14/x14/p14
    uzantıları şema hatası sayılmaz (Office de bunları anlamayan okuyucuya böyle sunar).
  - ODF (odt/ods/odp): OASIS ODF 1.3 RelaxNG, jing (odf_rng.py).
Office'in kendi yazdığı ve kendisinin sorunsuz okuduğu bilinen sapmalar xsd-izin.json'daki gerekçeli izin
listesiyle ayıklanır; geri kalan her hata "izin dışı"dır. İzin dışı hata sayıları temel çizgiyle
(temel.json "sema") karşılaştırılır: bir çıktıda yeni hata ya da artış her kipte HATA, azalma iyileşme
(kurallar temel.mjs'de: semaKarsilastir). Özgün dosyalar da doğrulanır: rapor, çıktıdaki hatanın özgünde
de bulunup bulunmadığını gösterir (taşınan içerik mi, yazıcının ürettiği mi).

Şemalar DEPOYA KONMAZ (ECMA-376 şemalarının yeniden dağıtım koşulu belirsiz): ilk kullanımda Ecma'dan
indirilir, SHA-256 doğrulanır ve önbellekte tutulur (OH_SEMA_ONBELLEK, varsayılan <kök>/.sema-onbellek;
git'e girmez; CI'da actions/cache). Ağ yoksa önbelleğe elle konabilir (aynı dosya adı ve SHA).

Kullanım:
  python3 scripts/derlem/xsd.py [--temelsiz] [--kip siki|baslangic] [--yalniz ooxml|odf] [--odf-zorunlu] [--ozgunsuz]
      derlem çıktıları; rapor .derlem/rapor/xsd.json, xsd.md. --odf-zorunlu: jing/java yoksa hata (CI).
      --ozgunsuz: özgün dosyalar doğrulanmaz (rapordaki "özgünde de var" sütunu boş kalır).
  python3 scripts/derlem/xsd.py --indir
      yalnız şemaları ve jing'i önbelleğe indirir (CI önbellek adımı); --anahtar: önbellek anahtarını yazar.
  python3 scripts/derlem/xsd.py <dosya|dizin>...
      verilen dosyaları doğrular ve listeler (izin listesi uygulanır; temel karşılaştırması yok).
Çıkış: 0 tamam, 1 temel çizgiye göre HATA, 2 doğrulama yapılamadı (şema/araç yok, izin listesi geçersiz).
Ortam: DERLEM_CIKTI, OH_SEMA_ONBELLEK, OH_SEMA_ECMA_URL (aynı SHA'lı yansı), JING.
"""
import io
import json
import multiprocessing
import os
import re
import shutil
import subprocess
import sys
import time
import zipfile
from concurrent.futures import ProcessPoolExecutor
from datetime import datetime, timezone
from pathlib import Path

try:
    import resource
except ImportError:  # Windows
    resource = None

sys.dont_write_bytecode = True  # odf_rng'yi içe aktarırken depoya __pycache__ yazılmasın
sys.path.insert(0, str(Path(__file__).resolve().parent))
import odf_rng  # noqa: E402
from odf_rng import KOK, ODF_UZANTI, ONBELLEK, SemaYok, ac, indir, kaldir, ns_of, parca_kalibi, xml_oku  # noqa: E402

from lxml import etree  # noqa: E402

CIKTI = Path(os.environ.get("DERLEM_CIKTI") or KOK / ".derlem").resolve()
RAPOR = CIKTI / "rapor"
IZIN = Path(__file__).resolve().parent / "xsd-izin.json"

# ECMA-376 5. baskı (Aralık 2016) Bölüm 4: içinde OfficeOpenXML-XMLSchema-Transitional.zip var.
ECMA_URL = "https://ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip"
ECMA_ZIP = "ECMA-376-4_5th_edition_december_2016.zip"
ECMA_SHA = "bd25da1109f73762356596918bf5ff8b74a1331642dba5f1c1d1dfc6bed34ecd"
ECMA_IC = "OfficeOpenXML-XMLSchema-Transitional.zip"
OOXML_UZANTI = {"docx", "docm", "dotx", "dotm", "xlsx", "xlsm", "xltx", "xltm", "pptx", "pptm", "potx", "potm", "ppsx", "ppsm"}

MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
XML_NS = "http://www.w3.org/XML/1998/namespace"
# xml: ad alanı şemalarda schemaLocation'sız içe aktarılıyor; lxml ağa çıkmasın diye bu yerel kopya kullanılır.
XML_XSD = """<?xml version="1.0"?>
<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema" targetNamespace="http://www.w3.org/XML/1998/namespace" xml:lang="en">
  <xs:attribute name="lang" type="xs:language"/>
  <xs:attribute name="space"><xs:simpleType><xs:restriction base="xs:NCName"><xs:enumeration value="default"/><xs:enumeration value="preserve"/></xs:restriction></xs:simpleType></xs:attribute>
  <xs:attribute name="base" type="xs:anyURI"/>
  <xs:attribute name="id" type="xs:ID"/>
</xs:schema>
"""

OXML = "http://schemas.openxmlformats.org/"
# Parçanın kök öğesinin ad alanı → şema dosyası (kök ad alanı burada olmayan parçalar doğrulanmaz:
# OPC [Content_Types]/.rels/core.xml Bölüm 2'de; Office'in 2010+ uzantı parçaları şemada yok).
KOK_SEMA = {
    OXML + "wordprocessingml/2006/main": "wml.xsd",
    OXML + "spreadsheetml/2006/main": "sml.xsd",
    OXML + "presentationml/2006/main": "pml.xsd",
    OXML + "drawingml/2006/main": "dml-main.xsd",
    OXML + "drawingml/2006/chart": "dml-chart.xsd",
    OXML + "drawingml/2006/chartDrawing": "dml-chartDrawing.xsd",
    OXML + "drawingml/2006/spreadsheetDrawing": "dml-spreadsheetDrawing.xsd",
    OXML + "drawingml/2006/diagram": "dml-diagram.xsd",
    OXML + "officeDocument/2006/extended-properties": "shared-documentPropertiesExtended.xsd",
    OXML + "officeDocument/2006/custom-properties": "shared-documentPropertiesCustom.xsd",
    OXML + "officeDocument/2006/customXml": "shared-customXmlDataProperties.xsd",
    OXML + "officeDocument/2006/bibliography": "shared-bibliography.xsd",
}
# İletilerde okunur ön ekler (imza ve izin listesi bu adlarla yazılır)
ONEK = {
    OXML + "wordprocessingml/2006/main": "w",
    OXML + "spreadsheetml/2006/main": "x",
    OXML + "presentationml/2006/main": "p",
    OXML + "drawingml/2006/main": "a",
    OXML + "drawingml/2006/chart": "c",
    OXML + "drawingml/2006/chartDrawing": "cdr",
    OXML + "drawingml/2006/spreadsheetDrawing": "xdr",
    OXML + "drawingml/2006/wordprocessingDrawing": "wp",
    OXML + "drawingml/2006/picture": "pic",
    OXML + "drawingml/2006/diagram": "dgm",
    OXML + "drawingml/2006/lockedCanvas": "lc",
    OXML + "officeDocument/2006/relationships": "r",
    OXML + "officeDocument/2006/math": "m",
    OXML + "officeDocument/2006/extended-properties": "ep",
    OXML + "officeDocument/2006/custom-properties": "op",
    OXML + "officeDocument/2006/docPropsVTypes": "vt",
    OXML + "officeDocument/2006/sharedTypes": "s",
    OXML + "officeDocument/2006/customXml": "ds",
    OXML + "officeDocument/2006/bibliography": "b",
    "urn:schemas-microsoft-com:vml": "v",
    "urn:schemas-microsoft-com:office:office": "o",
    "urn:schemas-microsoft-com:office:word": "w10",
    "urn:schemas-microsoft-com:office:excel": "xvml",
    "urn:schemas-microsoft-com:office:powerpoint": "pvml",
    XML_NS: "xml",
    MC: "mc",
    OXML + "drawingml/2006/compatibility": "com",
    "http://www.w3.org/2001/XMLSchema": "xsd",
}


# ------------------------------------------------------------------ şemalar

def ecma_semalari():
    """Önbellekteki Transitional XSD dizini; yoksa Ecma zip'i indirilir (SHA-256 sabit) ve açılır."""
    d = ONBELLEK / "ecma376-4-transitional"
    isaret = d / ".sha256"
    if isaret.exists() and isaret.read_text().strip() == ECMA_SHA:
        return d
    z = indir(os.environ.get("OH_SEMA_ECMA_URL") or ECMA_URL, ONBELLEK / ECMA_ZIP, ECMA_SHA)
    shutil.rmtree(d, ignore_errors=True)
    d.mkdir(parents=True)
    with zipfile.ZipFile(z) as dis, zipfile.ZipFile(io.BytesIO(dis.read(ECMA_IC))) as ic:
        for ad in ic.namelist():
            if not re.fullmatch(r"[A-Za-z0-9-]+\.xsd", ad):
                continue
            veri = re.sub(rb'<xsd:import\s+namespace="http://www\.w3\.org/XML/1998/namespace"\s*/>',
                          b'<xsd:import namespace="http://www.w3.org/XML/1998/namespace" schemaLocation="xml.xsd"/>', ic.read(ad))
            (d / ad).write_bytes(veri)
    (d / "xml.xsd").write_text(XML_XSD, encoding="utf-8")
    isaret.write_text(ECMA_SHA)
    return d


class Xsd:
    """Şema dizini: dosya başına derlenmiş XMLSchema (tembel) ve anlaşılan ad alanları."""

    def __init__(self, dizin):
        self.dizin = Path(dizin)
        self._sema = {}
        ayr = etree.XMLParser(no_network=True, resolve_entities=False)
        self.ayr = ayr
        self.anlasilan = {XML_NS, MC}
        for f in self.dizin.glob("*.xsd"):
            tns = etree.parse(str(f), ayr).getroot().get("targetNamespace")
            if tns:
                self.anlasilan.add(tns)

    def __call__(self, ad):
        if ad not in self._sema:
            self._sema[ad] = etree.XMLSchema(etree.parse(str(self.dizin / ad), self.ayr))
        return self._sema[ad]


# ------------------------------------------------------------------ MCE ön işleme

MCN = {"mc": MC}


def _agacta(el, kok):
    """el hâlâ kok'un ağacında mı (üst öğesiyle birlikte silinmiş olabilir)?"""
    return el is kok or any(u is kok for u in el.iterancestors())


def mce(kok, anlasilan, sorunlar, ham=None):
    """ECMA-376 Bölüm 3: kok'un ağacını anlaşılan ad alanlarına indirger (yerinde).

    Hızlı yol: arama libxml2'de (iter/XPath) yapılır; büyük çalışma sayfalarında öğe öğe Python
    dolaşmasından onlarca kat hızlı. mc:ProcessContent / mc:MustUnderstand varsa tam dolaşma (mce_tam).
    `ham` (parçanın baytları) verilirse MCE ad alanı geçmeyen parça hiç dolaşılmaz."""
    if ham is not None and MC.encode() not in ham:
        return
    # Uzantı (extLst/ext) içeriği tümüyle yok sayılan ad alanındaysa ext boş kalır, oysa şema ext'te bir
    # çocuk ister. Okuyucu anlamadığı uzantıyı hiç yokmuş gibi görür: MCE'nin boşalttığı ext de atılır.
    dolu_ext = kok.xpath("//*[local-name()='ext'][*]") if ham is None or re.search(rb"<(\w+:)?ext[\s>]", ham) else []
    if (ham is None or b"ProcessContent" in ham or b"MustUnderstand" in ham) and kok.xpath(
            "boolean(//@mc:ProcessContent | //@mc:MustUnderstand)", namespaces=MCN):
        mce_tam(kok, frozenset(), frozenset(), anlasilan, sorunlar)
    else:
        _mce_hizli(kok, anlasilan)
    for e in dolu_ext:
        if len(e) == 0 and _agacta(e, kok):
            kaldir(e.getparent(), e)


def _mce_hizli(kok, anlasilan):
    for alt in list(kok.iter(f"{{{MC}}}AlternateContent")):
        ust = alt.getparent()
        if ust is None or not _agacta(alt, kok):
            continue  # seçilmeyen dalla birlikte gitti
        secilen = None
        for ch in alt:
            if ch.tag == f"{{{MC}}}Choice":
                if all(ch.nsmap.get(p) in anlasilan for p in (ch.get("Requires") or "").split()):
                    secilen = ch
                    break
            elif ch.tag == f"{{{MC}}}Fallback":
                secilen = ch
                break
        ac(ust, alt, list(secilen) if secilen is not None else [])
    for el in kok.xpath("descendant-or-self::*[@mc:Ignorable]", namespaces=MCN):
        if not _agacta(el, kok):
            continue
        for ns in sorted({el.nsmap.get(p) for p in el.get(f"{{{MC}}}Ignorable").split()} - {None} - anlasilan):
            for c in list(el.iter(f"{{{ns}}}*")):
                if c is not el and _agacta(c, kok):
                    kaldir(c.getparent(), c)
            for a in el.xpath("descendant-or-self::*/@n:*", namespaces={"n": ns}):
                del a.getparent().attrib[a.attrname]
    for a in kok.xpath("descendant-or-self::*/@mc:*", namespaces=MCN):
        del a.getparent().attrib[a.attrname]


def mce_tam(el, yoksay, isle, anlasilan, sorunlar):
    """mce'nin öğe öğe dolaşan tam biçimi (mc:ProcessContent ve mc:MustUnderstand dahil)."""
    def coz(deger):
        return {el.nsmap.get(p) for p in deger.split()} - {None}

    ig = el.get(f"{{{MC}}}Ignorable")
    if ig:
        yoksay = yoksay | coz(ig)
    pc = el.get(f"{{{MC}}}ProcessContent")
    if pc:
        ek = set()
        for q in pc.split():
            p, _, yerel = q.partition(":")
            if el.nsmap.get(p):
                ek.add(f"{{{el.nsmap[p]}}}{yerel}")
        isle = isle | ek
    mu = el.get(f"{{{MC}}}MustUnderstand")
    if mu:
        for ns in sorted(coz(mu) - anlasilan):
            sorunlar.append(f"mc:MustUnderstand anlaşılmayan ad alanı ister: {ns}")
    for a in list(el.attrib):
        if a.startswith("{"):
            ns = ns_of(a)
            if ns == MC or (ns in yoksay and ns not in anlasilan):
                del el.attrib[a]
    i = 0
    while i < len(el):
        c = el[i]
        if not isinstance(c.tag, str):
            i += 1
            continue
        if c.tag == f"{{{MC}}}AlternateContent":
            secilen = None
            for ch in c:
                if ch.tag == f"{{{MC}}}Choice":
                    if all(ch.nsmap.get(p) in anlasilan for p in (ch.get("Requires") or "").split()):
                        secilen = ch
                        break
                elif ch.tag == f"{{{MC}}}Fallback":
                    secilen = ch
                    break
            ac(el, c, [g for g in secilen] if secilen is not None else [])
            continue  # yerleşen öğeler aynı döngüde işlenir
        ns = ns_of(c.tag)
        if ns in yoksay and ns not in anlasilan:
            if c.tag in isle or f"{{{ns}}}*" in isle:
                ac(el, c)
            else:
                kaldir(el, c)
            continue
        mce_tam(c, yoksay, isle, anlasilan, sorunlar)
        i += 1


# ------------------------------------------------------------------ hata imzası ve izin listesi

def kisa(ad):
    """'{uri}yerel' → 'önek:yerel' (bilinen ad alanları)."""
    return re.sub(r"\{([^}\s']+)\}(?=[A-Za-z_])", lambda m: (ONEK.get(m.group(1)) or "{" + m.group(1) + "}") + ":", ad)


def xsd_hata(parca, e):
    """lxml hata kaydı → {parca, satir, oge, oznitelik, deger, govde, ileti}."""
    ileti = kisa(e.message)
    m = re.match(r"^Element '([^']+)'(?:, attribute '([^']+)')?: (.*)$", ileti, re.S)
    oge, oz, govde = (m.group(1), m.group(2), m.group(3)) if m else (None, None, ileti)
    d = re.search(r"The value '(.*?)' is", govde) or re.match(r"^'(.*?)' is not a valid value", govde)
    govde = re.sub(r"\s*Expected is (one of )?\(.*?\)\.?", "", govde, flags=re.S)
    govde = re.sub(r"The value '.*?' is", "The value '…' is", govde)
    govde = re.sub(r"^'.*?' is not a valid value", "'…' is not a valid value", govde)
    govde = re.sub(r"pattern '.*'\.", "pattern '…'.", govde)
    govde = re.sub(r"set \{.*\}\.", "set {…}.", govde)
    govde = re.sub(r"\('[^']*'\)", "('…')", govde)
    govde = re.sub(r"length of '\d+'", "length of '…'", govde)
    return {"parca": parca, "satir": e.line, "yol": e.path, "oge": oge, "oznitelik": oz, "deger": d.group(1) if d else None, "govde": govde.strip(), "ileti": ileti[:400]}


def imza(h):
    yer = (h.get("oge") or "") + (f"@{h['oznitelik']}" if h.get("oznitelik") else "")
    return f"{parca_kalibi(h['parca'])} | {yer}: {h['govde']}" if yer else f"{parca_kalibi(h['parca'])} | {h['govde']}"


def izin_oku(yol=IZIN):
    """İzin listesini okur ve denetler: her girdinin kimliği, gerekçesi ve en az bir eşleşme alanı olmalı."""
    veri = json.loads(Path(yol).read_text(encoding="utf-8"))
    izinler = veri.get("izinler", [])
    sorun = []
    kimlikler = set()
    for i, z in enumerate(izinler):
        k = z.get("kimlik")
        if not k or k in kimlikler:
            sorun.append(f"{i}. girdi: kimlik yok ya da yinelenmiş ({k})")
        kimlikler.add(k)
        if len((z.get("gerekce") or "").strip()) < 20:
            sorun.append(f"{k}: gerekçe yok ya da çok kısa")
        if not (z.get("oge") or z.get("ileti")):
            sorun.append(f"{k}: en az 'oge' ya da 'ileti' gerekli (çok geniş izin)")
        if z.get("dogrulayici", "xsd") not in ("xsd", "rng"):
            sorun.append(f"{k}: dogrulayici 'xsd' ya da 'rng' olmalı")
        for alan in ("parca", "ileti", "deger"):
            if z.get(alan):
                try:
                    re.compile(z[alan])
                except re.error as e:
                    sorun.append(f"{k}: {alan} düzenli ifadesi geçersiz ({e})")
    if sorun:
        raise ValueError("xsd-izin.json geçersiz:\n  " + "\n  ".join(sorun))
    return izinler


def izinli_mi(h, uzanti, dogrulayici, izinler):
    """Hatayı kapsayan ilk izin girdisinin kimliği (yoksa None)."""
    for z in izinler:
        if z.get("dogrulayici", "xsd") != dogrulayici:
            continue
        if z.get("uzanti") and uzanti not in z["uzanti"]:
            continue
        if z.get("parca") and not re.search(z["parca"], h["parca"]):
            continue
        if z.get("oge") and h.get("oge") not in ([z["oge"]] if isinstance(z["oge"], str) else z["oge"]):
            continue
        if "oznitelik" in z and z["oznitelik"] != h.get("oznitelik"):
            continue
        if z.get("ileti") and not re.search(z["ileti"], h["govde"]):
            continue
        if z.get("deger") and (h.get("deger") is None or not re.fullmatch(z["deger"], h["deger"])):
            continue
        return z["kimlik"]
    return None


# ------------------------------------------------------------------ doğrulama

def ooxml_dogrula(yol, xsd):
    """{parca: doğrulanan parça sayısı, hatalar: [...]} — bir parçada aynı düğüm/öznitelik bir kez sayılır."""
    hatalar = []
    n = 0
    try:
        z = zipfile.ZipFile(yol)
    except Exception as e:
        return {"parca": 0, "hatalar": [{"parca": "", "satir": 0, "govde": "ZIP açılamadı", "ileti": str(e)[:300]}]}
    with z:
        for ad in z.namelist():
            if not ad.endswith(".xml") or ad == "[Content_Types].xml":
                continue
            ham = z.read(ad)
            try:
                kok = xml_oku(ham)
            except Exception as e:
                hatalar.append({"parca": ad, "satir": 0, "govde": "XML okunamadı", "ileti": str(e)[:300]})
                continue
            sema_adi = KOK_SEMA.get(ns_of(kok.tag))
            if not sema_adi:
                continue
            n += 1
            sorunlar = []
            mce(kok, xsd.anlasilan, sorunlar, ham)
            for s in sorunlar:
                hatalar.append({"parca": ad, "satir": 0, "govde": s, "ileti": s})
            sc = xsd(sema_adi)
            if sc.validate(kok):
                continue
            goruldu = set()
            for e in sc.error_log:
                h = xsd_hata(ad, e)
                anahtar = (e.path, h["oge"], h["oznitelik"])
                if anahtar in goruldu:
                    continue  # aynı değer için kalıp + atomik tür iletisi
                goruldu.add(anahtar)
                hatalar.append(h)
    return {"parca": n, "hatalar": hatalar}


def siniflandir(sonuc, uzanti, dogrulayici, izinler):
    """Hataları izinli / izin dışı ayırır; imza sayıları ve örnekler."""
    imzalar, izinli, ornekler = {}, {}, []
    for h in sonuc["hatalar"]:
        k = izinli_mi(h, uzanti, dogrulayici, izinler)
        if k:
            izinli[k] = izinli.get(k, 0) + 1
            continue
        s = imza(h)
        imzalar[s] = imzalar.get(s, 0) + 1
        if len(ornekler) < 25:
            ornekler.append({**{a: h.get(a) for a in ("parca", "satir", "ileti", "deger", "baglam") if h.get(a) not in (None, "")}, "imza": s})
    return {"parca": sonuc["parca"], "imzalar": dict(sorted(imzalar.items())), "izinli": dict(sorted(izinli.items())), "ornekler": ornekler}


def uzantisi(yol):
    return Path(yol).suffix[1:].lower()


_isci = {}


def _isci_dogrula(yol):
    """Süreç havuzu işçisi: şemalar fork'tan önce ana süreçte derlenir, işçi onları paylaşır."""
    return yol, siniflandir(ooxml_dogrula(yol, _isci["xsd"]), uzantisi(yol), "xsd", _isci["izinler"])


def ooxml_toplu(yollar, xsd, izinler):
    """OOXML dosyalarını doğrular; Linux'ta (fork) çekirdek sayısı kadar süreçle, büyük dosyalar önce."""
    _isci.update(xsd=xsd, izinler=izinler)
    paralel = min(4, os.cpu_count() or 1, len(yollar) // 16)
    if paralel < 2 or "fork" not in multiprocessing.get_all_start_methods():
        return dict(_isci_dogrula(y) for y in yollar)
    sirali = sorted(yollar, key=lambda y: -os.path.getsize(y))
    with ProcessPoolExecutor(paralel, mp_context=multiprocessing.get_context("fork")) as h:
        return dict(h.map(_isci_dogrula, sirali, chunksize=2))


def _cpu():
    """Bu süreç ve beklenmiş alt süreçlerin (işçiler, jing) toplam işlemci süresi; makine yükünden az etkilenir."""
    if resource is None:
        return time.process_time()
    c = resource.getrusage(resource.RUSAGE_CHILDREN)
    return time.process_time() + c.ru_utime + c.ru_stime


def dosyalari_dogrula(yollar, izinler, *, yalniz=None, odf_zorunlu=False, olcum=None):
    """yollar → {yol: sınıflanmış sonuç}; ODF yapılamazsa (jing/java yok) ODF yolları sonuçta yer almaz."""
    olcum = olcum if olcum is not None else {}
    sonuc = {}
    ooxml = [y for y in yollar if uzantisi(y) in OOXML_UZANTI] if yalniz != "odf" else []
    odf = [y for y in yollar if uzantisi(y) in ODF_UZANTI] if yalniz != "ooxml" else []

    def odf_atla(e):
        if odf_zorunlu:
            raise e
        olcum["odf"] = f"atlandı: {e}"
        print(f"Uyarı: ODF RelaxNG doğrulaması atlandı: {e}", file=sys.stderr)

    T0, C0 = time.perf_counter(), _cpu()
    odf_is = None
    if odf:
        try:  # jing arka planda çalışırken XSD doğrulanır
            odf_is = odf_rng.baslat(odf, odf_rng.hazirla())
        except SemaYok as e:
            odf_atla(e)
    try:
        if ooxml:
            t0 = time.perf_counter()
            xsd = Xsd(ecma_semalari())
            for ad in set(KOK_SEMA.values()):
                xsd(ad)  # derleme süresi ayrı ölçülsün
            olcum["xsdYuklemeMs"] = round((time.perf_counter() - t0) * 1000)
            t0, c0 = time.perf_counter(), _cpu()
            sonuc.update(ooxml_toplu(ooxml, xsd, izinler))
            olcum["xsdMs"] = round((time.perf_counter() - t0) * 1000)
            olcum["xsdCpuMs"] = round((_cpu() - c0) * 1000)
            olcum["xsdDosya"] = len(ooxml)
    except BaseException:
        if odf_is:
            odf_rng.iptal(odf_is)
        raise
    if odf_is:
        try:
            for y, r in odf_rng.bitir(odf_is).items():
                sonuc[y] = siniflandir(r, uzantisi(y), "rng", izinler)
            olcum["odfMs"] = round((time.perf_counter() - T0) * 1000)
            olcum["odfDosya"] = len(odf)
            olcum["odf"] = "tamam"
        except SemaYok as e:
            odf_atla(e)
    olcum["toplamMs"] = round((time.perf_counter() - T0) * 1000)
    olcum["toplamCpuMs"] = round((_cpu() - C0) * 1000)
    olcum["dosya"] = len(ooxml) + (len(odf) if olcum.get("odf") == "tamam" else 0)
    return sonuc


# ------------------------------------------------------------------ rapor

def git_commit():
    try:
        return subprocess.run(["git", "rev-parse", "--short", "HEAD"], cwd=KOK, capture_output=True, text=True).stdout.strip() or None
    except OSError:
        return None


def sure_ozeti(o):
    """Duvar saati ve işlemci süresi (yüklü makinede işlemci süresi daha kararlı)."""
    sure = []
    if o.get("dosya"):
        sure.append(f"toplam {o['dosya']} dosya {o['toplamMs'] / 1000:.2f} sn = {o['toplamMs'] / 10 / o['dosya']:.2f} sn/100 dosya (işlemci {o['toplamCpuMs'] / 1000:.2f} sn)")
    if o.get("xsdDosya"):
        sure.append(f"XSD {o['xsdDosya']} dosya {o['xsdMs'] / 1000:.2f} sn = {o['xsdMs'] / 10 / o['xsdDosya']:.2f} sn/100 dosya "
                    f"(işlemci {o['xsdCpuMs'] / 10 / o['xsdDosya']:.2f} sn/100; şema derleme {o['xsdYuklemeMs'] / 1000:.2f} sn)")
    if o.get("odfDosya"):
        sure.append(f"ODF {o['odfDosya']} dosya {o['odfMs'] / 1000:.2f} sn (jing, XSD ile eş zamanlı)")
    if o.get("odf", "").startswith("atlandı"):
        sure.append(f"ODF {o['odf']}")
    return " · ".join(sure)


def md_rapor(rapor):
    S = rapor["satirlar"]
    o = rapor["olcum"]
    L = ["# Şema doğrulaması (ECMA-376 Transitional XSD · ODF 1.3 RelaxNG)", ""]
    L.append(f"Şemalar: ECMA-376 5. baskı Bölüm 4 Transitional (SHA-256 `{ECMA_SHA[:12]}…`), OASIS ODF 1.3 (jing {odf_rng.JING_SURUM}). İzin listesi: `scripts/derlem/xsd-izin.json`.")
    L.append("")
    L += [sure_ozeti(o) + " (çıktılar ve özgünleri)", ""]
    L += ["| uygulama | biçim | çıktı | izin dışı hatalı | izin dışı hata | izinli sapma | özgünü izin dışı hatalı |", "|---|---|---|---|---|---|---|"]
    gruplar = {}
    for s in S:
        gruplar.setdefault((s["app"], s["bicim"]), []).append(s)
    for (app, bicim), G in sorted(gruplar.items()):
        L.append(f"| {app} | {bicim} | {len(G)} | {sum(1 for s in G if s['imzalar'])} | {sum(sum(s['imzalar'].values()) for s in G)} | {sum(sum(s['izinli'].values()) for s in G)} | {sum(1 for s in G if s.get('ozgunImzalar'))} |")
    L.append("")
    kullanim = {}
    for s in S:
        for k, n in s["izinli"].items():
            kullanim[k] = kullanim.get(k, 0) + n
    L += ["## İzin listesi", "", "| kimlik | çıktılarda eşleşen | özgünlerde eşleşen | gerekçe |", "|---|---|---|---|"]
    for z in rapor["izinler"]:
        L.append(f"| `{z['kimlik']}` | {kullanim.get(z['kimlik'], 0)} | {z.get('ozgunlerde', 0)} | {z['gerekce'].replace('|', '/')} |")
    L.append("")
    hatali = [s for s in S if s["imzalar"]]
    L += [f"## İzin dışı hatalar ({len(hatali)} çıktı)", ""]
    if hatali:
        L += ["| çıktı | hata (imza) | adet | özgünde de var | örnek |", "|---|---|---|---|---|"]
        for s in hatali:
            oz = s.get("ozgunImzalar") or {}
            for i, (im, n) in enumerate(s["imzalar"].items()):
                orn = next((x for x in s["ornekler"] if x["imza"] == im), {})
                ornek = (orn.get("baglam") or orn.get("deger") or orn.get("ileti") or "").replace("|", "/").replace("\n", " ")[:140]
                var = "—" if s.get("ozgunImzalar") is None else ("evet" if im in oz else "hayır")
                L.append(f"| {s['anahtar'] if i == 0 else '〃'} | `{im.replace('|', '/')}` | {n} | {var} | {ornek} |")
    L.append("")
    return "\n".join(L) + "\n"


def derlem_cikti_listesi(tur, yalniz=None):
    isler = []
    for f in tur["dosyalar"]:
        for t in f.get("turlar") or []:
            b = t.get("bicim")
            if t.get("hata") or not t.get("ciktiYolu") or not Path(t["ciktiYolu"]).exists():
                continue
            if b in OOXML_UZANTI and yalniz != "odf" or b in ODF_UZANTI and yalniz != "ooxml":
                isler.append((f"{f['anahtar']} →{b}", f["app"], b, t["ciktiYolu"], f["yol"]))
    return isler


def argumanlar(argv):
    a = {"_": []}
    i = 0
    while i < len(argv):
        s = argv[i]
        if not s.startswith("--"):
            a["_"].append(s)
        elif "=" in s:
            k, v = s[2:].split("=", 1)
            a[k] = v
        elif i + 1 < len(argv) and not argv[i + 1].startswith("--") and s[2:] in ("kip", "yalniz"):
            a[s[2:]] = argv[i + 1]
            i += 1
        else:
            a[s[2:]] = True
        i += 1
    return a


def ana(argv):
    A = argumanlar(argv)
    if A.get("anahtar"):
        # CI önbellek anahtarı: sabit SHA'lar değişince önbellek yenilenir
        print("sema-" + "-".join(x[:12] for x in [ECMA_SHA, *odf_rng.ODF_SEMALAR.values(), odf_rng.JING_SHA]))
        return 0
    if A.get("indir"):
        try:
            d = ecma_semalari()
            h = odf_rng.hazirla()
        except SemaYok as e:
            print(f"İndirilemedi: {e}", file=sys.stderr)
            return 2
        print(f"Şemalar hazır: {d} · {h['ana'].parent} · jing: {' '.join(h['jing'])}")
        return 0
    try:
        izinler = izin_oku()
    except (ValueError, json.JSONDecodeError) as e:
        print(e, file=sys.stderr)
        return 2
    yalniz = A.get("yalniz") if A.get("yalniz") in ("ooxml", "odf") else None

    if A["_"]:
        yollar = []
        for p in map(Path, A["_"]):
            if p.is_dir():
                yollar += sorted(str(x) for x in p.rglob("*") if uzantisi(x) in OOXML_UZANTI | ODF_UZANTI)
            elif p.exists():
                yollar.append(str(p))
        try:
            sonuc = dosyalari_dogrula(yollar, izinler, yalniz=yalniz)
        except SemaYok as e:
            print(f"Şema doğrulaması yapılamadı: {e}", file=sys.stderr)
            return 2
        toplam = 0
        for y, r in sonuc.items():
            n = sum(r["imzalar"].values())
            toplam += n
            izinli = sum(r["izinli"].values())
            print(f"{y}: {r['parca']} parça, {'tamam' if not n else f'{n} izin dışı hata'}{f' ({izinli} izinli sapma)' if izinli else ''}")
            for im, k in list(r["imzalar"].items())[:8]:
                print(f"    {k} × {im}")
        print(f"\n{toplam} izin dışı şema hatası")
        return 0

    rapor_yolu = RAPOR / "derlem.json"
    if not rapor_yolu.exists():
        print("Önce `pnpm derlem` çalıştırın (kaydedilen dosyalar ve derlem.json gerekli).", file=sys.stderr)
        return 2
    tur = json.loads(rapor_yolu.read_text(encoding="utf-8"))
    isler = derlem_cikti_listesi(tur, yalniz)
    ozgunler = sorted({oz for *_, oz in isler if Path(oz).exists() and uzantisi(oz) in OOXML_UZANTI | ODF_UZANTI})
    olcum = {}
    t0 = time.perf_counter()
    try:
        # Özgünler yalnız bilgi için (hata çıktıda mı doğdu, taşındı mı); aynı çağrıda (tek jing, tek havuz).
        ozgunler = [] if A.get("ozgunsuz") else ozgunler
        sonuc = dosyalari_dogrula([c for _, _, _, c, _ in isler] + ozgunler, izinler, yalniz=yalniz, odf_zorunlu=bool(A.get("odf-zorunlu")), olcum=olcum)
        ozgun_sonuc = sonuc
    except SemaYok as e:
        print(f"Şema doğrulaması yapılamadı: {e}", file=sys.stderr)
        return 2
    satirlar = []
    for anahtar, app, bicim, cikti, ozgun in isler:
        r = sonuc.get(cikti)
        if r is None:
            continue  # ODF atlandı
        oz = ozgun_sonuc.get(ozgun)
        satirlar.append({"anahtar": anahtar, "app": app, "bicim": bicim, "dogrulayici": "rng" if bicim in ODF_UZANTI else "xsd",
                         "cikti": os.path.relpath(cikti, KOK), **r, "ozgunImzalar": oz["imzalar"] if oz else None})
    rapor = {
        "tarih": datetime.now(timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "Z"),
        "derlemTarih": tur.get("tarih"),
        "commit": git_commit(),
        "kapsamli": tur.get("kapsamli", True) and yalniz is None and olcum.get("odf", "tamam") == "tamam",
        "sureMs": round((time.perf_counter() - t0) * 1000),
        "olcum": olcum,
        "izinler": [{"kimlik": z["kimlik"], "gerekce": z["gerekce"],
                     "ozgunlerde": sum(ozgun_sonuc[o]["izinli"].get(z["kimlik"], 0) for o in ozgunler if o in ozgun_sonuc)} for z in izinler],
        "satirlar": satirlar,
    }
    RAPOR.mkdir(parents=True, exist_ok=True)
    (RAPOR / "xsd.json").write_text(json.dumps(rapor, ensure_ascii=False, indent=1), encoding="utf-8")
    (RAPOR / "xsd.md").write_text(md_rapor(rapor), encoding="utf-8")
    hatali = [s for s in satirlar if s["imzalar"]]
    print(f"\nŞema doğrulaması: {len(satirlar)} çıktı, izin dışı hatalı {len(hatali)}, izin dışı hata {sum(sum(s['imzalar'].values()) for s in satirlar)}, izinli sapma {sum(sum(s['izinli'].values()) for s in satirlar)}.")
    for app in ("belge", "tablo", "sunum"):
        X = [s for s in satirlar if s["app"] == app]
        if X:
            ox = [s for s in X if s["dogrulayici"] == "xsd"]
            od = [s for s in X if s["dogrulayici"] == "rng"]
            print(f"  {app}: XSD {sum(1 for s in ox if s['imzalar'])}/{len(ox)} çıktı izin dışı hatalı · ODF {sum(1 for s in od if s['imzalar'])}/{len(od)}")
    print(f"  süre: {sure_ozeti(olcum)}")
    print(f"Rapor: {RAPOR / 'xsd.md'}")
    if A.get("temelsiz"):
        return 0
    komut = ["node", str(KOK / "scripts/derlem/temel.mjs"), "--sema"]
    if A.get("kip"):
        komut += ["--kip", str(A["kip"])]
    return subprocess.run(komut, cwd=KOK).returncode


if __name__ == "__main__":
    sys.exit(ana(sys.argv[1:]))
