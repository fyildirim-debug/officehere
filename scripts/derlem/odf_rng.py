#!/usr/bin/env python3
"""ODF (odt/ods/odp) çıktılarını OASIS ODF 1.3 RelaxNG şemasıyla doğrular (jing).

Yabancı öğe ve öznitelikler (ODF 1.3 Bölüm 3 §3.17; LibreOffice'in loext/calcext/officeooo… ad alanları,
OfficeHere'in ohtablo: işaretleri) doğrulamadan önce standardın tarif ettiği gibi atılır: text:p / text:h
içindekilerin (ya da office:process-content="true" olanların) içeriği korunur, diğerleri tümüyle silinir.
Geriye kalan belge ODF 1.3 şemasına uymalıdır. content/styles/meta/settings.xml (gömülü nesnelerinkiler
dahil) ana şemayla, META-INF/manifest.xml manifest şemasıyla doğrulanır.

Şemalar ve jing depoya konmaz; ilk kullanımda önbelleğe indirilir (SHA-256 sabit). Önbellek dizini
OH_SEMA_ONBELLEK (varsayılan <kök>/.sema-onbellek, git'e girmez). jing sırası: JING ortam değişkeni
(jing.jar ya da çalıştırılabilir), önbellekteki sabit sürüm (java gerekir; yoksa indirilir), PATH'teki jing.
CI'da java hazırdır; yerelde: `sudo apt-get install default-jre-headless` (ya da `apt-get install jing`).

Bu modül xsd.py'nin ortak yardımcılarını da taşır (önbellek, indirme, imza). Ana giriş xsd.py'dir
(pnpm derlem:xsd hem OOXML hem ODF'yi doğrular); tek başına kullanım yalnız listeler:
  python3 scripts/derlem/odf_rng.py <dosya.odt|ods|odp|dizin>...
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.request
import zipfile
from hashlib import sha256
from pathlib import Path

from lxml import etree

KOK = Path(__file__).resolve().parents[2]
ONBELLEK = Path(os.environ.get("OH_SEMA_ONBELLEK") or KOK / ".sema-onbellek").resolve()
ODF_UZANTI = {"odt", "ott", "ods", "ots", "odp", "otp"}

# OASIS ODF 1.3 (OASIS Standard, 27 Nisan 2021). Şema dosyaları OASIS telif bildirimiyle serbestçe
# kopyalanabilir; yine de depoya konmaz (ECMA-376 ile aynı düzen).
ODF_KAYNAK = "https://docs.oasis-open.org/office/OpenDocument/v1.3/os/schemas/"
ODF_SEMALAR = {
    "OpenDocument-v1.3-schema.rng": "40bad03efdbb02825230d357da0aa6ac679934c5bf56c6281752c0c24d58e4e6",
    "OpenDocument-v1.3-manifest-schema.rng": "8aee71f03484be112af972d622cc9031280c007b0a454ae8815e8e333c9bdd17",
}
# jing (BSD-3-Clause; Thai Open Source Software Center). Yalnız araç olarak çalıştırılır, dağıtılmaz.
JING_SURUM = "20220510"
JING_URL = f"https://github.com/relaxng/jing-trang/releases/download/V{JING_SURUM}/jing-{JING_SURUM}.zip"
JING_SHA = "c1887f5ea6bd90067baafb4cd0d215294781f6fd121b7905c89829d43ea9e280"

OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
XML_NS = "http://www.w3.org/XML/1998/namespace"
PARAGRAF = {f"{{{TEXT}}}p", f"{{{TEXT}}}h"}
ISLE = f"{{{OFFICE}}}process-content"


class SemaYok(Exception):
    """Şema ya da doğrulayıcı hazırlanamadı (ağ yok, SHA uyuşmuyor, java yok)."""


# ------------------------------------------------------------------ ortak: önbellek, indirme, imza

def dosya_sha(yol):
    h = sha256()
    with open(yol, "rb") as f:
        for parca in iter(lambda: f.read(1 << 20), b""):
            h.update(parca)
    return h.hexdigest()


def indir(url, hedef, sha, deneme=3):
    """url'yi hedefe indirir ve SHA-256'yı doğrular; hedef zaten doğruysa dokunmaz."""
    hedef = Path(hedef)
    if hedef.exists() and dosya_sha(hedef) == sha:
        return hedef
    hedef.parent.mkdir(parents=True, exist_ok=True)
    gecici = hedef.with_suffix(hedef.suffix + ".indiriliyor")
    son = None
    for i in range(deneme):
        try:
            # Ecma sunucusu bazı istemci kimliklerinde bağlantıyı kesiyor; sıradan bir kimlik yeterli.
            istek = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (OfficeHere derlem; sema)"})
            with urllib.request.urlopen(istek, timeout=120) as r, open(gecici, "wb") as f:
                shutil.copyfileobj(r, f)
            gercek = dosya_sha(gecici)
            if gercek != sha:
                gecici.unlink(missing_ok=True)
                raise SemaYok(f"{url}: SHA-256 uyuşmuyor (beklenen {sha}, gelen {gercek}). Kaynak değişmiş olabilir; elle doğrulanmadan sabit değer güncellenmemeli.")
            gecici.replace(hedef)
            return hedef
        except SemaYok:
            raise
        except Exception as e:  # ağ hatası: kısa bekleyip yeniden dene
            son = e
            time.sleep(2 * (i + 1))
    gecici.unlink(missing_ok=True)
    raise SemaYok(f"{url} indirilemedi: {son}")


def parca_kalibi(ad):
    """Parça adındaki sıra numaralarını '#' yapar: ppt/slides/slide12.xml → ppt/slides/slide#.xml."""
    return re.sub(r"\d+", "#", ad)


def kaldir(ust, c):
    """Öğeyi kaldırır; ardından gelen metni (tail) korur."""
    if c.tail:
        onceki = c.getprevious()
        if onceki is not None:
            onceki.tail = (onceki.tail or "") + c.tail
        else:
            ust.text = (ust.text or "") + c.tail
    ust.remove(c)


def ac(ust, c, cocuklar=None):
    """Öğeyi yerinde çocuklarıyla (ve metniyle) değiştirir; cocuklar verilirse yalnız onlar yerleşir."""
    i = ust.index(c)
    if cocuklar is None:
        cocuklar = list(c)
        if c.text:
            onceki = c.getprevious()
            if onceki is not None:
                onceki.tail = (onceki.tail or "") + c.text
            else:
                ust.text = (ust.text or "") + c.text
    kuyruk = c.tail
    ust.remove(c)
    for k, g in enumerate(cocuklar):
        ust.insert(i + k, g)
    if kuyruk:
        son = cocuklar[-1] if cocuklar else (ust[i - 1] if i > 0 else None)
        if son is not None:
            son.tail = (son.tail or "") + kuyruk
        else:
            ust.text = (ust.text or "") + kuyruk


GUVENLI = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=True, remove_blank_text=False)


def xml_oku(veri):
    return etree.fromstring(veri, GUVENLI)


def ns_of(ad):
    return ad[1:].split("}")[0] if ad.startswith("{") else ""


# ------------------------------------------------------------------ ODF şemaları ve jing

def odf_semalari():
    """Önbellekteki ODF 1.3 şemaları (yoksa indirilir): (ana şema yolu, manifest şema yolu)."""
    d = ONBELLEK / "odf-1.3"
    yollar = [indir(ODF_KAYNAK + ad, d / ad, sha) for ad, sha in ODF_SEMALAR.items()]
    return yollar[0], yollar[1]


def _java():
    return shutil.which("java")


# Kısa ömürlü JVM: yalnız C1 derleyici ve tek iş parçacıklı çöp toplayıcı (derlemde işlemci süresi ~%30 az).
JVM = ["-XX:+UseSerialGC", "-XX:TieredStopAtLevel=1"]


def jing_komutu(indirme=True):
    """jing'i çalıştıran komut (liste). Bulunamazsa SemaYok."""
    ortam = os.environ.get("JING")
    if ortam:
        return ["java", *JVM, "-jar", ortam] if ortam.endswith(".jar") else [ortam]
    jar = ONBELLEK / f"jing-{JING_SURUM}" / "bin" / "jing.jar"
    java = _java()
    if java and not jar.exists() and indirme:
        z = indir(JING_URL, ONBELLEK / f"jing-{JING_SURUM}.zip", JING_SHA)
        with zipfile.ZipFile(z) as arsiv:
            for ad in arsiv.namelist():
                if ad.startswith(f"jing-{JING_SURUM}/bin/") and ad.endswith(".jar"):
                    hedef = ONBELLEK / ad
                    hedef.parent.mkdir(parents=True, exist_ok=True)
                    hedef.write_bytes(arsiv.read(ad))
    if java and jar.exists():
        return [java, *JVM, "-jar", str(jar)]
    if shutil.which("jing"):
        return ["jing"]
    raise SemaYok("jing çalıştırılamıyor: java bulunamadı (CI'da hazır; yerelde `sudo apt-get install default-jre-headless`) ve PATH'te jing yok.")


def hazirla(indirme=True):
    """Şemaları ve jing'i hazırlar (CI önbellek adımı bunu çağırır)."""
    ana, man = odf_semalari()
    return {"ana": ana, "manifest": man, "jing": jing_komutu(indirme)}


_ns_cache = {}


def sema_ad_alanlari(rng):
    """Şemanın tanıdığı ad alanları (kök öğedeki bildirimler + xml); bunların dışı 'yabancı'dır."""
    if rng not in _ns_cache:
        kok = etree.parse(str(rng), GUVENLI).getroot()
        _ns_cache[rng] = {u for p, u in kok.nsmap.items() if u != "http://relaxng.org/ns/structure/1.0"} | {XML_NS}
    return _ns_cache[rng]


def _paragraf_icerigi(el):
    """el paragraf içeriği mi: text:p / text:h ya da onların içindeki text: öğeleri (text:span, text:a…)."""
    while el is not None:
        if el.tag in PARAGRAF:
            return True
        if ns_of(el.tag) != TEXT:
            return False
        el = el.getparent()
    return False


def yabancilari_at(kok, bilinen, ham=None):
    """ODF 1.3 §3.17: yabancı öznitelikler silinir; yabancı öğe paragraf içeriğindeyse ya da
    office:process-content="true" ise içeriğiyle değiştirilir, değilse tümüyle silinir.

    Yabancı ad alanları belgedeki bildirimlerden bulunur ve arama libxml2'de (XPath) yapılır; büyük
    çalışma sayfalarında öğe öğe dolaşmaktan çok daha hızlı. `ham`: parçanın baytları (yoksa yeniden üretilir)."""
    ham = etree.tostring(kok) if ham is None else ham
    yabanci = sorted({u.decode() for u in re.findall(rb'xmlns(?::[\w.-]+)?="([^"]*)"', ham)} - bilinen)
    if yabanci:
        nsler = {f"y{i}": u for i, u in enumerate(yabanci)}
        for a in kok.xpath(" | ".join(f"//@{p}:*" for p in nsler), namespaces=nsler):
            del a.getparent().attrib[a.attrname]
        # Belge sırasıyla: dıştaki önce; silinen öğenin içindekiler artık ağaçta değildir.
        for c in kok.xpath(" | ".join(f"//{p}:*" for p in nsler), namespaces=nsler):
            ust = c.getparent()
            if ust is None or not any(u is kok for u in c.iterancestors()):
                continue
            isle = c.get(ISLE)
            if isle == "true" or (isle is None and _paragraf_icerigi(ust)):
                ac(ust, c)
            else:
                kaldir(ust, c)
    for a in kok.xpath("//@o:process-content", namespaces={"o": OFFICE}):
        del a.getparent().attrib[a.attrname]


def _jing_cikti(cikti, harita):
    """jing çıktısı: '<dosya>:<satır>:<sütun>: error: <ileti>' → {(yol, parça): [(satır, ileti)]}.
    Doğrulanan dosyalara ait olmayan hata (şema okunamadı vb.) SemaYok'tur."""
    sonuc = {}
    for s in cikti.splitlines():
        m = re.match(r"^(.*?):(\d+):(\d+): (?:error|fatal): (.*)$", s)
        if not m:
            continue
        anahtar = harita.get(os.path.realpath(m.group(1)))
        if anahtar is None:
            raise SemaYok(f"jing: {s[:400]}")
        sonuc.setdefault(anahtar, []).append((int(m.group(2)), int(m.group(3)), m.group(4)))
    return sonuc


def rng_imza(ileti):
    """jing iletisinden kararlı imza: beklenen listesi ve değerler atılır."""
    govde = ileti.split("; expected")[0].split("; must be")[0]
    govde = re.sub(r'value "[^"]*"', 'value "…"', govde)
    oge = re.search(r'element "([^"]+)"', ileti)
    oz = re.search(r'attribute "([^"]+)"', ileti)
    return govde.strip(), (oge.group(1) if oge else None), (oz.group(1) if oz else None)


def _dengele(dosyalar, n):
    """Dosyaları boyutça dengeli en çok n gruba böler (büyükler önce, en hafif gruba)."""
    gruplar = [[] for _ in range(max(1, min(n, len(dosyalar))))]
    yuk = [0] * len(gruplar)
    for d in sorted(dosyalar, key=lambda d: -os.path.getsize(d)):
        i = yuk.index(min(yuk))
        gruplar[i].append(d)
        yuk[i] += os.path.getsize(d)
    return [g for g in gruplar if g]


def baslat(yollar, hazir=None, paralel=None):
    """Parçaları hazırlar (yabancı öğeler atılır) ve jing'i arka planda, boyutça dengeli birkaç süreçte
    başlatır. Sonuç bitir() ile toplanır; arada başka iş (XSD doğrulaması) yapılabilir."""
    hazir = hazir or hazirla()
    bilinen = sema_ad_alanlari(hazir["ana"])
    bilinen_man = sema_ad_alanlari(hazir["manifest"])
    sonuc = {}
    gecici = tempfile.mkdtemp(prefix="oh-odf-")
    ana_dosyalar, man_dosyalar, harita = [], [], {}
    for i, yol in enumerate(yollar):
        sonuc[yol] = {"parca": 0, "hatalar": []}
        try:
            z = zipfile.ZipFile(yol)
        except Exception as e:
            sonuc[yol]["hatalar"].append({"parca": "", "satir": 0, "ileti": f"ZIP açılamadı: {e}", "govde": "ZIP açılamadı"})
            continue
        with z:
            for ad in z.namelist():
                taban = ad.rsplit("/", 1)[-1]
                if ad == "META-INF/manifest.xml":
                    liste, ns = man_dosyalar, bilinen_man
                elif taban in ("content.xml", "styles.xml", "meta.xml", "settings.xml"):
                    liste, ns = ana_dosyalar, bilinen
                else:
                    continue
                ham = z.read(ad)
                try:
                    kok = xml_oku(ham)
                except Exception as e:
                    sonuc[yol]["hatalar"].append({"parca": ad, "satir": 0, "ileti": f"XML okunamadı: {e}", "govde": "XML okunamadı"})
                    continue
                if liste is ana_dosyalar and ns_of(kok.tag) != OFFICE:
                    continue
                yabancilari_at(kok, ns, ham)
                hedef = Path(gecici) / str(i) / ad.replace("/", "__")
                hedef.parent.mkdir(parents=True, exist_ok=True)
                hedef.write_bytes(etree.tostring(kok, xml_declaration=True, encoding="UTF-8", pretty_print=True))
                harita[os.path.realpath(hedef)] = (yol, ad)
                liste.append(str(hedef))
                sonuc[yol]["parca"] += 1
    # Her JVM şemayı yeniden derler (~1 sn işlemci): XSD ile eş zamanlı tek süreç en hızlısı (ölçüldü: 1/2/3 süreçte
    # 3,3/3,9/4,6 sn). Süreç başına en çok 400 dosya (komut satırı sınırı).
    n = paralel or max(1, -(-len(ana_dosyalar) // 400))
    isler = [(hazir["ana"], g) for g in _dengele(ana_dosyalar, n)] + ([(hazir["manifest"], man_dosyalar)] if man_dosyalar else [])
    is_ = {"sonuc": sonuc, "harita": harita, "surecler": [], "gecici": gecici}
    try:
        for k, (sema, dosyalar) in enumerate(isler):
            cikti = open(Path(gecici) / f"jing-{k}.txt", "w+", encoding="utf-8", errors="replace")
            # -i: ID/IDREF denetimi kapalı (ODF şeması DTD uyumlu ID türleri kullanmıyor; jing şemayı reddeder)
            p = subprocess.Popen([*hazir["jing"], "-i", str(sema), *dosyalar], stdout=cikti, stderr=subprocess.STDOUT)
            is_["surecler"].append((p, cikti))
    except OSError as e:
        iptal(is_)
        raise SemaYok(f"jing başlatılamadı: {e}") from e
    return is_


def iptal(is_):
    """Süren jing süreçlerini durdurur ve geçici dosyaları siler."""
    for p, f in is_["surecler"]:
        if p.poll() is None:
            p.kill()
        p.wait()
        f.close()
    shutil.rmtree(is_["gecici"], ignore_errors=True)


def bitir(is_):
    """jing'i bekler: {yol: {"parca": n, "hatalar": [{parca, satir, ileti, oge, oznitelik, govde, baglam}]}}."""
    sonuc, harita = is_["sonuc"], is_["harita"]
    try:
        for p, f in is_["surecler"]:
            kod = p.wait()
            f.seek(0)
            cikti = f.read()
            if kod not in (0, 1):
                raise SemaYok(f"jing çalışmadı (çıkış {kod}): {cikti[:400]}")
            for (yol, ad), hatalar in _jing_cikti(cikti, harita).items():
                dosya = next(d for d, v in harita.items() if v == (yol, ad))
                satirlar = Path(dosya).read_text(encoding="utf-8", errors="replace").splitlines()
                for satir, sutun, ileti in hatalar:
                    govde, oge, oz = rng_imza(ileti)
                    metin = satirlar[satir - 1] if 0 < satir <= len(satirlar) else ""
                    # jing sütunu başlangıç etiketinin sonunu gösterir: o etiketten itibaren kısa bir kesit
                    bas = metin.rfind("<", 0, max(sutun - 1, 1))
                    if metin.startswith("</", bas):  # karakter içeriği hatası: kapanıştan önceki başlangıç etiketi
                        bas = metin.rfind("<", 0, bas)
                    baglam = metin[max(bas, 0):][:160].strip()
                    sonuc[yol]["hatalar"].append({"parca": ad, "satir": satir, "ileti": ileti, "govde": govde, "oge": oge, "oznitelik": oz, "baglam": baglam})
    finally:
        iptal(is_)
    return sonuc


def dogrula(yollar, hazir=None):
    """ODF dosyalarını doğrular (baslat + bitir)."""
    return bitir(baslat(yollar, hazir))


def _ana():
    yollar = []
    for a in sys.argv[1:]:
        p = Path(a)
        if p.is_dir():
            yollar += sorted(str(x) for x in p.rglob("*") if x.suffix[1:].lower() in ODF_UZANTI)
        elif p.suffix[1:].lower() in ODF_UZANTI:
            yollar.append(str(p))
    if not yollar:
        print(__doc__)
        return 0
    try:
        sonuc = dogrula(yollar)
    except SemaYok as e:
        print(f"ODF doğrulaması yapılamadı: {e}", file=sys.stderr)
        return 2
    toplam = 0
    for yol, r in sonuc.items():
        toplam += len(r["hatalar"])
        print(f"{yol}: {r['parca']} parça, {'tamam' if not r['hatalar'] else str(len(r['hatalar'])) + ' hata'}")
        for h in r["hatalar"][:8]:
            print(f"    {h['parca']}:{h['satir']}: {h['ileti'][:200]}")
    print(f"\n{toplam} şema hatası")
    return 0


if __name__ == "__main__":
    sys.exit(_ana())
