#!/usr/bin/env bash
# OfficeHere apt deposunu reprepro ile üretir ve imzalar (Release, Release.gpg, InRelease).
#
# Kullanım: scripts/apt/depo-uret.sh <çıktı-dizini> <paket.deb>...
#
# Ortam:
#   APT_ANAHTAR  İmzalayan anahtarın parmak izi (zorunlu). Gizli anahtar GNUPGHOME'daki anahtar
#                zincirinde olmalı; CI'da yalnız imzalama alt anahtarı, parolasız ve onaylı bir
#                GitHub ortamının gizlisinden içe aktarılır (docs/dagitim.md). Anahtar depoya girmez.
#   APT_ADRES    Deponun yayımlanacağı adres (ör. https://apt.ornek.org). officehere.sources'a yazılır;
#                verilmezse yer tutucu kalır.
#
# Depo her sürümde baştan, verilen paketlerden üretilir (durumsuz): dizinde yalnız son sürüm listelenir.
# Eski .deb dosyaları uzak depoda kalabilir; yukle.sh pool/ altını silmez.
#
# Çıktı (<çıktı-dizini>):
#   dists/, pool/            yayımlanacak depo
#   officehere.pgp           açık anahtar (ikili; Signed-By bunu gösterir)
#   officehere.asc           açık anahtar (ASCII zırhlı; elle doğrulama için)
#   officehere.sources       deb822 kaynak dosyası
set -euo pipefail

if [[ $# -lt 2 ]]; then
  echo "Kullanım: $0 <çıktı-dizini> <paket.deb>..." >&2
  exit 2
fi
cikti=$1
shift
: "${APT_ANAHTAR:?APT_ANAHTAR (imzalayan anahtarın parmak izi) gerekli}"
adres=${APT_ADRES:-https://apt.ALAN-ADI}

for arac in reprepro gpg gpgv dpkg-deb; do
  command -v "$arac" >/dev/null || { echo "Eksik araç: $arac (Debian/Pardus: apt-get install reprepro gnupg gpgv)" >&2; exit 1; }
done
gpg --batch --list-secret-keys "$APT_ANAHTAR" >/dev/null 2>&1 || { echo "Gizli anahtar bulunamadı: $APT_ANAHTAR" >&2; exit 1; }

for deb in "$@"; do
  [[ -f "$deb" ]] || { echo "Paket yok: $deb" >&2; exit 1; }
  ad=$(dpkg-deb -f "$deb" Package)
  [[ "$ad" == officehere ]] || { echo "Beklenmeyen paket adı: $ad ($deb)" >&2; exit 1; }
done

taban=$(mktemp -d)
trap 'rm -rf "$taban"' EXIT
mkdir -p "$taban/conf" "$cikti"
rm -rf "$cikti/dists" "$cikti/pool"

# NotAutomatic + ButAutomaticUpgrades: depo önceliği 100 olur; dağıtımın kendi paketlerinin yerine
# hiçbir şey geçmez, kurulu officehere ise güncellenir (Debian UseThirdParty önerisi).
cat > "$taban/conf/distributions" <<EOF
Origin: OfficeHere
Label: OfficeHere
Suite: stable
Codename: stable
Architectures: amd64 arm64
Components: main
Description: OfficeHere masaüstü paketleri
SignWith: $APT_ANAHTAR
NotAutomatic: yes
ButAutomaticUpgrades: yes
EOF

for deb in "$@"; do
  reprepro --silent -b "$taban" --outdir "$cikti" includedeb stable "$deb"
done

gpg --batch --yes --export "$APT_ANAHTAR" > "$cikti/officehere.pgp"
gpg --batch --yes --armor --export "$APT_ANAHTAR" > "$cikti/officehere.asc"

cat > "$cikti/officehere.sources" <<EOF
Types: deb
URIs: $adres
Suites: stable
Components: main
Architectures: amd64
Signed-By: /etc/apt/keyrings/officehere.pgp
EOF

# Yayından önce imza ve özetler yalnız açık anahtarla doğrulanır.
gpgv --keyring "$cikti/officehere.pgp" "$cikti/dists/stable/InRelease" 2>/dev/null
gpgv --keyring "$cikti/officehere.pgp" "$cikti/dists/stable/Release.gpg" "$cikti/dists/stable/Release" 2>/dev/null
paketler="$cikti/dists/stable/main/binary-amd64/Packages"
for deb in "$@"; do
  ozet=$(sha256sum "$deb" | cut -d' ' -f1)
  grep -q "^SHA256: $ozet$" "$paketler" || { echo "Dizinde özet yok: $deb" >&2; exit 1; }
done
grep -q "^ $(sha256sum "$paketler" | cut -d' ' -f1) " "$cikti/dists/stable/Release" || { echo "Release, Packages özetini içermiyor" >&2; exit 1; }

echo "Depo hazır: $cikti"
grep -E '^(Package|Version|Filename):' "$paketler"
