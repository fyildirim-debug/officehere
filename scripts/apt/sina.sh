#!/usr/bin/env bash
# apt deposunu uçtan uca sınar: geçici bir TEST anahtarıyla depo üretilir (depo-uret.sh), yerel bir
# HTTP sunucusundan yayımlanır, deb822 .sources + Signed-By ile eklenir ve "apt-get install officehere"
# çalıştırılır. Ardından imzanın gerçekten denetlendiği sınanır: başka anahtarla ve bozulmuş dizinle
# "apt-get update" başarısız olmalıdır. Gerçek anahtar ve alan adı kullanılmaz.
#
# Kullanım (root, Debian/Pardus; .github/workflows/pardus.yml): scripts/apt/sina.sh <paket.deb>
# Gerekenler: reprepro gnupg gpgv python3 (yerel HTTP sunucusu için)
set -euo pipefail

deb=$(readlink -f "$1")
port=${OH_APT_PORT:-8765}
burasi=$(cd "$(dirname "$0")" && pwd)
[[ $(id -u) -eq 0 ]] || { echo "root olarak çalıştırın (apt yapılandırması değişir)" >&2; exit 1; }

is=$(mktemp -d)
export GNUPGHOME="$is/gnupg"
mkdir -m 700 "$GNUPGHOME"
sunucu=""
temizle() {
  [[ -n "$sunucu" ]] && kill "$sunucu" 2>/dev/null || true
  rm -f /etc/apt/sources.list.d/officehere.sources /etc/apt/keyrings/officehere.pgp
  gpgconf --kill all 2>/dev/null || true
  rm -rf "$is"
}
trap temizle EXIT

anahtar_uret() {
  gpg --batch --passphrase '' --quick-gen-key "$1" ed25519 sign never 2>/dev/null
  gpg --batch --with-colons --list-secret-keys "$1" | awk -F: '/^fpr:/ { print $10; exit }'
}

echo "== Test anahtarı ve depo"
APT_ANAHTAR=$(anahtar_uret "OfficeHere apt TEST <test@officehere.invalid>")
export APT_ANAHTAR
APT_ADRES="http://127.0.0.1:$port" "$burasi/depo-uret.sh" "$is/depo" "$deb"

python3 -m http.server "$port" --bind 127.0.0.1 --directory "$is/depo" >/dev/null 2>&1 &
sunucu=$!
for _ in $(seq 1 50); do
  python3 -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:$port/dists/stable/InRelease')" 2>/dev/null && break
  sleep 0.2
done

echo "== deb822 kaynağı ve Signed-By"
install -d -m 0755 /etc/apt/keyrings
install -m 0644 "$is/depo/officehere.pgp" /etc/apt/keyrings/officehere.pgp
install -m 0644 "$is/depo/officehere.sources" /etc/apt/sources.list.d/officehere.sources
cat /etc/apt/sources.list.d/officehere.sources

# Yalnız bu kaynağı tazeler (dağıtımın listeleri önceden güncellenmiş olmalı).
yalniz_bu=(-o Dir::Etc::SourceList=/etc/apt/sources.list.d/officehere.sources -o Dir::Etc::SourceParts=- -o APT::Get::List-Cleanup=0)
apt-get update "${yalniz_bu[@]}"

echo "== apt-get install officehere"
apt-cache policy officehere
DEBIAN_FRONTEND=noninteractive apt-get install -y officehere
beklenen=$(dpkg-deb -f "$deb" Version)
kurulu=$(dpkg-query -W -f '${Version}' officehere)
[[ "$kurulu" == "$beklenen" ]] || { echo "Kurulu sürüm $kurulu, beklenen $beklenen" >&2; exit 1; }
# NotAutomatic + ButAutomaticUpgrades: depo önceliği 100 (dağıtım paketlerinin önüne geçmez).
apt-cache policy officehere | grep -q " 100 http://127.0.0.1:$port" || { echo "Depo önceliği 100 değil" >&2; exit 1; }

# Olumsuz sınamalarda önbellekteki eski dizin kullanılmasın ve imza uyarısı hata sayılsın
# (apt, eski dizin varken imza hatasında yalnız uyarır ve 0 döner).
olumsuz=("${yalniz_bu[@]}" --error-on=any)

echo "== Olumsuz sınama: başka anahtar"
gpg --batch --yes --export "$(anahtar_uret "Yabanci <yabanci@officehere.invalid>")" > /etc/apt/keyrings/officehere.pgp
rm -rf /var/lib/apt/lists/127.0.0.1_*
if apt-get update "${olumsuz[@]}" >"$is/yabanci.log" 2>&1; then
  cat "$is/yabanci.log" >&2
  echo "Başka anahtarla imzalı depo KABUL EDİLDİ: Signed-By denetimi çalışmıyor" >&2
  exit 1
fi
grep -m1 -E "NO_PUBKEY|Missing key|not signed" "$is/yabanci.log" || { cat "$is/yabanci.log" >&2; echo "Beklenen imza hatası yok" >&2; exit 1; }
install -m 0644 "$is/depo/officehere.pgp" /etc/apt/keyrings/officehere.pgp

echo "== Olumsuz sınama: imzadan sonra değiştirilmiş InRelease"
cp "$is/depo/dists/stable/InRelease" "$is/InRelease.ozgun"
sed -i 's/^Label: OfficeHere$/Label: Degisti/' "$is/depo/dists/stable/InRelease"
rm -rf /var/lib/apt/lists/127.0.0.1_*
if apt-get update "${olumsuz[@]}" >"$is/bozuk.log" 2>&1; then
  cat "$is/bozuk.log" >&2
  echo "Bozulmuş imza KABUL EDİLDİ" >&2
  exit 1
fi
grep -m1 -E "BADSIG|[Bb]ad signature|verification failed|Signature by" "$is/bozuk.log" || { cat "$is/bozuk.log" >&2; echo "Beklenen imza hatası yok" >&2; exit 1; }
cp "$is/InRelease.ozgun" "$is/depo/dists/stable/InRelease"

echo "apt deposu sınaması geçti: officehere $kurulu imzalı depodan kuruldu; yabancı anahtar ve bozuk imza reddedildi."
