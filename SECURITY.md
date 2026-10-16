# Güvenlik ilkesi

## Desteklenen sürümler

| Sürüm | Güvenlik düzeltmesi |
| --- | --- |
| Son kararlı sürüm (0.1.x) | Evet |
| Son beta (0.2.0-beta.N) | Evet |
| Daha eski sürümler | Hayır — lütfen güncelleyin |

## Güvenlik açığı bildirme

Güvenlik açıklarını **herkese açık sorun (issue) olarak açmayın.** Şu yollardan birini kullanın:

1. GitHub'da bu deponun **Security › Report a vulnerability** (özel güvenlik bildirimi) sayfası.
2. E-posta: posta@furkanyildirim.com (konu: "OfficeHere güvenlik").

Bildiriminizde şunlar bulunsun: etkilenen sürüm ve işletim sistemi, açığın kısa tanımı, yeniden
üretme adımları ve varsa kavram kanıtı dosyası (kişisel veri içermeyen). Bildirim alındığında en geç
7 gün içinde yanıt verilir; düzeltme yayımlanınca bildiren kişi (isterse) sürüm notunda anılır.

## Kapsam

Özellikle değerli bildirimler: kötü niyetli bir dosyayı açınca betik çalışması ya da yerel dosyalara
erişim, ağ isteği yapılması (README › Ağ erişimi listesindekiler dışında), dosyayı açınca takılma ya da
bellek tükenmesi, güncelleyicinin imzasız ya da sahte bir paket kurması.

OfficeHere makroları (VBA, LibreOffice Basic) **çalıştırmaz**; makrolu dosyalar yalnız korunarak
kaydedilir. İndirdiğiniz paketin bütünlüğünü README › "İndirilen paketi doğrulama" adımlarıyla
denetleyebilirsiniz.
