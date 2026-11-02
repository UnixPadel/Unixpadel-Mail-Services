# 📧 UnixPadel Mail Services — Kullanım Kılavuzu

Bu kılavuz, **UnixPadel Mail Services** ile toplu ve soğuk e-postalarınızı nasıl güvenli, öngörülebilir ve spama düşmeden göndereceğinizi anlatır.

> Uygulama iki dilde çalışır. Sağ üstteki (giriş ekranı) veya sol menünün altındaki (ana ekran) **Türkçe / English** açılır menüsünden dili değiştirebilirsiniz.

---

## 1. İlk Kurulum

1. **Uygulamayı indirin:** `UnixPadel Mail Services.exe` dosyasını çalıştırın. Kurulum gerektirmez.
2. **Otomatik güncelleme:** Uygulama her açılışta arka planda yeni sürüm olup olmadığına bakar. Yeni sürüm varsa sessizce indirir ve kendini yeniden başlatır — sol alttaki sürüm etiketinde ilerlemeyi görürsünüz.
3. **Giriş yapın:** Kullanıcı adı ve şifrenizi girip **GİRİŞ YAP**'a basın. Her seferinde yazmamak için **Beni Hatırla**'yı işaretleyin.

### Dosyalarınız nerede?

Hepsi tek bir klasörde tutulur:

```
%LOCALAPPDATA%\UnixpadelITMailServices
```

| Dosya | İçerik |
|---|---|
| `settings.json` | Tüm ayarlarınız (sunucu, şifre, gönderim tercihleri) |
| `logs\kampanya_*.log` | Her gönderim kampanyasının tam kaydı — hatalar dahil |
| `tracking_log.jsonl` | Başarılı gönderimlerin takip kayıtları (satır başına bir mail) |

> 💡 Bu yolu ezberlemenize gerek yok: log kutusunun üstündeki **📂 Klasörü Aç** butonu klasörü doğrudan açar.

---

## 2. Sunucu Ayarları & Domain Sağlık Denetimi

Sol menüden **⚙ Sunucu Ayarları**'na gidin.

1. **E-Posta ve Şifre:** En üstteki iki alana e-posta adresinizi ve şifrenizi yazın. Bunlar aşağıdaki SMTP ve IMAP alanlarına otomatik olarak yayılır — yapıştırarak girseniz bile.
   - Şifrenin doğru yazıldığını **👁** butonuyla kontrol edebilirsiniz.
2. **🔌 SMTP Test:** Gönderim bağlantısını sınar.
3. **📥 IMAP Test:** Gönderilenler klasörüne erişimi sınar ve sunucudaki klasör listesini loga yazar (doğru "Sent" adını buradan öğrenirsiniz).
4. **🔍 DNS & Spam Sağlık Testi:** Alan adınızın `SPF`, `DKIM`, `DMARC` ve `MX` kayıtlarını tek tıkla denetler ve **100 üzerinden bir teslim edilebilirlik skoru** verir. Eksik kayıt varsa ne eklemeniz gerektiğini de yazar. 100/100 almak, e-postalarınızın doğrudan gelen kutusuna düşme şansını belirgin şekilde artırır.
5. **💾 Kaydet** ile ayarlarınızı saklayın.

> Port alanına sayı dışında bir şey yazarsanız uygulama kaydederken uyarır ve alanı kırmızıya boyar. Hatalı bir portla gönderim başlatılamaz.

---

## 3. Kişiselleştirme: Spintax ve Değişkenler

### A. Spintax — metin varyasyonu

Aynı metnin yüzlerce kez gitmesi spam filtrelerinin en kolay yakaladığı şeydir. `{Seçenek 1|Seçenek 2|Seçenek 3}` yazarsanız sistem her alıcı için rastgele birini seçer.

* **Konu:** `{Merhaba|Selamlar|İyi günler} {{isim | Sayın Müşterimiz}}`
* **Gövde:** `{2026 projeleriniz|Padel kort yatırımlarınız} için {hazırladığımız özel teklif|fiyat listemiz} ektedir.`

> HTML şablonlarınızdaki CSS ve JavaScript süslü parantezleri (`.baslik{color:red}` gibi) etkilenmez — spintax yalnızca içinde `|` bulunan parantezleri işler.

### B. Değişkenler ve boşluk kurtarma (fallback)

Excel'deki sütun adlarını `{{sütun_adı | Varsayılan Değer}}` şeklinde kullanın. Hücre boşsa varsayılan devreye girer, böylece maillerde asla boşluk kalmaz.

| Yazdığınız | Excel'de değer varsa | Excel'de boşsa |
|---|---|---|
| `{{isim \| Sayın Yetkili}}` | `Ahmet Bey` | `Sayın Yetkili` |
| `{{sirket \| Şirketiniz}}` | `Örnek A.Ş.` | `Şirketiniz` |

İkisini birlikte de kullanabilirsiniz: `{Merhaba|Selamlar} {{isim | Sayın Yetkili}},`

---

## 4. Excel Dosyanız Nasıl Olmalı?

**E-posta adresleri A sütununda olmalıdır.** Diğer bütün sütunlar, başlık adlarıyla birebir merge tag olarak kullanılabilir.

| A (e-posta) | isim | sirket | sehir |
|---|---|---|---|
| ahmet@ornek.com | Ahmet Bey | Örnek A.Ş. | İstanbul |
| info@deneme.com |  | Deneme Ltd. | Ankara |

Bu dosyayla `Sayın {{isim | Yetkili}}, {{sehir}} ofisiniz için...` yazarsanız ikinci satıra `Sayın Yetkili, Ankara ofisiniz için...` gider.

Dosyayı seçtiğiniz anda uygulama listeyi çözümler ve dosya adının yanında şunu gösterir:

```
✓ teklif_listesi.xlsx — 412 geçerli, 18 tekrar, 7 geçersiz · kullanılabilir sütunlar: isim, sirket, sehir
```

* **Tekrar eden adresler** otomatik olarak ayıklanır — aynı kişiye iki kez mail gitmez.
* **Geçersiz adresler** gönderimden önce listeden çıkarılır.
* **Kullanılabilir sütunlar**, `{{...}}` içinde yazabileceğiniz adlardır.

Alternatif olarak **Liste kaynağı**nı **Manuel liste** yapıp her satıra bir adres yapıştırabilirsiniz (bu modda `{{sütun}}` değişkenleri kullanılamaz, yalnızca spintax çalışır). Excel ve manuel liste artık aynı sayfadadır; radyo düğmesiyle geçiş yaparsınız ve yazdıklarınız sayfa değiştirseniz de korunur.

---

## 5. Gönderim Kontrolleri & Soğuk Mail Ayarları

**Gelişmiş Anti-Spam & Gönderim Ayarları** bölümündeki seçenekler:

1. **🎲 Akıllı Rastgele Gecikme (Jitter ±%25):** Belirlediğiniz süreye insan davranışı dalgalanması katar (120 sn seçilirse 90–150 sn arası bekler).
2. **🛡️ Canlı MX Doğrulaması:** Listedeki adreslerin alan adında gerçekten posta sunucusu olup olmadığını kontrol eder, olmayanları atlar. Böylece hard-bounce yemez, domain itibarınız korunur.
   *DNS'e ulaşılamazsa adres atlanmaz — geçici bir ağ sorunu yüzünden listenizin çöpe gitmesini engellemek için sistem "şüphede kalırsa gönder" davranır.*
3. **⏰ Sadece Mesai Saatlerinde Gönder (09:00 – 18:00):** Mesai bitince gönderim duraklar, sabah 09:00'da kaldığı yerden devam eder. Durum çubuğunda "⏰ Mesai saati bekleniyor" yazar.
4. **📅 Hafta Sonu Gönderimini Duraklat:** Cumartesi–Pazar durur, Pazartesi otomatik devam eder.
5. **🔥 Günlük Maksimum Limit (Warmup):** Günde en fazla kaç mail gideceğini belirler (örn. `40`). Yanında bugün kaç mail gittiğini de görürsünüz. `0` = limitsiz.
6. **Bülten Modu:** Maillere `List-Unsubscribe` başlığı ekler — büyük sağlayıcılar bunu olumlu bir sinyal olarak okur.
7. **Spam Kelime Kontrolü:** Konuyu ve **gerçekten gönderilecek gövdeyi** (seçtiğiniz HTML şablonu dahil) tarar; riskli kelime bulursa hangi kelimenin nerede geçtiğini gösterip devam etmek isteyip istemediğinizi sorar.

---

## 6. Adım Adım Gönderim

Her şey tek bir **📤 Gönderim** sayfasındadır.

1. **Mesaj formatını seçin:** HTML şablon dosyası seçin **veya** "Düz Metin"e geçip metninizi yazın.
2. **👁 Önizle ile kontrol edin.** Değişkenler ve spintax çözülmüş hâlini gönderimden önce görürsünüz. HTML şablonu tarayıcınızda açılır, düz metin pencerede gösterilir. **🔄 Yenile** ile farklı spintax varyasyonlarını görebilirsiniz.
3. **Kişileri ekleyin:** *Liste kaynağı*nı seçin — Excel dosyası veya manuel liste.
4. **Konu başlığını yazın:** Birden fazla konu için virgülle ayırın; spintax ve `{{değişken}}` de kullanabilirsiniz.
5. **Gönderim aralığını ayarlayın:** kaydırıcıyı sürükleyin ya da hazır düğmeleri kullanın (30 sn · 1 dk · 2 dk · 5 dk). Seçtiğiniz değer kaydedilir.
6. **🧪 Test Gönder ile kendinize bir deneme yapın.** Listeye dokunmaz, günlük limiti harcamaz, mesai saatini beklemez — sadece tek bir mail gönderir. Şablonun gerçekte nasıl göründüğünü görmenin en güvenli yolu budur.
7. **Onay kutusunu işaretleyin:** "Gönderim onayı alındı".
8. **▶ İŞLEMİ BAŞLAT**'a basın.

### Klavye kısayolları

| Kısayol | İşlev |
|---|---|
| `Ctrl` + `Enter` | Gönderimi başlat |
| `Esc` | Gönderimi durdur (kampanya sürerken) |
| `Ctrl` + `S` | Ayarları kaydet |
| `Esc` | Açık pencereyi kapat (özet, önizleme, DNS testi) |
| `Enter` | Giriş ekranında oturum aç |

> Gönderim özeti penceresinde odak bilinçli olarak **Vazgeç** düğmesindedir; refleksle Enter'a basıp yüzlerce maili yola çıkarmanız mümkün değildir.

### Gönderim Özeti ekranı

Butona bastığınızda mailler hemen gitmez — önce bir özet çıkar:

```
📤  Gönderim Özeti
Alıcı           412 geçerli   (18 tekrar çıkarıldı, 7 geçersiz atlandı)
Gönderen        Unixpadel <support@unixpadel.com>
Konu            3 varyasyon
Şablon          teklif_2026.html
Aralık          2 dakika  (jitter ±%25)
Tahmini bitiş   08.09.2026 03:14   (≈ 14 saat 20 dakika)
Günlük limit    40 (bugün gönderilen: 12 → bu kampanyada en fazla 28 gidecek)

                          [ Vazgeç ]   [ Gönderimi Başlat ]
```

Buradaki sayıları kontrol edin. Yanlış bir liste veya yanlış bir şablon seçtiyseniz **tek fark edeceğiniz yer burasıdır**.

### Gönderim sırasında

* İlerleme çubuğunun altında `47 / 412 (%11) · kalan ≈ 12 saat 10 dakika · tahmini bitiş 03:14` yazar.
* Log kutusu her adımı canlı gösterir; aynı kayıt `logs\` klasörüne de yazılır.
* **⏹ DURDUR** ile istediğiniz an durdurabilirsiniz. O ana kadar gönderilenler gönderilmiş sayılır.
* Gönderim sürerken **sekme değiştiremezsiniz** — uygulama sizi uyarır. Bu, kampanyanın yarıda kesilmesini önlemek içindir.
* Pencereyi kapatmaya çalışırsanız kaç kişiye gönderildiğini gösteren bir onay sorusu çıkar.

### Log araç çubuğu

Log kutusunun üstündeki üç buton:

* **💾 Logu Kaydet** — ekrandaki logu istediğiniz yere `.txt` olarak kaydeder.
* **🗑 Temizle** — ekranı temizler (dosyadaki kayıt silinmez).
* **📂 Klasörü Aç** — tüm kampanya loglarının bulunduğu klasörü açar.

Log alanı pencereyle birlikte büyür; pencereyi büyütürseniz daha fazla satır görürsünüz.

---

## 7. Kampanya Geçmişi

Sol menüdeki **📜 Kampanya Geçmişi** sayfası, o güne kadar yaptığınız tüm gönderimleri en yeniden eskiye listeler:

```
09.09.2026  14:20      ✓ 412   ✗ 0   ⚠ 18            [ Aç ]
08.09.2026  09:05      ✓ 120   ✗ 3   ⚠ 5             [ Aç ]
07.09.2026  16:41      Tamamlanmadı  ·  84 satır      [ Aç ]
```

* **Tamamlanmadı**, kampanyanın durdurulduğunu veya uygulamanın kapandığını gösterir.
* **Aç** düğmesi o kampanyanın tam log dosyasını açar — kimlere gönderildiğini buradan görebilirsiniz.

---

## 💡 Sık Karşılaşılan Sorunlar

| Gördüğünüz mesaj | Anlamı ve çözümü |
|---|---|
| **"Kullanıcı adı veya şifre reddedildi"** | E-posta şifreniz yanlış. Ayarlar → *E-Posta ve Şifre* alanını **👁** ile kontrol edin. |
| **"Sunucu adresi bulunamadı"** | *Sunucu* alanı hatalı yazılmış (örn. `mail.unixpadel.com` olmalı). |
| **"Sunucuya ulaşılamadı"** | Port veya güvenlik ayarı yanlış. SMTP için `587 + STARTTLS` ya da `465 + SSL` deneyin. |
| **"Sunucu bağlantıyı reddetti"** | Port numarası yanlış. SMTP: 587 / 465, IMAP: 993. |
| **"Güvenli bağlantı kurulamadı"** | Güvenlik ayarını SSL yerine STARTTLS (veya tersi) yapın. |
| **"❌ SMTP ayarları eksik"** | Sunucu, port, kullanıcı adı veya şifre alanlarından biri boş. |
| **"Port geçerli bir port değil"** | Port alanına sayı dışında bir şey yazılmış. |
| **"⚠ Sent klasörü bulunamadı"** | *Sent Klasör* adı yanlış. **📥 IMAP Test**'i çalıştırın; doğru klasör adı loga yazılır (`INBOX.Sent`, `Sent`, `Gönderilmiş Öğeler` …). |
| **"too many messages in this connection"** | Normaldir. Uygulama *Oturum Başına Maks. Mesaj* değerine ulaşınca bağlantıyı kendisi kapatıp yeniden açar. |
| **"❌ Excel okunamadı"** | Dosya açık olabilir; Excel'de kapatıp tekrar deneyin. E-postaların **A sütununda** olduğundan emin olun. |
| **"DNS kütüphanesi yüklü değil"** | Kaynaktan çalıştırıyorsanız `pip install -r requirements.txt` yapın. MX doğrulaması ve DNS testi bu paketi gerektirir. |

**Diğer sorular**

* **Gönderimi ortasında kapatırsam ne olur?** O ana kadar gönderilenler gitmiştir, kalanlara gönderim yapılmaz. Kimlere gittiğini `logs\` klasöründeki kampanya dosyasından görebilirsiniz. Uygulama kaldığı yerden **devam edemez** — kalan listeyi yeni bir dosyaya alıp yeniden başlatmanız gerekir.
* **Aynı kişiye iki kez mail gider mi?** Tek bir kampanya içinde hayır — tekrar eden adresler otomatik ayıklanır. Farklı kampanyalarda kontrol sizde.
* **Ayarlarım nerede saklanıyor?** `%LOCALAPPDATA%\UnixpadelITMailServices\settings.json`
