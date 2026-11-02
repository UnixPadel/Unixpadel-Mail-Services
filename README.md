# UnixPadel Mail Services

Windows masaüstü uygulaması: Excel listelerinden veya elle girilen adreslerden **toplu ve soğuk e-posta** gönderir. Kişiselleştirme (merge tag + spintax), teslim edilebilirlik denetimi (SPF/DKIM/DMARC/MX) ve soğuk mail disiplini (jitter, mesai penceresi, günlük warmup limiti) yerleşik gelir.

Python + [CustomTkinter](https://customtkinter.tomschimansky.com/) ile yazılmış, PyInstaller ile tek `.exe` olarak dağıtılır. Arayüz **Türkçe ve İngilizce**dir.

> **Son kullanıcıysanız** bu dosyaya değil, [kullanim_kilavuzu.md](kullanim_kilavuzu.md) dosyasına bakın.

---

## Özellikler

- **Kişiselleştirme** — `{{sütun | varsayılan}}` merge tag'leri ve `{a|b|c}` spintax varyasyonu (konuda ve gövdede)
- **Teslim edilebilirlik** — tek tıkla SPF / DKIM / DMARC / MX denetimi ve 100 üzerinden skor
- **Hard-bounce koruması** — gönderim öncesi canlı MX doğrulaması (DNS'e ulaşılamazsa fail-open)
- **Soğuk mail disiplini** — akıllı jitter (±%25), mesai saati penceresi, hafta sonu duraklatma, günlük warmup limiti
- **Gönderim öncesi özet** — alıcı sayısı, tekrar/geçersiz ayıklaması, tahmini bitiş saati; onaylamadan tek mail gitmez
- **Test gönderimi ve önizleme** — listeye dokunmadan kendinize tek deneme maili; değişkenler ve spintax çözülmüş hâlin önizlemesi
- **Kampanya geçmişi** — geçmiş gönderimler, sonuç sayaçları ve tam log dosyaları uygulama içinden erişilebilir
- **Dayanıklı SMTP** — oturum başına mesaj limitinde proaktif yeniden bağlanma, geçici hatalarda artan beklemeli tekrar deneme
- **IMAP** — her gönderimin bir kopyasını Gönderilenler klasörüne yazar
- **Kampanya logları** — ekranda canlı, diskte kalıcı (hatalar dahil)
- **Otomatik güncelleme** — GitHub Releases üzerinden sessiz indirme ve yeniden başlatma

---

## Kaynaktan çalıştırma

Python 3.11+ gerekir.

```bash
pip install -r requirements.txt
```

```bash
python "mail system unic/UnixPadel Mail Services.py"
```

`dnspython` olmadan da açılır, ancak MX doğrulaması ve DNS sağlık testi devre dışı kalır.

## Exe derleme

```bash
cd "mail system unic" && pyinstaller --noconfirm "UnixPadel Mail Services.spec"
```

Çıktı: `mail system unic/dist/UnixPadel Mail Services.exe`

> `.spec` dosyasındaki `hiddenimports` listesi önemlidir: `dns.resolver` dinamik import edildiği için PyInstaller onu kendiliğinden bulamaz. Listeden çıkarılırsa MX ve DNS özellikleri **yalnızca exe'de** sessizce kapanır — kaynaktan çalıştırırken sorun görünmez.

## Sürüm çıkma

`v` ile başlayan bir etiket push edildiğinde [`.github/workflows/release.yml`](.github/workflows/release.yml) devreye girer: exe'yi derler ve `UnixPadel/Unixpadel-Mail-Services-Releases` deposuna release olarak yayınlar. Uygulamanın otomatik güncelleyicisi bu depoyu dinler.

```bash
git tag v5.2.4 && git push origin v5.2.4
```

Etiketi atmadan önce [`mail system unic/updater.py`](mail%20system%20unic/updater.py) içindeki `APP_VERSION` değerini güncelleyin — güncelleyici karşılaştırmayı bu değerle yapar.

Workflow `RELEASE_TOKEN` secret'ına ihtiyaç duyar (hedef depoya yazma yetkisi olan bir token).

---

## Proje yapısı

```
mail system unic/
  UnixPadel Mail Services.py    Uygulamanın tamamı (arayüz + gönderim motoru)
  updater.py                    Otomatik güncelleme modülü
  UnixPadel Mail Services.spec  PyInstaller yapılandırması
requirements.txt                Sabitlenmiş bağımlılıklar
kullanim_kilavuzu.md            Son kullanıcı kılavuzu
.github/workflows/release.yml   Derleme ve yayınlama
```

## Çalışma zamanı dosyaları

Kullanıcı verileri `%LOCALAPPDATA%\UnixpadelITMailServices` altında tutulur:

| Dosya | İçerik |
|---|---|
| `settings.json` | Sunucu ayarları ve gönderim tercihleri |
| `logs\kampanya_*.log` | Kampanya başına tam log |
| `tracking_log.jsonl` | Başarılı gönderimlerin takip kayıtları (satır başına bir JSON) |

---

## Bilinen sınırlamalar

- **Kampanya kaldığı yerden devam edemez.** Uygulama gönderim ortasında kapanırsa kalan listeyi elle ayıklamak gerekir. Kimlere gönderildiği Kampanya Geçmişi'ndeki log dosyasından görülebilir.
- **Suppression / abonelikten çıkma listesi yok.** Bülten modu `List-Unsubscribe` başlığı ekler ama gelen çıkış taleplerini uygulama takip etmez.
- **Koyu tema yok.** Renkler kodda sabit hex değerler; `ctk.set_appearance_mode("Light")` sabittir.
- **Kimlik bilgileri kaynak kodda gömülü.** Veritabanı ve varsayılan SMTP bilgileri düz metin olarak kodda duruyor ve dağıtılan exe'nin içine giriyor; giriş doğrulaması şifreleri düz metin karşılaştırıyor. Bu bilgiler fiilen kamuya açık kabul edilmeli ve dışarı taşınmalıdır.
- **Tek dosyalık kod.** Uygulama mantığının tamamı ~3400 satırlık tek bir modülde; çeviri sözlüğü de gömülü.
