# 📧 UnixPadel IT Toplu Mail Gönderim Sistemi — Kullanım Kılavuzu

Bu kılavuz, **UnixPadel Mail Services (v5.2)** uygulamasını kullanarak e-postalarınızı nasıl kolay, güvenli ve spama düşmeden göndereceğinizi anlatmaktadır. Kılavuz, tamamen günlük kullanıma uygun ve anlaşılır şekilde hazırlanmıştır.


## LOGLARA ULAŞMAK İÇİN:
   C:\Users\Lenovo_It_2\AppData\Local\UnixpadelITMailServices yolundaki tracking_log.json

---

## 1. Uygulamaya Giriş

1. **Uygulamayı Açın:** Bilgisayarınızdaki **UnixPadel Mail Services** programına çift tıklayın.
2. **Giriş Yapın:** Size verilen kullanıcı adı ve şifreyi girerek **GİRİŞ YAP** butonuna tıklayın. Her seferinde şifre girmemek için "Beni Hatırla" seçeneğini kullanabilirsiniz.

---

## 2. Gönderim Öncesi Hazırlık (Sunucu Ayarları)

Uygulamanın sol menüsünden **⚙ Sunucu Ayarları** sekmesine tıklayarak e-posta hesabınızı bağlayın. Bu işlemi sadece ilk kullanımda veya e-posta şifreniz değiştiğinde yapmanız yeterlidir.

1. E-posta adresinizi ve e-posta şifrenizi ilgili alanlara girin.
2. E-postalarınızın spama düşmesini engellemek için sayfanın altındaki **Plain-Text (Düz Metin)** kutusuna e-postanızın yazısız, en sade halini ekleyin.
3. **💾 Kaydet** butonuna basın.
4. Bilgilerinizin doğruluğunu test etmek için **🔌 SMTP Test** ve **📥 IMAP Test** butonlarına tıklayın. Ekranda "Başarılı" mesajını görüyorsanız her şey hazır demektir!

---

## 3. E-Posta Gönderme İşlemi

Gönderim yapmak için sol menüden elinizdeki listeye uygun olan seçeneği tıklayın:
* **📊 Excel Gönderim:** Binlerce kişiye Excel listeniz üzerinden toplu gönderim yapmak için.
* **📝 Manuel Liste:** Sadece birkaç kişiye elle e-posta adreslerini yazarak gönderim yapmak için.

### Adım Adım Gönderim:

1. **Şablonunuzu Seçin:** "HTML Şablonu" yanındaki "Dosya Seç" butonuna tıklayarak göndermek istediğiniz e-postanın görsel tasarımını seçin.
2. **Kişileri Belirleyin:**
   - *Excel Modu:* "Excel Listesi" yanındaki "Dosya Seç" ile e-postaların alt alta bulunduğu (A sütununda) Excel dosyanızı seçin.
   - *Manuel Mod:* E-posta adreslerini kopyalayıp, her satıra bir tane gelecek şekilde büyük boş kutucuğa yapıştırın.
3. **Konu Başlığını Yazın:** E-postanızın konu başlığını girin. *(İpucu: Birden fazla konu başlığı yazarak aralarına virgül koyarsanız, sistem bu konuları sırayla değiştirerek kullanır ve spama düşme riskini iyice azaltır).*
4. **Gönderim Aralığını Seçin:** İki e-posta arasında kaç dakika bekleneceğini sürgü (slider) üzerinden ayarlayın. Spama düşmemek için **en az 3-5 dakika** önerilir.
5. **Onay Verin:** Güvenlik gereği "Elshan Bey'den gönderim onayı alındı" kutucuğunu işaretleyin.
6. **Başlatın:** **▶ İŞLEMİ BAŞLAT** butonuna basarak gönderimi başlatın!

*Alt kısımdaki siyah ekrandan e-postaların başarıyla gidip gitmediğini anlık takip edebilirsiniz. Gönderimi durdurmak isterseniz **⏹ DURDUR** butonunu kullanabilirsiniz.*

---

## 🛡️ Otomatik Anti-Spam (Spam Önleme) Özellikleri

Maillerinizin alıcıların "Gereksiz (Spam)" kutusu yerine doğrudan "Gelen Kutusuna (Inbox)" düşmesi için uygulamamız arka planda şu gelişmiş özellikleri otomatik olarak uygular:

- **Bekleme Süresi (Gönderim Aralığı):** E-postaları aynı saniye içerisinde binlerce kişiye atmak yerine, sizin belirlediğiniz dakika aralıklarıyla yavaş yavaş gönderir. Bu sayede e-posta sağlayıcıları sizi "robot" olarak algılamaz.
- **Değişken Konu Başlıkları:** Konu kısmına virgülle ayırarak birden fazla başlık girdiğinizde, uygulama bu başlıkları rastgele seçerek gönderir. Hep aynı başlıkla e-posta çıkışı olmadığı için spam filtreleri tetiklenmez.
- **Tek Tıkla Abonelikten Çıkma Desteği:** E-postalara otomatik "List-Unsubscribe" (Abonelikten Çıkma) altyapısı ekler. Bu durum Gmail ve Outlook gibi sağlayıcıların gözünde güvenilirliğinizi büyük ölçüde artırır.
- **Akıllı E-Posta Doğrulama:** Excel listenizdeki hatalı, yanlış yazılmış veya geçersiz e-posta adreslerini gönderim yapmadan önce fark edip atlar. Böylece sunucu "itibarınız" korunur.
- **Çift Katmanlı İçerik (Sade Metin Desteği):** E-postaların hem görsel hem de düz metin versiyonunu aynı anda göndererek spam filtrelerinden geçer not almanızı sağlar.
- **Görünmez Takip Kodları:** E-postanın güvenilir görünmesi için maillere özel görünmez güvenlik ve takip kimlikleri (Tracking ID) eklenir.

---

## 💡 İpuçları ve Sık Sorulan Sorular

- **"Bağlantı Hatası" alıyorum:** Sunucu ayarlarınızdaki şifrenin doğru olduğundan emin olun. Şifrenizi güncellediyseniz uygulamada da güncellemeniz gerekir.
- **"too many messages in this connection" hatası:** Mail sunucuları, tek bir bağlantı üzerinden gönderilebilecek mesaj sayısını sınırlar (genellikle 1000). Uygulama artık bunu otomatik yönetir: Sunucu Ayarları'ndaki **"Oturum Başına Maks. Mesaj"** değerine ulaşınca bağlantıyı kendiliğinden kapatıp yeniden açar. Logda `♻ Bağlantı tazeleniyor` satırını gördüğünüzde bu normaldir, gönderim kesintisiz devam eder. Varsayılan değer **100**'dür; sunucunuz daha düşük bir limit uyguluyorsa bu sayıyı azaltın.
- **Logda `🔄 Bağlantı sorunu … yeniden bağlanılıyor` görüyorum:** İnternet veya sunucu kaynaklı anlık bir kopma olmuş demektir. Uygulama otomatik yeniden bağlanır ve o maili tekrar dener — mail kaybolmaz.
- **"Excel dosyasından okunamadı" hatası:** E-posta adreslerinizin Excel dosyasının **ilk sütununda (A sütunu)** olduğundan emin olun.
- **"Klasör bulunamadı" uyarısı:** Uygulama gönderilen mailleri hesabınızın "Sent" (Gönderilmiş Öğeler) klasörüne kaydetmeye çalışır. Eğer kaydedemezse mail gitmiş demektir, sadece bir kopyasını kendinize ayıramamıştır.
- **Gönderimi Kapattığımda Ne Olur?:** Gönderim esnasında programı kapatır veya durdurursanız, o ana kadar gönderilmiş olanlar başarılı sayılır, listedeki geri kalan e-postalara gönderim yapılmaz. Sonrasında kaldığınız yerden devam etmek için Excel dosyanızı düzenleyebilirsiniz.
