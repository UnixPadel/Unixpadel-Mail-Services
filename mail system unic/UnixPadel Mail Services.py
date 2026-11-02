"""
Unixpadel IT Mail Services — v5.2
Centered window • Smaller login/main window • Remember me • Persistent SMTP/IMAP settings
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import os, sys, json, queue, threading, time, smtplib, imaplib, ssl, re, uuid, random, pymysql, socket, urllib.request, webbrowser
from datetime import datetime, timedelta
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate, make_msgid
from email.header import Header

try:
    from updater import APP_VERSION, start_silent_auto_update
except ImportError:
    APP_VERSION = "5.2.0"
    def start_silent_auto_update(status_callback=None, app_instance=None):
        pass

DB_CONFIG = {
    "host": "b5l1shhnklxmq4ogmpjc-mysql.services.clever-cloud.com",
    "port": 21446,
    "user": "udhzqgatlxfof1ji",
    "password": "97PRxh88Uohomd51sVF",
    "database": "b5l1shhnklxmq4ogmpjc",
    "connect_timeout": 8
}

# ── Appearance ────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

# ── Palette ───────────────────────────────────────────────────────────────────
C = {
    "navy": "#0f2554",
    "navy_mid": "#1a3a7a",
    "navy_light": "#2563eb",
    "accent": "#3b82f6",
    "bg": "#f0f4f8",
    "card": "#ffffff",
    "border": "#e2e8f0",
    "text": "#1e293b",
    "muted": "#64748b",
    "green": "#16a34a",
    "green_h": "#15803d",
    "red": "#dc2626",
    "red_h": "#b91c1c",
    "amber": "#d97706",
    "purple": "#7c3aed",
    "purple_h": "#6d28d9",
    "teal": "#0891b2",
    "teal_h": "#0e7490",
    "log_bg": "#0d1117",
    "log_fg": "#4ade80",
}

# ── Fonts ─────────────────────────────────────────────────────────────────────
F = {
    "title": ("Georgia", 26, "bold"),
    "head": ("Georgia", 18, "bold"),
    "subhead": ("Trebuchet MS", 13, "bold"),
    "body": ("Trebuchet MS", 12),
    "small": ("Trebuchet MS", 11),
    "mono": ("Consolas", 11),
    "logo": ("Georgia", 16, "bold"),
    "nav": ("Trebuchet MS", 12),
    "badge": ("Trebuchet MS", 11),  # 10 px küçük metin C["muted"] ile okunmuyordu
}

EMAIL_REGEX = re.compile(r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$")

TRANSLATIONS = {
    "tr": {
        # General / Common
        "app_title": f"Unixpadel IT Mail Services — v{APP_VERSION}",
        "save": "💾  Kaydet",
        "select_file": "Dosya Seç",
        "not_selected": "Henüz seçilmedi",
        "warning": "Uyarı",
        "error": "Hata",
        "success": "Başarılı",
        "yes": "Evet",
        "no": "Hayır",
        
        # Login
        "login_title": "Mail Services",
        "username": "Kullanıcı Adı",
        "password": "Şifre",
        "remember_me": "Beni Hatırla",
        "login_btn": "GİRİŞ YAP",
        "login_empty_err": "Lütfen kullanıcı adı ve şifrenizi girin.",
        "login_fail_err": "Kullanıcı adı veya şifre yanlış.",
        "db_err": "Veritabanına bağlanılamadı:\n{}",
        
        # Sidebar
        "menu": "  MENÜ",
        "nav_excel": "📊   Excel Gönderim",
        "nav_manual": "📝   Manuel Liste",
        "nav_settings": "⚙     Sunucu Ayarları",
        "logout": "🚪   Çıkış Yap",
        
        # Sections
        "sec_template_file": "📁  Şablon ve Dosya Seçimi",
        "sec_template_only": "📁  Şablon Seçimi",
        "msg_format": "Mesaj Formatı",
        "fmt_html": "HTML Şablonu",
        "fmt_text": "Düz Metin",
        "your_plain_text": "Düz Metin Mesajınız:",
        "excel_list": "Excel Listesi",
        "manual_list_title": "✉️  E-Posta Listesi — Her satıra bir adres",
        
        # Spintax & Personalization Guide
        "sec_spintax_guide": "💡  Spintax & Kişiselleştirme Rehberi",
        "spintax_guide_text": "{Merhaba|Selamlar|İyi günler} → Konuda veya gövdede her alıcıya rastgele bir kelime seçilir. E-postaların spam filtrelerine takılmasını önler.",
        "fallback_guide_text": "{{isim | Sayın Yetkili}} → Excel'de 'isim' sütunu doluysa ismi yazar; boşsa 'Sayın Yetkili' yazar. Asla boşluk kalmaz.",
        "sec_advanced_opts": "⚙️  Gelişmiş Anti-Spam & Gönderim Ayarları",

        # Send Controls
        "sec_send_ctrl": "⚡  Gönderim Kontrolü",
        "subjects_csv": "Konu Başlıkları (Virgüllü & Spintax Destekli)",
        "send_interval": "Gönderim Aralığı",
        "newsletter_mode": "📰 Bülten Modu (List-Unsubscribe Ekle)",
        "spam_check": "🔎 Spam Kelime Kontrolü Yap",
        "approval_check": "Gönderim onayı alındı",
        "approval_hint": "Bu kutucuk, listeyi ve şablonu kontrol ettiğinizi onaylar. İşaretlemeden gönderim başlatılamaz.",
        "smart_jitter": "🎲 Akıllı Rastgele Gecikme (Jitter ±%25)",
        "verify_mx": "🛡️ Canlı MX Doğrulaması (Hard-Bounce Önleme)",
        "business_hours": "⏰ Sadece Mesai Saatlerinde Gönder (09:00 - 18:00)",
        "pause_weekends": "📅 Hafta Sonu Gönderimini Duraklat",
        "warmup_limit_lbl": "🔥 Günlük Maks. Limit (Warmup):",
        "warmup_unlimited_hint": "(0 = limitsiz)",
        "status_ready": "Hazır",
        "btn_start": "▶   İŞLEMİ BAŞLAT",
        "btn_stop": "⏹   DURDUR",
        "lbl_sent": "✓  Gönderilen: {}",
        "lbl_fail": "✗  Hatalı: {}",
        "lbl_skip": "⚠  Atlanan: {}",
        
        # Settings
        "settings_title": "Sunucu & Gönderim Ayarları",
        "set_basic": "🔑  E-Posta ve Şifre",
        "set_email": "E-Posta Adresi",
        "set_smtp": "🔒  Gelişmiş: SMTP — Giden Posta",
        "set_server": "Sunucu",
        "set_port": "Port",
        "set_sec": "Güvenlik (SSL / STARTTLS / NONE)",
        "set_sender_mail": "Gönderici E-Posta",
        "set_smtp_user": "SMTP Kullanıcı Adı",
        "set_sender_name": "Gönderici Adı",
        "set_reply_to": "Reply-To",
        "set_def_subject": "Varsayılan Konu",
        "set_max_conn": "Oturum Başına Maks. Mesaj",
        
        "set_imap": "📥  Gelişmiş: IMAP — Gönderilenlere Kaydet",
        "set_imap_server": "IMAP Sunucu",
        "set_imap_port": "IMAP Port",
        "set_imap_sec": "IMAP Güvenlik (SSL / STARTTLS / NONE)",
        "set_imap_user": "IMAP Kullanıcı Adı",
        "set_imap_pass": "IMAP Şifre",
        "set_sent_folder": "Sent Klasör",
        "set_save_sent_cb": "Her gönderimden sonra Sent klasörüne kopyala",
        "set_plain_text": "📄  Plain-Text Versiyonu (Spintax & Fallback Destekli)",
        "btn_smtp_test": "🔌  SMTP Test",
        "btn_imap_test": "📥  IMAP Test",
        "btn_dns_health": "🔍  DNS & Spam Sağlık Testi",
        "settings_saved": "✓  Ayarlar kaydedildi",
        "settings_fail": "Kaydedilemedi:\n{}",
        
        # Tests & Alerts
        "test_smtp_ok": "Bağlantı ve giriş başarılı!",
        "test_imap_ok": "Bağlantı başarılı!\nSent klasörü: \"{}\"\n\nKlasör listesi loglara yazıldı.",
        "alert_approval": "Devam etmek için \"Gönderim onayı alındı\" kutucuğunu işaretleyin.",
        "alert_empty_list": "❌ E-posta listesi boş.",
        "alert_no_html": "❌ HTML şablonu seçilmedi. (Düz metin kullanmak için 'Mesaj Formatı'nı 'Düz Metin' yapın).",
        "spam_warn_title": "Spam Filtresi Uyarısı",
        "spam_warn_desc": "Mesajınızda spam filtrelerine takılabilecek kelimeler tespit edildi:\n{}\n\nYine de devam etmek istiyor musunuz?",
        
        # DNS Health Modal
        "dns_modal_title": "Domain & DNS Teslimat Sağlığı Denetimi",
        "dns_scanning": "⏳ DNS kayıtları (SPF, DKIM, DMARC, MX) taranıyor…",
        "dns_score_lbl": "Teslim Edilebilirlik Skoru: {} / 100",
        
        # Logs
        "log_stopped": "⏹ Kullanıcı tarafından durduruldu.",
        "log_invalid": "[{}/{}] ⚠ Girdiğiniz mail geçersiz: {}  → atlandı",
        "log_mx_fail": "[{}/{}] ⚠ Domain MX sunucusu bulunamadı: {}  → atlandı (Hard-bounce önlendi)",
        "log_warmup_cap": "🛑 Günlük gönderim limitine ({}) ulaşıldı! (Bugün gönderilen: {})",
        "log_wait_business": "⏳ Mesai saatleri dışında (09:00 - 18:00 bekleniyor)…",
        "log_wait_weekend": "⏳ Hafta sonu duraklatıldı — Pazartesi 09:00 bekleniyor…",
        "log_refresh": "♻  Bağlantı tazeleniyor (bu oturumda {} mesaj gönderildi)",
        "log_new_smtp": "✓  Yeni SMTP oturumu açıldı",
        "log_conn_err": "[{}/{}] 🔄 Bağlantı sorunu ({}) — yeniden bağlanılıyor…",
        "log_sent": "[{}/{}] ✉  Gönderildi → {}  🔖 {}",
        "log_sent_ok": "[{}/{}] 📥  Sent'e kaydedildi",
        "log_sent_fail": "[{}/{}] ⚠  Sent klasörü bulunamadı — adı kontrol edin",
        "log_imap_fail": "[{}/{}] ⚠  IMAP yeniden bağlanamadı: {}",
        "log_err": "[{}/{}] ❌ Hata ({}): {}",
        "log_wait": "⏳ {} bekleniyor…",
        "log_wait_jitter": "⏳ {} bekleniyor (Akıllı Jitter: ±%25)…",
        "log_done": "🏁 TAMAMLANDI  ·  ✓ {}  ✗ {}  ⚠ {}",
        "log_critical": "❌ KRİTİK HATA: {}",
        "log_stop_req": "⏹ Durdurma isteği alındı…",
        "time_min_sec": "{} dakika {} saniye",
        "time_min": "{} dakika",
        "time_sec": "{} saniye",
        "time_hour_min": "{} saat {} dakika",
        "time_hour": "{} saat",

        # DNS & Teslimat Sağlığı (check_domain_dns_health açıklamaları)
        "close": "Kapat",
        "dns_no_domain": "Gönderici e-posta adresi ayarlı değil.",
        "dns_unavailable": "DNS kütüphanesi (dnspython) yüklü değil — kayıtlar denetlenemedi.",
        "dns_mx_ok": "Posta sunucusu bulundu: {}",
        "dns_mx_fail": "MX kaydı yok — bu domaine mail ulaşamaz. Barındırma sağlayıcınızdan MX kaydı ekletin.",
        "dns_spf_ok": "SPF kaydı mevcut: {}",
        "dns_spf_fail": "SPF kaydı yok — DNS'e şu TXT kaydını ekleyin: v=spf1 mx a ~all",
        "dns_dkim_ok": "DKIM imzası bulundu (selector: {})",
        "dns_dkim_fail": "DKIM kaydı bulunamadı — mail sunucunuzdan DKIM imzalamayı açıp verilen TXT kaydını ekleyin.",
        "dns_dmarc_ok": "DMARC politikası mevcut: {}",
        "dns_dmarc_fail": "DMARC kaydı yok — _dmarc.{} adresine şu TXT kaydını ekleyin: v=DMARC1; p=none; rua=mailto:postmaster@{}",
        "dns_tip_perfect": "✨ Domain yapılandırmanız soğuk mailler için hazır!",
        "dns_tip_issues": "⚠️ Eksik DNS kayıtları e-postalarınızın spama düşme riskini artırır.",

        # Gönderim öncesi özet
        "confirm_title": "📤  Gönderim Özeti",
        "confirm_recipients": "Alıcı",
        "confirm_recipients_val": "{} geçerli   ({} tekrar çıkarıldı, {} geçersiz atlandı)",
        "confirm_sender": "Gönderen",
        "confirm_subject": "Konu",
        "confirm_subject_val": "{} varyasyon",
        "confirm_template": "Şablon",
        "confirm_template_plain": "Düz metin mesajı",
        "confirm_interval": "Aralık",
        "confirm_interval_jitter": "{}  (jitter ±%25)",
        "confirm_eta": "Tahmini bitiş",
        "confirm_eta_val": "{}   (≈ {})",
        "confirm_daily": "Günlük limit",
        "confirm_daily_val": "{} (bugün gönderilen: {} → bu kampanyada en fazla {} gidecek)",
        "confirm_daily_none": "Limitsiz",
        "btn_cancel": "Vazgeç",
        "btn_confirm_send": "Gönderimi Başlat",

        # Durum & ilerleme
        "status_sending": "Gönderiliyor…",
        "status_waiting_business": "⏰ Mesai saati bekleniyor",
        "status_waiting_weekend": "📅 Hafta sonu duraklatıldı",
        "status_limit": "🛑 Günlük limit doldu",
        "status_stopped": "⏹ Durduruldu",
        "status_done": "🏁 Tamamlandı",
        "prog_eta": "{} / {}  (%{})  ·  kalan ≈ {}  ·  tahmini bitiş {}",
        "pct_fmt": "%{}",

        # Log araç çubuğu
        "btn_save_log": "💾  Logu Kaydet",
        "btn_clear_log": "🗑  Temizle",
        "btn_open_folder": "📂  Klasörü Aç",
        "log_saved": "✓  Log kaydedildi: {}",
        "log_save_fail": "Log kaydedilemedi:\n{}",

        # Liste özeti
        "excel_hint": "E-posta adresleri A sütununda olmalı. Diğer sütunlar {{sütun_adı}} ile mailde kullanılabilir.",
        "list_summary": "{} geçerli, {} tekrar, {} geçersiz",
        "list_summary_cols": "{} geçerli, {} tekrar, {} geçersiz  ·  kullanılabilir sütunlar: {}",

        # Anlaşılır hata mesajları
        "err_auth": "Kullanıcı adı veya şifre reddedildi.\nAyarlar → \"E-Posta ve Şifre\" alanlarını kontrol edin.",
        "err_dns": "Sunucu adresi bulunamadı.\n\"Sunucu\" alanını kontrol edin (örn. mail.unixpadel.com).",
        "err_timeout": "Sunucuya ulaşılamadı.\nPort veya güvenlik ayarı (SSL / STARTTLS) yanlış olabilir.",
        "err_refused": "Sunucu bağlantıyı reddetti.\nPort numarasını kontrol edin (SMTP: 587 veya 465, IMAP: 993).",
        "err_ssl": "Güvenli bağlantı kurulamadı.\nGüvenlik ayarını SSL yerine STARTTLS (ya da tersi) olarak deneyin.",
        "err_detail_in_log": "\n\nTeknik ayrıntı loglara yazıldı.",

        # Kilit & kapanış
        "nav_locked": "Gönderim sürerken sekme değiştirilemez.\nÖnce ⏹ DURDUR butonuna basın.",
        "close_title": "Gönderim sürüyor",
        "close_body": "{} alıcıdan {} tanesine gönderildi.\nKapatırsanız kalanlara gönderim yapılmayacak.\n\nYine de kapatılsın mı?",
        "port_invalid": "\"{}\" geçerli bir port değil. 1 ile 65535 arası bir sayı girin.",
        "warmup_invalid": "\"{}\" geçerli bir limit değil. 0 veya pozitif bir tam sayı girin (0 = limitsiz).",

        # Daha önce çeviri dışında kalmış metinler
        "warn_settings_load": "Ayar dosyası okunamadı, varsayılan ayarlara dönüldü.\n\n{}",
        "log_schedule_resumed": "✓ Mesai penceresi açıldı, gönderim devam ediyor…",
        "today_sent_inline": "Bugün gönderilen: {}",
        "err_stopped_by_user": "Kullanıcı tarafından durduruldu",
        "log_smtp_conn_fail": "⚠  SMTP bağlantısı kurulamadı (deneme {}/{}): {}",
        "log_smtp_retry": "⏳ {} sn sonra tekrar denenecek…",
        "log_tracking_fail": "⚠️ Takip kaydı yazılamadı: {}",
        "log_imap_connected": "📥 IMAP bağlandı → \"{}\"",
        "log_imap_none": "⚠  IMAP bağlantısı yok (gönderim devam eder): {}",
        "log_smtp_missing": "❌ SMTP ayarları eksik — Sunucu Ayarları sayfasını kontrol edin.",
        "log_connecting": "🔗 Bağlanılıyor → {}:{} [{}]",
        "log_smtp_login_ok": "✓  SMTP girişi başarılı  ·  oturum başına en fazla {} mesaj",
        "log_smtp_test_start": "🔌 SMTP test → {}:{} [{}]",
        "log_smtp_test_ok": "✓  SMTP bağlantı testi BAŞARILI",
        "log_smtp_test_fail": "❌ SMTP test BAŞARISIZ: {}",
        "log_imap_test_start": "📥 IMAP test başlıyor…",
        "log_imap_test_ok": "✓  IMAP bağlantı testi BAŞARILI — Mevcut klasörler:",
        "log_imap_test_fail": "❌ IMAP test BAŞARISIZ: {}",
        "log_excel_err": "❌ Excel okunamadı: {}",
        "log_html_picked": "📄 HTML şablonu: {}",
        "log_excel_picked": "📊 Excel listesi: {}",
        "badge_spintax": "🎲 Spintax",
        "badge_fallback": "🧩 Değişken / Fallback",
        "show_password": "Şifreyi göster / gizle",
        "loc_subject": "konuda",
        "loc_body": "gövdede",

        # Birleşik gönderim sayfası
        "nav_send": "📤   Gönderim",
        "nav_history": "📜   Kampanya Geçmişi",
        "send_title": "📤  Gönderim",
        "sec_source": "📋  Liste Kaynağı",
        "list_source": "Liste kaynağı",
        "src_excel": "Excel dosyası",
        "src_manual": "Manuel liste",
        "manual_list_hint": "Her satıra bir e-posta adresi. Bu modda {{sütun}} değişkenleri kullanılamaz.",
        "empty_list_hint": "Excel dosyanızı seçin veya manuel listeye geçin.",
        "empty_log": "Gönderim başladığında canlı kayıt burada görünecek.",
        "sender_account": "Gönderen",

        # Gönderim aralığı ön ayarları
        "interval_presets": "Hazır:",

        # Test gönderimi
        "btn_test_send": "🧪  Test Gönder",
        "test_send_title": "Test Gönderimi",
        "test_send_prompt": "Test maili hangi adrese gönderilsin?",
        "test_send_running": "🧪 [TEST] Gönderiliyor → {}",
        "test_send_ok": "Test maili gönderildi:\n{}\n\nGelen kutunuzu kontrol edin.",
        "test_send_log_ok": "🧪 [TEST] ✓ Gönderildi → {}",
        "test_send_log_fail": "🧪 [TEST] ❌ Başarısız: {}",
        "test_send_no_content": "Önce bir HTML şablonu seçin veya düz metin mesajınızı yazın.",

        # Önizleme
        "btn_preview": "👁  Önizle",
        "preview_title": "Mesaj Önizleme",
        "preview_subject": "Konu",
        "preview_body": "Gövde",
        "preview_refresh": "🔄  Yenile",
        "preview_note": "Listenin ilk kaydıyla, değişkenler ve spintax çözülmüş hâli.",
        "preview_note_nolist": "Liste seçilmediği için değişkenler varsayılan değerleriyle gösteriliyor.",
        "preview_html_opened": "HTML önizleme tarayıcınızda açıldı.",
        "preview_html_fail": "Önizleme açılamadı:\n{}",

        # Kampanya geçmişi
        "history_title": "📜  Kampanya Geçmişi",
        "history_empty": "Henüz kampanya yok.\nİlk gönderiminizi yaptığınızda burada listelenecek.",
        "history_result": "✓ {}   ✗ {}   ⚠ {}",
        "history_incomplete": "Tamamlanmadı",
        "history_lines": "{} satır",
        "btn_history_open": "Aç",

        # Tooltip metinleri
        "tip_newsletter": "Maillere List-Unsubscribe başlığı ekler. Gmail ve Outlook bunu olumlu bir sinyal olarak okur.",
        "tip_spam_check": "Konuyu ve gerçekten gönderilecek gövdeyi (HTML şablonu dahil) riskli kelimeler için tarar.",
        "tip_jitter": "Bekleme süresine ±%25 rastgelelik katar; sabit aralık robot gibi görünür.",
        "tip_verify_mx": "Alan adında posta sunucusu olmayan adresleri atlar. DNS'e ulaşılamazsa adres atlanmaz.",
        "tip_business_hours": "09:00 dışında gönderim duraklar, sabah kaldığı yerden devam eder.",
        "tip_pause_weekends": "Cumartesi ve Pazar durur, Pazartesi otomatik devam eder.",
        "tip_warmup": "Günde en fazla kaç mail gideceği. Yeni bir domainde itibar oluşturmak için düşük başlayın.",
        "tip_max_per_conn": "Sunucular tek bağlantıda sınırlı sayıda mesaja izin verir. Bu sayıya ulaşılınca bağlantı yenilenir.",
        "tip_test_send": "Listeye dokunmadan kendinize tek bir deneme maili gönderir.",
        "tip_preview": "Değişkenler ve spintax çözülmüş hâlini gönderimden önce gösterir.",

        "lang_name": "Türkçe"
    },
    "en": {
        # General / Common
        "app_title": f"Unixpadel IT Mail Services — v{APP_VERSION}",
        "save": "💾  Save",
        "select_file": "Select File",
        "not_selected": "Not selected yet",
        "warning": "Warning",
        "error": "Error",
        "success": "Success",
        "yes": "Yes",
        "no": "No",
        
        # Login
        "login_title": "Mail Services",
        "username": "Username",
        "password": "Password",
        "remember_me": "Remember Me",
        "login_btn": "LOGIN",
        "login_empty_err": "Please enter your username and password.",
        "login_fail_err": "Invalid username or password.",
        "db_err": "Could not connect to database:\n{}",
        
        # Sidebar
        "menu": "  MENU",
        "nav_excel": "📊   Excel Campaign",
        "nav_manual": "📝   Manual List",
        "nav_settings": "⚙     Server Settings",
        "logout": "🚪   Logout",
        
        # Sections
        "sec_template_file": "📁  Template and File Selection",
        "sec_template_only": "📁  Template Selection",
        "msg_format": "Message Format",
        "fmt_html": "HTML Template",
        "fmt_text": "Plain Text",
        "your_plain_text": "Your Plain Text Message:",
        "excel_list": "Excel List",
        "manual_list_title": "✉️  Email List — One address per line",
        
        # Spintax & Personalization Guide
        "sec_spintax_guide": "💡  Spintax & Personalization Guide",
        "spintax_guide_text": "{Hello|Hi|Greetings} → Selects a random variation for each recipient in subject or body to prevent spam filtering.",
        "fallback_guide_text": "{{name | Valued Customer}} → Uses recipient name if present in Excel; falls back to 'Valued Customer' if blank.",
        "sec_advanced_opts": "⚙️  Advanced Anti-Spam & Sending Options",

        # Send Controls
        "sec_send_ctrl": "⚡  Send Controls",
        "subjects_csv": "Subjects (Comma separated & Spintax enabled)",
        "send_interval": "Send Interval",
        "newsletter_mode": "📰 Newsletter Mode (Add List-Unsubscribe)",
        "spam_check": "🔎 Check for Spam Words",
        "approval_check": "Sending approval obtained",
        "approval_hint": "This box confirms you have checked the list and the template. Sending cannot start without it.",
        "smart_jitter": "🎲 Smart Random Delay (Jitter ±25%)",
        "verify_mx": "🛡️ Live MX Verification (Prevent Hard-Bounce)",
        "business_hours": "⏰ Send Only During Business Hours (09:00 - 18:00)",
        "pause_weekends": "📅 Pause On Weekends",
        "warmup_limit_lbl": "🔥 Daily Max Limit (Warmup):",
        "warmup_unlimited_hint": "(0 = unlimited)",
        "status_ready": "Ready",
        "btn_start": "▶   START PROCESS",
        "btn_stop": "⏹   STOP",
        "lbl_sent": "✓  Sent: {}",
        "lbl_fail": "✗  Failed: {}",
        "lbl_skip": "⚠  Skipped: {}",
        
        # Settings
        "settings_title": "Server & Sending Settings",
        "set_basic": "🔑  Email and Password",
        "set_email": "Email Address",
        "set_smtp": "🔒  Advanced: SMTP — Outgoing Mail",
        "set_server": "Server",
        "set_port": "Port",
        "set_sec": "Security (SSL / STARTTLS / NONE)",
        "set_sender_mail": "Sender Email",
        "set_smtp_user": "SMTP Username",
        "set_sender_name": "Sender Name",
        "set_reply_to": "Reply-To",
        "set_def_subject": "Default Subject",
        "set_max_conn": "Max Msgs per Connection",
        
        "set_imap": "📥  Advanced: IMAP — Save to Sent",
        "set_imap_server": "IMAP Server",
        "set_imap_port": "IMAP Port",
        "set_imap_sec": "IMAP Security (SSL / STARTTLS / NONE)",
        "set_imap_user": "IMAP Username",
        "set_imap_pass": "IMAP Password",
        "set_sent_folder": "Sent Folder",
        "set_save_sent_cb": "Copy to Sent folder after each send",
        "set_plain_text": "📄  Plain-Text Version (Spintax & Fallback Enabled)",
        "btn_smtp_test": "🔌  SMTP Test",
        "btn_imap_test": "📥  IMAP Test",
        "btn_dns_health": "🔍  DNS & Spam Health Test",
        "settings_saved": "✓  Settings saved",
        "settings_fail": "Failed to save:\n{}",
        
        # Tests & Alerts
        "test_smtp_ok": "Connection and login successful!",
        "test_imap_ok": "Connection successful!\nSent folder: \"{}\"\n\nFolder list written to logs.",
        "alert_approval": "Tick the \"Sending approval obtained\" box to continue.",
        "alert_empty_list": "❌ Email list is empty.",
        "alert_no_html": "❌ HTML template not selected. (To use plain text, change 'Message Format' to 'Plain Text').",
        "spam_warn_title": "Spam Filter Warning",
        "spam_warn_desc": "Words that might trigger spam filters were detected in your message:\n{}\n\nDo you still want to continue?",
        
        # DNS Health Modal
        "dns_modal_title": "Domain & DNS Deliverability Health Audit",
        "dns_scanning": "⏳ Scanning DNS records (SPF, DKIM, DMARC, MX)…",
        "dns_score_lbl": "Deliverability Score: {} / 100",
        
        # Logs
        "log_stopped": "⏹ Stopped by user.",
        "log_invalid": "[{}/{}] ⚠ Invalid email entered: {}  → skipped",
        "log_mx_fail": "[{}/{}] ⚠ No valid MX record for: {}  → skipped (Prevented hard-bounce)",
        "log_warmup_cap": "🛑 Daily sending limit ({}) reached! (Sent today: {})",
        "log_wait_business": "⏳ Outside business hours (waiting for 09:00 - 18:00)…",
        "log_wait_weekend": "⏳ Paused for weekend — waiting for Monday 09:00…",
        "log_refresh": "♻  Refreshing connection ({} messages sent in this session)",
        "log_new_smtp": "✓  New SMTP session opened",
        "log_conn_err": "[{}/{}] 🔄 Connection issue ({}) — reconnecting…",
        "log_sent": "[{}/{}] ✉  Sent → {}  🔖 {}",
        "log_sent_ok": "[{}/{}] 📥  Saved to Sent",
        "log_sent_fail": "[{}/{}] ⚠  Sent folder not found — check the name",
        "log_imap_fail": "[{}/{}] ⚠  IMAP failed to reconnect: {}",
        "log_err": "[{}/{}] ❌ Error ({}): {}",
        "log_wait": "⏳ Waiting {}…",
        "log_wait_jitter": "⏳ Waiting {} (Smart Jitter: ±25%)…",
        "log_done": "🏁 COMPLETED  ·  ✓ {}  ✗ {}  ⚠ {}",
        "log_critical": "❌ CRITICAL ERROR: {}",
        "log_stop_req": "⏹ Stop request received…",
        "time_min_sec": "{} minute(s) {} second(s)",
        "time_min": "{} minute(s)",
        "time_sec": "{} second(s)",
        "time_hour_min": "{} hour(s) {} minute(s)",
        "time_hour": "{} hour(s)",

        # DNS & Deliverability Health (check_domain_dns_health descriptions)
        "close": "Close",
        "dns_no_domain": "No sender email address is configured.",
        "dns_unavailable": "The DNS library (dnspython) is not installed — records could not be checked.",
        "dns_mx_ok": "Mail server found: {}",
        "dns_mx_fail": "No MX record — mail cannot reach this domain. Ask your hosting provider to add one.",
        "dns_spf_ok": "SPF record present: {}",
        "dns_spf_fail": "No SPF record — add this TXT record to your DNS: v=spf1 mx a ~all",
        "dns_dkim_ok": "DKIM signature found (selector: {})",
        "dns_dkim_fail": "No DKIM record found — enable DKIM signing on your mail server and publish the TXT record it gives you.",
        "dns_dmarc_ok": "DMARC policy present: {}",
        "dns_dmarc_fail": "No DMARC record — add this TXT record at _dmarc.{}: v=DMARC1; p=none; rua=mailto:postmaster@{}",
        "dns_tip_perfect": "✨ Your domain configuration is ready for cold email!",
        "dns_tip_issues": "⚠️ Missing DNS records increase the risk of your emails landing in spam.",

        # Pre-send summary
        "confirm_title": "📤  Send Summary",
        "confirm_recipients": "Recipients",
        "confirm_recipients_val": "{} valid   ({} duplicates removed, {} invalid skipped)",
        "confirm_sender": "Sender",
        "confirm_subject": "Subject",
        "confirm_subject_val": "{} variation(s)",
        "confirm_template": "Template",
        "confirm_template_plain": "Plain text message",
        "confirm_interval": "Interval",
        "confirm_interval_jitter": "{}  (jitter ±25%)",
        "confirm_eta": "Estimated finish",
        "confirm_eta_val": "{}   (≈ {})",
        "confirm_daily": "Daily limit",
        "confirm_daily_val": "{} (sent today: {} → at most {} will go out in this campaign)",
        "confirm_daily_none": "Unlimited",
        "btn_cancel": "Cancel",
        "btn_confirm_send": "Start Sending",

        # Status & progress
        "status_sending": "Sending…",
        "status_waiting_business": "⏰ Waiting for business hours",
        "status_waiting_weekend": "📅 Paused for the weekend",
        "status_limit": "🛑 Daily limit reached",
        "status_stopped": "⏹ Stopped",
        "status_done": "🏁 Completed",
        "prog_eta": "{} / {}  ({}%)  ·  ≈ {} left  ·  finishes around {}",
        "pct_fmt": "{}%",

        # Log toolbar
        "btn_save_log": "💾  Save Log",
        "btn_clear_log": "🗑  Clear",
        "btn_open_folder": "📂  Open Folder",
        "log_saved": "✓  Log saved: {}",
        "log_save_fail": "Could not save log:\n{}",

        # List summary
        "excel_hint": "Email addresses must be in column A. Other columns can be used in the mail as {{column_name}}.",
        "list_summary": "{} valid, {} duplicate, {} invalid",
        "list_summary_cols": "{} valid, {} duplicate, {} invalid  ·  available columns: {}",

        # Human-readable errors
        "err_auth": "The username or password was rejected.\nCheck Settings → \"Email and Password\".",
        "err_dns": "Server address not found.\nCheck the \"Server\" field (e.g. mail.unixpadel.com).",
        "err_timeout": "Could not reach the server.\nThe port or security setting (SSL / STARTTLS) may be wrong.",
        "err_refused": "The server refused the connection.\nCheck the port number (SMTP: 587 or 465, IMAP: 993).",
        "err_ssl": "Could not establish a secure connection.\nTry STARTTLS instead of SSL (or the other way around).",
        "err_detail_in_log": "\n\nTechnical details were written to the log.",

        # Lock & shutdown
        "nav_locked": "You cannot switch tabs while sending.\nPress the ⏹ STOP button first.",
        "close_title": "Sending in progress",
        "close_body": "{} of {} recipients have been sent to.\nIf you close now, the rest will not receive the email.\n\nClose anyway?",
        "port_invalid": "\"{}\" is not a valid port. Enter a number between 1 and 65535.",
        "warmup_invalid": "\"{}\" is not a valid limit. Enter 0 or a positive whole number (0 = unlimited).",

        # Strings that previously bypassed translation
        "warn_settings_load": "The settings file could not be read; defaults were restored.\n\n{}",
        "log_schedule_resumed": "✓ Business-hours window opened, sending resumes…",
        "today_sent_inline": "Sent today: {}",
        "err_stopped_by_user": "Stopped by user",
        "log_smtp_conn_fail": "⚠  Could not open SMTP connection (attempt {}/{}): {}",
        "log_smtp_retry": "⏳ Retrying in {} s…",
        "log_tracking_fail": "⚠️ Could not write tracking record: {}",
        "log_imap_connected": "📥 IMAP connected → \"{}\"",
        "log_imap_none": "⚠  No IMAP connection (sending continues): {}",
        "log_smtp_missing": "❌ SMTP settings incomplete — check the Server Settings page.",
        "log_connecting": "🔗 Connecting → {}:{} [{}]",
        "log_smtp_login_ok": "✓  SMTP login successful  ·  up to {} messages per session",
        "log_smtp_test_start": "🔌 SMTP test → {}:{} [{}]",
        "log_smtp_test_ok": "✓  SMTP connection test SUCCEEDED",
        "log_smtp_test_fail": "❌ SMTP test FAILED: {}",
        "log_imap_test_start": "📥 Starting IMAP test…",
        "log_imap_test_ok": "✓  IMAP connection test SUCCEEDED — available folders:",
        "log_imap_test_fail": "❌ IMAP test FAILED: {}",
        "log_excel_err": "❌ Could not read Excel file: {}",
        "log_html_picked": "📄 HTML template: {}",
        "log_excel_picked": "📊 Excel list: {}",
        "badge_spintax": "🎲 Spintax",
        "badge_fallback": "🧩 Variable / Fallback",
        "show_password": "Show / hide password",
        "loc_subject": "in subject",
        "loc_body": "in body",

        # Unified send page
        "nav_send": "📤   Send",
        "nav_history": "📜   Campaign History",
        "send_title": "📤  Send",
        "sec_source": "📋  List Source",
        "list_source": "List source",
        "src_excel": "Excel file",
        "src_manual": "Manual list",
        "manual_list_hint": "One email address per line. {{column}} variables are not available in this mode.",
        "empty_list_hint": "Choose your Excel file, or switch to the manual list.",
        "empty_log": "Live progress will appear here once sending starts.",
        "sender_account": "Sender",

        # Send interval presets
        "interval_presets": "Presets:",

        # Test send
        "btn_test_send": "🧪  Send Test",
        "test_send_title": "Test Send",
        "test_send_prompt": "Which address should the test email go to?",
        "test_send_running": "🧪 [TEST] Sending → {}",
        "test_send_ok": "Test email sent to:\n{}\n\nCheck your inbox.",
        "test_send_log_ok": "🧪 [TEST] ✓ Sent → {}",
        "test_send_log_fail": "🧪 [TEST] ❌ Failed: {}",
        "test_send_no_content": "Select an HTML template first, or write your plain text message.",

        # Preview
        "btn_preview": "👁  Preview",
        "preview_title": "Message Preview",
        "preview_subject": "Subject",
        "preview_body": "Body",
        "preview_refresh": "🔄  Refresh",
        "preview_note": "Rendered with the first record of your list — variables and spintax resolved.",
        "preview_note_nolist": "No list selected, so variables are shown with their default values.",
        "preview_html_opened": "The HTML preview was opened in your browser.",
        "preview_html_fail": "Could not open the preview:\n{}",

        # Campaign history
        "history_title": "📜  Campaign History",
        "history_empty": "No campaigns yet.\nYour first send will be listed here.",
        "history_result": "✓ {}   ✗ {}   ⚠ {}",
        "history_incomplete": "Did not finish",
        "history_lines": "{} lines",
        "btn_history_open": "Open",

        # Tooltip texts
        "tip_newsletter": "Adds a List-Unsubscribe header. Gmail and Outlook read this as a positive signal.",
        "tip_spam_check": "Scans the subject and the body that will actually be sent (HTML template included) for risky words.",
        "tip_jitter": "Adds ±25% randomness to the wait; a fixed interval looks robotic.",
        "tip_verify_mx": "Skips addresses whose domain has no mail server. Nothing is skipped if DNS is unreachable.",
        "tip_business_hours": "Sending pauses outside 09:00–18:00 and resumes in the morning where it left off.",
        "tip_pause_weekends": "Stops on Saturday and Sunday, resumes automatically on Monday.",
        "tip_warmup": "Maximum emails per day. Start low on a new domain to build reputation.",
        "tip_max_per_conn": "Servers allow a limited number of messages per connection. The connection is refreshed at this count.",
        "tip_test_send": "Sends a single trial email to yourself without touching the list.",
        "tip_preview": "Shows the message with variables and spintax resolved, before you send.",

        "lang_name": "English"
    }
}

# TR ve EN anahtar kümeleri ayrışırsa arayüz sessizce yarı Türkçe kalıyor
# (bu dosyada geçmişte tam olarak bu oldu). Kaynaktan çalışırken erken uyar.
_TRANSLATION_GAPS = set(TRANSLATIONS["tr"]) ^ set(TRANSLATIONS["en"])
if _TRANSLATION_GAPS and not getattr(sys, "frozen", False):
    print(f"[i18n] Eşleşmeyen çeviri anahtarları: {sorted(_TRANSLATION_GAPS)}", file=sys.stderr)


# ── SMTP oturum yönetimi ──────────────────────────────────────────────────────
# Sunucular tek bir SMTP oturumunda gönderilebilecek mesaj sayısını sınırlar
# (Exim/cPanel varsayılanı: 1000 → "421 too many messages in this connection").
# Limite dayanmadan bağlantıyı proaktif olarak tazeliyoruz.
SMTP_MAX_PER_CONNECTION = 100   # Kaç mesajda bir bağlantı kapatılıp yeniden açılsın
SMTP_RECONNECT_RETRIES = 3      # Yeniden bağlanma deneme sayısı
SMTP_RECONNECT_BACKOFF = 10     # Denemeler arası bekleme (saniye), her denemede katlanır

# Yeniden bağlanmayı hak eden geçici SMTP yanıt kodları
SMTP_TRANSIENT_CODES = (421, 450, 451, 452)

# Yakalanıp değerlendirilecek hata tipleri (smtplib.SMTPException zaten OSError türevidir)
SMTP_RETRYABLE_ERRORS = (smtplib.SMTPException, OSError)

# Spam filtresi uyarı listesi. Sıradan iş kelimeleri ("kampanya", "fırsat",
# "indirim", "offer") bilinçli olarak çıkarıldı — her maili işaretleyen bir uyarı
# kullanıcıyı uyarıyı yok saymaya alıştırır.
SPAM_WORDS = (
    "ücretsiz", "bedava", "kazandınız", "acele edin", "son şans", "hemen tıklayın",
    "risksiz", "garantili kazanç", "kredi kartı numaranız", "şifrenizi girin",
    "free money", "act now", "click here", "risk-free", "guaranteed income",
    "you have won", "winner", "no obligation", "100% free", "verify your account",
)


def valid_email(e: str) -> bool:
    if not e:
        return False

    e = str(e).strip()

    # mailto: varsa temizle
    if e.lower().startswith("mailto:"):
        e = e[7:].strip()

    # sonundaki noktalı virgül / virgül gibi karakterleri temizle
    e = e.rstrip(";, ")

    # içinde boşluk varsa geçersiz say
    if " " in e:
        return False

    return bool(EMAIL_REGEX.match(e))

def normalize_email(e: str) -> str:
    if e is None:
        return ""

    e = str(e).strip()

    if e.lower().startswith("mailto:"):
        e = e[7:].strip()

    e = e.rstrip(";, ")
    return e


# ── DNS (opsiyonel bağımlılık) ────────────────────────────────────────────────
# dnspython kurulu değilse MX/DNS özellikleri "kullanılamıyor" olarak raporlanır
# ve gönderim engellenmez (fail-open).
try:
    import dns.resolver as _dns_resolver
    DNS_AVAILABLE = True
except Exception:
    _dns_resolver = None
    DNS_AVAILABLE = False

DNS_TIMEOUT = 5.0
DKIM_SELECTORS = ("default", "mail", "google", "selector1", "selector2", "k1", "dkim", "s1")

# Spintax regex'i boru işaretini ZORUNLU kılar: HTML şablonlarındaki CSS/JS süslü
# parantezleri (örn. ".a{color:red}") bu sayede bozulmaz.
SPINTAX_RE = re.compile(r"\{([^{}|]*\|[^{}]*)\}")
MERGE_TAG_RE = re.compile(r"\{\{\s*([^{}|]+?)\s*(?:\|\s*([^{}]*?)\s*)?\}\}")

_MX_CACHE = {}
_MX_CACHE_LOCK = threading.Lock()


def apply_merge_tags(text: str, record: dict) -> str:
    """{{sütun | varsayılan}} ve {{sütun}} kalıplarını Excel satırındaki değerle değiştirir.

    Değer yoksa ya da boşsa varsayılana düşer; varsayılan da yoksa boş bırakır.
    Anahtar araması küçük harfe indirgenmiştir (get_emails zaten öyle üretiyor).
    """
    if not text or "{{" not in text:
        return text or ""

    lookup = {}
    if record:
        for k, v in record.items():
            lookup[str(k).strip().lower()] = "" if v is None else str(v)

    def _sub(m):
        key = (m.group(1) or "").strip().lower()
        default = m.group(2) or ""
        value = lookup.get(key, "")
        return value.strip() if value.strip() else default

    return MERGE_TAG_RE.sub(_sub, text)


def resolve_spintax(text: str, rng=random) -> str:
    """{a|b|c} kalıplarından rastgele birini seçer; en içteki gruptan dışa doğru çözer."""
    if not text or "{" not in text:
        return text or ""

    for _ in range(50):  # iç içe spintax için sınırlı döngü
        new_text = SPINTAX_RE.sub(lambda m: rng.choice(m.group(1).split("|")), text)
        if new_text == text:
            break
        text = new_text
    return text


def _dns_resolver_instance():
    r = _dns_resolver.Resolver()
    r.timeout = DNS_TIMEOUT
    r.lifetime = DNS_TIMEOUT
    return r


def _domain_accepts_mail(domain: str) -> bool:
    """True → domain posta kabul edebilir. False → kesinlikle edemez.

    DNS'e ulaşılamıyorsa fail-open (True) döner: tek bir ağ dalgalanması yüzünden
    tüm liste sessizce atlanmasın.
    """
    if not DNS_AVAILABLE:
        return True

    resolver = _dns_resolver_instance()

    try:
        answers = resolver.resolve(domain, "MX")
        if any(str(r.exchange).strip(".") for r in answers):
            return True
    except _dns_resolver.NXDOMAIN:
        return False
    except (_dns_resolver.NoAnswer, _dns_resolver.NoNameservers):
        pass  # MX yok — aşağıda A kaydı fallback'ine bakılır
    except Exception:
        return True  # timeout / ağ hatası → fail-open

    # MX yoksa RFC 5321 gereği A kaydı implicit MX sayılır
    try:
        resolver.resolve(domain, "A")
        return True
    except (_dns_resolver.NXDOMAIN, _dns_resolver.NoAnswer, _dns_resolver.NoNameservers):
        return False
    except Exception:
        return True


def verify_email_mx(email: str) -> bool:
    """Adresin domaini posta kabul edebiliyor mu? Domain bazlı cache'li."""
    email = normalize_email(email)
    if "@" not in email:
        return False

    domain = email.rsplit("@", 1)[-1].strip().lower().strip(".")
    if not domain:
        return False

    with _MX_CACHE_LOCK:
        if domain in _MX_CACHE:
            return _MX_CACHE[domain]

    result = _domain_accepts_mail(domain)

    with _MX_CACHE_LOCK:
        _MX_CACHE[domain] = result
    return result


def _txt_records(resolver, name) -> list:
    try:
        answers = resolver.resolve(name, "TXT")
    except Exception:
        return []

    out = []
    for r in answers:
        try:
            out.append(b"".join(r.strings).decode("utf-8", "replace"))
        except Exception:
            out.append(str(r).strip('"'))
    return out


def check_domain_dns_health(domain: str, t=None) -> dict:
    """Domainin MX / SPF / DKIM / DMARC kayıtlarını denetler.

    {"score": 0-100, "details": [(etiket, tamam_mı, açıklama), ...]} döndürür —
    open_dns_health_modal tam olarak bu yapıyı bekliyor.
    """
    def _s(key, fallback, *args):
        if t:
            try:
                return t(key, *args)
            except Exception:
                pass
        return fallback.format(*args) if args else fallback

    domain = (domain or "").strip().lower().strip(".")
    if not domain:
        return {"score": 0, "details": [("DNS", False, _s("dns_no_domain", "Gönderici e-posta adresi ayarlı değil."))]}

    if not DNS_AVAILABLE:
        return {
            "score": 0,
            "details": [("DNS", False, _s(
                "dns_unavailable",
                "DNS kütüphanesi (dnspython) yüklü değil — kayıtlar denetlenemedi.",
            ))],
        }

    resolver = _dns_resolver_instance()
    details = []
    score = 0

    # ── MX ────────────────────────────────────────────────────────────────────
    hosts = []
    try:
        for r in resolver.resolve(domain, "MX"):
            host = str(r.exchange).strip(".")
            if host:
                hosts.append(host)
    except Exception:
        pass

    if hosts:
        score += 25
        details.append(("MX", True, _s("dns_mx_ok", "Posta sunucusu bulundu: {}", ", ".join(hosts[:2]))))
    else:
        details.append(("MX", False, _s(
            "dns_mx_fail",
            "MX kaydı yok — bu domaine mail ulaşamaz. Barındırma sağlayıcınızdan MX kaydı ekletin.",
        )))

    # ── SPF ───────────────────────────────────────────────────────────────────
    spf = next(
        (v for v in _txt_records(resolver, domain) if v.strip().lower().startswith("v=spf1")),
        None,
    )
    if spf:
        score += 25
        details.append(("SPF", True, _s("dns_spf_ok", "SPF kaydı mevcut: {}", spf[:70])))
    else:
        details.append(("SPF", False, _s(
            "dns_spf_fail",
            "SPF kaydı yok — DNS'e şu TXT kaydını ekleyin: v=spf1 mx a ~all",
        )))

    # ── DKIM ──────────────────────────────────────────────────────────────────
    dkim_selector = None
    for selector in DKIM_SELECTORS:
        for value in _txt_records(resolver, selector + "._domainkey." + domain):
            low = value.lower()
            if "v=dkim1" in low or "p=" in low:
                dkim_selector = selector
                break
        if dkim_selector:
            break

    if dkim_selector:
        score += 25
        details.append(("DKIM", True, _s("dns_dkim_ok", "DKIM imzası bulundu (selector: {})", dkim_selector)))
    else:
        details.append(("DKIM", False, _s(
            "dns_dkim_fail",
            "DKIM kaydı bulunamadı — mail sunucunuzdan DKIM imzalamayı açıp verilen TXT kaydını ekleyin.",
        )))

    # ── DMARC ─────────────────────────────────────────────────────────────────
    dmarc = next(
        (v for v in _txt_records(resolver, "_dmarc." + domain) if v.strip().lower().startswith("v=dmarc1")),
        None,
    )
    if dmarc:
        score += 25
        details.append(("DMARC", True, _s("dns_dmarc_ok", "DMARC politikası mevcut: {}", dmarc[:70])))
    else:
        details.append(("DMARC", False, _s(
            "dns_dmarc_fail",
            "DMARC kaydı yok — _dmarc.{} adresine şu TXT kaydını ekleyin: v=DMARC1; p=none; rua=mailto:postmaster@{}",
            domain, domain,
        )))

    return {"score": score, "details": details}


class UnixpadelApp(ctk.CTk):

    def t(self, key, *args):
        settings = getattr(self, "settings", None)
        lang = "tr"
        if isinstance(settings, dict):
            lang = settings.get("language", "tr")
        if lang not in TRANSLATIONS:
            lang = "tr"
        text = TRANSLATIONS.get(lang, {}).get(key, TRANSLATIONS.get("tr", {}).get(key, key))
        if args:
            try:
                return text.format(*args)
            except Exception:
                return text
        return text

    def __init__(self):
        self.settings = {}
        self.selected_html = None
        self.selected_excel = None
        self._stop_flag = threading.Event()

        # settings.json'a hem ana thread hem worker thread yazabiliyor
        self._settings_lock = threading.Lock()

        # Worker thread'lerin arayüz güncellemeleri buradan geçer (bkz. _ui / _pump_ui)
        self._ui_queue = queue.Queue()

        # Kampanya durumu — sekme kilidi, kapanış onayı ve ETA hesabı için
        self._campaign_running = False
        self._campaign_total = 0
        self._campaign_cur = 0
        self._campaign_sent = 0
        self._campaign_interval = 0
        self._status_key = None
        self._log_file_path = None

        # Pencere genişliğine göre sarılacak metin etiketleri (bkz. _bind_responsive_wrap)
        self._wrap_targets = []

        # Seçilen listenin çözümlenmiş hâli + Excel önbelleği (dosya
        # değişmedikçe her sayfa kurulumunda yeniden okunmasın)
        self._list_stats = None
        self._excel_key = None
        self._excel_rows = []
        self._excel_cols = []

        self.app_dir = os.path.join(os.path.expanduser("~"), "AppData", "Local", "UnixpadelITMailServices")
        os.makedirs(self.app_dir, exist_ok=True)
        self.log_dir = os.path.join(self.app_dir, "logs")
        os.makedirs(self.log_dir, exist_ok=True)
        self.config_path = os.path.join(self.app_dir, "settings.json")
        self.old_config_path = os.path.join(self.app_dir, "config.json")

        self.load_settings()

        super().__init__()
        self.title(self.t("app_title"))
        self.setup_window()

        self.configure(fg_color=C["bg"])
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        if self.settings.get("is_logged_in"):
            self.build_main_ui()
        else:
            self.show_login_screen()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self._bind_shortcuts()

        self._pump_ui()

        # Arka planda sessiz otomatik güncelleme kontrolü
        start_silent_auto_update(status_callback=self.update_version_label, app_instance=self)

    # ──────────────────────── Klavye ──────────────────────────────────────────

    def _bind_shortcuts(self):
        """Uygulamada tek bir klavye bağlaması vardı (şifre alanında Enter);
        her işlem fare gerektiriyordu."""

        def _start(_e=None):
            if hasattr(self, "run_btn") and self.run_btn.winfo_exists() and not self._campaign_running:
                self.start_engine()
            return "break"

        def _stop(_e=None):
            if self._campaign_running:
                self._request_stop()
            return "break"

        def _save(_e=None):
            if hasattr(self, "main_content"):
                self.save_settings()
            return "break"

        for seq, fn in (
            ("<Control-Return>", _start),
            ("<Escape>", _stop),
            ("<Control-s>", _save),
            ("<Control-S>", _save),
        ):
            try:
                self.bind_all(seq, fn)
            except Exception:
                pass

    # ──────────────────────── Thread → arayüz köprüsü ─────────────────────────

    def _ui(self, fn):
        """Bir arayüz güncellemesini ana thread'e devreder.

        Tk API'si thread-güvenli değildir ve worker thread'den after() çağırmak
        da güvenilir değil — iş kuyruğa konur, ana thread _pump_ui ile boşaltır.
        """
        if threading.current_thread() is threading.main_thread():
            try:
                fn()
            except Exception:
                pass
            return

        try:
            self._ui_queue.put_nowait(fn)
        except Exception:
            pass

    def _pump_ui(self):
        """Kuyruktaki arayüz güncellemelerini ana thread'de uygular."""
        while True:
            try:
                fn = self._ui_queue.get_nowait()
            except queue.Empty:
                break
            except Exception:
                break
            try:
                fn()
            except Exception:
                pass  # widget sekme değişiminde yok edilmiş olabilir

        try:
            self.after(50, self._pump_ui)
        except Exception:
            pass

    def update_version_label(self, status_text: str):
        def _update():
            if hasattr(self, 'version_lbl') and self.version_lbl.winfo_exists():
                self.version_lbl.configure(text=status_text)

        self._ui(_update)


    # Dil açılır menüsü kullanıcıya "Türkçe / English" gösterir, ayarlarda kod tutulur.
    @staticmethod
    def _lang_display(code: str) -> str:
        return TRANSLATIONS.get(code, {}).get("lang_name", code)

    @staticmethod
    def _lang_values() -> list:
        return [TRANSLATIONS[c].get("lang_name", c) for c in TRANSLATIONS]

    @staticmethod
    def _lang_code(display: str) -> str:
        for code, table in TRANSLATIONS.items():
            if table.get("lang_name", code) == display or code == display:
                return code
        return "tr"

    def _change_lang_login(self, choice):
        self.settings["language"] = self._lang_code(choice)
        self.save_settings(silent=True)
        self.show_login_screen()

    def _change_lang_main(self, choice):
        # Dil değişimi tüm arayüzü yeniden kuruyor (log_view dahil) — gönderim
        # sürerken buna izin vermek kampanyayı öldürür.
        if self._campaign_running:
            messagebox.showwarning(self.t("warning"), self.t("nav_locked"))
            self.lang_var.set(self._lang_display(self.settings.get("language", "tr")))
            return

        self.settings["language"] = self._lang_code(choice)
        self.save_settings(silent=True)
        self.build_main_ui()

    # ──────────────────────── Window helpers ──────────────────────────────────

    def setup_window(self):
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()

        width = min(1080, screen_w - 140)
        height = min(720, screen_h - 140)

        width = max(width, 880)
        height = max(height, 600)

        self.minsize(880, 600)
        self.center_window(width, height)

    def center_window(self, width=None, height=None):
        self.update_idletasks()

        if width is None:
            width = self.winfo_width()
        if height is None:
            height = self.winfo_height()

        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()

        x = int((screen_w - width) / 2)
        y = int((screen_h - height) / 2)

        if x < 0:
            x = 0
        if y < 0:
            y = 0

        self.geometry(f"{width}x{height}+{x}+{y}")

    # ──────────────────────── Settings ────────────────────────────────────────

    def load_settings(self):
        d = {
            "language": "tr",
            "remember_me": False,
            "remembered_user": "",
            "remembered_pass": "",

            "server": "mail.unixpadel.com",
            "port": "587",
            "security": "STARTTLS",
            "mail": "support@unixpadel.com",
            "pass": "UJKF[hs3Ktzd?QG6",
            "login_username": "support@unixpadel.com",
            "from_name": "Unixpadel",
            "reply_to": "support@unixpadel.com",
            "subject": "Unixpadel Bilgilendirme",
            "max_per_conn": str(SMTP_MAX_PER_CONNECTION),

            "imap_server": "mail.unixpadel.com",
            "imap_port": "993",
            "imap_security": "SSL",
            "imap_username": "support@unixpadel.com",
            "imap_pass": "UJKF[hs3Ktzd?QG6",
            "sent_folder": "INBOX.Sent",
            "save_sent": True,

            # Gönderim aralığı eskiden hiç saklanmıyordu; her sayfa kurulumunda
            # 120 sn'ye dönüyor, kullanıcının seçimi sessizce kayboluyordu.
            "send_interval": "120",
            "list_source": "excel",
            "manual_list": "",

            "smart_jitter": True,
            "verify_mx": True,
            "business_hours": False,
            "pause_weekends": False,
            "warmup_limit": "0",
            "daily_sent_stats": {},

            "plain_text": (
              "Hello,\n\n"

                "I am reaching out regarding your current quotation and 2026 project planning.\n"

"Due to recent raw material cost changes, we may need to update pricing for upcoming orders. If you would like us to review your current quotation or pending approval before any update, you can reply directly to this email."

"Best regards,\n"
"Unixpadel\n"
"support@unixpadel.com"
            ),
        }

        target_path = self.config_path if os.path.exists(self.config_path) else getattr(self, "old_config_path", self.config_path)
        if os.path.exists(target_path):
            try:
                with open(target_path, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                    d.update(loaded)
            except Exception as e:
                messagebox.showwarning(self.t("warning"), self.t("warn_settings_load", e))

        self.settings = d

    def _persist_settings(self):
        """settings sözlüğünü diske yazar.

        Tk widget'larına HİÇ dokunmaz, bu yüzden worker thread'den de güvenle
        çağrılabilir — gönderim sayacı bunu kullanır.
        """
        with self._settings_lock:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(self.settings, f, indent=4, ensure_ascii=False)

    def _flag_invalid(self, attr: str):
        """İlgili giriş kutusunu kırmızıya boyar ve odağı oraya taşır.

        Kullanıcı düzeltmeye başlar başlamaz işaret kalkar — eskiden bir kez
        kırmızıya boyanan alan sayfa yeniden kurulana kadar kırmızı kalıyordu.
        """
        w = getattr(self, attr, None)
        try:
            if w is None or not w.winfo_exists():
                return
            w.configure(border_color=C["red"])
            w.focus_set()
            w.bind("<KeyRelease>", lambda _e, a=attr: self._clear_invalid(a), add="+")
        except Exception:
            pass

    def _clear_invalid(self, *attrs):
        """Hata işaretini kaldırır (kenarlığı normale döndürür)."""
        for attr in attrs:
            w = getattr(self, attr, None)
            try:
                if w is not None and w.winfo_exists():
                    w.configure(border_color=C["border"])
            except Exception:
                pass

    def _validate_numeric_settings(self) -> str:
        """Sayısal ayarları doğrular; sorun varsa gösterilecek metni döndürür.

        Geçersiz bir port eskiden ancak gönderim sırasında ValueError'a dönüşüp
        kampanyayı "KRİTİK HATA" ile öldürüyordu — artık gönderim hiç başlamıyor.
        """
        # Her doğrulamada önce tüm işaretleri temizle; aksi halde düzeltilen alan
        # da kırmızı kalıyordu.
        self._clear_invalid("s_port", "s_imap_port", "warmup_limit_entry")

        for key, widget_attr in (("port", "s_port"), ("imap_port", "s_imap_port")):
            raw = str(self.settings.get(key, "")).strip()
            try:
                if not (1 <= int(raw) <= 65535):
                    raise ValueError
            except (TypeError, ValueError):
                self._flag_invalid(widget_attr)
                return self.t("port_invalid", raw)

        raw = str(self.settings.get("warmup_limit", "0")).strip() or "0"
        try:
            if int(raw) < 0:
                raise ValueError
        except (TypeError, ValueError):
            self._flag_invalid("warmup_limit_entry")
            return self.t("warmup_invalid", raw)

        return ""

    def save_settings(self, silent=False):
        """Ekrandaki değerleri toplayıp diske yazar. YALNIZCA ana thread'den çağrılmalı."""
        try:
            widget_map = {
                "server": "s_server", "port": "s_port", "security": "s_security",
                "mail": "s_mail", "login_username": "s_login_username", "pass": "s_pass",
                "from_name": "s_from_name", "reply_to": "s_reply_to",
                "max_per_conn": "s_max_per_conn", "imap_server": "s_imap_server",
                "imap_port": "s_imap_port", "imap_security": "s_imap_security",
                "imap_username": "s_imap_username", "imap_pass": "s_imap_pass",
                "sent_folder": "s_sent_folder", "easy_mail": "easy_mail", "easy_pass": "easy_pass"
            }
            # winfo_exists() ŞART: sayfa değişince widget yok ediliyor ama öznitelik
            # duruyor. Kontrolsüz .get() çağrısı TclError atıyor ve tüm save_settings
            # sessizce iptal oluyordu — bu yüzden ayarlar kaydedilmeden kayboluyordu.
            for key, attr in widget_map.items():
                w = getattr(self, attr, None)
                if w is not None and w.winfo_exists():
                    self.settings[key] = w.get().strip()

            if hasattr(self, "s_save_sent"):
                self.settings["save_sent"] = bool(self.s_save_sent.get())

            if hasattr(self, "s_plain_text") and self.s_plain_text.winfo_exists():
                self.settings["plain_text"] = self.s_plain_text.get("1.0", "end").strip()

            if hasattr(self, "subject_entry") and self.subject_entry.winfo_exists():
                self.settings["subject"] = self.subject_entry.get().strip()

            if hasattr(self, "template_mode"):
                self.settings["template_mode"] = self.template_mode.get()
            
            if hasattr(self, "direct_plain_text") and self.direct_plain_text.winfo_exists():
                self.settings["direct_plain_text"] = self.direct_plain_text.get("1.0", "end").strip()

            if hasattr(self, "newsletter_var"):
                self.settings["newsletter_mode"] = self.newsletter_var.get()
            
            if hasattr(self, "spam_check_var"):
                self.settings["spam_check"] = self.spam_check_var.get()

            if hasattr(self, "smart_jitter_var"):
                self.settings["smart_jitter"] = bool(self.smart_jitter_var.get())

            if hasattr(self, "verify_mx_var"):
                self.settings["verify_mx"] = bool(self.verify_mx_var.get())

            if hasattr(self, "business_hours_var"):
                self.settings["business_hours"] = bool(self.business_hours_var.get())

            if hasattr(self, "pause_weekends_var"):
                self.settings["pause_weekends"] = bool(self.pause_weekends_var.get())

            if hasattr(self, "warmup_limit_entry") and self.warmup_limit_entry.winfo_exists():
                self.settings["warmup_limit"] = self.warmup_limit_entry.get().strip()

            if hasattr(self, "slider") and self.slider.winfo_exists():
                self.settings["send_interval"] = str(int(self.slider.get()))

            if hasattr(self, "list_source"):
                self.settings["list_source"] = self.list_source.get()

            if hasattr(self, "manual_text") and self.manual_text.winfo_exists():
                self.settings["manual_list"] = self.manual_text.get("1.0", "end").strip()

            if hasattr(self, "remember_me_var"):
                self.settings["remember_me"] = bool(self.remember_me_var.get())

            if self.settings.get("remember_me"):
                if hasattr(self, "user_entry") and self.user_entry.winfo_exists():
                    self.settings["remembered_user"] = self.user_entry.get().strip()
                if hasattr(self, "pass_entry") and self.pass_entry.winfo_exists():
                    self.settings["remembered_pass"] = self.pass_entry.get()
            else:
                self.settings["remembered_user"] = ""
                self.settings["remembered_pass"] = ""

            if hasattr(self, "lang_var"):
                # lang_var görünen adı ("Türkçe") tutar; ayarlara kod yazılmalı
                code = self._lang_code(self.lang_var.get())
                if code != self.settings.get("language", "tr"):
                    self.settings["language"] = code

            # Yazılanı kaybetmemek için önce kaydet, sonra uyar.
            self._persist_settings()

            problem = self._validate_numeric_settings()
            if problem:
                if not silent:
                    messagebox.showerror(self.t("error"), problem)
                return

            self._refresh_sender_label()

            if not silent and hasattr(self, "main_content"):
                self._toast(self.t("settings_saved"), C["green"])

        except Exception as e:
            if not silent:
                messagebox.showerror(self.t("error"), self.t("settings_fail", e))

    # ──────────────────────── Cold Email & Scheduling Helpers ─────────────────

    def _get_today_sent_count(self) -> int:
        today_str = datetime.now().strftime("%Y-%m-%d")
        daily_stats = self.settings.get("daily_sent_stats", {})
        return int(daily_stats.get(today_str, 0))

    def _increment_today_sent_count(self):
        today_str = datetime.now().strftime("%Y-%m-%d")
        daily_stats = self.settings.get("daily_sent_stats", {})
        if len(daily_stats) > 14:
            daily_stats = {k: v for k, v in list(daily_stats.items())[-7:]}
        daily_stats[today_str] = int(daily_stats.get(today_str, 0)) + 1
        self.settings["daily_sent_stats"] = daily_stats
        # Worker thread'den çağrılıyor: save_settings tüm Tk widget'larını okuduğu
        # için buradan çağrılamaz — yalnızca dosyaya yaz.
        try:
            self._persist_settings()
        except Exception:
            pass

    def _check_schedule_window(self) -> tuple:
        now = datetime.now()
        if bool(self.settings.get("pause_weekends", False)) and now.weekday() in (5, 6):
            return False, "weekend"
        if bool(self.settings.get("business_hours", False)):
            if now.hour < 9 or now.hour >= 18:
                return False, "business_hours"
        return True, ""

    def _wait_for_schedule_window(self) -> bool:
        was_waiting = False
        while not self._stop_flag.is_set():
            is_ok, reason = self._check_schedule_window()
            if is_ok:
                if was_waiting:
                    self.log(self.t("log_schedule_resumed"))
                    self._set_status("status_sending")
                return True

            if not was_waiting:
                # Durum etiketine de yaz: kullanıcı uygulamanın donduğunu sanmasın
                if reason == "weekend":
                    self.log(self.t("log_wait_weekend"))
                    self._set_status("status_waiting_weekend")
                else:
                    self.log(self.t("log_wait_business"))
                    self._set_status("status_waiting_business")
                was_waiting = True

            if not self._sleep_interruptible(10):
                return False
        return False

    def _calculate_sleep_interval(self, base_seconds: int) -> int:
        if base_seconds <= 0:
            return 0
        if not bool(self.settings.get("smart_jitter", True)):
            return base_seconds
        jitter_factor = random.uniform(0.75, 1.25)
        return max(5, int(base_seconds * jitter_factor))

    def open_dns_health_modal(self):
        self.save_settings(silent=True)
        from_email = self.settings.get("mail", "").strip()
        domain = from_email.split("@")[-1] if "@" in from_email else "unixpadel.com"

        modal = ctk.CTkToplevel(self)
        modal.title(self.t("dns_modal_title"))
        modal.geometry("560x520")
        modal.resizable(False, False)
        modal.transient(self)
        modal.grab_set()
        modal.bind("<Escape>", lambda _e: modal.destroy())

        self._center_modal(modal)

        top_f = ctk.CTkFrame(modal, fg_color=C["navy"], height=60, corner_radius=0)
        top_f.pack(fill="x")
        ctk.CTkLabel(top_f, text=f"🔍 {domain}", font=F["head"], text_color="white").pack(side="left", padx=20, pady=15)
        
        status_lbl = ctk.CTkLabel(modal, text=self.t("dns_scanning"), font=F["subhead"], text_color=C["accent"])
        status_lbl.pack(pady=(15, 5))

        content_frame = ctk.CTkFrame(modal, fg_color="transparent")
        content_frame.pack(fill="both", expand=True, padx=20, pady=10)

        def _run_audit():
            health = check_domain_dns_health(domain, t=self.t)
            
            def _render():
                if not modal.winfo_exists():
                    return
                score = health["score"]
                status_lbl.configure(
                    text=self.t("dns_score_lbl", score),
                    text_color=C["green"] if score >= 75 else (C["amber"] if score >= 50 else C["red"])
                )

                for item in content_frame.winfo_children():
                    item.destroy()

                for tag, is_ok, desc in health["details"]:
                    row = ctk.CTkFrame(content_frame, fg_color=C["card"], corner_radius=8, border_width=1, border_color=C["border"])
                    row.pack(fill="x", pady=4)
                    
                    color = C["green"] if is_ok else C["red"]
                    
                    badge = ctk.CTkLabel(row, text=f" {tag} ", width=65, font=("Trebuchet MS", 11, "bold"), fg_color=color, text_color="white", corner_radius=6)
                    badge.pack(side="left", padx=10, pady=8)

                    desc_lbl = ctk.CTkLabel(row, text=desc, font=F["small"], text_color=C["text"], anchor="w", wraplength=400, justify="left")
                    desc_lbl.pack(side="left", fill="x", expand=True, padx=(5, 10), pady=8)

                tip_text = self.t("dns_tip_perfect") if score == 100 else self.t("dns_tip_issues")
                ctk.CTkLabel(
                    content_frame, text=tip_text, font=F["badge"], text_color=C["muted"],
                    wraplength=480, justify="left",
                ).pack(pady=(10, 5))

                ctk.CTkButton(
                    content_frame, text=self.t("close"), fg_color=C["navy"], hover_color=C["navy_light"],
                    width=120, height=34, font=F["body"], corner_radius=8, command=modal.destroy,
                ).pack(pady=5)

            # Worker thread'den modal.after() çağırmak güvenilir değil — arayüz
            # kuyruğuna koy, ana thread uygulasın (bkz. _ui / _pump_ui).
            self._ui(_render)

        threading.Thread(target=_run_audit, daemon=True).start()

    # ──────────────────────── Login ───────────────────────────────────────────

    def show_login_screen(self):
        for w in self.winfo_children():
            w.destroy()

        self.configure(fg_color=C["navy"])

        win_w = min(900, self.winfo_screenwidth() - 160)
        win_h = min(620, self.winfo_screenheight() - 160)
        win_w = max(win_w, 760)
        win_h = max(win_h, 520)

        self.center_window(win_w, win_h)

        # Akışkan yerleşim: kart içeriğine göre büyür. Eski sürüm sabit 380×430
        # bir kart içinde mutlak piksel ofsetleri kullanıyordu ve Windows
        # %125/%150 ölçeklemesinde elemanlar üst üste binip kırpılıyordu.
        card = ctk.CTkFrame(self, fg_color=C["card"], corner_radius=20)
        card.place(relx=0.5, rely=0.5, anchor="center")

        strip = ctk.CTkFrame(card, height=70, fg_color=C["navy"], corner_radius=0)
        strip.pack(fill="x")
        strip.pack_propagate(False)

        ctk.CTkLabel(
            strip, text="UNIXPADEL IT", font=("Georgia", 22, "bold"), text_color="white"
        ).pack(side="left", padx=24)

        self.lang_var = ctk.StringVar(value=self._lang_display(self.settings.get("language", "tr")))
        ctk.CTkOptionMenu(
            strip,
            variable=self.lang_var,
            values=self._lang_values(),
            width=100,
            height=28,
            fg_color=C["navy_mid"],
            button_color=C["navy_mid"],
            font=F["badge"],
            command=self._change_lang_login,
        ).pack(side="right", padx=18)

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=46, pady=(20, 26))

        ctk.CTkLabel(
            inner, text=self.t("login_title"), font=("Trebuchet MS", 12), text_color=C["muted"]
        ).pack()
        ctk.CTkFrame(inner, height=1, fg_color=C["border"]).pack(fill="x", pady=(12, 22))

        self.user_entry = ctk.CTkEntry(
            inner, placeholder_text=self.t("username"), width=290, height=38,
            border_color=C["border"], font=F["body"],
        )
        self.user_entry.pack(fill="x", pady=(0, 12))
        # Enter yalnızca şifre alanında çalışıyordu
        self.user_entry.bind("<Return>", lambda _: self.check_login())

        self.pass_entry = ctk.CTkEntry(
            inner, placeholder_text=self.t("password"), width=290, height=38, show="●",
            border_color=C["border"], font=F["body"],
        )
        self.pass_entry.pack(fill="x")
        self.pass_entry.bind("<Return>", lambda _: self.check_login())

        self.remember_me_var = ctk.BooleanVar(value=bool(self.settings.get("remember_me", False)))
        ctk.CTkCheckBox(
            inner,
            text=self.t("remember_me"),
            variable=self.remember_me_var,
            font=F["body"],
            text_color=C["text"],
            fg_color=C["navy"],
            hover_color=C["navy_light"],
            checkmark_color="white",
        ).pack(anchor="w", pady=(16, 20))

        if self.settings.get("remember_me"):
            self.user_entry.insert(0, self.settings.get("remembered_user", ""))
            self.pass_entry.insert(0, self.settings.get("remembered_pass", ""))

        ctk.CTkButton(
            inner,
            text=self.t("login_btn"),
            command=self.check_login,
            fg_color=C["navy"],
            hover_color=C["navy_light"],
            height=42,
            font=("Trebuchet MS", 12, "bold"),
            corner_radius=10,
        ).pack(fill="x")

        ctk.CTkLabel(
            inner, text="© 2026 Unixpadel · IT Department", font=F["badge"], text_color=C["muted"]
        ).pack(pady=(22, 0))

        self.user_entry.focus_set()
        self.after(80, lambda: self.center_window(win_w, win_h))

    def verify_db_login(self, username, password):
        try:
            conn = pymysql.connect(**DB_CONFIG)
            with conn.cursor() as cursor:
                sql = "SELECT 1 FROM KullaniciBilgileri WHERE KullaniciAdi = %s AND Sifre = %s"
                cursor.execute(sql, (username, password))
                result = cursor.fetchone()
            conn.close()
            return result is not None
        except Exception as e:
            messagebox.showerror(self.t("error"), self.t("db_err", self._friendly_error(e)))
            return False

    def check_login(self):
        username = self.user_entry.get().strip()
        password = self.pass_entry.get()

        if not username or not password:
            self.pass_entry.configure(border_color=C["red"])
            messagebox.showerror(self.t("error"), self.t("login_empty_err"))
            return

        if self.verify_db_login(username, password):
            self.settings["is_logged_in"] = True
            self.settings["logged_in_user"] = username
            self.save_settings(silent=True)
            self.configure(fg_color=C["bg"])
            self.build_main_ui()
        else:
            self.pass_entry.configure(border_color=C["red"])
            messagebox.showerror(self.t("error"), self.t("login_fail_err"))

    # ──────────────────────── Main Layout ─────────────────────────────────────

    def build_main_ui(self):
        # Pencere boyutu YALNIZCA ilk kurulumda ayarlanır. Eskiden dil her
        # değiştiğinde kullanıcının boyutlandırdığı pencere 1080×720'ye sıfırlanıyordu.
        first_build = not getattr(self, "_main_ui_sized", False)
        if first_build:
            self.center_window(1080, 720)
            self._main_ui_sized = True

        for w in self.winfo_children():
            w.destroy()

        self.configure(fg_color=C["bg"])

        self._sidebar = ctk.CTkFrame(self, width=210, corner_radius=0, fg_color=C["navy"])
        self._sidebar.grid(row=0, column=0, sticky="nsew")
        self._sidebar.grid_propagate(False)
        self._sidebar.grid_rowconfigure(10, weight=1)

        logo_bar = ctk.CTkFrame(self._sidebar, height=58, fg_color=C["navy_mid"], corner_radius=0)
        logo_bar.pack(fill="x")
        ctk.CTkLabel(logo_bar, text="UNIXPADEL", font=F["logo"], text_color="white").pack(side="left", padx=16, pady=16)

        ctk.CTkFrame(self._sidebar, height=1, fg_color="#1e4080").pack(fill="x")

        ctk.CTkLabel(
            self._sidebar,
            text=self.t("menu"),
            font=("Trebuchet MS", 11, "bold"),
            text_color="#7fa3d4"
        ).pack(anchor="w", padx=18, pady=(14, 6))

        self._nav_btns = {}
        nav_items = [
            ("send", self.t("nav_send"), self.show_send),
            ("history", self.t("nav_history"), self.show_history),
            ("settings", self.t("nav_settings"), self.show_settings),
        ]
        for key, label, cmd in nav_items:
            btn = ctk.CTkButton(
                self._sidebar,
                text=label,
                anchor="w",
                fg_color="transparent",
                hover_color="#1a3a7a",
                font=F["nav"],
                height=38,
                corner_radius=8,
                command=lambda k=key, c=cmd: self._nav_click(k, c)
            )
            btn.pack(fill="x", padx=10, pady=2)
            self._nav_btns[key] = btn

        self.lang_var = ctk.StringVar(value=self._lang_display(self.settings.get("language", "tr")))
        lang_cb_main = ctk.CTkOptionMenu(
            self._sidebar,
            variable=self.lang_var,
            values=self._lang_values(),
            width=110,
            height=28,
            fg_color=C["navy_mid"],
            button_color=C["navy_mid"],
            font=F["badge"],
            command=self._change_lang_main
        )
        lang_cb_main.pack(side="bottom", pady=10)

        # Eski #3d6199 lacivert zemin üzerinde ~2.1:1 kontrast veriyordu (okunmuyordu)
        self.version_lbl = ctk.CTkLabel(self._sidebar, text=f"v{APP_VERSION}", font=F["badge"], text_color="#8fb0dd")
        self.version_lbl.pack(side="bottom", pady=(0, 6))

        # Hangi hesaptan gönderildiği hiçbir ekranda yazmıyordu; yanlış hesaptan
        # gönderim geri alınamaz bir hata, bu yüzden her zaman görünür olmalı.
        self.sender_lbl = ctk.CTkButton(
            self._sidebar,
            text="",
            anchor="w",
            fg_color="transparent",
            hover_color="#1a3a7a",
            text_color="#8fb0dd",
            font=F["badge"],
            height=26,
            corner_radius=6,
            command=lambda: self._nav_click("settings", self.show_settings),
        )
        self.sender_lbl.pack(side="bottom", fill="x", padx=10, pady=(6, 2))
        self._refresh_sender_label()

        ctk.CTkButton(
            self._sidebar,
            text=self.t("logout"),
            anchor="w",
            fg_color="transparent",
            hover_color="#7f1d1d",
            text_color="#f87171",
            font=F["nav"],
            height=38,
            corner_radius=8,
            command=self.logout
        ).pack(side="bottom", fill="x", padx=10, pady=(2, 4))

        self.main_content = ctk.CTkFrame(self, fg_color=C["bg"], corner_radius=0)
        self.main_content.grid(row=0, column=1, sticky="nsew")
        self.main_content.grid_columnconfigure(0, weight=1)
        # Satır 0 toast şeridine ayrıldı; içerik satır 1'de. Eskiden ikisi de
        # satır 0'daydı ve toast kaydırma çerçevesinin üstüne biniyordu.
        self.main_content.grid_rowconfigure(1, weight=1)

        self._toast_lbl = None
        self._nav_click("send", self.show_send)

        if first_build:
            self.after(80, lambda: self.center_window(1080, 720))

    def logout(self):
        if self._campaign_running:
            messagebox.showwarning(self.t("warning"), self.t("nav_locked"))
            return

        self._stop_flag.set()
        self.settings["is_logged_in"] = False
        self.save_settings(silent=True)
        self.after(100, self.show_login_screen)

    def _nav_click(self, key: str, cmd):
        # Sekme değişimi log_view/progress_bar'ı yok ediyor; gönderim sürerken
        # buna izin vermek kampanyayı öldürür.
        if self._campaign_running:
            messagebox.showwarning(self.t("warning"), self.t("nav_locked"))
            return

        # clear_main widget'ları yok ediyor; kaydedilmemiş her şey (SMTP alanları,
        # konu, aralık, manuel liste) sessizce kayboluyordu. save_settings yalnızca
        # var olan widget'ları okur, bu yüzden her geçişte çağrılması güvenli.
        try:
            self.save_settings(silent=True)
        except Exception:
            pass

        for k, b in self._nav_btns.items():
            b.configure(
                fg_color="#1e3d80" if k == key else "transparent",
                font=("Trebuchet MS", 12, "bold") if k == key else F["nav"]
            )
        cmd()

    # ──────────────────────── Scroll Container ────────────────────────────────

    def _scroll(self):
        sf = ctk.CTkScrollableFrame(
            self.main_content,
            fg_color=C["bg"],
            scrollbar_button_color="#cbd5e1",
            scrollbar_button_hover_color="#94a3b8"
        )
        sf.grid(row=1, column=0, sticky="nsew")
        sf.grid_columnconfigure(0, weight=1)
        self._bind_responsive_wrap(sf)
        return sf

    # ──────────────────────── Pages ───────────────────────────────────────────

    def show_send(self):
        """Tek gönderim sayfası.

        Eskiden "Excel Gönderim" ve "Manuel Liste" iki ayrı sekmeydi ve içeriklerinin
        çoğu birebir aynıydı; her sekme geçişi tüm gönderim panelini baştan kurduğu
        için seçilen aralık, konu ve liste sıfırlanıyordu.
        """
        self.clear_main()
        sf = self._scroll()
        self._ptitle(sf, self.t("send_title"))
        self._card(sf, self._section_template)
        self._card(sf, self._section_source)
        self._card(sf, self._section_spintax_guide)
        self._card(sf, self._section_send_controls)
        self._build_log_pane()

    def _build_log_pane(self):
        """Log'u kaydırma çerçevesinin dışında, alt bölmede kurar.

        Sabit yükseklikli bir kutu olarak kaydırma çerçevesi içindeyken pencere
        büyüdüğünde hiç büyümüyordu (ölçüldü: 1080×720 ve 1500×1000'de aynı 263 px).
        """
        pane = ctk.CTkFrame(self.main_content, fg_color=C["card"], corner_radius=0)
        pane.grid(row=2, column=0, sticky="nsew")
        pane.grid_columnconfigure(0, weight=1)
        pane.grid_rowconfigure(1, weight=1)

        bar = ctk.CTkFrame(pane, fg_color="transparent")
        bar.grid(row=0, column=0, sticky="ew", padx=24, pady=(8, 4))

        self.lbl_sent = self._chip(bar, self.t("lbl_sent", 0), C["green"])
        self.lbl_fail = self._chip(bar, self.t("lbl_fail", 0), C["red"])
        self.lbl_skip = self._chip(bar, self.t("lbl_skip", 0), C["amber"])

        for key, cmd in (
            ("btn_open_folder", self._open_log_folder),
            ("btn_clear_log", self._clear_log),
            ("btn_save_log", self._save_log),
        ):
            ctk.CTkButton(
                bar, text=self.t(key), width=120, height=28, corner_radius=8,
                fg_color=C["border"], hover_color="#cbd5e1", text_color=C["text"],
                font=F["small"], command=cmd,
            ).pack(side="right", padx=(6, 0))

        self.log_view = ctk.CTkTextbox(
            pane, fg_color=C["log_bg"], text_color=C["log_fg"], font=F["mono"],
            border_width=0, corner_radius=10,
        )
        self.log_view.grid(row=1, column=0, sticky="nsew", padx=24, pady=(0, 16))

        # Boş durum: simsiyah boş bir dikdörtgen yerine ne olacağını söyle
        self._log_placeholder = ctk.CTkLabel(
            pane, text=self.t("empty_log"), font=F["small"], text_color="#3f5a4a",
        )
        self._show_log_placeholder()

        # Üst kısım kaydırılır, log pencereyle birlikte büyür
        self.main_content.grid_rowconfigure(1, weight=3)
        self.main_content.grid_rowconfigure(2, weight=2, minsize=180)

    def _show_log_placeholder(self):
        lbl = getattr(self, "_log_placeholder", None)
        if lbl is None or not lbl.winfo_exists():
            return
        try:
            empty = not self.log_view.get("1.0", "end").strip()
            if empty:
                lbl.place(in_=self.log_view, relx=0.5, rely=0.5, anchor="center")
            else:
                lbl.place_forget()
        except Exception:
            pass

    def show_history(self):
        """Kampanya logları diske yazılıyordu ama uygulama içinden hiç görülemiyordu."""
        self.clear_main()
        sf = self._scroll()
        self._ptitle(sf, self.t("history_title"))

        entries = self._read_campaign_history()

        if not entries:
            card = self._card_raw(sf)
            ctk.CTkLabel(
                card, text=self.t("history_empty"), font=F["body"], text_color=C["muted"],
                justify="center",
            ).grid(row=0, column=0, padx=24, pady=48)
            return

        for item in entries:
            card = self._card_raw(sf)
            card.grid_columnconfigure(1, weight=1)

            ctk.CTkLabel(
                card, text=item["when"], font=F["subhead"], text_color=C["navy"], anchor="w", width=170,
            ).grid(row=0, column=0, sticky="w", padx=(20, 10), pady=14)

            ctk.CTkLabel(
                card, text=item["summary"], font=F["body"], text_color=C["text"], anchor="w",
            ).grid(row=0, column=1, sticky="w", pady=14)

            ctk.CTkButton(
                card, text=self.t("btn_history_open"), width=80, height=30, corner_radius=8,
                fg_color=C["border"], hover_color="#cbd5e1", text_color=C["text"], font=F["small"],
                command=lambda p=item["path"]: self._open_campaign_log(p),
            ).grid(row=0, column=2, padx=(10, 20), pady=14)

    def _read_campaign_history(self) -> list:
        """logs/ klasöründeki kampanya dosyalarını okuyup özetler (en yeni önce)."""
        entries = []
        try:
            names = sorted(os.listdir(self.log_dir), reverse=True)
        except OSError:
            return entries

        for name in names:
            if not (name.startswith("kampanya_") and name.endswith(".log")):
                continue

            path = os.path.join(self.log_dir, name)
            try:
                with open(path, "r", encoding="utf-8") as f:
                    lines = f.read().splitlines()
            except OSError:
                continue

            stamp = name[len("kampanya_"):-len(".log")]
            try:
                when = datetime.strptime(stamp, "%Y%m%d_%H%M%S").strftime("%d.%m.%Y  %H:%M")
            except ValueError:
                when = stamp

            # Bitiş satırındaki "✓ n  ✗ n  ⚠ n" sayaçlarını ayrıştır
            summary = self.t("history_incomplete") + f"  ·  {self.t('history_lines', len(lines))}"
            for line in reversed(lines):
                if "✓" in line and "✗" in line and "⚠" in line:
                    nums = re.findall(r"[✓✗⚠]\s*(\d+)", line)
                    if len(nums) == 3:
                        summary = self.t("history_result", *nums)
                    break

            entries.append({"when": when, "summary": summary, "path": path})

        return entries

    def _open_campaign_log(self, path: str):
        try:
            os.startfile(path)
        except Exception as e:
            messagebox.showerror(self.t("error"), str(e))

    def show_settings(self):
        self.clear_main()
        sf = self._scroll()
        self._ptitle(sf, self.t("settings_title"))

        def make_collapsible(card, title, color):
            header = ctk.CTkButton(card, text=f"▶  {title}", anchor="w", fg_color="transparent", text_color=color, hover_color=C["border"], font=F["subhead"])
            header.pack(fill="x", padx=12, pady=8)
            content = ctk.CTkFrame(card, fg_color="transparent")
            def toggle():
                if content.winfo_ismapped():
                    content.pack_forget()
                    header.configure(text=f"▶  {title}")
                else:
                    content.pack(fill="x", padx=0, pady=(0, 12))
                    header.configure(text=f"▼  {title}")
            header.configure(command=toggle)
            return content

        # E-Posta ve Şifre
        easy_card = self._card_raw(sf)
        self._section_label(easy_card, self.t("set_basic"), C["accent"])
        easy_content = ctk.CTkFrame(easy_card, fg_color="transparent")
        easy_content.grid(row=1, column=0, columnspan=2, sticky="ew", padx=0, pady=(0, 12))
        self._srow(easy_content, 1, self.t("set_email"), "easy_mail", "easy_mail")
        self._srow(easy_content, 2, self.t("password"), "easy_pass", "easy_pass", is_pass=True)

        # StringVar + trace_add: yapıştırma (sağ tık → yapıştır), temizleme ve
        # programatik değişiklikleri de yakalar. Eski <KeyRelease> bağlaması
        # yalnızca klavyeyi görüyordu; fareyle yapıştıran kullanıcı SMTP/IMAP
        # alanları boş kalmış hâlde sessizce yanlış ayar kaydediyordu.
        self._easy_mail_var = ctk.StringVar(
            value=self.settings.get("easy_mail") or self.settings.get("mail", "")
        )
        self._easy_pass_var = ctk.StringVar(
            value=self.settings.get("easy_pass") or self.settings.get("pass", "")
        )
        self.easy_mail.configure(textvariable=self._easy_mail_var)
        self.easy_pass.configure(textvariable=self._easy_pass_var)

        def _sync_mail(*_):
            self._push_to(("s_mail", "s_login_username", "s_imap_username"),
                          self._easy_mail_var.get().strip())

        def _sync_pass(*_):
            self._push_to(("s_pass", "s_imap_pass"), self._easy_pass_var.get())

        self._easy_mail_var.trace_add("write", _sync_mail)
        self._easy_pass_var.trace_add("write", _sync_pass)

        # SMTP
        smtp_card = self._card_raw(sf)
        smtp_content = make_collapsible(smtp_card, self.t("set_smtp"), C["navy"])
        self._srow(smtp_content, 1, self.t("set_server"), "s_server", "server")
        self._srow(smtp_content, 2, self.t("set_port"), "s_port", "port")
        self._srow(smtp_content, 3, self.t("set_sec"), "s_security", "security")
        self._srow(smtp_content, 4, self.t("set_sender_mail"), "s_mail", "mail")
        self._srow(smtp_content, 5, self.t("set_smtp_user"), "s_login_username", "login_username")
        self._srow(smtp_content, 6, self.t("password"), "s_pass", "pass", is_pass=True)
        self._srow(smtp_content, 7, self.t("set_sender_name"), "s_from_name", "from_name")
        self._srow(smtp_content, 8, self.t("set_reply_to"), "s_reply_to", "reply_to")
        self._srow(smtp_content, 9, self.t("set_def_subject"), "subject_entry", "subject")
        self._srow(smtp_content, 10, self.t("set_max_conn"), "s_max_per_conn", "max_per_conn",
                   tip=self.t("tip_max_per_conn"))

        # IMAP
        imap_card = self._card_raw(sf)
        imap_content = make_collapsible(imap_card, self.t("set_imap"), C["teal"])
        self._srow(imap_content, 1, self.t("set_imap_server"), "s_imap_server", "imap_server")
        self._srow(imap_content, 2, self.t("set_imap_port"), "s_imap_port", "imap_port")
        self._srow(imap_content, 3, self.t("set_imap_sec"), "s_imap_security", "imap_security")
        self._srow(imap_content, 4, self.t("set_imap_user"), "s_imap_username", "imap_username")
        self._srow(imap_content, 5, self.t("set_imap_pass"), "s_imap_pass", "imap_pass", is_pass=True)
        self._srow(imap_content, 6, self.t("set_sent_folder"), "s_sent_folder", "sent_folder")

        self.s_save_sent = ctk.BooleanVar(value=bool(self.settings.get("save_sent", True)))
        ctk.CTkCheckBox(
            imap_content,
            text=self.t("set_save_sent_cb"),
            variable=self.s_save_sent,
            font=F["body"],
            text_color=C["green"],
            checkmark_color="white",
            fg_color=C["green"],
            hover_color=C["green_h"]
        ).grid(row=7, column=0, columnspan=2, sticky="w", padx=24, pady=(6, 10))

        # Plain-Text
        pt_card = self._card_raw(sf)
        pt_content = make_collapsible(pt_card, self.t("set_plain_text"), C["amber"])
        self.s_plain_text = ctk.CTkTextbox(pt_content, height=110, border_width=1, font=F["mono"], border_color=C["border"])
        self.s_plain_text.insert("1.0", self.settings.get("plain_text", ""))
        self.s_plain_text.grid(row=1, column=0, columnspan=2, sticky="ew", padx=24, pady=(4, 16))
        pt_content.grid_columnconfigure(0, weight=1)
        pt_content.grid_columnconfigure(1, weight=1)

        # Buttons
        btn_card = self._card_raw(sf)
        bf = ctk.CTkFrame(btn_card, fg_color="transparent")
        bf.pack(pady=18, padx=24)
        
        # Buton hiyerarşisi: eskiden dördü de aynı ağırlıkta, dört ayrı doygun
        # renkteydi — renk anlam taşımıyor, hepsi eşit derecede öne çıkıyordu.
        # Kaydet birincil (dolu), testler ikincil (çerçeveli).
        self._btn(bf, self.t("save"), C["green"], C["green_h"], self.save_settings).pack(side="left", padx=6)

        for key, cmd in (
            ("btn_smtp_test", self.test_connection),
            ("btn_imap_test", self.test_imap),
            ("btn_dns_health", self.open_dns_health_modal),
        ):
            ctk.CTkButton(
                bf, text=self.t(key), height=40, width=160, corner_radius=9,
                fg_color="transparent", hover_color=C["border"], text_color=C["navy"],
                border_width=1, border_color=C["border"], font=F["body"], command=cmd,
            ).pack(side="left", padx=6)

    # ──────────────────────── Section builders ────────────────────────────────

    def _section_template(self, p):
        self._section_label(p, self.t("sec_template_only"), C["navy"])
        self._build_template_selector(p, row_start=1)
        ctk.CTkLabel(p, text="").grid(row=4)

    def _section_source(self, p):
        """Excel ve manuel listeyi tek kartta toplar (eskiden iki ayrı sekmeydi)."""
        self._section_label(p, self.t("sec_source"), C["navy"])

        self.list_source = ctk.StringVar(value=self.settings.get("list_source", "excel"))

        src_fr = ctk.CTkFrame(p, fg_color="transparent")
        src_fr.grid(row=1, column=0, sticky="ew", padx=24, pady=(2, 8))
        ctk.CTkLabel(
            src_fr, text=self.t("list_source"), width=130, anchor="w",
            font=F["subhead"], text_color=C["text"],
        ).grid(row=0, column=0)
        ctk.CTkRadioButton(
            src_fr, text=self.t("src_excel"), variable=self.list_source, value="excel",
            command=self._toggle_source, font=F["body"],
        ).grid(row=0, column=1, padx=(0, 15))
        ctk.CTkRadioButton(
            src_fr, text=self.t("src_manual"), variable=self.list_source, value="manual",
            command=self._toggle_source, font=F["body"],
        ).grid(row=0, column=2)

        # ── Excel kaynağı ─────────────────────────────────────────────────────
        self.excel_fr = ctk.CTkFrame(p, fg_color="transparent")
        self.excel_fr.grid(row=2, column=0, sticky="ew")
        self.excel_fr.grid_columnconfigure(0, weight=1)

        self._frow(self.excel_fr, 0, self.t("excel_list"), self.load_excel, "excel_label")

        # A sütunu kuralı yalnızca kılavuzda yazıyordu; kullanıcı burada görsün
        self.excel_hint_lbl = ctk.CTkLabel(
            self.excel_fr, text=self.t("excel_hint"), font=F["small"], text_color=C["muted"],
            anchor="w", justify="left", wraplength=620,
        )
        self.excel_hint_lbl.grid(row=1, column=0, sticky="w", padx=24, pady=(0, 12))
        self._wrap_targets.append(self.excel_hint_lbl)

        # ── Manuel kaynak ─────────────────────────────────────────────────────
        self.manual_fr = ctk.CTkFrame(p, fg_color="transparent")
        self.manual_fr.grid(row=3, column=0, sticky="ew")
        self.manual_fr.grid_columnconfigure(0, weight=1)

        self.manual_text = ctk.CTkTextbox(
            self.manual_fr, height=140, border_width=1, font=F["mono"], border_color=C["border"],
        )
        self.manual_text.grid(row=0, column=0, sticky="ew", padx=24, pady=(2, 4))
        self.manual_text.insert("1.0", self.settings.get("manual_list", ""))

        self.manual_hint_lbl = ctk.CTkLabel(
            self.manual_fr, text=self.t("manual_list_hint"), font=F["small"], text_color=C["muted"],
            anchor="w", justify="left", wraplength=620,
        )
        self.manual_hint_lbl.grid(row=1, column=0, sticky="w", padx=24, pady=(0, 12))
        self._wrap_targets.append(self.manual_hint_lbl)

        self._toggle_source()
        self._refresh_list_summary()

    def _toggle_source(self):
        if self.list_source.get() == "manual":
            self.excel_fr.grid_remove()
            self.manual_fr.grid()
        else:
            self.manual_fr.grid_remove()
            self.excel_fr.grid()
            self._refresh_list_summary()

    def _build_template_selector(self, parent, row_start):
        self.template_mode = ctk.StringVar(value=self.settings.get("template_mode", "html"))
        
        mode_fr = ctk.CTkFrame(parent, fg_color="transparent")
        mode_fr.grid(row=row_start, column=0, sticky="ew", padx=24, pady=5)
        mode_fr.grid_columnconfigure(3, weight=1)
        
        ctk.CTkLabel(mode_fr, text=self.t("msg_format"), width=130, anchor="w", font=F["subhead"], text_color=C["text"]).grid(row=0, column=0)
        
        ctk.CTkRadioButton(mode_fr, text=self.t("fmt_html"), variable=self.template_mode, value="html", command=self._toggle_template, font=F["body"]).grid(row=0, column=1, padx=(0, 15))
        ctk.CTkRadioButton(mode_fr, text=self.t("fmt_text"), variable=self.template_mode, value="text", command=self._toggle_template, font=F["body"]).grid(row=0, column=2)
        
        self.html_sel_fr = ctk.CTkFrame(parent, fg_color="transparent")
        self.html_sel_fr.grid(row=row_start+1, column=0, sticky="ew", padx=24, pady=0)
        self.html_sel_fr.grid_columnconfigure(2, weight=1)
        
        ctk.CTkLabel(self.html_sel_fr, text=self.t("fmt_html"), width=130, anchor="w", font=F["subhead"], text_color=C["text"]).grid(row=0, column=0)
        ctk.CTkButton(
            self.html_sel_fr, text=self.t("select_file"), width=110, height=32,
            fg_color=C["navy"], hover_color=C["navy_light"], font=F["small"], corner_radius=8,
            command=self.load_html
        ).grid(row=0, column=1, padx=12)
        
        initial_label_text = self.t("not_selected")
        if hasattr(self, "selected_html") and self.selected_html:
            import os
            initial_label_text = os.path.basename(self.selected_html)
            
        self.html_label = ctk.CTkLabel(self.html_sel_fr, text=initial_label_text, anchor="w", font=F["small"], text_color=C["muted"])
        self.html_label.grid(row=0, column=2, sticky="w")

        preview_btn = ctk.CTkButton(
            self.html_sel_fr, text=self.t("btn_preview"), width=110, height=32, corner_radius=8,
            fg_color=C["border"], hover_color="#cbd5e1", text_color=C["text"], font=F["small"],
            command=self.open_preview,
        )
        preview_btn.grid(row=0, column=3, padx=(8, 0))
        self._tip(preview_btn, self.t("tip_preview"))
        
        self.plain_text_fr = ctk.CTkFrame(parent, fg_color="transparent")
        self.plain_text_fr.grid(row=row_start+2, column=0, sticky="ew", padx=24, pady=0)
        self.plain_text_fr.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(self.plain_text_fr, text=self.t("your_plain_text"), anchor="w", font=F["subhead"], text_color=C["text"]).grid(row=0, column=0, sticky="w", pady=(0, 5))

        pt_preview = ctk.CTkButton(
            self.plain_text_fr, text=self.t("btn_preview"), width=110, height=28, corner_radius=8,
            fg_color=C["border"], hover_color="#cbd5e1", text_color=C["text"], font=F["small"],
            command=self.open_preview,
        )
        pt_preview.grid(row=0, column=1, sticky="e", pady=(0, 5))
        self._tip(pt_preview, self.t("tip_preview"))
        self.direct_plain_text = ctk.CTkTextbox(self.plain_text_fr, height=150, border_width=1, font=F["mono"], border_color=C["border"])
        self.direct_plain_text.grid(row=1, column=0, columnspan=2, sticky="ew")
        self.direct_plain_text.insert("1.0", self.settings.get("direct_plain_text", ""))
        
        self._toggle_template()

    def _toggle_template(self):
        mode = self.template_mode.get()
        if mode == "html":
            self.html_sel_fr.grid()
            self.plain_text_fr.grid_remove()
        else:
            self.html_sel_fr.grid_remove()
            self.plain_text_fr.grid()

    def _section_spintax_guide(self, p):
        self._section_label(p, self.t("sec_spintax_guide"), C["purple"])

        guide_box = ctk.CTkFrame(p, fg_color="#f8fafc", corner_radius=10, border_width=1, border_color="#e2e8f0")
        guide_box.grid(row=1, column=0, columnspan=2, sticky="ew", padx=24, pady=(2, 14))
        guide_box.grid_columnconfigure(1, weight=1)

        # Spintax row
        sp_icon = ctk.CTkLabel(guide_box, text=self.t("badge_spintax"), width=140, font=("Trebuchet MS", 11, "bold"), fg_color="#ede9fe", text_color=C["purple"], corner_radius=6)
        sp_icon.grid(row=0, column=0, padx=(12, 10), pady=(10, 4), sticky="nw")

        sp_text = ctk.CTkLabel(
            guide_box,
            text=self.t("spintax_guide_text"),
            font=F["small"],
            text_color=C["text"],
            justify="left",
            anchor="w",
            wraplength=620,
        )
        self._wrap_targets.append(sp_text)
        sp_text.grid(row=0, column=1, padx=(0, 12), pady=(8, 4), sticky="w")

        # Fallback row
        fb_icon = ctk.CTkLabel(guide_box, text=self.t("badge_fallback"), width=140, font=("Trebuchet MS", 11, "bold"), fg_color="#dbeafe", text_color=C["accent"], corner_radius=6)
        fb_icon.grid(row=1, column=0, padx=(12, 10), pady=(4, 10), sticky="nw")

        fb_text = ctk.CTkLabel(
            guide_box,
            text=self.t("fallback_guide_text"),
            font=F["small"],
            text_color=C["text"],
            justify="left",
            anchor="w",
            wraplength=620,
        )
        self._wrap_targets.append(fb_text)
        fb_text.grid(row=1, column=1, padx=(0, 12), pady=(4, 10), sticky="w")

    def _section_send_controls(self, p):
        p.grid_columnconfigure(0, weight=1)
        self._section_label(p, self.t("sec_send_ctrl"), C["navy"])

        sr = ctk.CTkFrame(p, fg_color="transparent")
        sr.grid(row=1, column=0, sticky="ew", padx=24, pady=4)
        sr.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(sr, text=self.t("subjects_csv"), width=175, anchor="w", font=F["subhead"], text_color=C["text"]).grid(row=0, column=0)
        self.subject_entry = ctk.CTkEntry(sr, height=34, font=F["body"], border_color=C["border"])
        self.subject_entry.insert(0, self.settings.get("subject", "Unixpadel Bilgilendirme"))
        self.subject_entry.grid(row=0, column=1, sticky="ew", padx=(12, 0))

        sl_r = ctk.CTkFrame(p, fg_color="transparent")
        sl_r.grid(row=2, column=0, sticky="ew", padx=24, pady=(10, 4))
        sl_r.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(sl_r, text=self.t("send_interval"), width=140, anchor="w", font=F["subhead"], text_color=C["text"]).grid(row=0, column=0)
        self.slider = ctk.CTkSlider(
            sl_r,
            from_=15,
            to=420,
            number_of_steps=27,
            button_color=C["accent"],
            progress_color=C["navy"],
            command=self._update_slider
        )
        # Ayarlardan oku — eskiden sabit 120 idi ve her sayfa kurulumunda
        # kullanıcının seçtiği değeri siliyordu.
        try:
            saved_interval = int(str(self.settings.get("send_interval", "120")).strip() or "120")
        except (TypeError, ValueError):
            saved_interval = 120
        saved_interval = max(15, min(420, saved_interval))
        self.slider.set(saved_interval)
        self.slider.grid(row=0, column=1, sticky="ew", padx=(12, 0))

        self.l_slider = ctk.CTkLabel(
            sl_r, text=self._fmt_interval(saved_interval), width=140, anchor="e",
            font=F["subhead"], text_color=C["accent"],
        )
        self.l_slider.grid(row=0, column=2, padx=(8, 0))

        # 27 adımlı bir kaydırıcıda tam değer tutturmak zor — sık kullanılanlar hazır
        preset_fr = ctk.CTkFrame(p, fg_color="transparent")
        preset_fr.grid(row=3, column=0, sticky="w", padx=24, pady=(0, 4))
        ctk.CTkLabel(
            preset_fr, text=self.t("interval_presets"), width=140, anchor="w",
            font=F["small"], text_color=C["muted"],
        ).pack(side="left")
        for seconds in (30, 60, 120, 300):
            ctk.CTkButton(
                preset_fr, text=self._fmt_interval(seconds), width=78, height=26, corner_radius=8,
                fg_color=C["border"], hover_color="#cbd5e1", text_color=C["text"], font=F["small"],
                command=lambda s=seconds: self._set_interval(s),
            ).pack(side="left", padx=(6, 0))

        # --- Gelişmiş Anti-Spam ve Gönderim Ayarları (Collapsible Card) ---
        adv_card = ctk.CTkFrame(p, fg_color="#f8fafc", corner_radius=10, border_width=1, border_color="#e2e8f0")
        adv_card.grid(row=4, column=0, sticky="ew", padx=24, pady=(10, 6))
        adv_card.grid_columnconfigure(0, weight=1)

        # Kart her açılışta kapalı başlar: içerik burada pack edilmez,
        # yalnızca başlığa tıklanınca açılır.
        adv_content = ctk.CTkFrame(adv_card, fg_color="transparent")
        adv_content.grid_columnconfigure(0, weight=1)
        adv_content.grid_columnconfigure(1, weight=1)

        def _toggle_adv():
            if adv_content.winfo_ismapped():
                adv_content.pack_forget()
                adv_toggle_btn.configure(text=f"▶  {self.t('sec_advanced_opts')}")
            else:
                adv_content.pack(fill="x", padx=12, pady=(0, 10))
                adv_toggle_btn.configure(text=f"▼  {self.t('sec_advanced_opts')}")

        adv_toggle_btn = ctk.CTkButton(
            adv_card,
            text=f"▶  {self.t('sec_advanced_opts')}",
            anchor="w",
            fg_color="transparent",
            text_color=C["navy"],
            hover_color=C["border"],
            font=F["subhead"],
            height=34,
            command=_toggle_adv
        )
        adv_toggle_btn.pack(fill="x", padx=8, pady=(4, 2))

        self.newsletter_var = ctk.BooleanVar(value=bool(self.settings.get("newsletter_mode", True)))
        self.spam_check_var = ctk.BooleanVar(value=bool(self.settings.get("spam_check", True)))
        self.smart_jitter_var = ctk.BooleanVar(value=bool(self.settings.get("smart_jitter", True)))
        self.verify_mx_var = ctk.BooleanVar(value=bool(self.settings.get("verify_mx", True)))
        self.business_hours_var = ctk.BooleanVar(value=bool(self.settings.get("business_hours", False)))
        self.pause_weekends_var = ctk.BooleanVar(value=bool(self.settings.get("pause_weekends", False)))

        # Her seçeneğin ne yaptığı hiçbir yerde yazmıyordu — üzerine gelince açıklasın
        adv_options = (
            (0, 0, "newsletter_mode", self.newsletter_var, "tip_newsletter"),
            (0, 1, "smart_jitter", self.smart_jitter_var, "tip_jitter"),
            (1, 0, "spam_check", self.spam_check_var, "tip_spam_check"),
            (1, 1, "verify_mx", self.verify_mx_var, "tip_verify_mx"),
            (2, 0, "business_hours", self.business_hours_var, "tip_business_hours"),
            (2, 1, "pause_weekends", self.pause_weekends_var, "tip_pause_weekends"),
        )
        for row, col, key, var, tip_key in adv_options:
            cb = ctk.CTkCheckBox(
                adv_content, text=self.t(key), variable=var,
                font=F["subhead"], text_color=C["text"], fg_color=C["navy"],
            )
            cb.grid(row=row, column=col, sticky="w", padx=12, pady=(6, 4) if row == 0 else 4)
            self._tip(cb, self.t(tip_key))

        # Warmup daily limit row
        warmup_fr = ctk.CTkFrame(adv_content, fg_color="transparent")
        warmup_fr.grid(row=3, column=0, columnspan=2, sticky="w", padx=12, pady=(6, 4))
        warmup_lbl = ctk.CTkLabel(warmup_fr, text=self.t("warmup_limit_lbl"), font=F["subhead"], text_color=C["amber"])
        warmup_lbl.pack(side="left", padx=(0, 8))
        self._tip(warmup_lbl, self.t("tip_warmup"))
        self.warmup_limit_entry = ctk.CTkEntry(warmup_fr, width=65, height=28, font=F["body"], border_color=C["border"])
        self.warmup_limit_entry.insert(0, str(self.settings.get("warmup_limit", "0")))
        self.warmup_limit_entry.pack(side="left", padx=(0, 8))
        today_count = self._get_today_sent_count()
        ctk.CTkLabel(
            warmup_fr,
            text=f"{self.t('warmup_unlimited_hint')}  ·  {self.t('today_sent_inline', today_count)}",
            font=F["small"], text_color=C["muted"],
        ).pack(side="left")

        ctk.CTkFrame(p, height=1, fg_color=C["border"]).grid(row=5, column=0, sticky="ew", padx=24, pady=10)

        # Kutucuk ve açıklaması kendi çerçevesinde: aynı grid hücresine iki widget
        # koymak tam da bölüm başlıklarındaki çakışma hatasına yol açıyordu.
        approval_fr = ctk.CTkFrame(p, fg_color="transparent")
        approval_fr.grid(row=6, column=0, sticky="ew", padx=24, pady=(0, 10))
        approval_fr.grid_columnconfigure(0, weight=1)

        self.app_var = ctk.BooleanVar()
        ctk.CTkCheckBox(
            approval_fr,
            text=self.t("approval_check"),
            variable=self.app_var,
            font=F["subhead"],
            text_color=C["red"],
            checkmark_color="white",
            fg_color=C["red"],
            hover_color=C["red_h"]
        ).grid(row=0, column=0, sticky="w")

        # Bu kutucuğun neden var olduğu hiçbir yerde yazmıyordu
        approval_hint = ctk.CTkLabel(
            approval_fr, text=self.t("approval_hint"), font=F["small"], text_color=C["muted"],
            anchor="w", justify="left", wraplength=620,
        )
        approval_hint.grid(row=1, column=0, sticky="w", padx=(26, 0), pady=(2, 0))
        self._wrap_targets.append(approval_hint)

        self.progress_bar = ctk.CTkProgressBar(p, height=10, corner_radius=5, progress_color=C["accent"], fg_color=C["border"])
        self.progress_bar.set(0)
        self.progress_bar.grid(row=7, column=0, sticky="ew", padx=24)

        self.progress_label = ctk.CTkLabel(p, text=self.t("status_ready"), font=F["small"], text_color=C["muted"])
        self.progress_label.grid(row=8, column=0, sticky="w", padx=24, pady=(2, 8))

        ab = ctk.CTkFrame(p, fg_color="transparent")
        ab.grid(row=9, column=0, pady=(4, 20))

        self.run_btn = ctk.CTkButton(
            ab,
            text=self.t("btn_start"),
            width=220,
            height=46,
            font=("Trebuchet MS", 13, "bold"),
            fg_color=C["navy"],
            hover_color=C["navy_light"],
            corner_radius=10,
            command=self.start_engine
        )
        self.run_btn.pack(side="left", padx=8)

        self.stop_btn = ctk.CTkButton(
            ab,
            text=self.t("btn_stop"),
            width=150,
            height=46,
            font=("Trebuchet MS", 13, "bold"),
            fg_color=C["red"],
            hover_color=C["red_h"],
            corner_radius=10,
            state="disabled",
            command=self._request_stop
        )
        self.stop_btn.pack(side="left", padx=8)

        # Listeye dokunmadan tek deneme maili — en büyük eksikti
        self.test_btn = ctk.CTkButton(
            ab,
            text=self.t("btn_test_send"),
            width=160,
            height=46,
            font=F["body"],
            fg_color="transparent",
            hover_color=C["border"],
            text_color=C["navy"],
            border_width=1,
            border_color=C["navy"],
            corner_radius=10,
            command=self.send_test_mail,
        )
        self.test_btn.pack(side="left", padx=8)
        self._tip(self.test_btn, self.t("tip_test_send"))

    # ──────────────────────── Widget helpers ──────────────────────────────────

    def _ptitle(self, parent, text: str):
        ctk.CTkLabel(parent, text=text, font=F["head"], text_color=C["navy"]).grid(sticky="w", padx=28, pady=(18, 4))

    def _card(self, parent, builder_fn):
        c = self._card_raw(parent)
        builder_fn(c)

    def _card_raw(self, parent):
        c = ctk.CTkFrame(parent, fg_color=C["card"], corner_radius=14, border_width=1, border_color=C["border"])
        c.grid(sticky="ew", padx=24, pady=8)
        c.grid_columnconfigure(0, weight=1)
        return c

    def _section_label(self, parent, text: str, color: str):
        # Çizgi ile başlık aynı hücrede. Eskiden ikisi de dikeyde ortalanıyordu ve
        # 3 px'lik renkli çizgi başlık metninin tam ortasından geçiyordu
        # (ölçüldü: çizgi y=35, metin y=20..55). sticky="sew" ile çizgi artık
        # başlığın altına oturuyor ve alt çizgi görevi görüyor.
        bar = ctk.CTkFrame(parent, height=3, fg_color=color, corner_radius=2)
        bar.grid(row=0, column=0, columnspan=2, sticky="sew", padx=24, pady=(16, 6))
        ctk.CTkLabel(parent, text=text, font=F["subhead"], text_color=color).grid(
            row=0, column=0, sticky="w", padx=28, pady=(16, 14)
        )

    def _push_to(self, attrs, value: str):
        """Bir değeri birden çok giriş kutusuna yazar (kolay ayar senkronizasyonu)."""
        for attr in attrs:
            w = getattr(self, attr, None)
            try:
                if w is not None and w.winfo_exists():
                    w.delete(0, "end")
                    w.insert(0, value)
            except Exception:
                pass

    def _srow(self, parent, row_idx: int, label: str, attr: str, key: str, is_pass=False, tip=None):
        lbl = ctk.CTkLabel(parent, text=label, anchor="e", width=270, font=F["body"], text_color=C["muted"])
        lbl.grid(row=row_idx, column=0, padx=(24, 8), pady=6, sticky="e")
        if tip:
            self._tip(lbl, tip)

        holder = ctk.CTkFrame(parent, fg_color="transparent")
        holder.grid(row=row_idx, column=1, padx=(0, 24), pady=6, sticky="ew")
        holder.grid_columnconfigure(0, weight=1)

        e = ctk.CTkEntry(holder, height=34, show="●" if is_pass else "", font=F["body"], border_color=C["border"])
        e.insert(0, self.settings.get(key, ""))
        e.grid(row=0, column=0, sticky="ew")

        if is_pass:
            # Yanlış yazılmış şifre en sık destek çağrısı sebebi — görülebilmeli
            def _toggle(entry=e):
                entry.configure(show="" if entry.cget("show") else "●")

            eye = ctk.CTkButton(
                holder, text="👁", width=34, height=34, corner_radius=8,
                fg_color=C["border"], hover_color="#cbd5e1", text_color=C["text"],
                font=F["body"], command=_toggle,
            )
            eye.grid(row=0, column=1, padx=(6, 0))
            # Tab akışını bozmasın: kullanıcı alandan alana geçerken araya girmesin
            try:
                eye.configure(takefocus=False)
            except Exception:
                pass
            self._tip(eye, self.t("show_password"))

        parent.grid_columnconfigure(1, weight=1)
        setattr(self, attr, e)

    def _frow(self, parent, row_idx: int, label: str, cmd, lbl_attr: str):
        fr = ctk.CTkFrame(parent, fg_color="transparent")
        fr.grid(row=row_idx, sticky="ew", padx=24, pady=5)
        fr.grid_columnconfigure(2, weight=1)

        ctk.CTkLabel(fr, text=label, width=130, anchor="w", font=F["subhead"], text_color=C["text"]).grid(row=0, column=0)
        ctk.CTkButton(
            fr,
            text=self.t("select_file"),
            width=110,
            height=32,
            fg_color=C["navy"],
            hover_color=C["navy_light"],
            font=F["small"],
            corner_radius=8,
            command=cmd
        ).grid(row=0, column=1, padx=12)

        lbl = ctk.CTkLabel(fr, text=self.t("not_selected"), anchor="w", font=F["small"], text_color=C["muted"])
        lbl.grid(row=0, column=2, sticky="w")
        setattr(self, lbl_attr, lbl)

    def _btn(self, parent, text: str, color: str, hover: str, cmd):
        return ctk.CTkButton(
            parent,
            text=text,
            height=40,
            width=160,
            fg_color=color,
            hover_color=hover,
            font=("Trebuchet MS", 12, "bold"),
            corner_radius=9,
            command=cmd
        )

    def _chip(self, parent, text: str, color: str):
        lbl = ctk.CTkLabel(parent, text=text, font=F["small"], text_color=color)
        lbl.pack(side="left", padx=(0, 20))
        return lbl

    def _toast(self, msg: str, color: str):
        bar = ctk.CTkFrame(self.main_content, fg_color=color, height=36, corner_radius=0)
        bar.grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(bar, text=msg, font=F["small"], text_color="white").pack(pady=8)
        self.after(2400, bar.destroy)

    # ---------- Tooltip altyapısı ----------
    # Eskiden her ipucu için ayrı bir CTkToplevel açılıyordu. CustomTkinter'ın
    # Toplevel'i overrideredirect/topmost ayarlarını gecikmeli olarak yeniden
    # uyguladığı için <Leave> kaçırıldığında pencere ekranda asılı kalıyor,
    # uygulama arkaya alınsa bile "topmost" olduğu için en üstte duruyordu.
    # Artık tek bir hafif tk.Toplevel paylaşılıyor ve fare konumu izlenerek
    # kapatılıyor.

    def _tip_window(self):
        win = getattr(self, "_tipwin", None)
        try:
            if win is not None and win.winfo_exists():
                return win
        except Exception:
            pass

        win = tk.Toplevel(self)
        win.withdraw()
        win.overrideredirect(True)
        try:
            win.attributes("-topmost", True)
        except Exception:
            pass
        win.configure(background=C["navy"])
        lbl = tk.Label(
            win, text="", background=C["navy"], foreground="white",
            wraplength=320, justify="left", padx=10, pady=7,
        )
        lbl.pack()
        self._tipwin, self._tiplbl = win, lbl
        return win

    def _tip_cancel(self):
        for attr in ("_tip_show_id", "_tip_watch_id"):
            job = getattr(self, attr, None)
            if job is not None:
                try:
                    self.after_cancel(job)
                except Exception:
                    pass
            setattr(self, attr, None)

    def _tip_hide(self, _=None):
        self._tip_cancel()
        self._tip_owner = None
        win = getattr(self, "_tipwin", None)
        if win is not None:
            try:
                win.withdraw()
            except Exception:
                pass

    def _tip_pointer_inside(self, widget) -> bool:
        try:
            if not widget.winfo_exists() or not widget.winfo_ismapped():
                return False
            px, py = self.winfo_pointerxy()
            wx, wy = widget.winfo_rootx(), widget.winfo_rooty()
            return (wx <= px < wx + widget.winfo_width()
                    and wy <= py < wy + widget.winfo_height())
        except Exception:
            return False

    def _tip_app_active(self) -> bool:
        """Uygulama simge durumundaysa/arka plandaysa ipucu gösterilmez."""
        try:
            return self.state() == "normal" and bool(self.winfo_viewable())
        except Exception:
            return False

    def _tip_watch(self, widget):
        """<Leave> olayı kaçırılırsa diye fareyi izleyip ipucunu kapatır."""
        if self._tip_owner is not widget:
            return
        if not self._tip_pointer_inside(widget) or not self._tip_app_active():
            self._tip_hide()
            return
        self._tip_watch_id = self.after(150, self._tip_watch, widget)

    def _tip_display(self, widget, text: str):
        if not self._tip_pointer_inside(widget) or not self._tip_app_active():
            self._tip_hide()
            return
        try:
            win = self._tip_window()
            self._tiplbl.configure(text=text)
            win.update_idletasks()
            x = widget.winfo_rootx() + 14
            y = widget.winfo_rooty() + widget.winfo_height() + 6
            sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
            x = max(0, min(x, sw - win.winfo_reqwidth() - 4))
            if y + win.winfo_reqheight() > sh:
                y = widget.winfo_rooty() - win.winfo_reqheight() - 6
            win.geometry(f"+{int(x)}+{int(y)}")
            win.deiconify()
            win.lift()
        except Exception:
            self._tip_hide()
            return
        self._tip_owner = widget
        self._tip_watch_id = self.after(150, self._tip_watch, widget)

    def _tip(self, widget, text: str):
        """Fare üzerine gelince açıklama gösterir.

        Uygulamada hiç yardım metni yoktu; "Oturum Başına Maks. Mesaj" gibi
        jargonlu ayarların ne işe yaradığı hiçbir yerde yazmıyordu.
        """
        if not text or widget is None:
            return

        if not getattr(self, "_tip_root_bound", False):
            self._tip_owner = None
            self._tip_show_id = None
            self._tip_watch_id = None
            for seq in ("<Unmap>", "<FocusOut>", "<Configure>", "<Destroy>"):
                try:
                    self.bind(seq, self._tip_hide, add="+")
                except Exception:
                    pass
            for seq in ("<ButtonPress>", "<MouseWheel>", "<Key>"):
                try:
                    self.bind_all(seq, self._tip_hide, add="+")
                except Exception:
                    pass
            self._tip_root_bound = True

        def _enter(_=None):
            self._tip_cancel()
            self._tip_owner = None
            self._tip_show_id = self.after(450, self._tip_display, widget, text)

        def _leave(_=None):
            self._tip_hide()

        # CTk bileşenleri iç içe canvas/label barındırdığı için olayları
        # alt widget'lara da bağlıyoruz; aksi halde <Leave> hiç gelmiyor.
        targets = [widget]
        try:
            targets += list(widget.winfo_children())
        except Exception:
            pass

        for target in targets:
            for seq, fn in (("<Enter>", _enter), ("<Leave>", _leave),
                            ("<Destroy>", _leave)):
                try:
                    target.bind(seq, fn, add="+")
                except Exception:
                    pass

    def _bind_responsive_wrap(self, container):
        """Metin sarma genişliğini kap genişliğine bağlar.

        wraplength üç yerde 620 px sabitti: pencere genişleyince metin kartın
        ortasında kalıyor, daralınca taşıyordu.
        """
        def _resize(event):
            width = max(260, event.width - 110)
            for lbl in list(self._wrap_targets):
                try:
                    if lbl.winfo_exists():
                        lbl.configure(wraplength=width)
                    else:
                        self._wrap_targets.remove(lbl)
                except Exception:
                    pass

        try:
            container.bind("<Configure>", _resize, add="+")
        except Exception:
            pass

    def _refresh_sender_label(self):
        lbl = getattr(self, "sender_lbl", None)
        try:
            if lbl is None or not lbl.winfo_exists():
                return
            addr = (self.settings.get("mail") or "").strip() or "—"
            lbl.configure(text=f"✉  {self._shorten(addr, 24)}")
            self._tip(lbl, f"{self.t('sender_account')}: {addr}")
        except Exception:
            pass

    # ──────────────────────── SMTP helpers ────────────────────────────────────

    def _smtp_connect(self, host, port, security):
        sec = (security or "").upper().strip()

        if sec == "SSL":
            ctx = ssl.create_default_context()
            srv = smtplib.SMTP_SSL(host, port, timeout=25, context=ctx)
            srv.ehlo()
            return srv

        srv = smtplib.SMTP(host, port, timeout=25)
        srv.ehlo()

        if sec == "STARTTLS":
            ctx = ssl.create_default_context()
            srv.starttls(context=ctx)
            srv.ehlo()

        return srv

    def _sleep_interruptible(self, seconds: int) -> bool:
        """Saniye saniye uyur; DURDUR basılırsa erken çıkar. False → durduruldu."""
        for _ in range(int(seconds)):
            if self._stop_flag.is_set():
                return False
            time.sleep(1)
        return True

    def _max_per_connection(self) -> int:
        """Oturum başına gönderilecek maksimum mesaj sayısı (ayarlardan okunabilir)."""
        try:
            val = int(str(self.settings.get("max_per_conn", "")).strip())
            return val if val > 0 else SMTP_MAX_PER_CONNECTION
        except (TypeError, ValueError):
            return SMTP_MAX_PER_CONNECTION

    def _is_retryable_smtp_error(self, e) -> bool:
        """Bağlantı kopması / geçici sunucu hatası mı, yoksa kalıcı red mi?"""
        if isinstance(e, smtplib.SMTPRecipientsRefused):
            return False  # alıcı kalıcı olarak reddedildi — tekrar denemenin anlamı yok
        if isinstance(e, smtplib.SMTPResponseException):
            return e.smtp_code in SMTP_TRANSIENT_CODES
        return True  # SMTPServerDisconnected, soket/timeout hataları → yeniden bağlan

    def _smtp_session_open(self):
        """Ayarlardaki bilgilerle SMTP oturumu açar + login olur.
        Başarısız olursa artan bekleme ile birkaç kez yeniden dener."""
        host = self.settings.get("server", "").strip()
        port = int(self.settings.get("port", "587"))
        sec = self.settings.get("security", "STARTTLS").upper().strip()
        user = self.settings.get("login_username", "").strip() or self.settings.get("mail", "").strip()
        pw = self.settings.get("pass", "")

        last_err = None
        for attempt in range(1, SMTP_RECONNECT_RETRIES + 1):
            if self._stop_flag.is_set():
                raise RuntimeError(self.t("err_stopped_by_user"))
            try:
                srv = self._smtp_connect(host, port, sec)
                srv.login(user, pw)
                return srv
            except Exception as e:
                last_err = e
                if attempt < SMTP_RECONNECT_RETRIES:
                    wait = SMTP_RECONNECT_BACKOFF * attempt
                    self.log(self.t("log_smtp_conn_fail", attempt, SMTP_RECONNECT_RETRIES, e))
                    self.log(self.t("log_smtp_retry", wait))
                    self._sleep_interruptible(wait)

        raise last_err

    def _smtp_close(self, smtp):
        """SMTP oturumunu nazikçe kapatır; kapanmazsa soketi zorla bırakır."""
        if not smtp:
            return
        try:
            smtp.quit()
        except Exception:
            try:
                smtp.close()
            except Exception:
                pass

    def generate_tracking_id(self):
        """Her mail için benzersiz bir takip ID'si üretir."""
        return f"UNX-{uuid.uuid4().hex[:12].upper()}"

    def log_tracking(self, tracking_id, to_email, subject, message_id):
        """Takip kaydını JSONL dosyasına ekler (satır başına bir kayıt).

        Eskiden tüm JSON okunup baştan yazılıyordu; 10.000 gönderimde bu O(n²)
        dosya trafiği demekti.
        """
        entry = {
            "tracking_id": tracking_id,
            "to_email": to_email,
            "subject": subject,
            "sent_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "message_id": message_id
        }
        try:
            log_file = os.path.join(self.app_dir, "tracking_log.jsonl")
            with open(log_file, "a", encoding="utf-8") as f:
                f.write(json.dumps(entry, ensure_ascii=False) + "\n")
        except Exception as e:
            self.log(self.t("log_tracking_fail", e))

    def _build_msg(self, to_email: str, subject: str, html_body: str, plain_body: str) -> tuple:
        from_email = self.settings.get("mail", "").strip()
        from_name = self.settings.get("from_name", "Unixpadel").strip()
        reply_to = self.settings.get("reply_to", "").strip()
        domain = from_email.split("@")[-1] if "@" in from_email else "unixpadel.com"

        # Tracking ID üret
        tracking_id = self.generate_tracking_id()

        # Düz Metin (Plain Text) modunu kontrol et
        is_plain_only = self.settings.get("template_mode", "html") == "text"

        if is_plain_only:
            msg = MIMEText(plain_body, "plain", "utf-8")
        else:
            msg = MIMEMultipart("alternative")

        msg["From"] = f"{str(Header(from_name, 'utf-8'))} <{from_email}>"
        msg["To"] = to_email
        msg["Subject"] = str(Header(subject, "utf-8"))
        msg["Date"] = formatdate(localtime=True)
        msg["Message-ID"] = make_msgid(domain=domain)
        msg["MIME-Version"] = "1.0"
        msg["X-Mailer"] = "Unixpadel IT Mail Services"
        msg["X-Tracking-ID"] = tracking_id

        # Anti-Spam: List-Unsubscribe Header (Sadece bülten modu aktifse ekle)
        if bool(self.settings.get("newsletter_mode", False)):
            unsubscribe_email = reply_to if reply_to else from_email
            msg["List-Unsubscribe"] = f"<mailto:{unsubscribe_email}?subject=Unsubscribe-{to_email}>"
            msg["List-Unsubscribe-Post"] = "List-Unsubscribe=One-Click"

        if reply_to:
            msg["Reply-To"] = reply_to

        if not is_plain_only:
            msg.attach(MIMEText(plain_body, "plain", "utf-8"))
            if html_body:
                msg.attach(MIMEText(html_body, "html", "utf-8"))

        return msg, msg.as_bytes(), tracking_id
    # ──────────────────────── IMAP helpers ────────────────────────────────────

    def _imap_connect(self):
        host = self.settings.get("imap_server", "").strip() or self.settings.get("server", "").strip()
        port = int(self.settings.get("imap_port", "993"))
        sec = self.settings.get("imap_security", "SSL").upper().strip()
        user = self.settings.get("imap_username", "").strip() or self.settings.get("login_username", "").strip() or self.settings.get("mail", "").strip()
        pw = self.settings.get("imap_pass", "") or self.settings.get("pass", "")

        if sec == "SSL":
            ctx = ssl.create_default_context()
            imap = imaplib.IMAP4_SSL(host, port, ssl_context=ctx)
        elif sec == "STARTTLS":
            imap = imaplib.IMAP4(host, port)
            imap.starttls(ssl_context=ssl.create_default_context())
        else:
            imap = imaplib.IMAP4(host, port)

        imap.login(user, pw)
        return imap

    def _imap_save_sent(self, imap, msg_bytes: bytes) -> bool:
        folder = self.settings.get("sent_folder", "INBOX.Sent")
        for name in (folder, f'"{folder}"'):
            try:
                res = imap.append(name, r"\Seen", imaplib.Time2Internaldate(time.time()), msg_bytes)
                if res[0] == "OK":
                    return True
            except Exception:
                pass
        return False

    # ──────────────────────── Mailing engine ──────────────────────────────────

    def _set_running(self, running: bool):
        """Kampanya durumunu işaretler; butonları ana thread üzerinden günceller.

        Doğrudan configure çağrısı worker thread'den Tk'ye dokunmak demekti.
        """
        self._campaign_running = running

        def _update():
            if hasattr(self, "run_btn") and self.run_btn.winfo_exists():
                self.run_btn.configure(state="disabled" if running else "normal")
            if hasattr(self, "stop_btn") and self.stop_btn.winfo_exists():
                self.stop_btn.configure(state="normal" if running else "disabled")

        self._ui(_update)

    def mailing_process(self, emails):
        self._stop_flag.clear()
        self._set_running(True)

        sent = fail = skip = 0
        total = len(emails)

        imap_conn = None
        smtp = None
        msgs_in_conn = 0
        max_per_conn = self._max_per_connection()

        warmup_limit = 0
        try:
            warmup_limit = int(str(self.settings.get("warmup_limit", "0")).strip())
        except (ValueError, TypeError):
            warmup_limit = 0

        today_sent = self._get_today_sent_count()
        if warmup_limit > 0 and today_sent >= warmup_limit:
            self.log(self.t("log_warmup_cap", warmup_limit, today_sent))
            self._set_status("status_limit")
            self._set_running(False)
            return

        if self.settings.get("save_sent", True):
            try:
                imap_conn = self._imap_connect()
                self.log(self.t("log_imap_connected", self.settings.get("sent_folder", "INBOX.Sent")))
            except Exception as e:
                self.log(self.t("log_imap_none", e))

        try:
            host = self.settings.get("server", "").strip()
            port = int(self.settings.get("port", "587"))
            sec = self.settings.get("security", "STARTTLS").upper()
            user = self.settings.get("login_username", "").strip() or self.settings.get("mail", "").strip()
            m_from = self.settings.get("mail", "").strip()
            m_pass = self.settings.get("pass", "")
            
            # Widget'tan değil ayarlardan oku: burası worker thread ve Tk'ye
            # dokunmak "main thread is not in main loop" ile kampanyayı öldürür.
            # start_engine, save_settings ile subject_entry'yi buraya kopyalar.
            raw_subject = str(self.settings.get("subject", "")).strip()
            subjects = [s.strip() for s in raw_subject.split(",") if s.strip()]
            if not subjects:
                subjects = ["Unixpadel Bilgilendirme"]

            if not all([host, port, user, m_from, m_pass]):
                self.log(self.t("log_smtp_missing"))
                return

            is_plain_only = self.settings.get("template_mode", "html") == "text"
            
            html_body = ""
            if not is_plain_only and self.selected_html:
                with open(self.selected_html, "r", encoding="utf-8") as f:
                    html_body = f.read()

            if is_plain_only:
                base_plain_text = self.settings.get("direct_plain_text", "").strip()
            else:
                base_plain_text = self.settings.get("plain_text", "").strip()

            self.log(self.t("log_connecting", host, port, sec))
            smtp = self._smtp_session_open()
            self.log(self.t("log_smtp_login_ok", max_per_conn))
            self._prog(0, total, sent, fail, skip)

            for i, record in enumerate(emails, 1):
                if self._stop_flag.is_set():
                    self.log(self.t("log_stopped"))
                    break

                # 1. Warmup Günlük Limit Kontrolü
                if warmup_limit > 0 and self._get_today_sent_count() >= warmup_limit:
                    self.log(self.t("log_warmup_cap", warmup_limit, self._get_today_sent_count()))
                    self._set_status("status_limit")
                    break

                # 2. Mesai Saatleri & Hafta Sonu Bekleme Kontrolü
                if not self._wait_for_schedule_window():
                    self.log(self.t("log_stopped"))
                    break

                email = record.get("email", "")

                # 3. E-Posta Biçim Doğrulaması
                if not valid_email(email):
                    self.log(self.t("log_invalid", i, total, email))
                    skip += 1
                    self._prog(i, total, sent, fail, skip)
                    continue

                # 4. Canlı MX ve Domain Doğrulaması (Hard-Bounce Önleme)
                if bool(self.settings.get("verify_mx", True)):
                    if not verify_email_mx(email):
                        self.log(self.t("log_mx_fail", i, total, email))
                        skip += 1
                        self._prog(i, total, sent, fail, skip)
                        continue

                try:
                    # Sunucunun oturum başına mesaj limitine dayanmadan bağlantıyı tazele
                    if smtp is None or msgs_in_conn >= max_per_conn:
                        if smtp is not None:
                            self.log(self.t("log_refresh", msgs_in_conn))
                            self._smtp_close(smtp)
                            smtp = None
                        smtp = self._smtp_session_open()
                        msgs_in_conn = 0
                        self.log(self.t("log_new_smtp"))

                    # 5. Konu Seçimi, Merge Tags (Fallback Destekli) ve Spintax
                    current_subject = random.choice(subjects)
                    current_subject = apply_merge_tags(current_subject, record)
                    current_subject = resolve_spintax(current_subject)

                    # 6. Gövde Merge Tags (Fallback Destekli) ve Spintax
                    merged_html = html_body
                    merged_plain = base_plain_text

                    if merged_html:
                        merged_html = apply_merge_tags(merged_html, record)
                        merged_html = resolve_spintax(merged_html)

                    if merged_plain:
                        merged_plain = apply_merge_tags(merged_plain, record)
                        merged_plain = resolve_spintax(merged_plain)

                    msg_obj, msg_bytes, tracking_id = self._build_msg(email, current_subject, merged_html, merged_plain)

                    try:
                        smtp.sendmail(m_from, [email], msg_bytes)
                    except SMTP_RETRYABLE_ERRORS as e:
                        if not self._is_retryable_smtp_error(e):
                            raise
                        self.log(self.t("log_conn_err", i, total, e))
                        self._smtp_close(smtp)
                        smtp = None
                        smtp = self._smtp_session_open()
                        msgs_in_conn = 0
                        smtp.sendmail(m_from, [email], msg_bytes)

                    msgs_in_conn += 1
                    sent += 1
                    self._increment_today_sent_count()
                    self.log(self.t("log_sent", i, total, email, tracking_id))

                    # Tracking log kaydı
                    self.log_tracking(tracking_id, email, current_subject, msg_obj["Message-ID"])

                    if imap_conn:
                        try:
                            ok = self._imap_save_sent(imap_conn, msg_bytes)
                            if ok:
                                self.log(self.t("log_sent_ok", i, total))
                            else:
                                self.log(self.t("log_sent_fail", i, total))
                        except Exception:
                            try:
                                imap_conn = self._imap_connect()
                                self._imap_save_sent(imap_conn, msg_bytes)
                            except Exception as ie:
                                self.log(self.t("log_imap_fail", i, total, ie))

                except Exception as e:
                    fail += 1
                    self.log(self.t("log_err", i, total, email, e))

                self._prog(i, total, sent, fail, skip)

                # 7. Akıllı Jitter Bekleme Süresi
                if i < total and not self._stop_flag.is_set():
                    # slider.get() de Tk erişimi — değer start_engine'de alınmıştı
                    base_wait = int(getattr(self, "_campaign_interval", 0) or 0)
                    if base_wait:
                        wait = self._calculate_sleep_interval(base_wait)
                        if wait != base_wait:
                            self.log(self.t("log_wait_jitter", self._fmt_interval(wait)))
                        else:
                            self.log(self.t("log_wait", self._fmt_interval(wait)))
                        self._sleep_interruptible(wait)

            self.log(self.t("log_done", sent, fail, skip))
            self._prog(total, total, sent, fail, skip)
            if not self._stop_flag.is_set() and getattr(self, "_status_key", None) != "status_limit":
                self._set_status("status_done")

        except Exception as e:
            self.log(self.t("log_critical", e))
            self._set_status("status_stopped")
        finally:
            self._smtp_close(smtp)
            if imap_conn:
                try:
                    imap_conn.logout()
                except Exception:
                    pass
            self._set_running(False)

    # ──────────────────────── Test connections ────────────────────────────────

    def _friendly_error(self, e) -> str:
        """Ham Python istisnasını kullanıcının düzeltebileceği bir yönergeye çevirir.

        Teknik ayrıntı loga yazılır; diyalogda "(535, b'5.7.8 …')" görünmez.
        """
        text = str(e).lower()

        if isinstance(e, smtplib.SMTPAuthenticationError) or "535" in text \
                or "authenticat" in text or "login failed" in text:
            key = "err_auth"
        elif isinstance(e, socket.gaierror) or "getaddrinfo" in text:
            key = "err_dns"
        elif isinstance(e, ConnectionRefusedError) or "refused" in text:
            key = "err_refused"
        elif isinstance(e, TimeoutError) or "timed out" in text or "timeout" in text:
            key = "err_timeout"
        elif isinstance(e, ssl.SSLError) or "wrong version number" in text:
            key = "err_ssl"
        else:
            return str(e)

        return self.t(key) + self.t("err_detail_in_log")

    def _alert(self, kind: str, title: str, message: str):
        """Diyalogu her zaman ana thread'de açar (worker thread'den de çağrılabilir)."""
        show = messagebox.showinfo if kind == "info" else messagebox.showerror
        self._ui(lambda: show(title, message))

    def test_connection(self):
        self.save_settings(silent=True)

        problem = self._validate_numeric_settings()
        if problem:
            messagebox.showerror(self.t("error"), problem)
            return

        def _run():
            try:
                host = self.settings["server"].strip()
                port = int(self.settings["port"])
                sec = self.settings["security"].upper()
                user = self.settings.get("login_username", "").strip() or self.settings["mail"].strip()
                pw = self.settings["pass"]

                self.log(self.t("log_smtp_test_start", host, port, sec))
                srv = self._smtp_connect(host, port, sec)
                srv.login(user, pw)
                srv.quit()
                self.log(self.t("log_smtp_test_ok"))
                self._alert("info", "SMTP ✓", self.t("test_smtp_ok"))
            except Exception as e:
                self.log(self.t("log_smtp_test_fail", e))
                self._alert("error", self.t("error"), self._friendly_error(e))

        threading.Thread(target=_run, daemon=True).start()

    def test_imap(self):
        self.save_settings(silent=True)

        problem = self._validate_numeric_settings()
        if problem:
            messagebox.showerror(self.t("error"), problem)
            return

        def _run():
            try:
                self.log(self.t("log_imap_test_start"))
                imap = self._imap_connect()
                _, lst = imap.list()
                imap.logout()
                self.log(self.t("log_imap_test_ok"))
                for item in (lst or []):
                    self.log(f"   {item.decode() if isinstance(item, bytes) else item}")
                self._alert("info", "IMAP ✓", self.t("test_imap_ok", self.settings.get("sent_folder")))
            except Exception as e:
                self.log(self.t("log_imap_test_fail", e))
                self._alert("error", self.t("error"), self._friendly_error(e))

        threading.Thread(target=_run, daemon=True).start()

    # ──────────────────────── Engine entry ────────────────────────────────────

    def _campaign_body_text(self) -> str:
        """Gerçekten gönderilecek gövde metni (HTML şablonu dahil)."""
        if self.settings.get("template_mode", "html") == "text":
            return self.settings.get("direct_plain_text", "")

        parts = [self.settings.get("plain_text", "")]
        if self.selected_html:
            try:
                with open(self.selected_html, "r", encoding="utf-8") as f:
                    parts.append(f.read())
            except Exception:
                pass
        return "\n".join(parts)

    def _spam_check_passes(self) -> bool:
        """Konu + gerçekten gönderilecek gövdeyi tarar.

        Eskiden HTML modunda yalnızca ayarlardaki plain-text taranıyor, seçilen
        HTML şablonuna hiç bakılmıyordu — kullanıcı denetlenmiş sanıyordu.
        """
        subject = (self.subject_entry.get() if hasattr(self, "subject_entry") else "").lower()
        body = self._campaign_body_text().lower()

        found = []
        for word in SPAM_WORDS:
            places = []
            if word in subject:
                places.append(self.t("loc_subject"))
            if word in body:
                places.append(self.t("loc_body"))
            if places:
                found.append(f"• {word}  ({', '.join(places)})")

        if not found:
            return True

        return messagebox.askyesno(self.t("spam_warn_title"), self.t("spam_warn_desc", "\n".join(found)))

    def _show_confirm_modal(self, rows: list) -> bool:
        result = {"ok": False}

        modal = ctk.CTkToplevel(self)
        modal.title(self.t("confirm_title"))
        modal.resizable(False, False)
        modal.transient(self)
        modal.configure(fg_color=C["bg"])

        top = ctk.CTkFrame(modal, fg_color=C["navy"], height=56, corner_radius=0)
        top.pack(fill="x")
        ctk.CTkLabel(top, text=self.t("confirm_title"), font=F["head"], text_color="white").pack(
            side="left", padx=20, pady=14
        )

        body = ctk.CTkFrame(modal, fg_color=C["card"], corner_radius=12, border_width=1, border_color=C["border"])
        body.pack(fill="both", expand=True, padx=20, pady=(16, 8))
        body.grid_columnconfigure(1, weight=1)

        for i, (label, value) in enumerate(rows):
            ctk.CTkLabel(
                body, text=label, font=F["subhead"], text_color=C["muted"], anchor="w", width=140
            ).grid(row=i, column=0, sticky="nw", padx=(18, 10), pady=7)
            ctk.CTkLabel(
                body, text=value, font=F["body"], text_color=C["text"],
                anchor="w", justify="left", wraplength=400,
            ).grid(row=i, column=1, sticky="w", padx=(0, 18), pady=7)

        btns = ctk.CTkFrame(modal, fg_color="transparent")
        btns.pack(pady=(6, 18))

        def _accept():
            result["ok"] = True
            modal.destroy()

        cancel_btn = ctk.CTkButton(
            btns, text=self.t("btn_cancel"), width=130, height=42, fg_color=C["muted"],
            hover_color="#475569", font=F["body"], corner_radius=9, command=modal.destroy,
        )
        cancel_btn.pack(side="left", padx=8)
        ctk.CTkButton(
            btns, text=self.t("btn_confirm_send"), width=200, height=42, fg_color=C["green"],
            hover_color=C["green_h"], font=("Trebuchet MS", 12, "bold"), corner_radius=9, command=_accept,
        ).pack(side="left", padx=8)

        self._center_modal(modal)
        modal.grab_set()
        modal.protocol("WM_DELETE_WINDOW", modal.destroy)

        # Esc = vazgeç. Odak bilinçli olarak "Vazgeç"te: Enter'a refleksle basan
        # kullanıcı yüzlerce maili yanlışlıkla yola çıkarmasın.
        modal.bind("<Escape>", lambda _e: modal.destroy())
        cancel_btn.focus_set()

        self.wait_window(modal)
        return result["ok"]

    def _center_modal(self, modal):
        """Modalı ana pencerenin ortasına yerleştirir."""
        try:
            modal.update_idletasks()
            w = modal.winfo_width() if modal.winfo_width() > 1 else modal.winfo_reqwidth()
            h = modal.winfo_height() if modal.winfo_height() > 1 else modal.winfo_reqheight()
            x = self.winfo_x() + (self.winfo_width() - w) // 2
            y = self.winfo_y() + (self.winfo_height() - h) // 2
            modal.geometry(f"+{max(0, x)}+{max(0, y)}")
        except Exception:
            pass

    def _confirm_send(self, recipients: list) -> bool:
        """Gönderim öncesi özeti gösterir; kullanıcı onaylarsa True döner.

        Tek tıkla yüzlerce mailin yola çıkmasını engelleyen asıl durak burasıdır.
        """
        stats = self._list_stats or {}
        total = len(recipients)

        subjects = [s.strip() for s in self.subject_entry.get().split(",") if s.strip()]
        interval = int(self.slider.get())
        eta_seconds = max(0, total - 1) * interval
        finish = datetime.now() + timedelta(seconds=eta_seconds)

        sender = f"{self.settings.get('from_name', '')} <{self.settings.get('mail', '')}>".strip()

        if self.settings.get("template_mode", "html") == "text":
            template = self.t("confirm_template_plain")
        else:
            template = os.path.basename(self.selected_html or "")

        interval_txt = self._fmt_interval(interval)
        if bool(self.settings.get("smart_jitter", True)):
            interval_txt = self.t("confirm_interval_jitter", interval_txt)

        try:
            warmup = int(str(self.settings.get("warmup_limit", "0")).strip() or "0")
        except (TypeError, ValueError):
            warmup = 0

        today = self._get_today_sent_count()
        if warmup > 0:
            daily_txt = self.t("confirm_daily_val", warmup, today, max(0, warmup - today))
        else:
            daily_txt = self.t("confirm_daily_none")

        rows = [
            (self.t("confirm_recipients"),
             self.t("confirm_recipients_val", total, stats.get("dupes", 0), stats.get("invalid", 0))),
            (self.t("confirm_sender"), sender),
            (self.t("confirm_subject"), self.t("confirm_subject_val", len(subjects) or 1)),
            (self.t("confirm_template"), template),
            (self.t("confirm_interval"), interval_txt),
            (self.t("confirm_eta"),
             self.t("confirm_eta_val", finish.strftime("%d.%m.%Y %H:%M"), self._fmt_interval(eta_seconds))),
            (self.t("confirm_daily"), daily_txt),
        ]
        return self._show_confirm_modal(rows)

    def _render_sample(self) -> tuple:
        """Listenin ilk kaydıyla konu ve gövdeyi çözer.

        Önizleme ve test gönderimi aynı yolu kullanır; gerçek kampanyadaki
        apply_merge_tags → resolve_spintax sırası birebir korunur.
        """
        try:
            records = self.get_emails()
        except Exception:
            records = []
        record = records[0] if records else {}

        raw_subject = str(self.settings.get("subject", "")).strip()
        subjects = [s.strip() for s in raw_subject.split(",") if s.strip()] or [""]
        subject = resolve_spintax(apply_merge_tags(random.choice(subjects), record))

        is_plain = self.settings.get("template_mode", "html") == "text"

        if is_plain:
            html_body = ""
            plain_src = self.settings.get("direct_plain_text", "")
        else:
            html_body = ""
            if self.selected_html:
                try:
                    with open(self.selected_html, "r", encoding="utf-8") as f:
                        html_body = f.read()
                except Exception as e:
                    self.log(self.t("preview_html_fail", e))
            plain_src = self.settings.get("plain_text", "")

        plain_body = resolve_spintax(apply_merge_tags(plain_src, record)) if plain_src else ""
        if html_body:
            html_body = resolve_spintax(apply_merge_tags(html_body, record))

        return subject, html_body, plain_body, record, bool(records)

    def open_preview(self):
        """Değişkenler ve spintax çözülmüş hâli gönderimden ÖNCE gösterir."""
        self.save_settings(silent=True)
        subject, html_body, plain_body, _record, has_list = self._render_sample()

        if not html_body and not plain_body:
            messagebox.showwarning(self.t("warning"), self.t("test_send_no_content"))
            return

        # Tk HTML render edemez — HTML şablonu tarayıcıda açılır
        if html_body:
            try:
                tmp = os.path.join(self.app_dir, "onizleme.html")
                with open(tmp, "w", encoding="utf-8") as f:
                    f.write(html_body)
                webbrowser.open(f"file://{tmp}")
                self.log(self.t("preview_html_opened"))
            except Exception as e:
                messagebox.showerror(self.t("error"), self.t("preview_html_fail", e))
                return

        self._show_preview_modal(subject, plain_body, has_list, bool(html_body))

    def _show_preview_modal(self, subject: str, plain_body: str, has_list: bool, html_opened: bool):
        modal = ctk.CTkToplevel(self)
        modal.title(self.t("preview_title"))
        modal.transient(self)
        modal.configure(fg_color=C["bg"])
        modal.geometry("640x520")

        top = ctk.CTkFrame(modal, fg_color=C["navy"], height=52, corner_radius=0)
        top.pack(fill="x")
        top.pack_propagate(False)
        ctk.CTkLabel(top, text=self.t("preview_title"), font=F["head"], text_color="white").pack(
            side="left", padx=20
        )

        note = self.t("preview_note") if has_list else self.t("preview_note_nolist")
        if html_opened:
            note += "  ·  " + self.t("preview_html_opened")
        ctk.CTkLabel(
            modal, text=note, font=F["small"], text_color=C["muted"],
            wraplength=580, justify="left",
        ).pack(anchor="w", padx=22, pady=(12, 6))

        ctk.CTkLabel(
            modal, text=self.t("preview_subject"), font=F["subhead"], text_color=C["navy"],
        ).pack(anchor="w", padx=22)
        subj_box = ctk.CTkTextbox(modal, height=52, font=F["body"], border_width=1, border_color=C["border"])
        subj_box.pack(fill="x", padx=22, pady=(2, 10))

        ctk.CTkLabel(
            modal, text=self.t("preview_body"), font=F["subhead"], text_color=C["navy"],
        ).pack(anchor="w", padx=22)
        body_box = ctk.CTkTextbox(modal, font=F["mono"], border_width=1, border_color=C["border"])
        body_box.pack(fill="both", expand=True, padx=22, pady=(2, 10))

        def _fill(subj, body):
            for box, value in ((subj_box, subj), (body_box, body)):
                box.configure(state="normal")
                box.delete("1.0", "end")
                box.insert("1.0", value)
                box.configure(state="disabled")

        _fill(subject, plain_body)

        def _refresh():
            s, _h, b, _r, _ok = self._render_sample()
            _fill(s, b)

        btns = ctk.CTkFrame(modal, fg_color="transparent")
        btns.pack(pady=(0, 14))
        ctk.CTkButton(
            btns, text=self.t("preview_refresh"), width=140, height=36, corner_radius=9,
            fg_color=C["border"], hover_color="#cbd5e1", text_color=C["text"], font=F["body"],
            command=_refresh,
        ).pack(side="left", padx=8)
        close_btn = ctk.CTkButton(
            btns, text=self.t("close"), width=140, height=36, corner_radius=9,
            fg_color=C["navy"], hover_color=C["navy_light"], font=F["body"], command=modal.destroy,
        )
        close_btn.pack(side="left", padx=8)

        self._center_modal(modal)
        modal.bind("<Escape>", lambda _e: modal.destroy())
        modal.grab_set()
        close_btn.focus_set()

    def send_test_mail(self):
        """Listeye dokunmadan tek deneme maili gönderir.

        Onay diyaloğunu, warmup sayacını ve mesai penceresini bilinçli olarak
        atlar — bu bir kampanya değil, bir kontrol.
        """
        if self._campaign_running:
            return

        self.save_settings(silent=True)

        problem = self._validate_numeric_settings()
        if problem:
            messagebox.showerror(self.t("error"), problem)
            return

        subject, html_body, plain_body, _record, _has = self._render_sample()
        if not html_body and not plain_body:
            messagebox.showwarning(self.t("warning"), self.t("test_send_no_content"))
            return

        default_addr = (self.settings.get("mail") or "").strip()
        target = simpledialog.askstring(
            self.t("test_send_title"), self.t("test_send_prompt"),
            initialvalue=default_addr, parent=self,
        )
        if not target:
            return

        target = normalize_email(target)
        if not valid_email(target):
            messagebox.showerror(self.t("error"), self.t("log_invalid", 1, 1, target))
            return

        self.test_btn.configure(state="disabled")
        self.log(self.t("test_send_running", target))

        def _run():
            smtp = None
            try:
                smtp = self._smtp_session_open()
                _msg, msg_bytes, _tid = self._build_msg(target, subject, html_body, plain_body)
                smtp.sendmail(self.settings.get("mail", "").strip(), [target], msg_bytes)
                self.log(self.t("test_send_log_ok", target))
                self._alert("info", self.t("test_send_title"), self.t("test_send_ok", target))
            except Exception as e:
                self.log(self.t("test_send_log_fail", e))
                self._alert("error", self.t("error"), self._friendly_error(e))
            finally:
                self._smtp_close(smtp)
                self._ui(lambda: self.test_btn.winfo_exists() and self.test_btn.configure(state="normal"))

        threading.Thread(target=_run, daemon=True).start()

    def start_engine(self):
        if self._campaign_running:
            return

        if not self.app_var.get():
            messagebox.showwarning(self.t("warning"), self.t("alert_approval"))
            return

        self.save_settings(silent=True)

        problem = self._validate_numeric_settings()
        if problem:
            messagebox.showerror(self.t("error"), problem)
            return

        is_plain_only = self.settings.get("template_mode", "html") == "text"
        if not is_plain_only and not self.selected_html:
            self.log(self.t("alert_no_html"))
            messagebox.showwarning(self.t("warning"), self.t("alert_no_html"))
            return

        # Liste yalnızca BİR kez okunur — eskiden büyük Excel iki kez açılıyordu
        emails = self.get_emails()
        if not emails:
            self.log(self.t("alert_empty_list"))
            messagebox.showwarning(self.t("warning"), self.t("alert_empty_list"))
            return

        if bool(self.settings.get("spam_check", True)) and not self._spam_check_passes():
            return

        if not self._confirm_send(emails):
            return

        self._campaign_interval = int(self.slider.get())
        self._set_status("status_sending")
        self.progress_bar.set(0)
        self._prog(0, len(emails), 0, 0, 0)

        self.log(f"📝 {self._start_log_file()}")

        threading.Thread(target=self.mailing_process, args=(emails,), daemon=True).start()

    @staticmethod
    def _shorten(text: str, limit: int = 34) -> str:
        """Uzun dosya adlarını ortadan kısaltır ki yerleşimi itmesin."""
        text = text or ""
        if len(text) <= limit:
            return text
        head = (limit - 1) // 2
        return text[:head] + "…" + text[-(limit - 1 - head):]

    def _dedupe_and_validate(self, records: list) -> tuple:
        """Tekrar eden ve geçersiz adresleri ayıklar.

        (temiz_kayıtlar, {"valid":…, "dupes":…, "invalid":…}) döndürür — böylece
        kullanıcı listedeki sorunları gönderimden önce görebilir.
        """
        seen = set()
        clean, dupes, invalid = [], 0, 0

        for record in records:
            email = normalize_email(record.get("email", ""))
            if not valid_email(email):
                invalid += 1
                continue

            key = email.lower()
            if key in seen:
                dupes += 1
                continue

            seen.add(key)
            record["email"] = email
            clean.append(record)

        return clean, {"valid": len(clean), "dupes": dupes, "invalid": invalid}

    def _active_source(self) -> str:
        """Seçili liste kaynağı: "excel" veya "manual".

        Widget varlığına bakılamaz: birleşik sayfada her iki widget da mevcut,
        pasif olan yalnızca grid_remove edilmiş durumda ve winfo_exists() hâlâ
        True döndürüyor.
        """
        var = getattr(self, "list_source", None)
        if var is not None:
            try:
                return var.get()
            except Exception:
                pass
        return self.settings.get("list_source", "excel")

    def get_emails(self):
        """Seçili kaynaktaki listeyi çözümler (tekrarlar ve geçersizler ayıklanmış).

        Özet sayılar _list_stats içine yazılır; gönderim öncesi onay diyaloğu ve
        dosya etiketi bunu kullanır.
        """
        records, columns = [], []
        source = self._active_source()

        if source == "manual":
            raw = ""
            if hasattr(self, "manual_text") and self.manual_text.winfo_exists():
                raw = self.manual_text.get("1.0", "end")
            else:
                raw = self.settings.get("manual_list", "")
            records = [{"email": line} for line in raw.splitlines() if line.strip()]

        elif self.selected_excel:
            key = None
            try:
                key = (self.selected_excel, os.path.getmtime(self.selected_excel))
            except OSError:
                pass

            if key is not None and self._excel_key == key:
                records = [dict(r) for r in self._excel_rows]
                columns = list(self._excel_cols)
            else:
                try:
                    df = pd.read_excel(self.selected_excel).fillna("")
                    columns = [str(c).strip().lower() for c in df.columns]
                    for _, row in df.iterrows():
                        record = {"email": str(row.iloc[0])}
                        for col_name in df.columns:
                            record[str(col_name).strip().lower()] = str(row[col_name])
                        records.append(record)
                except Exception as e:
                    self.log(self.t("log_excel_err", e))
                    self._list_stats = None
                    self._excel_key = None
                    return []

                if key is not None:
                    self._excel_key = key
                    self._excel_rows = [dict(r) for r in records]
                    self._excel_cols = list(columns)

        clean, stats = self._dedupe_and_validate(records)
        stats["columns"] = [c for c in columns[1:] if c]  # ilk sütun e-posta adresi
        self._list_stats = stats
        return clean

    def _list_summary_text(self) -> str:
        stats = self._list_stats or {}
        cols = stats.get("columns") or []
        if cols:
            return self.t(
                "list_summary_cols",
                stats.get("valid", 0), stats.get("dupes", 0), stats.get("invalid", 0),
                ", ".join(cols[:6]),
            )
        return self.t("list_summary", stats.get("valid", 0), stats.get("dupes", 0), stats.get("invalid", 0))

    def _refresh_list_summary(self):
        """Seçilen Excel'i hemen çözümleyip özetini etikete yazar.

        Kullanıcı böylece geçersiz/tekrar eden adresleri ve merge tag olarak
        kullanabileceği sütun adlarını gönderimden ÖNCE görür.
        """
        if not self.selected_excel:
            # Boş durum: "Henüz seçilmedi" ne yapılacağını söylemiyordu
            if hasattr(self, "excel_label") and self.excel_label.winfo_exists():
                self.excel_label.configure(text=self.t("empty_list_hint"), text_color=C["muted"])
            return

        self.get_emails()
        stats = self._list_stats or {}

        if hasattr(self, "excel_label") and self.excel_label.winfo_exists():
            name = self._shorten(os.path.basename(self.selected_excel), 28)
            self.excel_label.configure(
                text=f"✓  {name}  —  {self._list_summary_text()}",
                text_color=C["green"] if stats.get("valid") else C["red"],
            )

    def load_html(self):
        p = filedialog.askopenfilename(filetypes=[("HTML Files", "*.html")])
        if p:
            self.selected_html = p
            n = os.path.basename(p)
            self.log(self.t("log_html_picked", n))
            if hasattr(self, "html_label") and self.html_label.winfo_exists():
                self.html_label.configure(text=f"✓  {self._shorten(n)}", text_color=C["green"])

    def load_excel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if p:
            self.selected_excel = p
            self.log(self.t("log_excel_picked", os.path.basename(p)))
            self._refresh_list_summary()

    # ──────────────────────── Misc helpers ────────────────────────────────────

    def _progress_text(self) -> str:
        """İlerleme etiketinin metni: sayaç + kalan süre + tahmini bitiş saati."""
        cur = getattr(self, "_campaign_cur", 0)
        tot = getattr(self, "_campaign_total", 0)
        status_key = getattr(self, "_status_key", None)

        if not tot:
            return self.t(status_key) if status_key else self.t("status_ready")

        pct = int(cur / tot * 100)
        head = f"{cur} / {tot}  ({self.t('pct_fmt', pct)})"  # TR: %2  ·  EN: 2%

        # Beklemede / durduruldu / bitti gibi durumlarda ETA anlamsız
        if status_key and status_key != "status_sending":
            return f"{head}  ·  {self.t(status_key)}"

        remaining = max(0, tot - cur)
        interval = int(getattr(self, "_campaign_interval", 0) or 0)
        if remaining and interval:
            secs = remaining * interval
            finish = datetime.now() + timedelta(seconds=secs)
            return self.t(
                "prog_eta", cur, tot, pct,
                self._fmt_interval(secs), finish.strftime("%d.%m %H:%M"),
            )
        return head

    def _set_status(self, status_key=None):
        """Durum etiketini günceller (Gönderiliyor / Mesai bekleniyor / Durduruldu …)."""
        self._status_key = status_key
        self._refresh_progress_label()

    def _refresh_progress_label(self):
        def _update():
            if hasattr(self, "progress_label") and self.progress_label.winfo_exists():
                self.progress_label.configure(text=self._progress_text())

        self._ui(_update)

    def _prog(self, cur, tot, sent, fail, skip):
        """Thread-güvenli ilerleme güncellemesi.

        Widget'lar sekme değişiminde yok edilmiş olabilir; güncelleme ana thread'e
        devredilir ve varlık kontrolü orada yapılır. Aksi halde worker thread
        TclError alır ve kampanya "KRİTİK HATA" ile ölür.
        """
        self._campaign_cur = cur
        self._campaign_total = tot
        self._campaign_sent = sent

        def _update():
            if tot and hasattr(self, "progress_bar") and self.progress_bar.winfo_exists():
                self.progress_bar.set(cur / tot)
            if hasattr(self, "progress_label") and self.progress_label.winfo_exists():
                self.progress_label.configure(text=self._progress_text())
            for attr, value in (("lbl_sent", sent), ("lbl_fail", fail), ("lbl_skip", skip)):
                w = getattr(self, attr, None)
                if w is not None and w.winfo_exists():
                    w.configure(text=self.t(attr, value))

        self._ui(_update)

    def _request_stop(self):
        self._stop_flag.set()
        self._set_status("status_stopped")
        self.log(self.t("log_stop_req"))

    def _fmt_interval(self, seconds):
        seconds = int(seconds)
        h, rem = divmod(seconds, 3600)
        m, s = divmod(rem, 60)
        if h:
            return self.t("time_hour_min", h, m) if m else self.t("time_hour", h)
        if m and s:
            return self.t("time_min_sec", m, s)
        if m:
            return self.t("time_min", m)
        return self.t("time_sec", s)

    def _update_slider(self, v):
        self.l_slider.configure(text=self._fmt_interval(v))

    def _set_interval(self, seconds: int):
        """Hazır aralık düğmeleri — kaydırıcıyı sürüklemeden tam değer seçmek için."""
        if hasattr(self, "slider") and self.slider.winfo_exists():
            self.slider.set(seconds)
            self._update_slider(seconds)

    # ──────────────────────── Logging ─────────────────────────────────────────

    def _start_log_file(self) -> str:
        """Yeni kampanya için log dosyası açar; yolunu döndürür."""
        name = f"kampanya_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        self._log_file_path = os.path.join(self.log_dir, name)
        return self._log_file_path

    def _write_log_file(self, line: str):
        path = getattr(self, "_log_file_path", None)
        if not path:
            return
        try:
            with open(path, "a", encoding="utf-8") as f:
                f.write(line + "\n")
        except Exception:
            pass  # log yazamamak gönderimi durdurmamalı

    def log(self, msg: str):
        """Thread-güvenli log; ekrana ve kampanya log dosyasına birlikte yazar.

        Tk'ye dokunmak her zaman ana thread'e devredilir — böylece kullanıcı
        gönderim sırasında sekme değiştirdiğinde satırlar sessizce kaybolmaz.
        """
        line = f"[{time.strftime('%H:%M:%S')}]  {msg}"
        self._write_log_file(line)

        def _append():
            if hasattr(self, "log_view") and self.log_view.winfo_exists():
                self.log_view.insert("end", line + "\n")
                self.log_view.see("end")
                self._show_log_placeholder()

        self._ui(_append)

    def _save_log(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            initialfile=f"unixpadel_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")],
        )
        if not path:
            return
        try:
            content = ""
            if hasattr(self, "log_view") and self.log_view.winfo_exists():
                content = self.log_view.get("1.0", "end")
            with open(path, "w", encoding="utf-8") as f:
                f.write(content)
            self.log(self.t("log_saved", path))
        except Exception as e:
            messagebox.showerror(self.t("error"), self.t("log_save_fail", e))

    def _clear_log(self):
        if hasattr(self, "log_view") and self.log_view.winfo_exists():
            self.log_view.delete("1.0", "end")
            self._show_log_placeholder()

    def _open_log_folder(self):
        try:
            os.startfile(self.log_dir)
        except Exception as e:
            messagebox.showerror(self.t("error"), str(e))

    def clear_main(self):
        if hasattr(self, "main_content"):
            for w in self.main_content.winfo_children():
                w.destroy()
            # Log bölmesi yalnızca gönderim sayfasında var; başka sayfada satır 2
            # ağırlıklı kalırsa ekranın altında boş bir şerit oluşuyor.
            self.main_content.grid_rowconfigure(1, weight=1)
            self.main_content.grid_rowconfigure(2, weight=0, minsize=0)

        self._wrap_targets = []

        for attr in [
            "s_mail", "s_server", "s_port", "s_security", "s_login_username",
            "s_pass", "s_from_name", "s_reply_to", "s_plain_text", "manual_text",
            "html_label", "excel_label", "s_imap_server", "s_imap_port",
            "s_imap_security", "s_imap_username", "s_imap_pass",
            "s_sent_folder", "s_save_sent", "excel_hint_lbl", "manual_hint_lbl",
            "excel_fr", "manual_fr", "_log_placeholder", "test_btn",
            # Bunlar listede yoktu: yok edilmiş widget'a işaret eden ölü
            # referanslar save_settings'i baştan patlatıyordu.
            "easy_mail", "easy_pass", "subject_entry", "warmup_limit_entry",
            "slider", "l_slider", "progress_bar", "progress_label",
            "log_view", "run_btn", "stop_btn", "lbl_sent", "lbl_fail", "lbl_skip",
        ]:
            if hasattr(self, attr):
                delattr(self, attr)

    def on_close(self):
        if self._campaign_running:
            keep_going = messagebox.askyesno(
                self.t("close_title"),
                self.t("close_body", self._campaign_total, self._campaign_sent),
            )
            if not keep_going:
                return
            self._stop_flag.set()

        try:
            self.save_settings(silent=True)
        except Exception:
            pass
        self.destroy()


if __name__ == "__main__":
    app = UnixpadelApp()
    app.mainloop()