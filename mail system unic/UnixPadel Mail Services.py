"""
Unixpadel IT Mail Services — v5.2
Centered window • Smaller login/main window • Remember me • Persistent SMTP/IMAP settings
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
import os, json, threading, time, smtplib, imaplib, ssl, re, uuid, random, pymysql
from datetime import datetime
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
    "badge": ("Trebuchet MS", 10),
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
        
        # Send Controls
        "sec_send_ctrl": "⚡  Gönderim Kontrolü",
        "subjects_csv": "Konu Başlıkları (Virgüllü)",
        "send_interval": "Gönderim Aralığı",
        "newsletter_mode": "Bülten Modu (List-Unsubscribe Ekle)",
        "spam_check": "Spam Kelime Kontrolü Yap",
        "approval_check": "Elshan Bey'den gönderim onayı alındı",
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
        "set_plain_text": "📄  Plain-Text Versiyonu (spam filtre için zorunlu)",
        "btn_smtp_test": "🔌  SMTP Test",
        "btn_imap_test": "📥  IMAP Test",
        "settings_saved": "✓  Ayarlar kaydedildi",
        "settings_fail": "Kaydedilemedi:\n{}",
        
        # Tests & Alerts
        "test_smtp_ok": "Bağlantı ve giriş başarılı!",
        "test_imap_ok": "Bağlantı başarılı!\nSent klasörü: \"{}\"\n\nKlasör listesi loglara yazıldı.",
        "alert_approval": "Elshan Bey onayı kutucuğunu işaretleyin.",
        "alert_empty_list": "❌ E-posta listesi boş.",
        "alert_no_html": "❌ HTML şablonu seçilmedi. (Düz metin kullanmak için 'Mesaj Formatı'nı 'Düz Metin' yapın).",
        "spam_warn_title": "Spam Filtresi Uyarısı",
        "spam_warn_desc": "Mesajınızda spam filtrelerine takılabilecek kelimeler tespit edildi:\n{}\n\nYine de devam etmek istiyor musunuz?",
        
        # Logs
        "log_stopped": "⏹ Kullanıcı tarafından durduruldu.",
        "log_invalid": "[{}/{}] ⚠ Girdiğiniz mail geçersiz: {}  → atlandı",
        "log_refresh": "♻  Bağlantı tazeleniyor (bu oturumda {} mesaj gönderildi)",
        "log_new_smtp": "✓  Yeni SMTP oturumu açıldı",
        "log_conn_err": "[{}/{}] 🔄 Bağlantı sorunu ({}) — yeniden bağlanılıyor…",
        "log_sent": "[{}/{}] ✉  Gönderildi → {}  🔖 {}",
        "log_sent_ok": "[{}/{}] 📥  Sent'e kaydedildi",
        "log_sent_fail": "[{}/{}] ⚠  Sent klasörü bulunamadı — adı kontrol edin",
        "log_imap_fail": "[{}/{}] ⚠  IMAP yeniden bağlanamadı: {}",
        "log_err": "[{}/{}] ❌ Hata ({}): {}",
        "log_wait": "⏳ {} bekleniyor…",
        "log_done": "🏁 TAMAMLANDI  ·  ✓ {}  ✗ {}  ⚠ {}",
        "log_critical": "❌ KRİTİK HATA: {}",
        "log_stop_req": "⏹ Durdurma isteği alındı…",
        "time_min_sec": "{} dakika {} saniye",
        "time_min": "{} dakika",
        "time_sec": "{} saniye",
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
        
        # Send Controls
        "sec_send_ctrl": "⚡  Send Controls",
        "subjects_csv": "Subjects (Comma separated)",
        "send_interval": "Send Interval",
        "newsletter_mode": "Newsletter Mode (Add List-Unsubscribe)",
        "spam_check": "Check for Spam Words",
        "approval_check": "Sending approval received from Elshan",
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
        "set_plain_text": "📄  Plain-Text Version (required for spam filters)",
        "btn_smtp_test": "🔌  SMTP Test",
        "btn_imap_test": "📥  IMAP Test",
        "settings_saved": "✓  Settings saved",
        "settings_fail": "Failed to save:\n{}",
        
        # Tests & Alerts
        "test_smtp_ok": "Connection and login successful!",
        "test_imap_ok": "Connection successful!\nSent folder: \"{}\"\n\nFolder list written to logs.",
        "alert_approval": "Please check the Elshan approval box.",
        "alert_empty_list": "❌ Email list is empty.",
        "alert_no_html": "❌ HTML template not selected. (To use plain text, change 'Message Format' to 'Plain Text').",
        "spam_warn_title": "Spam Filter Warning",
        "spam_warn_desc": "Words that might trigger spam filters were detected in your message:\n{}\n\nDo you still want to continue?",
        
        # Logs
        "log_stopped": "⏹ Stopped by user.",
        "log_invalid": "[{}/{}] ⚠ Invalid email entered: {}  → skipped",
        "log_refresh": "♻  Refreshing connection ({} messages sent in this session)",
        "log_new_smtp": "✓  New SMTP session opened",
        "log_conn_err": "[{}/{}] 🔄 Connection issue ({}) — reconnecting…",
        "log_sent": "[{}/{}] ✉  Sent → {}  🔖 {}",
        "log_sent_ok": "[{}/{}] 📥  Saved to Sent",
        "log_sent_fail": "[{}/{}] ⚠  Sent folder not found — check the name",
        "log_imap_fail": "[{}/{}] ⚠  IMAP failed to reconnect: {}",
        "log_err": "[{}/{}] ❌ Error ({}): {}",
        "log_wait": "⏳ Waiting {}…",
        "log_done": "🏁 COMPLETED  ·  ✓ {}  ✗ {}  ⚠ {}",
        "log_critical": "❌ CRITICAL ERROR: {}",
        "log_stop_req": "⏹ Stop request received…",
        "time_min_sec": "{} minute(s) {} second(s)",
        "time_min": "{} minute(s)",
        "time_sec": "{} second(s)",
        "lang_name": "English"
    }
}

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


class UnixpadelApp(ctk.CTk):

    def t(self, key, *args):
        lang = self.settings.get("language", "tr")
        if lang not in TRANSLATIONS:
            lang = "tr"
        text = TRANSLATIONS[lang].get(key, key)
        if args:
            try:
                return text.format(*args)
            except:
                return text
        return text

    def __init__(self):
        super().__init__()
        self.title(self.t("app_title"))

        self.app_dir = os.path.join(os.path.expanduser("~"), "AppData", "Local", "UnixpadelITMailServices")
        os.makedirs(self.app_dir, exist_ok=True)
        self.config_path = os.path.join(self.app_dir, "settings.json")
        self.old_config_path = os.path.join(self.app_dir, "config.json")

        self.selected_html = None
        self.selected_excel = None
        self.settings = {}
        self._stop_flag = threading.Event()

        self.load_settings()
        self.setup_window()

        self.configure(fg_color=C["bg"])
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        if self.settings.get("is_logged_in"):
            self.build_main_ui()
        else:
            self.show_login_screen()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        # Arka planda sessiz otomatik güncelleme kontrolü
        start_silent_auto_update(status_callback=self.update_version_label, app_instance=self)

    def update_version_label(self, status_text: str):
        def _update():
            if hasattr(self, 'version_lbl') and self.version_lbl.winfo_exists():
                self.version_lbl.configure(text=status_text)
        try:
            self.after(0, _update)
        except Exception:
            pass


    def _change_lang_login(self, choice):
        self.settings["language"] = choice
        self.save_settings(silent=True)
        self.show_login_screen()

    def _change_lang_main(self, choice):
        self.settings["language"] = choice
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
                messagebox.showwarning(self.t("warning"), f"Ayar dosyası yüklenemedi, varsayılanlara dönüldü.\n\n{e}")

        self.settings = d

    def save_settings(self, silent=False):
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
            for key, attr in widget_map.items():
                if hasattr(self, attr):
                    self.settings[key] = getattr(self, attr).get().strip()

            if hasattr(self, "s_save_sent"):
                self.settings["save_sent"] = bool(self.s_save_sent.get())

            if hasattr(self, "s_plain_text"):
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

            if hasattr(self, "lang_var") and self.lang_var.get() != self.settings.get("language", "tr"):
                self.settings["language"] = self.lang_var.get()


            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(self.settings, f, indent=4, ensure_ascii=False)

            if not silent and hasattr(self, "main_content"):
                self._toast(self.t("settings_saved"), C["green"])

        except Exception as e:
            if not silent:
                messagebox.showerror(self.t("error"), self.t("settings_fail", e))

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

        login_w = 380
        login_h = 430

        shadow = ctk.CTkFrame(
            self,
            width=login_w + 8,
            height=login_h + 8,
            fg_color="#0a1a40",
            corner_radius=22
        )
        shadow.place(relx=0.5, rely=0.5, anchor="center", x=4, y=4)

        card = ctk.CTkFrame(
            self,
            width=login_w,
            height=login_h,
            fg_color=C["card"],
            corner_radius=20
        )
        card.place(relx=0.5, rely=0.5, anchor="center")

        strip = ctk.CTkFrame(card, width=login_w, height=70, fg_color=C["navy"], corner_radius=0)
        strip.place(x=0, y=0)

        ctk.CTkLabel(
            strip,
            text="UNIXPADEL IT",
            font=("Georgia", 22, "bold"),
            text_color="white"
        ).place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(
            card,
            text="Mail Services",
            font=("Trebuchet MS", 12),
            text_color=C["muted"]
        ).place(relx=0.5, y=92, anchor="center")

        ctk.CTkFrame(card, width=290, height=1, fg_color=C["border"]).place(relx=0.5, y=115, anchor="center")

                # Language selector
        self.lang_var = ctk.StringVar(value=self.settings.get("language", "tr"))
        lang_cb = ctk.CTkOptionMenu(
            card,
            variable=self.lang_var,
            values=["tr", "en"],
            width=70,
            height=28,
            fg_color=C["navy"],
            button_color=C["navy_mid"],
            font=F["badge"],
            command=self._change_lang_login
        )
        lang_cb.place(relx=0.9, y=35, anchor="center")

        self.user_entry = ctk.CTkEntry(
            card,
            placeholder_text=self.t("username"),
            width=280,
            height=38,
            border_color=C["border"],
            font=F["body"]
        )
        self.user_entry.place(relx=0.5, y=155, anchor="center")

        self.pass_entry = ctk.CTkEntry(
            card,
            placeholder_text=self.t("password"),
            width=280,
            height=38,
            show="●",
            border_color=C["border"],
            font=F["body"]
        )
        self.pass_entry.place(relx=0.5, y=205, anchor="center")
        self.pass_entry.bind("<Return>", lambda _: self.check_login())

        self.remember_me_var = ctk.BooleanVar(value=bool(self.settings.get("remember_me", False)))
        remember_cb = ctk.CTkCheckBox(
            card,
            text=self.t("remember_me"),
            variable=self.remember_me_var,
            font=F["body"],
            text_color=C["text"],
            fg_color=C["navy"],
            hover_color=C["navy_light"],
            checkmark_color="white"
        )
        remember_cb.place(relx=0.5, y=248, anchor="center")

        if self.settings.get("remember_me"):
            self.user_entry.insert(0, self.settings.get("remembered_user", ""))
            self.pass_entry.insert(0, self.settings.get("remembered_pass", ""))

        ctk.CTkButton(
            card,
            text=self.t("login_btn"),
            command=self.check_login,
            fg_color=C["navy"],
            hover_color=C["navy_light"],
            width=280,
            height=40,
            font=("Trebuchet MS", 12, "bold"),
            corner_radius=10
        ).place(relx=0.5, y=300, anchor="center")

        ctk.CTkLabel(
            card,
            text="© 2026 Unixpadel · IT Department",
            font=F["badge"],
            text_color=C["muted"]
        ).place(relx=0.5, y=392, anchor="center")

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
            messagebox.showerror(self.t("error"), self.t("db_err", e))
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
        self.center_window(1080, 720)

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
            font=("Trebuchet MS", 9, "bold"),
            text_color="#4d7fbc"
        ).pack(anchor="w", padx=18, pady=(14, 6))

        self._nav_btns = {}
        nav_items = [
            ("excel", self.t("nav_excel"), self.show_bulk),
            ("manuel", self.t("nav_manual"), self.show_manual),
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

        self.lang_var = ctk.StringVar(value=self.settings.get("language", "tr"))
        lang_cb_main = ctk.CTkOptionMenu(
            self._sidebar,
            variable=self.lang_var,
            values=["tr", "en"],
            width=70,
            height=28,
            fg_color=C["navy"],
            button_color=C["navy_mid"],
            font=F["badge"],
            command=self._change_lang_main
        )
        lang_cb_main.pack(side="bottom", pady=10)

        self.version_lbl = ctk.CTkLabel(self._sidebar, text=f"v{APP_VERSION}", font=F["badge"], text_color="#3d6199")
        self.version_lbl.pack(side="bottom", pady=6)

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
        self.main_content.grid_rowconfigure(0, weight=1)

        self._toast_lbl = None
        self._nav_click("excel", self.show_bulk)

        self.after(80, lambda: self.center_window(1080, 720))

    def logout(self):
        self.settings["is_logged_in"] = False
        self.save_settings(silent=True)
        self.after(100, self.show_login_screen)

    def _nav_click(self, key: str, cmd):
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
        sf.grid(row=0, column=0, sticky="nsew")
        sf.grid_columnconfigure(0, weight=1)
        return sf

    # ──────────────────────── Pages ───────────────────────────────────────────

    def show_bulk(self):
        self.clear_main()
        sf = self._scroll()
        self._ptitle(sf, self.t("nav_excel"))
        self._card(sf, self._section_bulk_files)
        self._card(sf, self._section_send_controls)

    def show_manual(self):
        self.clear_main()
        sf = self._scroll()
        self._ptitle(sf, self.t("nav_manual"))
        self._card(sf, self._section_html_only)
        self._card(sf, self._section_manual_emails)
        self._card(sf, self._section_send_controls)

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

        def _sync_mail(e):
            val = self.easy_mail.get()
            for attr in ["s_mail", "s_login_username", "s_imap_username"]:
                if hasattr(self, attr) and getattr(self, attr).winfo_exists():
                    widget = getattr(self, attr)
                    widget.delete(0, 'end')
                    widget.insert(0, val)
        def _sync_pass(e):
            val = self.easy_pass.get()
            for attr in ["s_pass", "s_imap_pass"]:
                if hasattr(self, attr) and getattr(self, attr).winfo_exists():
                    widget = getattr(self, attr)
                    widget.delete(0, 'end')
                    widget.insert(0, val)

        self.easy_mail.bind("<KeyRelease>", _sync_mail)
        self.easy_pass.bind("<KeyRelease>", _sync_pass)

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
        self._srow(smtp_content, 10, self.t("set_max_conn"), "s_max_per_conn", "max_per_conn")

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
        
        self._btn(bf, self.t("save"), C["green"], C["green_h"], self.save_settings).pack(side="left", padx=6)
        self._btn(bf, self.t("btn_smtp_test"), C["navy"], C["navy_light"], self.test_connection).pack(side="left", padx=6)
        self._btn(bf, self.t("btn_imap_test"), C["teal"], C["teal_h"], self.test_imap).pack(side="left", padx=6)

    # ──────────────────────── Section builders ────────────────────────────────

    def _section_bulk_files(self, p):
        self._section_label(p, self.t("sec_template_file"), C["navy"])
        self._build_template_selector(p, row_start=1)
        self._frow(p, 4, self.t("excel_list"), self.load_excel, "excel_label")
        ctk.CTkLabel(p, text="").grid(row=5)

    def _section_html_only(self, p):
        self._section_label(p, self.t("sec_template_only"), C["navy"])
        self._build_template_selector(p, row_start=1)
        ctk.CTkLabel(p, text="").grid(row=4)

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
        
        self.plain_text_fr = ctk.CTkFrame(parent, fg_color="transparent")
        self.plain_text_fr.grid(row=row_start+2, column=0, sticky="ew", padx=24, pady=0)
        self.plain_text_fr.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(self.plain_text_fr, text=self.t("your_plain_text"), anchor="w", font=F["subhead"], text_color=C["text"]).grid(row=0, column=0, sticky="w", pady=(0, 5))
        self.direct_plain_text = ctk.CTkTextbox(self.plain_text_fr, height=150, border_width=1, font=F["mono"], border_color=C["border"])
        self.direct_plain_text.grid(row=1, column=0, sticky="ew")
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

    def _section_manual_emails(self, p):
        self._section_label(p, self.t("manual_list_title"), C["navy"])
        self.manual_text = ctk.CTkTextbox(p, height=150, border_width=1, font=F["mono"], border_color=C["border"])
        self.manual_text.grid(row=1, column=0, sticky="ew", padx=24, pady=(4, 16))

    def _section_send_controls(self, p):
        p.grid_columnconfigure(0, weight=1)
        self._section_label(p, self.t("sec_send_ctrl"), C["navy"])

        sr = ctk.CTkFrame(p, fg_color="transparent")
        sr.grid(row=1, sticky="ew", padx=24, pady=4)
        sr.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(sr, text=self.t("subjects_csv"), width=175, anchor="w", font=F["subhead"], text_color=C["text"]).grid(row=0, column=0)
        self.subject_entry = ctk.CTkEntry(sr, height=34, font=F["body"], border_color=C["border"])
        self.subject_entry.insert(0, self.settings.get("subject", "Unixpadel Bilgilendirme"))
        self.subject_entry.grid(row=0, column=1, sticky="ew", padx=(12, 0))

        sl_r = ctk.CTkFrame(p, fg_color="transparent")
        sl_r.grid(row=2, sticky="ew", padx=24, pady=(10, 4))
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
        self.slider.set(120)
        self.slider.grid(row=0, column=1, sticky="ew", padx=(12, 0))

        self.l_slider = ctk.CTkLabel(sl_r, text=self.t("time_min", 2), width=140, anchor="e", font=F["subhead"], text_color=C["accent"])
        self.l_slider.grid(row=0, column=2, padx=(8, 0))

        # --- Yeni Eklenen Ayarlar (Modlar) ---
        opts_f = ctk.CTkFrame(p, fg_color="transparent")
        opts_f.grid(row=3, sticky="ew", padx=24, pady=5)
        
        self.newsletter_var = ctk.BooleanVar(value=bool(self.settings.get("newsletter_mode", True)))
        self.spam_check_var = ctk.BooleanVar(value=bool(self.settings.get("spam_check", True)))

        ctk.CTkCheckBox(opts_f, text=self.t("newsletter_mode"), variable=self.newsletter_var, font=F["subhead"], text_color=C["text"], fg_color=C["navy"]).grid(row=0, column=0, sticky="w", padx=(0, 20), pady=4)
        ctk.CTkCheckBox(opts_f, text=self.t("spam_check"), variable=self.spam_check_var, font=F["subhead"], text_color=C["text"], fg_color=C["navy"]).grid(row=1, column=0, sticky="w", pady=4)

        ctk.CTkFrame(p, height=1, fg_color=C["border"]).grid(row=4, sticky="ew", padx=24, pady=10)

        self.app_var = ctk.BooleanVar()
        ctk.CTkCheckBox(
            p,
            text=self.t("approval_check"),
            variable=self.app_var,
            font=F["subhead"],
            text_color=C["red"],
            checkmark_color="white",
            fg_color=C["red"],
            hover_color=C["red_h"]
        ).grid(row=5, sticky="w", padx=24, pady=(0, 10))

        self.progress_bar = ctk.CTkProgressBar(p, height=10, corner_radius=5, progress_color=C["accent"], fg_color=C["border"])
        self.progress_bar.set(0)
        self.progress_bar.grid(row=6, sticky="ew", padx=24)

        self.progress_label = ctk.CTkLabel(p, text=self.t("status_ready"), font=F["small"], text_color=C["muted"])
        self.progress_label.grid(row=7, sticky="w", padx=24, pady=(2, 8))

        ab = ctk.CTkFrame(p, fg_color="transparent")
        ab.grid(row=8, pady=4)

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

        sc = ctk.CTkFrame(p, fg_color="transparent")
        sc.grid(row=9, sticky="w", padx=24, pady=(6, 4))
        self.lbl_sent = self._chip(sc, self.t("lbl_sent", 0), C["green"])
        self.lbl_fail = self._chip(sc, self.t("lbl_fail", 0), C["red"])
        self.lbl_skip = self._chip(sc, self.t("lbl_skip", 0), C["amber"])

        self.log_view = ctk.CTkTextbox(
            p,
            height=210,
            fg_color=C["log_bg"],
            text_color=C["log_fg"],
            font=F["mono"],
            border_width=0,
            corner_radius=10
        )
        self.log_view.grid(row=10, sticky="ew", padx=24, pady=(6, 22))

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
        bar = ctk.CTkFrame(parent, height=3, fg_color=color, corner_radius=2)
        bar.grid(row=0, column=0, columnspan=2, sticky="ew", padx=24, pady=(16, 8))
        ctk.CTkLabel(parent, text=text, font=F["subhead"], text_color=color).grid(row=0, column=0, sticky="w", padx=28, pady=(16, 8))

    def _srow(self, parent, row_idx: int, label: str, attr: str, key: str, is_pass=False):
        ctk.CTkLabel(parent, text=label, anchor="e", width=270, font=F["body"], text_color=C["muted"]).grid(
            row=row_idx, column=0, padx=(24, 8), pady=6, sticky="e"
        )
        e = ctk.CTkEntry(parent, height=34, show="●" if is_pass else "", font=F["body"], border_color=C["border"])
        e.insert(0, self.settings.get(key, ""))
        e.grid(row=row_idx, column=1, padx=(0, 24), pady=6, sticky="ew")
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
                raise RuntimeError("Kullanıcı tarafından durduruldu")
            try:
                srv = self._smtp_connect(host, port, sec)
                srv.login(user, pw)
                return srv
            except Exception as e:
                last_err = e
                if attempt < SMTP_RECONNECT_RETRIES:
                    wait = SMTP_RECONNECT_BACKOFF * attempt
                    self.log(f"⚠  SMTP bağlantısı kurulamadı (deneme {attempt}/{SMTP_RECONNECT_RETRIES}): {e}")
                    self.log(f"⏳ {wait} sn sonra tekrar denenecek…")
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
        """Tracking bilgisini JSON log dosyasına kaydeder."""
        log_file = os.path.join(self.app_dir, "tracking_log.json")
        entry = {
            "tracking_id": tracking_id,
            "to_email": to_email,
            "subject": subject,
            "sent_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "message_id": message_id
        }
        try:
            if os.path.exists(log_file):
                with open(log_file, 'r', encoding='utf-8') as f:
                    logs = json.load(f)
            else:
                logs = []
            logs.append(entry)
            with open(log_file, 'w', encoding='utf-8') as f:
                json.dump(logs, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.log(f"⚠️ Tracking log kaydedilemedi: {e}")

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

    def mailing_process(self, emails):
        self._stop_flag.clear()
        self.run_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")

        sent = fail = skip = 0
        total = len(emails)

        imap_conn = None
        smtp = None
        msgs_in_conn = 0
        max_per_conn = self._max_per_connection()

        if self.settings.get("save_sent", True):
            try:
                imap_conn = self._imap_connect()
                self.log(f"📥 IMAP bağlandı → \"{self.settings.get('sent_folder', 'INBOX.Sent')}\"")
            except Exception as e:
                self.log(f"⚠  IMAP bağlantısı yok (gönderim devam eder): {e}")

        try:
            host = self.settings.get("server", "").strip()
            port = int(self.settings.get("port", "587"))
            sec = self.settings.get("security", "STARTTLS").upper()
            user = self.settings.get("login_username", "").strip() or self.settings.get("mail", "").strip()
            m_from = self.settings.get("mail", "").strip()
            m_pass = self.settings.get("pass", "")
            
            raw_subject = self.subject_entry.get().strip()
            subjects = [s.strip() for s in raw_subject.split(",") if s.strip()]
            if not subjects:
                subjects = ["Unixpadel Bilgilendirme"]

            if not all([host, port, user, m_from, m_pass]):
                self.log("❌ SMTP ayarları eksik — Sunucu Ayarları sayfasını kontrol edin.")
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

            self.log(f"🔗 Bağlanılıyor → {host}:{port} [{sec}]")
            smtp = self._smtp_session_open()
            self.log(f"✓  SMTP girişi başarılı  ·  oturum başına en fazla {max_per_conn} mesaj")
            self._prog(0, total, sent, fail, skip)

            for i, record in enumerate(emails, 1):
                if self._stop_flag.is_set():
                    self.log(self.t("log_stopped"))
                    break

                email = record.get("email", "")

                if not valid_email(email):
                    self.log(self.t("log_invalid", i, total, email))
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

                    current_subject = random.choice(subjects)
                    merged_html = html_body
                    merged_plain = base_plain_text

                    # Mail Merge: {{isim}} vs değişimi
                    for key, val in record.items():
                        if key == "email": continue
                        placeholder = f"{{{{{key}}}}}"
                        current_subject = current_subject.replace(placeholder, val)
                        if merged_html:
                            merged_html = merged_html.replace(placeholder, val)
                        if merged_plain:
                            merged_plain = merged_plain.replace(placeholder, val)

                    msg_obj, msg_bytes, tracking_id = self._build_msg(email, current_subject, merged_html, merged_plain)

                    try:
                        smtp.sendmail(m_from, [email], msg_bytes)
                    except SMTP_RETRYABLE_ERRORS as e:
                        # Kalıcı red (550/553 gibi) ise tekrar denemenin anlamı yok
                        if not self._is_retryable_smtp_error(e):
                            raise
                        self.log(self.t("log_conn_err", i, total, e))
                        self._smtp_close(smtp)
                        smtp = None
                        smtp = self._smtp_session_open()
                        msgs_in_conn = 0
                        smtp.sendmail(m_from, [email], msg_bytes)  # tek tekrar denemesi

                    msgs_in_conn += 1
                    sent += 1
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

                if i < total and not self._stop_flag.is_set():
                    wait = int(self.slider.get())
                    if wait:
                        self.log(self.t("log_wait", self._fmt_interval(wait)))
                        self._sleep_interruptible(wait)

            self.log(self.t("log_done", sent, fail, skip))
            self._prog(total, total, sent, fail, skip)

        except Exception as e:
            self.log(self.t("log_critical", e))
        finally:
            self._smtp_close(smtp)
            if imap_conn:
                try:
                    imap_conn.logout()
                except Exception:
                    pass
            self.run_btn.configure(state="normal")
            self.stop_btn.configure(state="disabled")

    # ──────────────────────── Test connections ────────────────────────────────

    def test_connection(self):
        self.save_settings(silent=True)

        def _run():
            try:
                host = self.settings["server"].strip()
                port = int(self.settings["port"])
                sec = self.settings["security"].upper()
                user = self.settings.get("login_username", "").strip() or self.settings["mail"].strip()
                pw = self.settings["pass"]

                self.log(f"🔌 SMTP test → {host}:{port} [{sec}]")
                srv = self._smtp_connect(host, port, sec)
                srv.login(user, pw)
                srv.quit()
                self.log("✓  SMTP bağlantı testi BAŞARILI")
                messagebox.showinfo("SMTP ✓", self.t("test_smtp_ok"))
            except Exception as e:
                self.log(f"❌ SMTP test BAŞARISIZ: {e}")
                messagebox.showerror(self.t("error"), str(e))

        threading.Thread(target=_run, daemon=True).start()

    def test_imap(self):
        self.save_settings(silent=True)

        def _run():
            try:
                self.log("📥 IMAP test başlıyor…")
                imap = self._imap_connect()
                _, lst = imap.list()
                imap.logout()
                self.log("✓  IMAP bağlantı testi BAŞARILI — Mevcut klasörler:")
                for item in (lst or []):
                    self.log(f"   {item.decode() if isinstance(item, bytes) else item}")
                messagebox.showinfo(
                    "IMAP ✓",
                    self.t("test_imap_ok", self.settings.get("sent_folder"))
                )
            except Exception as e:
                self.log(f"❌ IMAP test BAŞARISIZ: {e}")
                messagebox.showerror(self.t("error"), str(e))

        threading.Thread(target=_run, daemon=True).start()

    # ──────────────────────── Engine entry ────────────────────────────────────

    def start_engine(self):
        if not self.app_var.get():
            messagebox.showwarning(self.t("warning"), self.t("alert_approval"))
            return

        emails = self.get_emails()
        if not emails:
            self.log(self.t("alert_empty_list"))
            return

        self.save_settings(silent=True)
        is_plain_only = self.settings.get("template_mode", "html") == "text"

        if not is_plain_only and not self.selected_html:
            self.log(self.t("alert_no_html"))
            return

        # Spam kontrolü
        if bool(self.settings.get("spam_check", True)):
            spam_words = ["ücretsiz", "indirim", "fırsat", "satın al", "kampanya", "free", "discount", "offer"]
            if is_plain_only:
                text_to_check = (self.settings.get("subject", "") + " " + self.settings.get("direct_plain_text", "")).lower()
            else:
                text_to_check = (self.settings.get("subject", "") + " " + self.settings.get("plain_text", "")).lower()
            found_spam = [w for w in spam_words if w in text_to_check]
            if found_spam:
                ans = messagebox.askyesno(self.t("spam_warn_title"), self.t("spam_warn_desc", ", ".join(found_spam)))
                if not ans:
                    return

        self.save_settings(silent=True)

        self.progress_bar.set(0)
        self.progress_label.configure(text=self.t("status_ready"))
        self.lbl_sent.configure(text=self.t("lbl_sent", 0))
        self.lbl_fail.configure(text=self.t("lbl_fail", 0))
        self.lbl_skip.configure(text=self.t("lbl_skip", 0))

        threading.Thread(target=self.mailing_process, args=(self.get_emails(),), daemon=True).start()

    def get_emails(self):
        if hasattr(self, "manual_text") and self.manual_text.winfo_exists():
            raw = self.manual_text.get("1.0", "end")
            return [{"email": normalize_email(e)} for e in raw.splitlines() if normalize_email(e)]

        if self.selected_excel:
            try:
                df = pd.read_excel(self.selected_excel)
                df = df.fillna("")
                records = []
                for _, row in df.iterrows():
                    email = normalize_email(str(row.iloc[0]))
                    if email:
                        record = {"email": email}
                        for col_name in df.columns:
                            record[str(col_name).lower()] = str(row[col_name])
                        records.append(record)
                return records
            except Exception as e:
                self.log(f"Excel hatası: {e}")

        return []

    def load_html(self):
        p = filedialog.askopenfilename(filetypes=[("HTML Files", "*.html")])
        if p:
            self.selected_html = p
            n = os.path.basename(p)
            self.log(f"📄 HTML: {n}")
            if hasattr(self, "html_label"):
                self.html_label.configure(text=f"✓  {n}", text_color=C["green"])

    def load_excel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if p:
            self.selected_excel = p
            n = os.path.basename(p)
            self.log(f"📊 Excel: {n}")
            if hasattr(self, "excel_label"):
                self.excel_label.configure(text=f"✓  {n}", text_color=C["green"])

    # ──────────────────────── Misc helpers ────────────────────────────────────

    def _prog(self, cur, tot, sent, fail, skip):
        if tot:
            self.progress_bar.set(cur / tot)
            self.progress_label.configure(text=f"{cur} / {tot}  ({int(cur / tot * 100)}%)")
        self.lbl_sent.configure(text=self.t("lbl_sent", sent))
        self.lbl_fail.configure(text=self.t("lbl_fail", fail))
        self.lbl_skip.configure(text=self.t("lbl_skip", skip))

    def _request_stop(self):
        self._stop_flag.set()
        self.log(self.t("log_stop_req"))

    def _fmt_interval(self, seconds):
        seconds = int(seconds)
        m, s = divmod(seconds, 60)
        if m and s:
            return self.t("time_min_sec", m, s)
        if m:
            return self.t("time_min", m)
        return self.t("time_sec", s)

    def _update_slider(self, v):
        self.l_slider.configure(text=self._fmt_interval(v))

    def log(self, msg: str):
        if hasattr(self, "log_view") and self.log_view.winfo_exists():
            self.log_view.insert("end", f"[{time.strftime('%H:%M:%S')}]  {msg}\n")
            self.log_view.see("end")

    def clear_main(self):
        if hasattr(self, "main_content"):
            for w in self.main_content.winfo_children():
                w.destroy()

        for attr in [
            "s_mail", "s_server", "s_port", "s_security", "s_login_username",
            "s_pass", "s_from_name", "s_reply_to", "s_plain_text", "manual_text",
            "html_label", "excel_label", "s_imap_server", "s_imap_port",
            "s_imap_security", "s_imap_username", "s_imap_pass",
            "s_sent_folder", "s_save_sent",
        ]:
            if hasattr(self, attr):
                delattr(self, attr)

    def on_close(self):
        try:
            self.save_settings(silent=True)
        except Exception:
            pass
        self.destroy()


if __name__ == "__main__":
    app = UnixpadelApp()
    app.mainloop()