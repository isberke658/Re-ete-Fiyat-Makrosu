# -*- coding: utf-8 -*-
"""
================================================================================
LOGO ERP - OTOMATIK RECETE CIKARMA VE AKILLI DURUS TESPIT MAKROSU (v1.0)
Hücre 38 & Adli Müfettiş Uyumlu - Saf Native Windows API (ctypes) + Tkinter GUI
================================================================================
Özellikler:
1. Sabit Koordinat Kilidi: Fare nereye kayarsa kaysın her turda tam kaydedilen
   sağ tık ve menü noktasına otomatik geri döner.
2. Boşluk (SPACE) Tuşu Pompası: Reçete aktarım penceresi açıkken seri ve güvenli
   boşluk tuşu basarak sırayı eritir.
3. Akıllı Bitiş & Duruş Tespiti: Ekranda veya pencerede hareket durduğunda (duruş
   süresi eşiği, varsayılan 2.0 sn) sıranın bittiğini anlar.
4. Adli Süre Kaydı: Her turun başlama, duruş ve toplam süresini saniye hassasiyetinde
   'logo_recete_makro_log.txt' dosyasına ve ekrana yazar.
5. Otomatik Döngü: İşlem bitince belirlenen bekleme sonrası mouse'u tekrar aynı
   noktaya çekip bir sonraki turu başlatır.
6. Kısayol Tuşları (Global Hotkeys):
   - F8: Sağ Tıklanacak Noktayı Kaydet (Farenin olduğu yer)
   - F9: Menü Seçeneğinin ("Tüm reçeteleri çıkart") Noktasını Kaydet
   - F6: Makroyu Başlat / Devam Et
   - F7: Makroyu Duraklat / Durdur
   - ESC: Acil Durdurma
"""

import sys
import os
import time
import json
import threading
import datetime
import ctypes
import ctypes.wintypes as wt
import tkinter as tk
from tkinter import ttk, messagebox

# Windows API Yüklemeleri
user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32
kernel32 = ctypes.windll.kernel32

# DPI Farkındalığı (Ekran koordinatlarının şaşmaması için kritik)
try:
    user32.SetProcessDPIAware()
except Exception:
    pass

# Windows Sabitleri
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP   = 0x0004
MOUSEEVENTF_RIGHTDOWN = 0x0008
MOUSEEVENTF_RIGHTUP   = 0x0010
VK_SPACE = 0x20
KEYEVENTF_KEYUP = 0x0002

# Hotkey Sanal Tuş Kodları
VK_F6  = 0x75
VK_F7  = 0x76
VK_F8  = 0x77
VK_F9  = 0x78
VK_F10 = 0x79
VK_ESCAPE = 0x1B

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo_makro_ayarlar.json")
LOG_FILE    = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo_recete_makro_log.txt")


def log_to_file(msg: str):
    """Adli zaman damgalı log kaydı düşer."""
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {msg}\n"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line)
    except Exception as e:
        print(f"Log yazma hatası: {e}")


class LogoMacroEngine:
    def __init__(self, ui_callback=None):
        self.ui_callback = ui_callback
        self.running = False
        self.paused = False
        self.thread = None

        # Varsayılan Koordinatlar ve Parametreler
        self.coord_right_click = (500, 300)
        self.coord_menu_click  = (550, 350)
        
        # Süre ve Eşik Ayarları (saniye)
        self.space_interval_ms = 40       # Boşluk tuşu basım aralığı (ms)
        self.dialog_wait_sec   = 0.8      # Sağ tık sonrası menü ve pencere açılış beklemesi
        self.idle_timeout_sec  = 2.2      # Ekranda/işlemde duruş süresi eşiği (bitiş kabulü)
        self.cycle_cooldown_sec= 1.5      # Bir tur bitince diğerine geçmeden önceki bekleme
        self.move_down_row     = False    # Her turda bir alt satıra geçilsin mi? (Aşağı ok)
        self.auto_space        = True     # Boşluk tuşuna makro kendisi mi bassın? (True=Otomatik, False=Kullanıcı bassın)
        self.target_space_sec  = 12.0     # Kaydedilen veya belirlenen Space basma süresi (sn)
        self.use_fixed_duration= True     # Belirlenen süre kadar bas ve otomatik sonraki tura geç
        self.max_cycles        = 0        # 0 = Sonsuz döngü (kullanıcı F7/ESC ile durdurana kadar)

        self.current_cycle = 0
        self.total_extracted_time = 0.0

        self.load_config()

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.coord_right_click = tuple(data.get("coord_right_click", self.coord_right_click))
                    self.coord_menu_click  = tuple(data.get("coord_menu_click", self.coord_menu_click))
                    self.space_interval_ms = data.get("space_interval_ms", self.space_interval_ms)
                    self.idle_timeout_sec  = data.get("idle_timeout_sec", self.idle_timeout_sec)
                    self.cycle_cooldown_sec= data.get("cycle_cooldown_sec", self.cycle_cooldown_sec)
                    self.dialog_wait_sec   = data.get("dialog_wait_sec", self.dialog_wait_sec)
                    self.move_down_row     = data.get("move_down_row", self.move_down_row)
                    self.auto_space        = data.get("auto_space", self.auto_space)
                    self.target_space_sec  = data.get("target_space_sec", self.target_space_sec)
                    self.use_fixed_duration= data.get("use_fixed_duration", self.use_fixed_duration)
            except Exception as e:
                print(f"Ayar yükleme hatası: {e}")

    def save_config(self):
        data = {
            "coord_right_click": list(self.coord_right_click),
            "coord_menu_click":  list(self.coord_menu_click),
            "space_interval_ms": self.space_interval_ms,
            "idle_timeout_sec":  self.idle_timeout_sec,
            "cycle_cooldown_sec":self.cycle_cooldown_sec,
            "dialog_wait_sec":   self.dialog_wait_sec,
            "move_down_row":     self.move_down_row,
            "auto_space":        self.auto_space,
            "target_space_sec":  self.target_space_sec,
            "use_fixed_duration":self.use_fixed_duration
        }
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"Ayar kaydetme hatası: {e}")

    def get_cursor_pos(self):
        pt = wt.POINT()
        user32.GetCursorPos(ctypes.byref(pt))
        return (pt.x, pt.y)

    def set_cursor_pos(self, x, y):
        user32.SetCursorPos(int(x), int(y))

    def mouse_right_click(self, x, y):
        self.set_cursor_pos(x, y)
        time.sleep(0.08)
        user32.mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
        time.sleep(0.04)
        user32.mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)

    def mouse_left_click(self, x, y):
        self.set_cursor_pos(x, y)
        time.sleep(0.08)
        user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.04)
        user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

    def send_space_tap(self):
        """Hardware scan code (0x39) ile gerçek klavye bas-çek döngüsü gönderir."""
        scan_code = user32.MapVirtualKeyW(VK_SPACE, 0)
        user32.keybd_event(VK_SPACE, scan_code, 0, 0)
        time.sleep(0.025)
        user32.keybd_event(VK_SPACE, scan_code, KEYEVENTF_KEYUP, 0)

    def send_down_arrow(self):
        """Aşağı ok tuşu bas-çek (Alt satıra geçmek için opsiyonel)."""
        VK_DOWN = 0x28
        user32.keybd_event(VK_DOWN, 0, 0, 0)
        time.sleep(0.02)
        user32.keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, 0)

    def get_screen_pixel_samples(self, center_x, center_y, radius=120, step=25):
        """
        Pencerenin / işlemin hareketli olup olmadığını anlamak için
        ekranın merkezinden veya aktif bölgeden 25-30 adet piksel rengi örnekler.
        """
        samples = []
        hdc = user32.GetDC(0)
        if not hdc:
            return samples
        try:
            for dx in range(-radius, radius + 1, step):
                for dy in range(-radius, radius + 1, step):
                    px = gdi32.GetPixel(hdc, center_x + dx, center_y + dy)
                    if px != -1:  # Geçersiz pikselleri filtrele
                        samples.append(px)
        finally:
            user32.ReleaseDC(0, hdc)
        return samples

    def get_active_window_info(self):
        hwnd = user32.GetForegroundWindow()
        buf = ctypes.create_unicode_buffer(512)
        user32.GetWindowTextW(hwnd, buf, 512)
        return hwnd, buf.value

    def is_window_valid(self, hwnd):
        if not hwnd:
            return False
        return bool(user32.IsWindow(hwnd) and user32.IsWindowVisible(hwnd))

    def start(self):
        if self.running:
            return
        self.running = True
        self.paused = False
        self.thread = threading.Thread(target=self._run_loop, daemon=True)
        self.thread.start()

    def stop(self):
        self.running = False
        self.paused = False
        if self.ui_callback:
            self.ui_callback("STATUS", "DURDURULDU", "Makro durduruldu.")

    def _run_loop(self):
        self.current_cycle = 0
        if self.ui_callback:
            self.ui_callback("STATUS", "CALISIYOR", "Makro aktif, Logo operasyonu başlıyor...")

        log_to_file("=== LOGO RECETE MAKROSU BASLATILDI ===")

        while self.running:
            self.current_cycle += 1
            if self.ui_callback:
                self.ui_callback("CYCLE", self.current_cycle, f"Tur #{self.current_cycle} Başlıyor...")

            # -------------------------------------------------------------
            # ADIM 0: Opsiyonel Alt Satıra Geçiş
            # -------------------------------------------------------------
            if self.move_down_row and self.current_cycle > 1:
                self.send_down_arrow()
                time.sleep(0.15)

            # -------------------------------------------------------------
            # ADIM 1: Mouse'u Sabit Sağ Tık Noktasına Geri Getir ve Tıkla
            # -------------------------------------------------------------
            rx, ry = self.coord_right_click
            if self.ui_callback:
                self.ui_callback("ACTION", f"Tur #{self.current_cycle}", f"Mouse sabit noktaya kilitleniyor: ({rx}, {ry}) -> Sağ Tık")
            
            self.mouse_right_click(rx, ry)
            time.sleep(0.25)  # Menünün açılması için bekleme

            if not self.running:
                break

            # -------------------------------------------------------------
            # ADIM 2: Menü Seçeneğine Tıkla ("Tüm reçeteleri çıkart")
            # -------------------------------------------------------------
            mx, my = self.coord_menu_click
            if self.ui_callback:
                self.ui_callback("ACTION", f"Tur #{self.current_cycle}", f"Menü seçeneği tıklanıyor: ({mx}, {my})")
            
            self.mouse_left_click(mx, my)

            # Dialogun/Pencerenin açılması için kısa süre tanı
            time.sleep(self.dialog_wait_sec)

            if not self.running:
                break

            # Aktarım başladığı andaki pencereyi kaydet
            dialog_hwnd, dialog_title = self.get_active_window_info()

            # -------------------------------------------------------------
            # ADIM 3: Boşluk Tuşu Pompalama ve Akıllı Duruş Tespiti
            # -------------------------------------------------------------
            cycle_start_time = time.time()
            last_activity_time = time.time()
            
            screen_w = user32.GetSystemMetrics(0)
            screen_h = user32.GetSystemMetrics(1)
            cx, cy = screen_w // 2, screen_h // 2
            
            prev_samples = self.get_screen_pixel_samples(cx, cy)
            prev_hwnd, prev_title = dialog_hwnd, dialog_title

            if self.auto_space:
                status_txt = "Otomatik Boşluk (SPACE) pompalanıyor, duruş takip ediliyor..."
            else:
                status_txt = "Kullanıcı Boşluk (SPACE) basışı dinleniyor, duruş takip ediliyor..."

            if self.ui_callback:
                self.ui_callback("ACTION", f"Tur #{self.current_cycle}", status_txt)

            space_count = 0
            check_interval = 0.20  # Piksel ve durum kontrol sıklığı (sn)
            last_check_time = time.time()

            while self.running:
                # Boşluk basımı veya dinlemesi
                if self.auto_space:
                    self.send_space_tap()
                    space_count += 1
                    time.sleep(self.space_interval_ms / 1000.0)
                else:
                    # Kullanıcı kendisi mi basıyor? Fiziksel klavye dinleme
                    is_pressed = bool(user32.GetAsyncKeyState(VK_SPACE) & 0x8000)
                    if is_pressed:
                        space_count += 1
                        last_activity_time = time.time()  # Kullanıcı bastıkça canlı tut
                    time.sleep(0.03)

                now = time.time()
                elapsed = now - cycle_start_time

                # 1. KONTROL: Kaydedilen/Belirlenen Space süresi doldu mu? (Sabit Süre Modu)
                if self.use_fixed_duration:
                    rem = max(0.0, self.target_space_sec - elapsed)
                    if self.ui_callback and (space_count % 5 == 0):
                        self.ui_callback("ACTION", f"Tur #{self.current_cycle}", f"⚡ Boşluk pompalanıyor... (Kalan: {rem:.1f} / {self.target_space_sec:.1f} sn | {space_count} tuş)")
                    if elapsed >= self.target_space_sec:
                        # Belirlenen süre doldu! Tur bitti, sonraki partiye geç
                        break

                else:
                    # 2. KONTROL: Sabit süre modu kapalıysa, pencere ve duruş takibini devreye sok
                    if now - last_check_time >= check_interval:
                        last_check_time = now
                        current_samples = self.get_screen_pixel_samples(cx, cy)
                        curr_hwnd, curr_title = self.get_active_window_info()

                        # Pencere kapandı mı?
                        if dialog_hwnd and (dialog_hwnd != curr_hwnd) and not self.is_window_valid(dialog_hwnd):
                            break

                        screen_changed = (current_samples != prev_samples) if (current_samples and prev_samples) else False
                        window_changed = (curr_hwnd != prev_hwnd or curr_title != prev_title)

                        if screen_changed or window_changed:
                            last_activity_time = now
                            prev_samples = current_samples
                            prev_hwnd, prev_title = curr_hwnd, curr_title
                        else:
                            idle_duration = now - last_activity_time
                            if self.ui_callback:
                                self.ui_callback("IDLE_PROGRESS", round(idle_duration, 1), f"Duruş Takibi: {idle_duration:.1f} / {self.idle_timeout_sec:.1f} sn")

                            if idle_duration >= self.idle_timeout_sec:
                                break

            if not self.running:
                break

            # -------------------------------------------------------------
            # ADIM 4: Süre Hesabı ve Adli Kayıt
            # -------------------------------------------------------------
            cycle_end_time = time.time()
            # Net işlem süresi (duruş beklemesi hariç)
            net_duration = max(0.1, (cycle_end_time - cycle_start_time) - self.idle_timeout_sec)
            total_duration = cycle_end_time - cycle_start_time

            log_msg = (
                f"Tur #{self.current_cycle} Bitti | "
                f"Net Islem Suresi: {net_duration:.2f} sn | "
                f"Toplam Sure: {total_duration:.2f} sn | "
                f"Gonderilen Bosluk: {space_count} adet | "
                f"Durus Esiği: {self.idle_timeout_sec} sn"
            )
            log_to_file(log_msg)

            if self.ui_callback:
                self.ui_callback("LOG", f"Tur #{self.current_cycle}", log_msg)
                self.ui_callback("ACTION", f"Tur #{self.current_cycle}", f"Tamamlandı! Net Süre: {net_duration:.1f} sn. Sıradaki tur için {self.cycle_cooldown_sec} sn bekleniyor...")

            # -------------------------------------------------------------
            # ADIM 5: Bekleme ve Sonraki Tur Döngüsü
            # -------------------------------------------------------------
            if self.max_cycles > 0 and self.current_cycle >= self.max_cycles:
                if self.ui_callback:
                    self.ui_callback("STATUS", "TAMAMLANDI", f"Belirlenen maksimum tur ({self.max_cycles}) tamamlandı.")
                self.running = False
                break

            # Cooldown beklemesi
            cd_start = time.time()
            while self.running and (time.time() - cd_start < self.cycle_cooldown_sec):
                time.sleep(0.1)

        log_to_file("=== LOGO RECETE MAKROSU DURDURULDU ===\n")


class LogoMacroApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Logo ERP - Reçete Çıkarma & Duruş Tespiti Makrosu")
        self.root.geometry("640x580")
        self.root.minsize(580, 500)
        self.root.attributes("-topmost", True)  # Always on top

        self.engine = LogoMacroEngine(ui_callback=self.on_engine_event)
        self.learning_active = False
        self.space_rec_active = False
        self.space_rec_start = 0.0

        self._setup_style()
        self._build_ui()

        # Global Hotkey Dinleyici Thread
        self.hotkey_running = True
        self.hotkey_thread = threading.Thread(target=self._hotkey_listener, daemon=True)
        self.hotkey_thread.start()

        # Kapanış Eventi
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _setup_style(self):
        style = ttk.Style()
        style.theme_use("clam")
        # Koyu / Modern Palet
        self.bg_color = "#1e1e2e"
        self.card_bg = "#282a36"
        self.fg_color = "#f8f8f2"
        self.accent_color = "#50fa7b"
        self.warn_color = "#ffb86c"
        self.danger_color = "#ff5555"
        self.cyan_color = "#8be9fd"

        self.root.configure(bg=self.bg_color)

    def _build_ui(self):
        # ÜST BAŞLIK KARTI
        header = tk.Frame(self.root, bg=self.card_bg, padx=15, pady=10)
        header.pack(fill="x", padx=10, pady=(10, 5))

        title = tk.Label(
            header,
            text="⚡ LOGO ERP OTOMATİK REÇETE & DURUŞ MAKROSU",
            font=("Segoe UI", 12, "bold"),
            fg=self.cyan_color,
            bg=self.card_bg
        )
        title.pack(anchor="w")

        subtitle = tk.Label(
            header,
            text="Sabit Koordinat Kilidi | Boşluk Pompası | Adli Duruş Tespiti & Süre Kaydı",
            font=("Segoe UI", 9),
            fg="#a0a0c0",
            bg=self.card_bg
        )
        subtitle.pack(anchor="w")

        # DURUM VE GÖSTERGE PANELİ
        status_card = tk.Frame(self.root, bg=self.card_bg, padx=15, pady=10)
        status_card.pack(fill="x", padx=10, pady=5)

        f_status = tk.Frame(status_card, bg=self.card_bg)
        f_status.pack(fill="x")

        tk.Label(f_status, text="DURUM:", font=("Segoe UI", 10, "bold"), fg=self.fg_color, bg=self.card_bg).pack(side="left")
        self.lbl_status = tk.Label(f_status, text="HAZIR (Bekliyor)", font=("Segoe UI", 10, "bold"), fg=self.warn_color, bg=self.card_bg)
        self.lbl_status.pack(side="left", padx=10)

        self.lbl_cycle = tk.Label(f_status, text="Tur: 0", font=("Segoe UI", 10, "bold"), fg=self.accent_color, bg=self.card_bg)
        self.lbl_cycle.pack(side="right")

        self.lbl_action = tk.Label(
            status_card,
            text="Başlamak için F6'ya veya 'Başlat' butonuna basın.",
            font=("Segoe UI", 9, "italic"),
            fg="#c0caf5",
            bg=self.card_bg,
            wraplength=580,
            justify="left"
        )
        self.lbl_action.pack(anchor="w", pady=(5, 0))

        # KOORDİNAT VE AYAR KARTI
        settings_card = tk.Frame(self.root, bg=self.card_bg, padx=15, pady=10)
        settings_card.pack(fill="x", padx=10, pady=5)

        # CANLI ÖĞRENME VE SIFIRLAMA BUTONLARI
        top_rec_frame = tk.Frame(settings_card, bg=self.card_bg)
        top_rec_frame.pack(fill="x", pady=(0, 8))

        self.btn_learn = tk.Button(
            top_rec_frame,
            text="🎯 CANLI KAYIT: Tıklayarak Öğren (Sağ Tıkla -> Menüye Tıkla)",
            command=self.start_live_learning,
            bg="#bd93f9", fg="#282a36", font=("Segoe UI", 10, "bold"),
            relief="flat", padx=10, pady=6, cursor="hand2"
        )
        self.btn_learn.pack(side="left", fill="x", expand=True, padx=(0, 5))

        self.btn_reset = tk.Button(
            top_rec_frame,
            text="🔄 Kayıtları Sıfırla",
            command=self.cmd_reset_all,
            bg="#ff5555", fg="white", font=("Segoe UI", 9, "bold"),
            relief="flat", padx=10, pady=6, cursor="hand2"
        )
        self.btn_reset.pack(side="right")

        # Koordinat Satırı 1 (Sağ Tık)
        r1 = tk.Frame(settings_card, bg=self.card_bg)
        r1.pack(fill="x", pady=2)
        tk.Label(r1, text="1. Sağ Tık Noktası [F8]:", width=22, anchor="w", font=("Segoe UI", 9, "bold"), fg=self.fg_color, bg=self.card_bg).pack(side="left")
        
        self.entry_rx = tk.Entry(r1, width=6, font=("Consolas", 9), justify="center")
        self.entry_rx.insert(0, str(self.engine.coord_right_click[0]))
        self.entry_rx.pack(side="left", padx=2)

        self.entry_ry = tk.Entry(r1, width=6, font=("Consolas", 9), justify="center")
        self.entry_ry.insert(0, str(self.engine.coord_right_click[1]))
        self.entry_ry.pack(side="left", padx=2)

        btn_set_r = tk.Button(r1, text="Farenin Yerini Al (F8)", command=self.save_coord_r, bg="#3b4261", fg=self.fg_color, relief="flat", padx=6)
        btn_set_r.pack(side="left", padx=5)

        # Koordinat Satırı 2 (Menü Tık)
        r2 = tk.Frame(settings_card, bg=self.card_bg)
        r2.pack(fill="x", pady=2)
        tk.Label(r2, text="2. Menü Tık Noktası [F9]:", width=22, anchor="w", font=("Segoe UI", 9, "bold"), fg=self.fg_color, bg=self.card_bg).pack(side="left")
        
        self.entry_mx = tk.Entry(r2, width=6, font=("Consolas", 9), justify="center")
        self.entry_mx.insert(0, str(self.engine.coord_menu_click[0]))
        self.entry_mx.pack(side="left", padx=2)

        self.entry_my = tk.Entry(r2, width=6, font=("Consolas", 9), justify="center")
        self.entry_my.insert(0, str(self.engine.coord_menu_click[1]))
        self.entry_my.pack(side="left", padx=2)

        btn_set_m = tk.Button(r2, text="Farenin Yerini Al (F9)", command=self.save_coord_m, bg="#3b4261", fg=self.fg_color, relief="flat", padx=6)
        btn_set_m.pack(side="left", padx=5)

        btn_cd_m = tk.Button(r2, text="⏱️ 3 Sn Sonra Al", command=self.countdown_save_menu, bg="#6272a4", fg=self.fg_color, relief="flat", padx=6)
        btn_cd_m.pack(side="left", padx=2)

        # Space Süresi ve Canlı Ölçüm Satırı
        r_dur = tk.Frame(settings_card, bg=self.card_bg)
        r_dur.pack(fill="x", pady=(6, 2))

        tk.Label(r_dur, text="⏱️ Space Basma Süresi (sn):", font=("Segoe UI", 9, "bold"), fg=self.accent_color, bg=self.card_bg).pack(side="left")
        self.entry_space_dur = tk.Entry(r_dur, width=6, font=("Consolas", 10, "bold"), justify="center", fg="#50fa7b", bg="#1a1b26")
        self.entry_space_dur.insert(0, str(self.engine.target_space_sec))
        self.entry_space_dur.pack(side="left", padx=5)

        self.btn_rec_space = tk.Button(
            r_dur, text="🔴 Space Süresini Canlı Kaydet (Ölç)",
            command=self.toggle_space_recording,
            bg="#ff5555", fg="white", font=("Segoe UI", 9, "bold"),
            relief="flat", padx=8, pady=3, cursor="hand2"
        )
        self.btn_rec_space.pack(side="left", padx=5)

        btn_test_space = tk.Button(
            r_dur, text="⌨️ Test Et (3 sn)",
            command=self.test_space_pumping,
            bg="#6272a4", fg="white", font=("Segoe UI", 9),
            relief="flat", padx=6, pady=3, cursor="hand2"
        )
        btn_test_space.pack(side="left", padx=5)

        # Parametreler (Duruş Eşiği, Boşluk Hızı, Cooldown)
        r3 = tk.Frame(settings_card, bg=self.card_bg)
        r3.pack(fill="x", pady=(6, 2))
        
        tk.Label(r3, text="Duruş Eşiği (sn):", font=("Segoe UI", 9), fg=self.fg_color, bg=self.card_bg).pack(side="left")
        self.entry_idle = tk.Entry(r3, width=5, font=("Consolas", 9), justify="center")
        self.entry_idle.insert(0, str(self.engine.idle_timeout_sec))
        self.entry_idle.pack(side="left", padx=(3, 15))

        tk.Label(r3, text="Boşluk Hızı (ms):", font=("Segoe UI", 9), fg=self.fg_color, bg=self.card_bg).pack(side="left")
        self.entry_space = tk.Entry(r3, width=5, font=("Consolas", 9), justify="center")
        self.entry_space.insert(0, str(self.engine.space_interval_ms))
        self.entry_space.pack(side="left", padx=(3, 15))

        tk.Label(r3, text="Tur Arası Bekleme (sn):", font=("Segoe UI", 9), fg=self.fg_color, bg=self.card_bg).pack(side="left")
        self.entry_cd = tk.Entry(r3, width=5, font=("Consolas", 9), justify="center")
        self.entry_cd.insert(0, str(self.engine.cycle_cooldown_sec))
        self.entry_cd.pack(side="left", padx=(3, 5))

        # Opsiyonel ve Mod Seçenekleri
        r4 = tk.Frame(settings_card, bg=self.card_bg)
        r4.pack(fill="x", pady=(6, 2))
        
        self.var_fixed_dur = tk.BooleanVar(value=self.engine.use_fixed_duration)
        chk_fixed_dur = tk.Checkbutton(
            r4, text="Belirlenen Space süresi kadar bas ve otomatik sonraki tura geç (Döngü / Loop)",
            variable=self.var_fixed_dur, bg=self.card_bg, fg=self.accent_color,
            selectcolor="#1a1b26", activebackground=self.card_bg, font=("Segoe UI", 9, "bold")
        )
        chk_fixed_dur.pack(anchor="w")

        self.var_auto_space = tk.BooleanVar(value=self.engine.auto_space)
        chk_auto_space = tk.Checkbutton(
            r4, text="Otomatik Boşluk Basımı (Açık: Makro pompalar | Kapalı: Siz basarsınız, makro takip eder)",
            variable=self.var_auto_space, bg=self.card_bg, fg=self.fg_color,
            selectcolor="#1a1b26", activebackground=self.card_bg, font=("Segoe UI", 9)
        )
        chk_auto_space.pack(anchor="w")

        self.var_down_row = tk.BooleanVar(value=self.engine.move_down_row)
        chk_down = tk.Checkbutton(
            r4, text="Her tur sonunda bir alt satıra geç (Aşağı Ok bas - Opsiyonel)",
            variable=self.var_down_row, bg=self.card_bg, fg=self.cyan_color,
            selectcolor="#1a1b26", activebackground=self.card_bg, font=("Segoe UI", 9)
        )
        chk_down.pack(anchor="w")

        # BUTONLAR
        btn_bar = tk.Frame(self.root, bg=self.bg_color, pady=5)
        btn_bar.pack(fill="x", padx=10)

        self.btn_start = tk.Button(
            btn_bar, text="▶ BAŞLAT (F6)", command=self.cmd_start,
            bg="#28a745", fg="white", font=("Segoe UI", 10, "bold"),
            relief="flat", padx=15, pady=6, cursor="hand2"
        )
        self.btn_start.pack(side="left", padx=5)

        self.btn_stop = tk.Button(
            btn_bar, text="⏹ DURDUR (F7)", command=self.cmd_stop,
            bg="#dc3545", fg="white", font=("Segoe UI", 10, "bold"),
            relief="flat", padx=15, pady=6, cursor="hand2"
        )
        self.btn_stop.pack(side="left", padx=5)

        btn_save = tk.Button(
            btn_bar, text="💾 Ayarları Kaydet", command=self.cmd_save_settings,
            bg="#44475a", fg=self.fg_color, font=("Segoe UI", 9),
            relief="flat", padx=10, pady=6
        )
        btn_save.pack(side="right", padx=5)

        # LOG PENCERESİ
        log_frame = tk.Frame(self.root, bg=self.card_bg, padx=10, pady=8)
        log_frame.pack(fill="both", expand=True, padx=10, pady=(5, 10))

        tk.Label(log_frame, text="📋 İŞLEM VE ADLİ SÜRE GÜNLÜĞÜ (Log):", font=("Segoe UI", 9, "bold"), fg=self.cyan_color, bg=self.card_bg).pack(anchor="w")

        self.txt_log = tk.Text(log_frame, bg="#11111b", fg="#cdd6f4", font=("Consolas", 8), height=9, relief="flat", padx=5, pady=5)
        self.txt_log.pack(fill="both", expand=True, pady=(4, 0))

        self.append_log("Sistem hazır. 'Canlı Kayıt' butonuna basarak koordinatları kolayca öğretebilirsiniz.")
        self.append_log(f"Kayıt dosyası: {LOG_FILE}")

    def append_log(self, text):
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.txt_log.insert("end", f"[{timestamp}] {text}\n")
        self.txt_log.see("end")

    def save_coord_r(self, pos=None):
        if pos is None:
            pos = self.engine.get_cursor_pos()
        self.engine.coord_right_click = pos
        self.entry_rx.delete(0, "end")
        self.entry_rx.insert(0, str(pos[0]))
        self.entry_ry.delete(0, "end")
        self.entry_ry.insert(0, str(pos[1]))
        self.append_log(f"1. Sağ tık noktası kaydedildi: {pos}")
        try:
            kernel32.Beep(900, 100)
        except Exception:
            pass
        self.engine.save_config()

    def save_coord_m(self, pos=None):
        if pos is None:
            pos = self.engine.get_cursor_pos()
        self.engine.coord_menu_click = pos
        self.entry_mx.delete(0, "end")
        self.entry_mx.insert(0, str(pos[0]))
        self.entry_my.delete(0, "end")
        self.entry_my.insert(0, str(pos[1]))
        self.append_log(f"2. Menü tıklama noktası kaydedildi: {pos}")
        try:
            kernel32.Beep(1400, 150)
        except Exception:
            pass
        self.engine.save_config()

    def countdown_save_menu(self):
        """3 saniye geri sayım yapar ve farenin o an durduğu koordinatı alır (Menüyü açık tutmak için)."""
        def _cd():
            self.append_log("⏱️ Geri sayım başladı! Fareyi menüdeki seçeneğin üzerine götürün...")
            for s in [3, 2, 1]:
                self.lbl_action.config(text=f"Menüye götürün... {s} sn sonra kaydedilecek!")
                try:
                    kernel32.Beep(800, 80)
                except Exception:
                    pass
                time.sleep(1.0)
            
            pos = self.engine.get_cursor_pos()
            self.root.after(0, self.save_coord_m, pos)
            self.lbl_action.config(text=f"2. Menü Noktası Alındı: {pos}")

        threading.Thread(target=_cd, daemon=True).start()

    def cmd_reset_all(self):
        """Kullanıcının yanlış kayıtlarını anında sıfırlar ve temizler."""
        self.learning_active = False
        self.space_rec_active = False
        self.engine.coord_right_click = (0, 0)
        self.engine.coord_menu_click = (0, 0)

        self.entry_rx.delete(0, "end"); self.entry_rx.insert(0, "0")
        self.entry_ry.delete(0, "end"); self.entry_ry.insert(0, "0")
        self.entry_mx.delete(0, "end"); self.entry_mx.insert(0, "0")
        self.entry_my.delete(0, "end"); self.entry_my.insert(0, "0")

        self.btn_learn.config(text="🎯 CANLI KAYIT: Tıklayarak Öğren (Sağ Tıkla -> Menüye Tıkla)", bg="#bd93f9")
        self.btn_rec_space.config(text="🔴 Space Süresini Canlı Kaydet (Ölç)", bg="#ff5555", fg="white")

        self.engine.save_config()
        self.append_log("🔄 Kayıtlar sıfırlandı (0, 0). Yanlış seçim temizlendi.")
        self.lbl_action.config(text="Kayıtlar sıfırlandı. Yeniden 'Canlı Kayıt' butonuna basarak seçim yapabilirsiniz.")
        try:
            kernel32.Beep(600, 150)
        except Exception:
            pass

    def toggle_space_recording(self):
        """Space basılacak süreyi canlı ölçer ve durdurunca kaydedip hafızaya alır."""
        if not self.space_rec_active:
            self.space_rec_active = True
            self.space_rec_start = time.time()
            self.btn_rec_space.config(text="⏹️ Kaydı Durdur (Süreyi Kaydet)", bg="#50fa7b", fg="#282a36")
            self.append_log("🔴 Space süresi ölçümü başladı! Logo'da reçeteler bitince bu butona tekrar basın.")

            def _timer_thread():
                while self.space_rec_active:
                    elapsed = time.time() - self.space_rec_start
                    self.lbl_action.config(text=f"⏱️ Space Ölçülüyor: {elapsed:.1f} sn... (Bitince butona basın)")
                    time.sleep(0.1)

            threading.Thread(target=_timer_thread, daemon=True).start()
        else:
            self.space_rec_active = False
            total_sec = round(max(0.5, time.time() - self.space_rec_start), 1)
            self.engine.target_space_sec = total_sec
            self.entry_space_dur.delete(0, "end")
            self.entry_space_dur.insert(0, str(total_sec))
            self.btn_rec_space.config(text="🔴 Space Süresini Canlı Kaydet (Ölç)", bg="#ff5555", fg="white")
            self.engine.save_config()
            self.append_log(f"✅ Space süresi kaydedildi: {total_sec} saniye! Makro bu süre kadar basıp otomatik döngüye (loop) girecek.")
            self.lbl_action.config(text=f"Space süresi: {total_sec} sn kaydedildi. F6 ile döngüyü başlatabilirsiniz.")
            try:
                kernel32.Beep(1200, 200)
            except Exception:
                pass

    def test_space_pumping(self):
        """Kullanıcının space tuşunu test etmesi için 3 saniye geri sayıp 3 saniye kesintisiz space basar."""
        def _test():
            self.append_log("⌨️ Space Testi: 3 saniye sonra space basılacak! Lütfen Logo'ya veya test alanına tıklayın...")
            for s in [3, 2, 1]:
                self.lbl_action.config(text=f"Test için Logo'ya odaklanın... {s} sn!")
                try:
                    kernel32.Beep(800, 80)
                except Exception:
                    pass
                time.sleep(1.0)

            self.lbl_action.config(text="⌨️ TEST: Space tuşları gönderiliyor...")
            t_end = time.time() + 3.0
            count = 0
            while time.time() < t_end:
                self.engine.send_space_tap()
                count += 1
                time.sleep(self.engine.space_interval_ms / 1000.0)

            self.append_log(f"✅ Space Testi Başarılı! 3 saniyede {count} adet Space tuşu gönderildi.")
            self.lbl_action.config(text=f"Space Testi Bitti ({count} adet basıldı).")
            try:
                kernel32.Beep(1400, 150)
            except Exception:
                pass

        threading.Thread(target=_test, daemon=True).start()

    def start_live_learning(self):
        """
        Kullanıcı hiçbir klavye tuşuna basmadan:
        1. Adım: Logo'da satıra SAĞ TIKLAR -> Nokta 1 kaydedilir, ses çalar.
        2. Adım: Açılan menüde seçeneğe SOL TIKLAR -> Nokta 2 kaydedilir, çift ses çalar.
        Menü ASLA kapanmaz çünkü klavye basılmaz!
        """
        self.learning_active = True

        def _learn():
            self.append_log("🎯 CANLI ÖĞRENME MODU AÇIK:")
            self.append_log("-> 1. Adım: Logo'daki satıra normal bir şekilde SAĞ TIKLAYIN.")
            self.lbl_action.config(text="1. ADIM: Logo'daki satıra SAĞ TIKLAYIN... (İptal için 'Sıfırla'ya basın)")
            self.btn_learn.config(text="⏳ Logo'da SAĞ TIK bekleniyor...", bg="#ffb86c")

            # 1. SAĞ TIK BEKLE
            VK_RBUTTON = 0x02
            while user32.GetAsyncKeyState(VK_RBUTTON) & 0x8000:
                time.sleep(0.05)

            pos_r = None
            while self.learning_active:
                if user32.GetAsyncKeyState(VK_RBUTTON) & 0x8000:
                    pos_r = self.engine.get_cursor_pos()
                    try:
                        kernel32.Beep(900, 120)
                    except Exception:
                        pass
                    break
                time.sleep(0.02)

            if not self.learning_active or not pos_r:
                return

            self.root.after(0, self.save_coord_r, pos_r)
            self.append_log(f"✅ Sağ Tık Alındı: {pos_r}")
            self.append_log("-> 2. Adım: Menüde 'Tüm reçeteleri çıkart' seçeneğine SOL TIKLAYIN.")
            self.lbl_action.config(text="2. ADIM: Açılan menüdeki seçeneğe SOL TIKLAYIN...")
            self.btn_learn.config(text="⏳ Menüde SOL TIK bekleniyor...", bg="#ff79c6")

            # Sağ tık butonunun bırakılmasını bekle
            while user32.GetAsyncKeyState(VK_RBUTTON) & 0x8000 and self.learning_active:
                time.sleep(0.05)

            time.sleep(0.2)

            # 2. SOL TIK BEKLE
            VK_LBUTTON = 0x01
            pos_m = None
            while self.learning_active:
                if user32.GetAsyncKeyState(VK_LBUTTON) & 0x8000:
                    pos_m = self.engine.get_cursor_pos()
                    try:
                        kernel32.Beep(1400, 150)
                        time.sleep(0.08)
                        kernel32.Beep(1800, 200)
                    except Exception:
                        pass
                    break
                time.sleep(0.02)

            if not self.learning_active or not pos_m:
                return

            self.learning_active = False
            self.root.after(0, self.save_coord_m, pos_m)
            self.append_log(f"🎉 Menü Tıklaması Alındı: {pos_m}")
            self.append_log("✨ İki koordinat da başarıyla hafızaya alındı! F6 ile başlatabilirsiniz.")
            self.lbl_action.config(text=f"Kayıt Tamam! Sağ Tık: {pos_r} | Menü: {pos_m}")
            self.btn_learn.config(text="🎯 CANLI KAYIT: Tıklayarak Öğren (Hazır)", bg="#bd93f9")

        threading.Thread(target=_learn, daemon=True).start()

    def cmd_save_settings(self, show_msg=True):
        try:
            rx = int(self.entry_rx.get())
            ry = int(self.entry_ry.get())
            mx = int(self.entry_mx.get())
            my = int(self.entry_my.get())
            self.engine.coord_right_click = (rx, ry)
            self.engine.coord_menu_click = (mx, my)

            self.engine.target_space_sec = float(self.entry_space_dur.get())
            self.engine.use_fixed_duration = bool(self.var_fixed_dur.get())

            self.engine.idle_timeout_sec = float(self.entry_idle.get())
            self.engine.space_interval_ms = int(self.entry_space.get())
            self.engine.cycle_cooldown_sec = float(self.entry_cd.get())
            self.engine.move_down_row = bool(self.var_down_row.get())
            self.engine.auto_space = bool(self.var_auto_space.get())
            self.engine.save_config()
            self.append_log("Ayarlar kaydedildi.")
            if show_msg:
                messagebox.showinfo("Başarılı", "Ayarlar kaydedildi.")
        except Exception as e:
            if show_msg:
                messagebox.showerror("Hata", f"Geçersiz ayar değeri: {e}")

    def cmd_start(self):
        self.cmd_save_settings(show_msg=False)
        self.engine.start()

    def cmd_stop(self):
        self.engine.stop()

    def on_engine_event(self, event_type, *args):
        # Thread güvenli UI güncellemesi
        self.root.after(0, self._handle_engine_event, event_type, args)

    def _handle_engine_event(self, event_type, args):
        if event_type == "STATUS":
            status, desc = args[0], args[1]
            self.lbl_status.config(text=status)
            if status == "CALISIYOR":
                self.lbl_status.config(fg=self.accent_color)
            elif status == "TAMAMLANDI":
                self.lbl_status.config(fg=self.cyan_color)
            else:
                self.lbl_status.config(fg=self.danger_color)
            self.lbl_action.config(text=desc)

        elif event_type == "CYCLE":
            cycle_no, text = args[0], args[1]
            self.lbl_cycle.config(text=f"Tur: {cycle_no}")
            self.lbl_action.config(text=text)

        elif event_type == "ACTION":
            tur, desc = args[0], args[1]
            self.lbl_action.config(text=f"[{tur}] {desc}")

        elif event_type == "IDLE_PROGRESS":
            sec, desc = args[0], args[1]
            self.lbl_action.config(text=desc)

        elif event_type == "LOG":
            tur, msg = args[0], args[1]
            self.append_log(msg)

    def _hotkey_listener(self):
        """Arka planda F6, F7, F8, F9, ESC tuşlarını menüyü bozmadan sessizce dinler."""
        last_press = {}

        def is_pressed(vk, cooldown=0.3):
            now = time.time()
            if user32.GetAsyncKeyState(vk) & 0x8000:
                if now - last_press.get(vk, 0) > cooldown:
                    last_press[vk] = now
                    return True
            return False

        while self.hotkey_running:
            try:
                # F8 -> Sağ Tık Koordinatını Sessizce Al
                if is_pressed(VK_F8):
                    pos = self.engine.get_cursor_pos()
                    self.root.after(0, self.save_coord_r, pos)

                # F9 -> Menü Koordinatını Sessizce Al
                if is_pressed(VK_F9):
                    pos = self.engine.get_cursor_pos()
                    self.root.after(0, self.save_coord_m, pos)

                # F6 -> Başlat
                if is_pressed(VK_F6):
                    self.root.after(0, self.cmd_start)

                # F7 -> Durdur
                if is_pressed(VK_F7):
                    self.root.after(0, self.cmd_stop)

                # ESC -> Acil Durdur
                if is_pressed(VK_ESCAPE):
                    if self.engine.running:
                        self.root.after(0, self.cmd_stop)

            except Exception:
                pass
            time.sleep(0.04)

    def on_close(self):
        self.hotkey_running = False
        self.engine.stop()
        self.root.destroy()


def main():
    root = tk.Tk()
    app = LogoMacroApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
