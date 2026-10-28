# -*- coding: utf-8 -*-
"""
================================================================================
NETUPP APP - OTOMATİK REÇETE SÜRÜKLE-BIRAK, FİRE(5) & TIGER AKTARIM MAKROSU
Sürüm: v3.3 (Tekil Dosya Kilitleme + Klasör İzolasyonu + Donanım Fire=5)
Teknoloji: Saf Native Windows API (ctypes) + Tkinter Modern Dark GUI
================================================================================
Kritik Mühendislik İyileştirmeleri (v3.3):
1. KLASÖR İZOLASYONU:
   - '_Aktarilanlar' klasörü artık reçete klasörünün İÇİNDE DEĞİL, bir üst dizinde
     (yanında) oluşturulur.
   - Böylece Gezgin penceresinde en üstte asla sarı bir klasör durmaz!
   - 1. satırda HER ZAMAN gerçek bir Excel reçete dosyası bulunur.
2. TEKİL DOSYA KİLİTLEME (Çoklu Seçim / Mavi Kutu Engelleyici):
   - Sürükleme öncesi hedef dosyaya önce tek bir sol tık yapılır (yalnızca o dosya seçilir).
   - OLE Drag nesnesi kilitlendikten sonra doğrudan sağa kaydırılır.
   - Asla dikdörtgen seçim kutusu (marquee) açılmaz, asla birden fazla dosya sürüklenmez!
3. DONANIM SEVİYESİNDE FİRE=5:
   - Mamül satırına çift tıklayıp fiziksel VK_5 ile '5' yazar ve Enter ile onaylar.
4. İKİ AŞAMALI SPACE ONAYI:
   - 1. Yes ve 2. Başarılı OK pencerelerini otomatik Space basarak kapatır.
================================================================================
"""

import sys
import os
import time
import json
import shutil
import glob
import threading
import datetime
import ctypes
import ctypes.wintypes as wt
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# Windows API
user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32
kernel32 = ctypes.windll.kernel32

try:
    user32.SetProcessDPIAware()
except Exception:
    pass

# Windows Sabitleri
MOUSEEVENTF_LEFTDOWN   = 0x0002
MOUSEEVENTF_LEFTUP     = 0x0004
MOUSEEVENTF_RIGHTDOWN  = 0x0008
MOUSEEVENTF_RIGHTUP    = 0x0010
KEYEVENTF_KEYUP        = 0x0002
KEYEVENTF_UNICODE      = 0x0004

VK_RETURN  = 0x0D
VK_SPACE   = 0x20
VK_BACK    = 0x08
VK_TAB     = 0x09
VK_CONTROL = 0x11
VK_A       = 0x41
VK_V       = 0x56
VK_5       = 0x35

# Hotkey Kodları
VK_F6      = 0x75
VK_F7      = 0x76
VK_ESCAPE  = 0x1B

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "netupp_makro_ayarlar.json")
LOG_FILE    = os.path.join(os.path.dirname(os.path.abspath(__file__)), "netupp_aktarim_log.txt")


def log_to_file(msg: str):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {msg}\n"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line)
    except Exception as e:
        print(f"Log yazma hatası: {e}")


class NetuppMacroEngine:
    def __init__(self, ui_callback=None):
        self.ui_callback = ui_callback
        self.running = False
        self.thread = None

        # Klasör Ayarları
        self.source_dir = ""

        # Koordinatlar (Sol Gezgin -> Sağ Netupp düzeni)
        self.coord_file_drag   = (200, 130)   # Sol: Gezgin 1. sıradaki dosya
        self.coord_netupp_drop  = (550, 65)    # Sağ: 'Excel Reçete Şablonunu Buraya Sürükleyip Bırakın'
        self.coord_fire_field  = (915, 135)   # Sağ: Mamülün yanındaki 1. satır Fire hücresi
        self.coord_btn_transfer= (930, 930)   # Sağ Alt: 'Reçeteyi Tiger'a Aktar' butonu

        # Fire Ayarları
        self.use_fire_step     = True         # Fire adımı varsayılan AÇIK
        self.fire_value        = "5"          # Yazılacak fire değeri
        self.fire_delay        = 0.20         # Fire tıklaması sonrası bekleme

        # Zamanlama & Onay Parametreleri (saniye)
        self.drag_duration_sec = 0.50         # Soldan sağa pürüzsüz sürükleme süresi
        self.after_drop_wait   = 1.1          # Dosya bırakıldıktan sonra Netupp tablo yükleme beklemesi
        self.transfer_timeout  = 45.0         # Tiger aktarımı için maksimum bekleme süresi
        self.cycle_cooldown    = 1.5          # Tur arası Gezgin yenilenme beklemesi
        self.auto_space_confirm= True         # 1. ve 2. diyaloglarda otomatik SPACE basılsın mı?
        self.tiger_wait_between_spaces = 6.0  # 1. Space (Yes) ile 2. Space (Başarılı) arasındaki Tiger bekleme süresi (sn)

        self.current_cycle = 0
        self.load_config()

    def get_archive_dir(self):
        """
        Kritik: _Aktarilanlar klasörünü reçete klasörünün İÇİNDE DEĞİL,
        yanında (bir üst dizinde) açar. Böylece Gezgin'de 1. sırada asla sarı klasör durmaz!
        """
        if not self.source_dir:
            return ""
        parent = os.path.dirname(os.path.abspath(self.source_dir))
        folder_name = os.path.basename(os.path.abspath(self.source_dir)) + "_Aktarilanlar"
        return os.path.join(parent, folder_name)

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.source_dir = data.get("source_dir", self.source_dir)
                    self.coord_file_drag    = tuple(data.get("coord_file_drag", self.coord_file_drag))
                    self.coord_netupp_drop   = tuple(data.get("coord_netupp_drop", self.coord_netupp_drop))
                    self.coord_fire_field   = tuple(data.get("coord_fire_field", self.coord_fire_field))
                    self.coord_btn_transfer = tuple(data.get("coord_btn_transfer", self.coord_btn_transfer))
                    
                    self.use_fire_step      = bool(data.get("use_fire_step", self.use_fire_step))
                    self.fire_value         = str(data.get("fire_value", self.fire_value))
                    self.fire_delay         = float(data.get("fire_delay", self.fire_delay))
                    
                    self.drag_duration_sec  = float(data.get("drag_duration_sec", self.drag_duration_sec))
                    self.after_drop_wait    = float(data.get("after_drop_wait", self.after_drop_wait))
                    self.transfer_timeout   = float(data.get("transfer_timeout", self.transfer_timeout))
                    self.cycle_cooldown     = float(data.get("cycle_cooldown", self.cycle_cooldown))
                    self.auto_space_confirm = bool(data.get("auto_space_confirm", self.auto_space_confirm))
                    self.tiger_wait_between_spaces = float(data.get("tiger_wait_between_spaces", self.tiger_wait_between_spaces))
            except Exception as e:
                print(f"Ayar yükleme hatası: {e}")

    def save_config(self):
        data = {
            "source_dir": self.source_dir,
            "coord_file_drag":    list(self.coord_file_drag),
            "coord_netupp_drop":  list(self.coord_netupp_drop),
            "coord_fire_field":   list(self.coord_fire_field),
            "coord_btn_transfer": list(self.coord_btn_transfer),
            "use_fire_step":      self.use_fire_step,
            "fire_value":         self.fire_value,
            "fire_delay":         self.fire_delay,
            "drag_duration_sec":  self.drag_duration_sec,
            "after_drop_wait":    self.after_drop_wait,
            "transfer_timeout":   self.transfer_timeout,
            "cycle_cooldown":     self.cycle_cooldown,
            "auto_space_confirm": self.auto_space_confirm,
            "tiger_wait_between_spaces": self.tiger_wait_between_spaces
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

    def mouse_click(self, x, y, double_click=False):
        self.set_cursor_pos(x, y)
        time.sleep(0.08)
        user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.05)
        user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        if double_click:
            time.sleep(0.08)
            user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            time.sleep(0.05)
            user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

    def single_file_drag_and_drop(self, x1, y1, x2, y2, steps=35):
        """
        Windows Explorer Saf Sürükle-Bırak (Çift Tıklama / Excel Açılma Kesin Önleyici):
        1. Öncesinde asla fazladan tıklama (bas-çek) yapmaz; çift tık ihtimalini sıfırlar.
        2. Sol tuşa basar basmaz anında 6 piksel sağa hareket ettirerek Windows OLE Drag
           modunu başlatır. Böylece Windows çift tık zamanlayıcısını anında iptal eder.
        3. Pürüzsüzce Netupp penceresindeki bırakma alanına taşır ve bırakır.
        """
        # Güvenlik: Sol tuşun serbest olduğundan emin ol
        user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(0.04)

        # 1. Fareyi dosyanın simgesine götür
        self.set_cursor_pos(x1, y1)
        time.sleep(0.10)

        # 2. Sol tuşa basılı tut
        user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.03)

        # 3. Anında 6 piksel sağa hareket et (Windows OLE Drag modunu derhal tetikle)
        self.set_cursor_pos(x1 + 6, y1)
        time.sleep(0.02)

        # 4. Sağa doğru pürüzsüz hareket et
        dt = self.drag_duration_sec / max(1, steps)
        for i in range(1, steps + 1):
            cx = int((x1 + 6) + (x2 - (x1 + 6)) * (i / steps))
            cy = int(y1 + (y2 - y1) * (i / steps))
            self.set_cursor_pos(cx, cy)
            time.sleep(dt)

        # 5. Netupp üzerinde bekle ve bırak
        time.sleep(0.20)
        user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(0.15)

    def send_space_key(self):
        """SPACE (Boşluk) tuşuna basar."""
        sc = user32.MapVirtualKeyW(VK_SPACE, 0)
        user32.keybd_event(VK_SPACE, sc, 0, 0)
        time.sleep(0.05)
        user32.keybd_event(VK_SPACE, sc, KEYEVENTF_KEYUP, 0)

    def send_enter_key(self):
        """Enter tuşuna basar."""
        sc = user32.MapVirtualKeyW(VK_RETURN, 0)
        user32.keybd_event(VK_RETURN, sc, 0, 0)
        time.sleep(0.04)
        user32.keybd_event(VK_RETURN, sc, KEYEVENTF_KEYUP, 0)

    def confirm_active_dialog(self):
        """Sadece SPACE (Boşluk) tuşuna basar. Enter tuşu gönderilmez (Excel dosyasının açılmasını kesin önler!)."""
        self.send_space_key()

    def write_fire_value(self, x, y, value="5"):
        """Mamül Fire hücresine donanım düzeyinde '5' yazar."""
        self.set_cursor_pos(x, y)
        time.sleep(0.10)
        user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.05)
        user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(0.12)

        # Çift tık
        user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.05)
        user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(0.06)
        user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.05)
        user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(0.15)

        # Ctrl + A
        sc_ctrl = user32.MapVirtualKeyW(VK_CONTROL, 0)
        sc_a    = user32.MapVirtualKeyW(VK_A, 0)
        user32.keybd_event(VK_CONTROL, sc_ctrl, 0, 0)
        time.sleep(0.03)
        user32.keybd_event(VK_A, sc_a, 0, 0)
        time.sleep(0.03)
        user32.keybd_event(VK_A, sc_a, KEYEVENTF_KEYUP, 0)
        user32.keybd_event(VK_CONTROL, sc_ctrl, KEYEVENTF_KEYUP, 0)
        time.sleep(0.06)

        # Backspace
        sc_back = user32.MapVirtualKeyW(VK_BACK, 0)
        user32.keybd_event(VK_BACK, sc_back, 0, 0)
        time.sleep(0.03)
        user32.keybd_event(VK_BACK, sc_back, KEYEVENTF_KEYUP, 0)
        time.sleep(0.08)

        # Donanım '5' tuşu
        for ch in str(value):
            if ch.isdigit():
                vk_code = 0x30 + int(ch)
                sc_code = user32.MapVirtualKeyW(vk_code, 0)
                user32.keybd_event(vk_code, sc_code, 0, 0)
                time.sleep(0.04)
                user32.keybd_event(vk_code, sc_code, KEYEVENTF_KEYUP, 0)
                time.sleep(0.04)

        time.sleep(0.10)
        sc_ret = user32.MapVirtualKeyW(VK_RETURN, 0)
        user32.keybd_event(VK_RETURN, sc_ret, 0, 0)
        time.sleep(0.04)
        user32.keybd_event(VK_RETURN, sc_ret, KEYEVENTF_KEYUP, 0)
        time.sleep(0.15)

    def get_active_window_info(self):
        hwnd = user32.GetForegroundWindow()
        buf = ctypes.create_unicode_buffer(512)
        user32.GetWindowTextW(hwnd, buf, 512)
        return hwnd, buf.value

    def clean_internal_archive_folder(self):
        """
        Eğer kullanıcının reçete klasörü içinde önceden açılmış '_Aktarilanlar'
        klasörü varsa, onu otomatik olarak bir üst dizine taşır.
        Böylece Gezgin'de 1. sırada asla sarı klasör durmaz!
        """
        old_internal = os.path.join(self.source_dir, "_Aktarilanlar")
        if os.path.exists(old_internal) and os.path.isdir(old_internal):
            new_external = self.get_archive_dir()
            os.makedirs(new_external, exist_ok=True)
            for f in os.listdir(old_internal):
                src = os.path.join(old_internal, f)
                dst = os.path.join(new_external, f)
                try:
                    if os.path.exists(dst):
                        os.remove(dst)
                    shutil.move(src, dst)
                except Exception:
                    pass
            try:
                os.rmdir(old_internal)
                log_to_file("Reçete klasörü içindeki eski '_Aktarilanlar' klasörü dışarı taşındı ve temizlendi.")
            except Exception:
                pass

    def get_pending_files(self):
        """Kaynak klasördeki bekleyen Excel dosyalarını alfabetik sırada getirir."""
        if not self.source_dir or not os.path.isdir(self.source_dir):
            return []
        files = []
        for f in os.listdir(self.source_dir):
            if f.lower().endswith((".xlsx", ".xls")) and not f.startswith("~$"):
                full_path = os.path.join(self.source_dir, f)
                if os.path.isfile(full_path):
                    files.append(f)
        files.sort()
        return files

    def start(self):
        if self.running:
            return
        if not self.source_dir or not os.path.isdir(self.source_dir):
            if self.ui_callback:
                self.ui_callback("ERROR", "Geçersiz Klasör", "Lütfen önce Excel reçetelerinin olduğu geçerli bir klasör seçin!")
            return

        self.running = True
        self.thread = threading.Thread(target=self._run_loop, daemon=True)
        self.thread.start()

    def stop(self):
        self.running = False
        if self.ui_callback:
            self.ui_callback("STATUS", "DURDURULDU", "Makro durduruldu.")

    def _run_loop(self):
        self.current_cycle = 0

        # 1. Eski iç klasör varsa dışarı taşı ve Gezgin'i temizle
        self.clean_internal_archive_folder()

        # 2. Yeni dış arşiv klasörünü hazırla
        archive_dir = self.get_archive_dir()
        if archive_dir and not os.path.exists(archive_dir):
            os.makedirs(archive_dir, exist_ok=True)

        if self.ui_callback:
            self.ui_callback("STATUS", "CALISIYOR", "Netupp aktarım süreci başladı...")

        log_to_file("=== NETUPP OTOMATIK RECETE AKTARIMI BASLATILDI (v3.3) ===")

        while self.running:
            pending = self.get_pending_files()
            if not pending:
                if self.ui_callback:
                    self.ui_callback("STATUS", "TAMAMLANDI", "🎉 Tebrikler! Klasördeki TÜM reçeteler başarıyla aktarıldı!")
                log_to_file("Klasördeki tüm dosyalar aktarıldı. Süreç başarıyla tamamlandı.")
                try:
                    kernel32.Beep(1000, 200)
                    time.sleep(0.1)
                    kernel32.Beep(1500, 300)
                except Exception:
                    pass
                self.running = False
                break

            current_file_name = pending[0]
            current_file_path = os.path.join(self.source_dir, current_file_name)
            self.current_cycle += 1

            if self.ui_callback:
                self.ui_callback("CYCLE", self.current_cycle, len(pending), current_file_name)

            # -------------------------------------------------------------
            # ADIM 1: Sadece Tekil Dosyayı Seç ve Sağdaki Netupp'a Sürükle
            # -------------------------------------------------------------
            fx, fy = self.coord_file_drag
            dx, dy = self.coord_netupp_drop

            if self.ui_callback:
                self.ui_callback("ACTION", f"Tur #{self.current_cycle}", f"'{current_file_name}' tekil kilitlenip sürükleniyor...")

            self.single_file_drag_and_drop(fx, fy, dx, dy)
            time.sleep(self.after_drop_wait)

            if not self.running:
                break

            # -------------------------------------------------------------
            # ADIM 2: Mamülün Yanındaki 1. Satır Fire Kutusuna '5' Yaz
            # -------------------------------------------------------------
            if self.use_fire_step:
                fire_x, fire_y = self.coord_fire_field
                if self.ui_callback:
                    self.ui_callback("ACTION", f"Tur #{self.current_cycle}", f"Mamül Fire kutusuna ({fire_x}, {fire_y}) '5' yazılıyor...")

                self.write_fire_value(fire_x, fire_y, self.fire_value)
                time.sleep(0.2)

            if not self.running:
                break

            # -------------------------------------------------------------
            # ADIM 3: 'Reçeteyi Tiger'a Aktar' Butonuna Tıkla
            # -------------------------------------------------------------
            tx, ty = self.coord_btn_transfer
            if self.ui_callback:
                self.ui_callback("ACTION", f"Tur #{self.current_cycle}", "'Tiger'a Aktar' butonuna tıklanıyor...")

            initial_hwnd, initial_title = self.get_active_window_info()
            self.mouse_click(tx, ty)
            transfer_start_time = time.time()

            if not self.running:
                break

            # -------------------------------------------------------------
            # ADIM 4 (1. AŞAMA ONAY): 'Tiger'a Aktar' Ön Onay Diyaloğu (Yes/No)
            # -------------------------------------------------------------
            if self.auto_space_confirm:
                if self.ui_callback:
                    self.ui_callback("ACTION", f"Tur #{self.current_cycle}", "1. Onay: 'Aktarım kaydı oluşturulsun mu? [Yes]' bekleniyor...")

                first_dialog_found = False
                wait_first = time.time()
                while self.running and (time.time() - wait_first < 6.0):
                    time.sleep(0.15)
                    c_hwnd, c_title = self.get_active_window_info()
                    title_l = c_title.lower()

                    # 1. Diyalog ("Tiger'a Aktar" başlıklı pencere)
                    if ("tiger" in title_l or "aktar" in title_l or (c_hwnd != initial_hwnd and c_hwnd != 0)) and (time.time() - wait_first >= 0.3):
                        first_dialog_found = True
                        time.sleep(0.3)
                        if self.ui_callback:
                            self.ui_callback("ACTION", f"Tur #{self.current_cycle}", "1. Onay: 'Yes' butonuna SPACE basılıyor...")
                        self.confirm_active_dialog()
                        time.sleep(0.5)
                        break

                if not first_dialog_found and self.running:
                    self.confirm_active_dialog()
                    time.sleep(0.4)

            if not self.running:
                break

            # -------------------------------------------------------------
            # ADIM 5 (2. AŞAMA ONAY): Tiger ERP Reçete İşleme & 'Başarılı' [OK] Diyaloğu
            # -------------------------------------------------------------
            if self.auto_space_confirm:
                wait_second = time.time()
                target_wait = max(2.5, self.tiger_wait_between_spaces)
                second_dialog_found = False

                log_to_file(f"Tur #{self.current_cycle}: Tiger ERP aktarımı başladı, {target_wait:.1f} sn bekleme ve 'Başarılı' penceresi bekleniyor...")

                while self.running and (time.time() - wait_second < self.transfer_timeout):
                    time.sleep(0.20)
                    elapsed = time.time() - wait_second

                    c_hwnd, c_title = self.get_active_window_info()
                    title_l = c_title.lower()

                    if self.ui_callback:
                        self.ui_callback("ACTION", f"Tur #{self.current_cycle}", f"⏳ Tiger ERP reçeteyi işliyor... ({elapsed:.1f} / {target_wait:.0f} sn)")

                    # Başarılı diyalog başlığı: SADECE "başarılı", "basarili", "tamamlandı", "tamamlandi", "tiger aktarımı"
                    # KESİNLİKLE 'recete', 'netupp', 'ok' gibi genel ana pencere kelimeleri aranmaz!
                    is_success_dialog = any(k in title_l for k in [
                        "başarılı", "basarili", "tamamlandı", "tamamlandi", "tiger aktarımı"
                    ])

                    # 1. Öncelik: Gerçek "Başarılı" penceresi geldiyse
                    if is_success_dialog and (elapsed >= 1.5):
                        second_dialog_found = True
                        time.sleep(0.35)
                        if self.ui_callback:
                            self.ui_callback("ACTION", f"Tur #{self.current_cycle}", "🎉 'Başarılı [OK]' penceresi tespit edildi! SPACE basılıyor...")
                        self.confirm_active_dialog()
                        time.sleep(0.6)
                        break

                    # 2. Öncelik: Kullanıcının belirlediği bekleme süresi dolduysa
                    if elapsed >= target_wait:
                        second_dialog_found = True
                        time.sleep(0.2)
                        if self.ui_callback:
                            self.ui_callback("ACTION", f"Tur #{self.current_cycle}", f"2. Onay: {target_wait:.0f} sn süre doldu, 'Başarılı [OK]' için SPACE basılıyor...")
                        self.confirm_active_dialog()
                        time.sleep(0.6)
                        break

                if not second_dialog_found and self.running:
                    self.confirm_active_dialog()
                    time.sleep(0.5)

            else:
                time.sleep(self.tiger_wait_between_spaces)

            if not self.running:
                break

            # -------------------------------------------------------------
            # ADIM 6: Aktarılan Dosyayı Yan Klasöre Taşı
            # -------------------------------------------------------------
            try:
                dest_path = os.path.join(archive_dir, current_file_name)
                if os.path.exists(dest_path):
                    os.remove(dest_path)
                shutil.move(current_file_path, dest_path)
                
                duration = time.time() - transfer_start_time
                log_msg = f"Tur #{self.current_cycle} Başarılı | '{current_file_name}' (Fire: {self.fire_value}) Tiger'a aktarıldı. (Süre: {duration:.1f} sn)"
                log_to_file(log_msg)

                if self.ui_callback:
                    self.ui_callback("LOG", f"Tur #{self.current_cycle}", log_msg)
            except Exception as e:
                err_msg = f"Dosya taşıma hatası ({current_file_name}): {e}"
                log_to_file(err_msg)
                if self.ui_callback:
                    self.ui_callback("LOG", "HATA", err_msg)

            # -------------------------------------------------------------
            # ADIM 7: Sonraki Tur İçin Cooldown Beklemesi
            # -------------------------------------------------------------
            if self.ui_callback:
                self.ui_callback("ACTION", f"Tur #{self.current_cycle}", f"Tamamlandı! Sıradaki dosya için {self.cycle_cooldown} sn bekleniyor...")

            cd_start = time.time()
            while self.running and (time.time() - cd_start < self.cycle_cooldown):
                time.sleep(0.1)

        log_to_file("=== NETUPP OTOMATIK RECETE AKTARIMI DURDURULDU ===\n")


class NetuppMacroApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Netupp App - Reçete Sürükle, Fire(5) & Tiger Aktarım Makrosu (v3.6)")
        self.root.geometry("740x760")
        self.root.minsize(700, 700)
        self.root.attributes("-topmost", True)

        self.engine = NetuppMacroEngine(ui_callback=self.on_engine_event)
        self.learning_active = False

        self._setup_style()
        self._build_ui()

        # Global Hotkey Listener
        self.hotkey_running = True
        self.hotkey_thread = threading.Thread(target=self._hotkey_listener, daemon=True)
        self.hotkey_thread.start()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.refresh_file_count()

    def _setup_style(self):
        style = ttk.Style()
        style.theme_use("clam")
        self.bg_color     = "#1e1e2e"
        self.card_bg    = "#282a36"
        self.fg_color     = "#f8f8f2"
        self.accent_color = "#50fa7b"
        self.warn_color   = "#ffb86c"
        self.danger_color = "#ff5555"
        self.cyan_color   = "#8be9fd"
        self.purple_color = "#bd93f9"

        self.root.configure(bg=self.bg_color)

    def _build_ui(self):
        # BAŞLIK
        header = tk.Frame(self.root, bg=self.card_bg, padx=15, pady=8)
        header.pack(fill="x", padx=10, pady=(8, 4))

        title = tk.Label(
            header, text="🚀 NETUPP APP - OTOMATİK REÇETE SÜRÜKLE, FİRE(5) & TIGER AKTARIM (v3.6)",
            font=("Segoe UI", 10, "bold"), fg=self.cyan_color, bg=self.card_bg
        )
        title.pack(anchor="w")

        subtitle = tk.Label(
            header, text="Akıllı Çift Space Zamanlaması | Sıfır Çift Tık (Excel Açılmaz) | Klasör İzolasyonu | Fire=5",
            font=("Segoe UI", 8), fg="#a0a0c0", bg=self.card_bg
        )
        subtitle.pack(anchor="w")

        # KLASÖR VE KUYRUK KARTI
        folder_card = tk.Frame(self.root, bg=self.card_bg, padx=15, pady=8)
        folder_card.pack(fill="x", padx=10, pady=4)

        f_top = tk.Frame(folder_card, bg=self.card_bg)
        f_top.pack(fill="x")

        tk.Label(f_top, text="📁 Reçeteler Klasörü (Soldaki Pencere):", font=("Segoe UI", 9, "bold"), fg=self.fg_color, bg=self.card_bg).pack(side="left")
        self.lbl_file_count = tk.Label(f_top, text="Kalan: 0 Dosya", font=("Segoe UI", 9, "bold"), fg=self.accent_color, bg=self.card_bg)
        self.lbl_file_count.pack(side="right")

        f_path = tk.Frame(folder_card, bg=self.card_bg)
        f_path.pack(fill="x", pady=(4, 0))

        self.entry_dir = tk.Entry(f_path, font=("Consolas", 9), fg="#50fa7b", bg="#1a1b26")
        self.entry_dir.insert(0, self.engine.source_dir)
        self.entry_dir.pack(side="left", fill="x", expand=True, padx=(0, 5))

        btn_browse = tk.Button(f_path, text="Gözat...", command=self.cmd_browse_folder, bg="#6272a4", fg="white", font=("Segoe UI", 9), relief="flat", padx=8)
        btn_browse.pack(side="right")

        # FIRE AYAR PANELİ
        fire_panel = tk.Frame(self.root, bg="#1a1b26", padx=15, pady=8, highlightbackground=self.warn_color, highlightthickness=1)
        fire_panel.pack(fill="x", padx=10, pady=4)

        fp_top = tk.Frame(fire_panel, bg="#1a1b26")
        fp_top.pack(fill="x")

        self.var_use_fire = tk.BooleanVar(value=self.engine.use_fire_step)
        tk.Checkbutton(
            fp_top, text="🔥 Mamülün Yanındaki 1. Satır Fire Kutusuna '5' Yaz (Aktif)",
            variable=self.var_use_fire, bg="#1a1b26", fg=self.warn_color,
            selectcolor="#282a36", activebackground="#1a1b26", font=("Segoe UI", 9, "bold")
        ).pack(side="left")

        tk.Label(fp_top, text="Değer:", font=("Segoe UI", 8, "bold"), fg=self.fg_color, bg="#1a1b26").pack(side="left", padx=(15, 2))
        self.entry_fire_val = tk.Entry(fp_top, width=4, font=("Consolas", 9, "bold"), fg="#ff5555", bg="#282a36", justify="center")
        self.entry_fire_val.insert(0, str(self.engine.fire_value))
        self.entry_fire_val.pack(side="left")

        fp_controls = tk.Frame(fire_panel, bg="#1a1b26")
        fp_controls.pack(fill="x", pady=(6, 2))

        tk.Label(fp_controls, text="Fire Koordinatı: X=", font=("Segoe UI", 8), fg="#c0caf5", bg="#1a1b26").pack(side="left")
        self.entry_fire_x = tk.Entry(fp_controls, width=5, font=("Consolas", 8), justify="center")
        self.entry_fire_x.insert(0, str(self.engine.coord_fire_field[0]))
        self.entry_fire_x.pack(side="left", padx=2)

        tk.Label(fp_controls, text="Y=", font=("Segoe UI", 8), fg="#c0caf5", bg="#1a1b26").pack(side="left")
        self.entry_fire_y = tk.Entry(fp_controls, width=5, font=("Consolas", 8), justify="center")
        self.entry_fire_y.insert(0, str(self.engine.coord_fire_field[1]))
        self.entry_fire_y.pack(side="left", padx=2)

        btn_pick_fire = tk.Button(
            fp_controls, text="🎯 MOUSE İLE FIRE KUTUSUNA TIKLA SEÇ", command=lambda: self.pick_single_coord("fire"),
            bg=self.warn_color, fg="#1e1e2e", font=("Segoe UI", 8, "bold"), relief="flat", padx=8, pady=3, cursor="hand2"
        )
        btn_pick_fire.pack(side="left", padx=(10, 6))

        btn_test_fire = tk.Button(
            fp_controls, text="⚡ ŞİMDİ 5 YAZ TESTİ YAP", command=self.cmd_test_fire_now,
            bg=self.accent_color, fg="#1e1e2e", font=("Segoe UI", 8, "bold"), relief="flat", padx=8, pady=3, cursor="hand2"
        )
        btn_test_fire.pack(side="left", padx=2)

        # DİĞER KOORDİNATLAR KARTI
        coord_card = tk.Frame(self.root, bg=self.card_bg, padx=15, pady=8)
        coord_card.pack(fill="x", padx=10, pady=4)

        top_rec_frame = tk.Frame(coord_card, bg=self.card_bg)
        top_rec_frame.pack(fill="x", pady=(0, 6))

        self.btn_learn = tk.Button(
            top_rec_frame, text="🎯 TÜM NOKTALARI SIRAYLA ÖĞREN (1.Dosya -> 2.Bırakma -> 3.Fire -> 4.Aktar)",
            command=self.start_live_learning, bg=self.purple_color, fg="#282a36",
            font=("Segoe UI", 9, "bold"), relief="flat", padx=8, pady=5, cursor="hand2"
        )
        self.btn_learn.pack(side="left", fill="x", expand=True, padx=(0, 5))

        btn_reset = tk.Button(
            top_rec_frame, text="🔄 Sıfırla", command=self.cmd_reset_all,
            bg=self.danger_color, fg="white", font=("Segoe UI", 9, "bold"),
            relief="flat", padx=8, pady=5, cursor="hand2"
        )
        btn_reset.pack(side="right")

        # 3 Temel Koordinat Satırları + Tekil '📍 Al' Butonları
        grid_f = tk.Frame(coord_card, bg=self.card_bg)
        grid_f.pack(fill="x")

        # 1. Dosya Tutma Yeri (Soldaki Gezgin 1. Sıra Yeşil Excel Simgesi)
        tk.Label(grid_f, text="1. Dosya (Yeşil Simge):", font=("Segoe UI", 8), fg=self.fg_color, bg=self.card_bg).grid(row=0, column=0, sticky="w", pady=2)
        self.entry_fx = tk.Entry(grid_f, width=5, font=("Consolas", 8), justify="center")
        self.entry_fx.insert(0, str(self.engine.coord_file_drag[0]))
        self.entry_fx.grid(row=0, column=1, padx=2)
        self.entry_fy = tk.Entry(grid_f, width=5, font=("Consolas", 8), justify="center")
        self.entry_fy.insert(0, str(self.engine.coord_file_drag[1]))
        self.entry_fy.grid(row=0, column=2, padx=2)
        tk.Button(grid_f, text="📍 Al", command=lambda: self.pick_single_coord("file"), bg="#44475a", fg="white", font=("Segoe UI", 7), relief="flat", padx=4).grid(row=0, column=3, padx=2)

        # 2. Netupp Bırakma Yeri (Sağdaki 'Sürükleyip Bırakın' alanı)
        tk.Label(grid_f, text="2. Bırakma Yeri (Sağ Netupp):", font=("Segoe UI", 8), fg=self.fg_color, bg=self.card_bg).grid(row=0, column=4, sticky="w", padx=(10, 0), pady=2)
        self.entry_dx = tk.Entry(grid_f, width=5, font=("Consolas", 8), justify="center")
        self.entry_dx.insert(0, str(self.engine.coord_netupp_drop[0]))
        self.entry_dx.grid(row=0, column=5, padx=2)
        self.entry_dy = tk.Entry(grid_f, width=5, font=("Consolas", 8), justify="center")
        self.entry_dy.insert(0, str(self.engine.coord_netupp_drop[1]))
        self.entry_dy.grid(row=0, column=6, padx=2)
        tk.Button(grid_f, text="📍 Al", command=lambda: self.pick_single_coord("drop"), bg="#44475a", fg="white", font=("Segoe UI", 7), relief="flat", padx=4).grid(row=0, column=7, padx=2)

        # 3. 'Tiger'a Aktar' Buton Yeri
        tk.Label(grid_f, text="3. 'Tiger'a Aktar' Butonu:", font=("Segoe UI", 8, "bold"), fg=self.accent_color, bg=self.card_bg).grid(row=1, column=0, sticky="w", pady=4)
        self.entry_tx = tk.Entry(grid_f, width=5, font=("Consolas", 8), justify="center")
        self.entry_tx.insert(0, str(self.engine.coord_btn_transfer[0]))
        self.entry_tx.grid(row=1, column=1, padx=2)
        self.entry_ty = tk.Entry(grid_f, width=5, font=("Consolas", 8), justify="center")
        self.entry_ty.insert(0, str(self.engine.coord_btn_transfer[1]))
        self.entry_ty.grid(row=1, column=2, padx=2)
        tk.Button(grid_f, text="📍 Al", command=lambda: self.pick_single_coord("transfer"), bg="#44475a", fg="white", font=("Segoe UI", 7), relief="flat", padx=4).grid(row=1, column=3, padx=2)

        # İki Aşamalı SPACE Onay Satırı
        r_confirm = tk.Frame(coord_card, bg="#1a1b26", padx=8, pady=5, highlightbackground="#50fa7b", highlightthickness=1)
        r_confirm.pack(fill="x", pady=(4, 2))

        rc_top = tk.Frame(r_confirm, bg="#1a1b26")
        rc_top.pack(fill="x")

        self.var_auto_space = tk.BooleanVar(value=self.engine.auto_space_confirm)
        tk.Checkbutton(
            rc_top, text="⚡ 2 Aşamalı Otomatik SPACE Onayı (1. Yes ➔ 2. Başarılı OK)",
            variable=self.var_auto_space, bg="#1a1b26", fg=self.accent_color,
            selectcolor="#282a36", activebackground="#1a1b26", font=("Segoe UI", 8, "bold")
        ).pack(side="left")

        btn_test_space = tk.Button(
            rc_top, text="⌨️ Şimdi SPACE Bas Testi", command=self.cmd_test_space_now,
            bg="#bd93f9", fg="#282a36", font=("Segoe UI", 8, "bold"), relief="flat", padx=6, pady=1, cursor="hand2"
        )
        btn_test_space.pack(side="right")

        rc_bottom = tk.Frame(r_confirm, bg="#1a1b26")
        rc_bottom.pack(fill="x", pady=(3, 0))

        tk.Label(rc_bottom, text="⏳ İki Space Arası Tiger Bekleme Süresi:", font=("Segoe UI", 8, "bold"), fg="#c0caf5", bg="#1a1b26").pack(side="left")
        self.entry_space_delay = tk.Entry(rc_bottom, width=5, font=("Consolas", 8, "bold"), fg="#50fa7b", bg="#282a36", justify="center")
        self.entry_space_delay.insert(0, str(self.engine.tiger_wait_between_spaces))
        self.entry_space_delay.pack(side="left", padx=4)
        tk.Label(rc_bottom, text="sn (Tiger reçeteyi işlerken beklenir, 'Başarılı' gelince SPACE basılır)", font=("Segoe UI", 8, "italic"), fg="#a0a0c0", bg="#1a1b26").pack(side="left")

        # DURUM PANELİ
        status_card = tk.Frame(self.root, bg=self.card_bg, padx=15, pady=8)
        status_card.pack(fill="x", padx=10, pady=4)

        f_stat = tk.Frame(status_card, bg=self.card_bg)
        f_stat.pack(fill="x")
        tk.Label(f_stat, text="DURUM:", font=("Segoe UI", 9, "bold"), fg=self.fg_color, bg=self.card_bg).pack(side="left")
        self.lbl_status = tk.Label(f_stat, text="HAZIR (Bekliyor)", font=("Segoe UI", 9, "bold"), fg=self.warn_color, bg=self.card_bg)
        self.lbl_status.pack(side="left", padx=10)

        self.lbl_action = tk.Label(status_card, text="Klasörü seçip koordinatları kaydettikten sonra F6 ile başlatabilirsiniz.", font=("Segoe UI", 8, "italic"), fg="#c0caf5", bg=self.card_bg)
        self.lbl_action.pack(anchor="w", pady=(3, 0))

        # BUTONLAR
        btn_bar = tk.Frame(self.root, bg=self.bg_color, pady=4)
        btn_bar.pack(fill="x", padx=10)

        self.btn_start = tk.Button(
            btn_bar, text="▶ AKTARIMI BAŞLAT (F6)", command=self.cmd_start,
            bg="#28a745", fg="white", font=("Segoe UI", 9, "bold"),
            relief="flat", padx=14, pady=5, cursor="hand2"
        )
        self.btn_start.pack(side="left", padx=4)

        self.btn_stop = tk.Button(
            btn_bar, text="⏹ DURDUR (F7)", command=self.cmd_stop,
            bg="#dc3545", fg="white", font=("Segoe UI", 9, "bold"),
            relief="flat", padx=14, pady=5, cursor="hand2"
        )
        self.btn_stop.pack(side="left", padx=4)

        btn_save = tk.Button(
            btn_bar, text="💾 Ayarları Kaydet", command=self.cmd_save_settings,
            bg="#44475a", fg=self.fg_color, font=("Segoe UI", 8),
            relief="flat", padx=10, pady=5
        )
        btn_save.pack(side="right", padx=4)

        # LOG PENCERESİ
        log_frame = tk.Frame(self.root, bg=self.card_bg, padx=10, pady=6)
        log_frame.pack(fill="both", expand=True, padx=10, pady=(4, 8))

        tk.Label(log_frame, text="📋 AKTARIM GÜNLÜĞÜ (Log):", font=("Segoe UI", 8, "bold"), fg=self.cyan_color, bg=self.card_bg).pack(anchor="w")

        self.txt_log = tk.Text(log_frame, bg="#11111b", fg="#cdd6f4", font=("Consolas", 8), height=6, relief="flat", padx=5, pady=5)
        self.txt_log.pack(fill="both", expand=True, pady=(3, 0))

        self.append_log("Netupp Reçete Aktarım Sistemi v3.3 hazır (Tekil Kilitleme + Klasör İzolasyonu).")

    def append_log(self, text):
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.txt_log.insert("end", f"[{timestamp}] {text}\n")
        self.txt_log.see("end")

    def refresh_file_count(self):
        folder = self.entry_dir.get().strip()
        if folder and os.path.isdir(folder):
            self.engine.source_dir = folder
            pending = self.engine.get_pending_files()
            self.lbl_file_count.config(text=f"Kalan: {len(pending)} Excel Dosyası")
        else:
            self.lbl_file_count.config(text="Geçersiz Klasör")

    def cmd_browse_folder(self):
        path = filedialog.askdirectory(title="Reçete Excel Dosyalarının Bulunduğu Klasörü Seçin")
        if path:
            self.entry_dir.delete(0, "end")
            self.entry_dir.insert(0, path)
            self.refresh_file_count()
            self.cmd_save_settings(show_msg=False)

    def cmd_reset_all(self):
        self.learning_active = False
        self.entry_fx.delete(0, "end"); self.entry_fx.insert(0, "0")
        self.entry_fy.delete(0, "end"); self.entry_fy.insert(0, "0")
        self.entry_dx.delete(0, "end"); self.entry_dx.insert(0, "0")
        self.entry_dy.delete(0, "end"); self.entry_dy.insert(0, "0")
        self.entry_fire_x.delete(0, "end"); self.entry_fire_x.insert(0, "0")
        self.entry_fire_y.delete(0, "end"); self.entry_fire_y.insert(0, "0")
        self.entry_tx.delete(0, "end"); self.entry_tx.insert(0, "0")
        self.entry_ty.delete(0, "end"); self.entry_ty.insert(0, "0")
        self.btn_learn.config(text="🎯 TÜM NOKTALARI SIRAYLA ÖĞREN", bg=self.purple_color)
        self.append_log("🔄 Koordinatlar sıfırlandı.")
        try:
            kernel32.Beep(600, 150)
        except Exception:
            pass

    def cmd_test_fire_now(self):
        """Kullanıcının seçtiği Fire koordinatına gidip '5' yazmasını test eder."""
        try:
            fx = int(self.entry_fire_x.get())
            fy = int(self.entry_fire_y.get())
            val = self.entry_fire_val.get().strip() or "5"
        except Exception:
            messagebox.showerror("Hata", "Lütfen önce Fire koordinatını (X, Y) seçin veya girin!")
            return

        def _test_worker():
            self.append_log(f"⏳ Fire testi 2 saniye sonra başlıyor... Netupp tablosunu açık tutun.")
            for i in range(2, 0, -1):
                self.lbl_action.config(text=f"Fire Testi Başlıyor: {i} sn...")
                try:
                    kernel32.Beep(800, 100)
                except Exception:
                    pass
                time.sleep(1.0)

            self.lbl_action.config(text=f"Fire kutusuna ({fx}, {fy}) çift tıklanıyor ve '{val}' yazılıyor...")
            self.engine.write_fire_value(fx, fy, val)
            self.append_log(f"⚡ Fire Testi Tamamlandı! ({fx}, {fy}) -> '{val}' yazıldı ve Enter basıldı.")
            self.lbl_action.config(text=f"Fire Testi Başarılı! ({fx}, {fy}) -> {val}")
            try:
                kernel32.Beep(1400, 200)
            except Exception:
                pass

        threading.Thread(target=_test_worker, daemon=True).start()

    def cmd_test_space_now(self):
        """Diyalog açıkken SPACE basışını 2 sn sonra test eder."""
        def _test_space():
            for i in range(2, 0, -1):
                self.lbl_action.config(text=f"SPACE Testi: {i} sn sonra diyalog penceresine basılacak...")
                try:
                    kernel32.Beep(900, 100)
                except Exception:
                    pass
                time.sleep(1.0)

            self.engine.confirm_active_dialog()
            try:
                kernel32.Beep(1400, 200)
            except Exception:
                pass
            self.append_log("⚡ SPACE tuşu gönderildi!")
            self.lbl_action.config(text="SPACE tuşu gönderildi!")

        threading.Thread(target=_test_space, daemon=True).start()

    def pick_single_coord(self, target_type):
        """Kullanıcının tek bir noktayı tıklayarak güncellemesini sağlar."""
        def _pick_worker():
            VK_LBUTTON = 0x01
            prompt_name = {
                "file": "1. sıradaki Excel'in YEŞİL SİMGESİNE [X] (Sol Gezgin)",
                "drop": "Netupp'taki 'Sürükleyip Bırakın' alanına (Sağ Pencere)",
                "fire": "Mamülün yanındaki 1. satır Fire kutusuna",
                "transfer": "Sağ alttaki 'Tiger'a Aktar' butonuna"
            }.get(target_type, "alana")

            self.lbl_action.config(text=f"Lütfen fare ile {prompt_name} SOL TIKLAYIN...")
            try:
                kernel32.Beep(800, 100)
            except Exception:
                pass

            while user32.GetAsyncKeyState(VK_LBUTTON) & 0x8000:
                time.sleep(0.05)

            pt = None
            while True:
                if user32.GetAsyncKeyState(VK_LBUTTON) & 0x8000:
                    pt = self.engine.get_cursor_pos()
                    try:
                        kernel32.Beep(1300, 150)
                    except Exception:
                        pass
                    break
                time.sleep(0.02)

            if not pt:
                return

            if target_type == "file":
                self.engine.coord_file_drag = pt
                self.entry_fx.delete(0, "end"); self.entry_fx.insert(0, str(pt[0]))
                self.entry_fy.delete(0, "end"); self.entry_fy.insert(0, str(pt[1]))
            elif target_type == "drop":
                self.engine.coord_netupp_drop = pt
                self.entry_dx.delete(0, "end"); self.entry_dx.insert(0, str(pt[0]))
                self.entry_dy.delete(0, "end"); self.entry_dy.insert(0, str(pt[1]))
            elif target_type == "fire":
                self.engine.coord_fire_field = pt
                self.entry_fire_x.delete(0, "end"); self.entry_fire_x.insert(0, str(pt[0]))
                self.entry_fire_y.delete(0, "end"); self.entry_fire_y.insert(0, str(pt[1]))
            elif target_type == "transfer":
                self.engine.coord_btn_transfer = pt
                self.entry_tx.delete(0, "end"); self.entry_tx.insert(0, str(pt[0]))
                self.entry_ty.delete(0, "end"); self.entry_ty.insert(0, str(pt[1]))

            self.append_log(f"✅ {prompt_name} seçildi: {pt}")
            self.lbl_action.config(text=f"Nokta kaydedildi: {pt}")
            self.cmd_save_settings(show_msg=False)

        threading.Thread(target=_pick_worker, daemon=True).start()

    def start_live_learning(self):
        """Kullanıcının kritik 4 noktayı sırayla tıklayarak kaydetmesini sağlar."""
        self.learning_active = True

        def _learn():
            VK_LBUTTON = 0x01
            use_fire = bool(self.var_use_fire.get())
            total_steps = 4 if use_fire else 3
            self.append_log(f"🎯 CANLI ÖĞRENME BAŞLADI ({total_steps} Adım):")
            
            # ADIM 1: Sol Gezgin Dosya Tutma Noktası
            self.lbl_action.config(text="1. ADIM: Soldaki Gezgin'de 1. sıradaki Excel'in YEŞİL SİMGESİNE [X] TIKLAYIN...")
            self.btn_learn.config(text="⏳ 1. Sol Dosya Simgesi bekleniyor...", bg="#ffb86c")
            
            while user32.GetAsyncKeyState(VK_LBUTTON) & 0x8000:
                time.sleep(0.05)

            p1 = None
            while self.learning_active:
                if user32.GetAsyncKeyState(VK_LBUTTON) & 0x8000:
                    p1 = self.engine.get_cursor_pos()
                    try:
                        kernel32.Beep(900, 100)
                    except Exception:
                        pass
                    break
                time.sleep(0.02)

            if not self.learning_active or not p1:
                return

            self.engine.coord_file_drag = p1
            self.entry_fx.delete(0, "end"); self.entry_fx.insert(0, str(p1[0]))
            self.entry_fy.delete(0, "end"); self.entry_fy.insert(0, str(p1[1]))
            self.append_log(f"✅ 1. Dosya Yeri Alındı: {p1}")

            while user32.GetAsyncKeyState(VK_LBUTTON) & 0x8000 and self.learning_active:
                time.sleep(0.05)
            time.sleep(0.3)

            # ADIM 2: Sağ Netupp Bırakma Noktası
            self.lbl_action.config(text="2. ADIM: Sağdaki Netupp penceresinde 'Sürükleyip Bırakın' alanına SOL TIKLAYIN...")
            self.btn_learn.config(text="⏳ 2. Bırakma Yeri bekleniyor...", bg="#ff79c6")

            p2 = None
            while self.learning_active:
                if user32.GetAsyncKeyState(VK_LBUTTON) & 0x8000:
                    p2 = self.engine.get_cursor_pos()
                    try:
                        kernel32.Beep(1100, 100)
                    except Exception:
                        pass
                    break
                time.sleep(0.02)

            if not self.learning_active or not p2:
                return

            self.engine.coord_netupp_drop = p2
            self.entry_dx.delete(0, "end"); self.entry_dx.insert(0, str(p2[0]))
            self.entry_dy.delete(0, "end"); self.entry_dy.insert(0, str(p2[1]))
            self.append_log(f"✅ 2. Bırakma Yeri Alındı: {p2}")

            while user32.GetAsyncKeyState(VK_LBUTTON) & 0x8000 and self.learning_active:
                time.sleep(0.05)
            time.sleep(0.3)

            # ADIM 3: Mamülün Yanındaki 1. Satır Fire Kutusu
            if use_fire:
                self.lbl_action.config(text="3. ADIM: Netupp tablosunda Mamülün yanındaki 1. Satır Fire kutusuna SOL TIKLAYIN...")
                self.btn_learn.config(text="⏳ 3. Mamül Fire Kutusu bekleniyor...", bg="#f1fa8c")

                p_fire = None
                while self.learning_active:
                    if user32.GetAsyncKeyState(VK_LBUTTON) & 0x8000:
                        p_fire = self.engine.get_cursor_pos()
                        try:
                            kernel32.Beep(1300, 120)
                        except Exception:
                            pass
                        break
                    time.sleep(0.02)

                if not self.learning_active or not p_fire:
                    return

                self.engine.coord_fire_field = p_fire
                self.entry_fire_x.delete(0, "end"); self.entry_fire_x.insert(0, str(p_fire[0]))
                self.entry_fire_y.delete(0, "end"); self.entry_fire_y.insert(0, str(p_fire[1]))
                self.append_log(f"✅ 3. Mamül Fire Kutusu Alındı: {p_fire}")

                while user32.GetAsyncKeyState(VK_LBUTTON) & 0x8000 and self.learning_active:
                    time.sleep(0.05)
                time.sleep(0.3)

            # SON ADIM: 'Tiger'a Aktar' Buton Noktası
            step_no = 4 if use_fire else 3
            self.lbl_action.config(text=f"{step_no}. ADIM: Sağ alttaki 'Tiger'a Aktar' butonuna SOL TIKLAYIN...")
            self.btn_learn.config(text=f"⏳ {step_no}. Aktar Butonu bekleniyor...", bg="#50fa7b")

            p_trans = None
            while self.learning_active:
                if user32.GetAsyncKeyState(VK_LBUTTON) & 0x8000:
                    p_trans = self.engine.get_cursor_pos()
                    try:
                        kernel32.Beep(1500, 150)
                        time.sleep(0.08)
                        kernel32.Beep(1900, 220)
                    except Exception:
                        pass
                    break
                time.sleep(0.02)

            if not self.learning_active or not p_trans:
                return

            self.learning_active = False
            self.engine.coord_btn_transfer = p_trans
            self.entry_tx.delete(0, "end"); self.entry_tx.insert(0, str(p_trans[0]))
            self.entry_ty.delete(0, "end"); self.entry_ty.insert(0, str(p_trans[1]))
            self.append_log(f"✅ 'Tiger'a Aktar' Butonu Alındı: {p_trans}")
            self.append_log("🎉 4 nokta da hafızaya alındı! F6 ile aktarımı başlatabilirsiniz.")
            self.lbl_action.config(text=f"Kayıt Tamam! Dosya: {p1} | Bırak: {p2} | Fire: {self.engine.coord_fire_field} | Aktar: {p_trans}")
            self.btn_learn.config(text="🎯 TÜM NOKTALARI SIRAYLA ÖĞREN (Hazır)", bg=self.purple_color)
            self.cmd_save_settings(show_msg=False)

        threading.Thread(target=_learn, daemon=True).start()

    def cmd_save_settings(self, show_msg=True):
        try:
            self.engine.source_dir = self.entry_dir.get().strip()
            self.engine.coord_file_drag    = (int(self.entry_fx.get()), int(self.entry_fy.get()))
            self.engine.coord_netupp_drop   = (int(self.entry_dx.get()), int(self.entry_dy.get()))
            self.engine.coord_fire_field   = (int(self.entry_fire_x.get()), int(self.entry_fire_y.get()))
            self.engine.coord_btn_transfer = (int(self.entry_tx.get()), int(self.entry_ty.get()))
            
            self.engine.use_fire_step      = bool(self.var_use_fire.get())
            self.engine.fire_value         = str(self.entry_fire_val.get().strip() or "5")
            self.engine.auto_space_confirm = bool(self.var_auto_space.get())
            try:
                self.engine.tiger_wait_between_spaces = float(self.entry_space_delay.get().strip() or "6.0")
            except Exception:
                pass
            
            self.engine.save_config()
            self.refresh_file_count()
            self.append_log("Ayarlar başarıyla kaydedildi.")
            if show_msg:
                messagebox.showinfo("Başarılı", "Tüm ayarlar ve koordinatlar kaydedildi.")
        except Exception as e:
            if show_msg:
                messagebox.showerror("Hata", f"Geçersiz ayar: {e}")

    def cmd_start(self):
        self.cmd_save_settings(show_msg=False)
        self.engine.start()

    def cmd_stop(self):
        self.engine.stop()

    def on_engine_event(self, event_type, *args):
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
            cycle_no, count, fname = args[0], args[1], args[2]
            self.lbl_file_count.config(text=f"Kalan: {count} Dosya")
            self.lbl_action.config(text=f"Tur #{cycle_no}: '{fname}' işleniyor...")

        elif event_type == "ACTION":
            tur, desc = args[0], args[1]
            self.lbl_action.config(text=f"[{tur}] {desc}")

        elif event_type == "LOG":
            tur, msg = args[0], args[1]
            self.append_log(msg)
            self.refresh_file_count()

        elif event_type == "ERROR":
            title, msg = args[0], args[1]
            messagebox.showerror(title, msg)

    def _hotkey_listener(self):
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
                if is_pressed(VK_F6):
                    self.root.after(0, self.cmd_start)
                if is_pressed(VK_F7):
                    self.root.after(0, self.cmd_stop)
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
    app = NetuppMacroApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
