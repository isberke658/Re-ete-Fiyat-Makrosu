# 🛠️ Geliştirici ve Mimari Kılavuzu (Developer Handbook)

Bu belge, **Tiger ERP Reçete Maliyeti Makrosu (v3.1)** projesini sonradan açıp kodları incelemek, değiştirmek, yeni özellikler eklemek veya yeniden derlemek istediğinizde size rehberlik etmek için hazırlanmıştır.

---

## 🏛️ Mimari Felsefe ve Tasarım Kararları

### Neden Saf Windows API (`ctypes`)?
Pek çok Python otomasyon projesinde `pyautogui`, `pynput` veya `keyboard` gibi üçüncü taraf kütüphaneler kullanılır. Ancak bu kütüphaneler kurumsal ERP sistemlerinde iki büyük sorun yaratır:
1. **DPI ve Ölçekleme Kayması:** Windows'un %125 veya %150 ölçekleme yaptığı ekranlarda koordinatlar sapar. Bu projede `SetProcessDPIAware()` doğrudan Win32 çekirdeğinde çağrılarak piksel hassasiyeti sağlanmıştır.
2. **DLL ve Hook Çakışmaları:** Harici klavye/fare dinleyicileri güvenlik duvarlarına takılabilir. Kodlarımızda `ctypes.windll.user32` ve `ctypes.windll.kernel32` üzerinden doğrudan Windows işletim sistemi çağrıları (`keybd_event`, `mouse_event`, `SendMessageW`, `GlobalAlloc`, `OpenClipboard`) kullanılmıştır.

Sonuç: **Sıfır ek DLL gereksinimi, mikrosaniye hızında tepki ve her Windows sürümünde (Windows 10, 11, Server) %100 yerel uyumluluk.**

---

## 📂 Ana Kod Dosyası ve Görev Dağılımı

### `recete_fiyat_kopyalama_makro.py` (Ana Motor - v3.1)
* **Görevi:** Tiger ERP Reçete Maliyeti ekranındaki LookUpEdit mamül seçiciyi sırayla gezdirir, birim fiyatı cımbızlar ve anlık olarak Excel'e yazar.
* **Kritik Fonksiyonlar:**
  * `parse_tr_price(val)`: `1.250,50 ₺`, `1,250.50`, `450 TL` gibi tüm Türkçe para formatlarını güvenli `float` değerine çevirir.
  * `parse_grid_row_clip(text)`: DevExpress tablosundaki bir satır kopyalandığında gelen tab (`\t`) ayrılmış metni parçalar; `109...` ile başlayan mamül kodunu, mamül adını ve birim fiyatı bulur.
  * `ReceteFiyatEngine.run_cycle()`: Ana otomasyon döngüsüdür:
    1. Mamül seçiciye tıklar (`coord_mamul_box`).
    2. `VK_F4` tuşu basarak DevExpress listesini zorla aşağı açar (Drop-down).
    3. `VK_DOWN` basarak sıradaki mamüle geçer, `VK_RETURN` ile seçimi onaylar.
    4. ERP'nin reçeteyi hesaplaması için belirlenen bekleme süresini uygular.
    5. Fiyat hücresine tıklar / seçer, Panoya kopyalar (Ctrl+C).
    6. Panodaki veriyi ayrıştırır ve anında `Reçete_Maliyeti_Canlı_Liste.xlsx` dosyasına yazar (`wb.save`).

---

## ⚙️ Geliştirme Ortamı Kurulumu

Projeyi bilgisayarınızda açıp geliştirmek için:

1. **Depoyu Klonlayın:**
   ```bash
   git clone https://github.com/isberke658/Re-ete-Fiyat-Makrosu.git
   cd Re-ete-Fiyat-Makrosu
   ```

2. **Python Sanal Ortamı (Virtualenv) Oluşturun:**
   ```bash
   python -m venv venv
   .\venv\Scripts\activate
   ```

3. **Gerekli Kütüphaneleri Yükleyin:**
   ```bash
   pip install -r requirements.txt
   ```
   *(Sadece `openpyxl` ve derleme için `pyinstaller` gereklidir).*

4. **Programı Çalıştırın:**
   ```bash
   python recete_fiyat_kopyalama_makro.py
   ```

---

## 🔨 Yeniden Standalone EXE Derleme

Kodu değiştirdikten sonra son kullanıcıya dağıtılacak tekil `.exe` dosyasını derlemek için:

1. Kök dizindeki **`build_exe.cmd`** dosyasına çift tıklayın.
2. Derleme tamamlandığında `dist\Recete_Fiyat_Kopyalama_Makro.exe` dosyası üretilecektir.

---

## 🛡️ Hücre 38 Veri Güvenliği İlkesi

* Program, kopyalanan her yeni satırda Excel dosyasını kapatıp açmaya gerek kalmadan dosya tanıtıcısını güncel tutar ve diske kaydeder.
* Kullanıcı işlemi durdursa veya elektrik kesilse dahi son kopyalanan mamül hafızada değil doğrudan diskteki Excel dosyasında mühürlü kalır.
