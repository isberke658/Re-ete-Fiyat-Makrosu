# 🛠️ Geliştirici ve Mimari Kılavuzu (Developer Handbook)

Bu belge, **Tiger & Logo ERP Reçete Otomasyon Kiti** projesini sonradan açıp kodları incelemek, değiştirmek, yeni özellikler eklemek veya yeniden derlemek istediğinizde size rehberlik etmek için hazırlanmıştır.

---

## 🏛️ Mimari Felsefe ve Tasarım Kararları

### Neden Saf Windows API (`ctypes`)?
Pek çok Python otomasyon projesinde `pyautogui`, `pynput` veya `keyboard` gibi üçüncü taraf kütüphaneler kullanılır. Ancak bu kütüphaneler kurumsal ERP sistemlerinde iki büyük sorun yaratır:
1. **DPI ve Ölçekleme Kayması:** Windows'un %125 veya %150 ölçekleme yaptığı ekranlarda koordinatlar sapar. Bu projede `SetProcessDPIAware()` doğrudan Win32 çekirdeğinde çağrılarak piksel hassasiyeti sağlanmıştır.
2. **DLL ve Hook Çakışmaları:** Harici klavye/fare dinleyicileri güvenlik duvarlarına takılabilir. Kodlarımızda `ctypes.windll.user32` ve `ctypes.windll.kernel32` üzerinden doğrudan Windows işletim sistemi çağrıları (`keybd_event`, `mouse_event`, `SendMessageW`, `GlobalAlloc`, `OpenClipboard`) kullanılmıştır.

Sonuç: **Sıfır ek DLL gereksinimi, mikrosaniye hızında tepki ve her Windows sürümünde (Windows 10, 11, Server) %100 yerel uyumluluk.**

---

## 📂 Kod Dosyaları ve Görev Dağılımı

### 1. `recete_fiyat_kopyalama_makro.py` (Ana Motor - v3.1)
* **Satır Sayısı:** ~2.850 satır.
* **Görevi:** Tiger ERP Reçete Maliyeti ekranındaki LookUpEdit mamül seçiciyi sırayla gezdirir, birim fiyatı cımbızlar ve anlık olarak Excel'e yazar.
* **Kritik Fonksiyonlar:**
  * `parse_tr_price(val)`: `1.250,50 ₺`, `1,250.50`, `450 TL` gibi tüm Türkçe para formatlarını güvenli `float` değerine çevirir.
  * `parse_grid_row_clip(text)`: DevExpress tablosundaki bir satır kopyalandığında gelen tab (`\t`) ayrılmış metni parçalar; `109...` ile başlayan mamül kodunu, mamül adını ve birim fiyatı bulur.
  * `ReceteFiyatEngine.run_cycle()`: Ana otomasyon döngüsüdür:
    1. Mamül seçiciye tıklar (`coord_mamul_box`).
    2. `VK_F4` tuşu basarak DevExpress listesini zorla aşağı açar (Drop-down).
    3. `VK_DOWN` basarak sıradaki mamüle geçer, `VK_RETURN` ile seçimi onaylar.
    4. ERP'nin reçeteyi patlatıp maliyeti hesaplaması için belirlenen bekleme süresini (`wait_after_down_sec`) uygular.
    5. Fiyat satırına (`coord_price_cell`) çift tık / Ctrl+A + Ctrl+C yapar.
    6. Panoyu `GetClipboardData` ile okur ve doğrular.
    7. `update_excel_row()` ile Excel'e yazar ve `wb.save()` ile diske kaydeder.
  * `EMBEDDED_MAMUL_182`: 182 adet mamülün kod, model, ad ve renk bilgilerini içeren dahili veri tabanıdır. Ortamda `reçete sıra.xlsx` olmasa dahi program sıfırdan kendi Excel dosyasını üretir.

### 2. `logo_recete_makro.py` (Otomatik Reçete Çıkarma - v1.0)
* **Görevi:** Logo ERP üzerinde toplu reçete çıkarma penceresinde sağ tık menüsünü tetikler ve seri space basarak kuyruğu eritir.
* **Kritik Özellikler:**
  * **Sabit Koordinat Kilidi:** Fare nereye kayarsa kaysın her turda `SetCursorPos` ile kaydedilen koordinata döner.
  * **Akıllı Duruş Tespiti:** Ekranda işlem bittiğinde (belirlenen duruş süresi eşiğinde) döngüyü güvenle tamamlar.

### 3. `netupp_recete_aktarim.py` (Reçete Aktarım & Fire - v3.3)
* **Görevi:** Klasördeki reçete Excel dosyalarını sürükle-bırak (OLE Drag & Drop) ile Netupp penceresine taşır, donanım seviyesinde `Fire=5` yazar ve pencereleri otomatik onaylar.
* **Kritik Özellikler:**
  * **Klasör İzolasyonu:** Başarıyla aktarılan dosyayı Gezgin listesinde karmaşa yaratmaması için anında `_Aktarilanlar` klasörüne taşır.
  * **Tekil Dosya Kilitleme:** Dosyayı sürüklemeden önce tek tıkla seçer, mavi kutu (marquee) oluşumunu engeller.

---

## 🚀 Geliştirme Ortamı Kurulumu

Projeyi klonladıktan sonra bilgisayarınızda geliştirmeye başlamak için:

### 1. Sanal Ortam (Virtual Environment) Oluşturma
```powershell
# Proje dizinine girin
cd Tiger-ERP-Recete-Fiyat-Makrosu

# Sanal ortamı oluşturun
python -m venv venv

# Sanal ortamı aktif edin
.\venv\Scripts\activate
```

### 2. Bağımlılıkları Yükleme
```powershell
pip install -r requirements.txt
```
*(Gereksinimler yalnızca `openpyxl` ve `pyinstaller`'dır; diğer her şey Python standart kütüphanesidir).*

### 3. Python Betiğini Çalıştırma
```powershell
python recete_fiyat_kopyalama_makro.py
```

---

## 🔨 Kodu Değiştirip Tekrar EXE Derleme (Build)

Kodda herhangi bir değişiklik yaptıktan sonra bağımsız `.exe` dosyasını üretmek için iki seçeneğiniz vardır:

### Kolay Yöntem (Tek Tık):
Proje kök dizinindeki `build_exe.cmd` (veya tüm araçlar için `build_all_tools.cmd`) dosyasına çift tıklayın.

### Manuel Komut Satırı:
```powershell
pyinstaller --noconfirm Recete_Fiyat_Kopyalama_Makro.spec
```
Derlenen çalıştırılabilir dosya `dist\Recete_Fiyat_Kopyalama_Makro.exe` konumunda hazır olacaktır.

### Spec Dosyasındaki Kritik Ayarlar (`Recete_Fiyat_Kopyalama_Makro.spec`):
* `console=False`: Arka planda siyah konsol penceresi açılmaz; yalnızca şık Tkinter GUI görünür.
* `upx=True`: Binary boyutunu küçültmek için UPX sıkıştırması aktiftir.
* `hiddenimports=['openpyxl']`: PyInstaller'ın dinamik içe aktarılan openpyxl paketini kaçırmaması için eklenmiştir.

---

## 🔍 Olası Sorunlar ve Kod Seviyesinde Müdahale Noktaları

| Durum | Sebebi | Kodda Müdahale Edilecek Yer |
| :--- | :--- | :--- |
| **Mamül listesi aşağı inmiyor** | ERP'nin LookUpEdit kontrolü kapalı kalmıştır. | `recete_fiyat_kopyalama_makro.py` içindeki `press_key(VK_F4)` bekleme süresini artırın (`time.sleep(0.15)`). |
| **Fiyat alanı yanlış kopyalanıyor** | Ekran çözünürlüğü veya ERP pencere konumu değişmiştir. | Program arayüzünde `F8` ve `F9` ile yeni koordinatı kaydedin veya `self.coord_price_cell` varsayılanını güncelleyin. |
| **Excel kayıt hatası (Permission Denied)** | Hedef Excel dosyası kullanıcının Excel uygulamasında açıktır. | `wb.save()` bloğu zaten `try-except` ile korunur ve `_GUNCEL.xlsx` oluşturur; gerekirse dosya adı mantığını `save_live_excel()` fonksiyonunda özelleştirin. |
| **Kopyalama çok hızlı geliyor, fiyat boş kalıyor** | ERP reçeteyi hesaplarken geç yanıt veriyordur. | Arayüzdeki **"Bekleme (sn)"** kutucuğunu `1.0` yerine `1.5` veya `2.0` yapın (kodda `self.wait_after_down_sec`). |

---

## 📜 Lisans & Katkı
Bu proje MIT lisansı altındadır. Dilediğiniz gibi çatallayabilir (fork), ERP'nizin ihtiyaçlarına göre uyarlayabilir ve dağıtabilirsiniz.
