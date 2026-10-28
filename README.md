# 🏭 Tiger & Logo ERP Reçete Maliyeti ve Otomasyon Kiti (v3.1)

[![Python Version](https://img.shields.io/badge/python-3.10%20%7C%203.11%20%7C%203.12-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows%2010%20%7C%2011%20%7C%20Server-lightgrey.svg)]()
[![Architecture](https://img.shields.io/badge/architecture-Pure%20Win32%20API%20(ctypes)-brightgreen.svg)]()

> **Kurumsal RPA ve Ürün Ağacı Otomasyonu:** Tiger ERP ve Logo ERP sistemlerindeki ürün reçeteleri, birim maliyetler ve aktarım süreçlerini sıfır insan hatasıyla saniyede 1 reçete hızında otomatize eden bağımsız masaüstü otomasyon araçları seti.

---

## ⚡ Hızlı İndirme (Kurulum Gerektirmez - Taşınabilir)

Kodlama veya Python kurulumu yapmadan doğrudan kullanmak için derlenmiş hazır paketleri indirebilirsiniz:

| İndirme Bağlantısı | Tür | Boyut | Açıklama |
| :--- | :---: | :---: | :--- |
| 🚀 **[Tiger_ERP_Recete_Fiyat_Makrosu_v3.1_Portable.zip](releases/Tiger_ERP_Recete_Fiyat_Makrosu_v3.1_Portable.zip)** | `.zip` | ~33.5 MB | **(Önerilen)** Ana makro, Excel şablonu ve kılavuz tek arşivde. |
| 💻 **[Recete_Fiyat_Kopyalama_Makro.exe](releases/Recete_Fiyat_Kopyalama_Makro.exe)** | `.exe` | ~33.8 MB | Ana makronun tekil çalıştırılabilir dosyası. |
| 📂 **[Logo_Recete_Makro_v1.0_Portable.zip](releases/Logo_Recete_Makro_v1.0_Portable.zip)** | `.zip` | ~10.2 MB | Logo ERP otomatik reçete çıkarma ve akıllı duruş tespit aracı. |
| 🚚 **[Netupp_Recete_Aktarim_v3.3_Portable.zip](releases/Netupp_Recete_Aktarim_v3.3_Portable.zip)** | `.zip` | ~10.2 MB | Reçete sürükle-bırak aktarımı ve Fire(5) işleme aracı. |
| 📑 **[Reçete_Maliyeti_Canlı_Liste_Sablon.xlsx](releases/Reçete_Maliyeti_Canlı_Liste_Sablon.xlsx)** | `.xlsx` | ~9 KB | Canlı liste çıktı şablonu. |

---

## 🖼️ Ekran Görüntüsü

<p align="center">
  <img src="assets/macro_ui_preview.jpg" alt="Makro Arayüzü ve Canlı Terminal" width="850">
</p>

---

## 🎯 Pakete Dahil Olan 3 Büyük Araç

### 1. Reçete Fiyat Kopyalama Makrosu (`Recete_Fiyat_Kopyalama_Makro.exe` - v3.1)
* **DevExpress LookUpEdit F4 Çözümü:** ERP açılır kutularında kapalıyken Aşağı Ok almayan listelere önce `VK_F4` göndererek tek hamlede sıradaki ürüne iner.
* **Canlı Excel Akış Kaydı:** Kopyalanan her satırı anında `Reçete_Maliyeti_Canlı_Liste.xlsx` dosyasına yazar ve diske mühürler (`wb.save`). Olası elektrik kesintisi veya durdurmalarda veri kaybı **sıfırdır**.
* **Çift Mod:** Hem seri otomatik tarama (**F6**) hem de tek tek kontrollü adım (**F10**) destekler.
* **Başa Sar Kalkanı:** Yeni bir partiye başlarken Excel'i tek tıkla sıfırlayıp 1. sıraya çeker.

### 2. Logo Reçete Çıkarma Makrosu (`Logo_Recete_Makro.exe` - v1.0)
* **Sabit Koordinat Kilidi:** Fare nereye kayarsa kaysın her turda menü ve sağ tık noktasına otomatik geri döner.
* **Seri Space Pompası:** Reçete çıkarma diyalog pencerelerini insan hızının 10 katı hızla seri boşluk tuşu basarak eritir.
* **Akıllı Duruş Tespiti:** Ekrandaki hareket veya log üretimi durduğunda serinin bittiğini anlar ve güvenle durur.

### 3. Netupp Reçete Aktarım Makrosu (`Netupp_Recete_Aktarim.exe` - v3.3)
* **OLE Sürükle-Bırak:** Excel reçete dosyalarını Gezgin'den Netupp penceresine otomatik taşır.
* **Donanım Seviyesinde Fire=5:** Mamül satırında donanım seviyesinde `Fire=5` girer ve onaylar.
* **Klasör İzolasyonu:** Aktarılan dosyaları anında `_Aktarilanlar` klasörüne ayırarak listenin daima temiz kalmasını sağlar.

---

## ⌨️ Kısayol Tuşları (Global Hotkeys)

Makro arka plandayken veya simge durumundayken dahi klavye tuşlarıyla kontrol edilir:

| Kısayol | Fonksiyon | Açıklama |
| :---: | :--- | :--- |
| `F6` | **Başlat / Devam Et** | Belirlenen sıra ve adette otomatik döngüyü başlatır. |
| `F7` | **Durdur / Duraklat** | Makroyu güvenli şekilde duraklatır; mevcut satırı korur. |
| `F8` | **Fiyat Koordinatını Kaydet** | Fare imlecinin bulunduğu ekran noktasını "Fiyat Satırı" olarak kaydeder. |
| `F9` | **Mamul Kutu Koordinatını Kaydet** | Fare imlecinin bulunduğu ekran noktasını "LookUpEdit Mamul Kutusu" olarak kaydeder. |
| `F10` | **Tek Satır Kaydet (Manuel Mod)** | Sıradaki 1 adet reçeteyi çeker, Excel'e işler ve durur. |
| `ESC` | **Acil Durdurma** | Tüm iş parçacıklarını anında sonlandırır. |

---

## 📖 Kullanım Kılavuzu (4 Adımda Başlatma)

1. **ERP Penceresini Açın:** Tiger ERP Reçete Maliyeti hesaplama ekranını açın.
2. **Koordinatları Doğrulayın (Opsiyonel):** Program standart Full HD (1920x1080) koordinatlarıyla açılır. Eğer pencereniz farklı bir yerdeyse:
   * Farenizi ERP'deki Mamul Seçim kutusunun üzerine getirin ➔ **`F9`** basın.
   * Farenizi tablodaki 1. satır (fiyat hücresi) üzerine getirin ➔ **`F8`** basın.
3. **Başlatın:** **`F6`** tuşuna basın. Makro sırayla tüm reçeteleri gezip birim fiyatları çekecektir.
4. **Çıktıyı Alın:** Veriler Masaüstünüzdeki `Reçete_Maliyeti_Canlı_Liste.xlsx` dosyasında hazır olacaktır.

---

## 🛠️ Sonradan Proje Üzerinde Çalışmak İsterseniz (Geliştirici Rehberi)

Kodu değiştirmek, yeni özellikler eklemek veya kendi ERP sisteminize uyarlamak isterseniz:

### 1. Depoyu Klonlayın ve Sanal Ortamı Kurun
```bash
git clone https://github.com/isberke658/Re-ete-Fiyat-Makrosu.git
cd Re-ete-Fiyat-Makrosu

# Sanal ortam oluşturup aktif edin
python -m venv venv
.\venv\Scripts\activate

# Gerekli tek kütüphaneyi yükleyin
pip install -r requirements.txt
```

### 2. Kodu Doğrudan Çalıştırın
```bash
python recete_fiyat_kopyalama_makro.py
```

### 3. Değişiklik Sonrası Yeniden EXE Derleyin
```bash
# Sadece Ana Makro için:
build_exe.cmd

# Tüm araçları birden derlemek için:
build_all_tools.cmd
```
Derlenen dosyalar `dist\` klasöründe hazır olacaktır.

> 💡 **Ayrıntılı Kod Mimarisi ve Fonksiyon Detayları için [GELİŞTİRİCİ KILAVUZU](GELISTIRICI_KILAVUZU.md) belgesini inceleyebilirsiniz.**

---

## 📄 Lisans
Bu proje [MIT Lisansı](LICENSE) kapsamında açık kaynak olarak paylaşılmıştır.
