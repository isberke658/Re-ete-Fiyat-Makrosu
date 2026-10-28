# ⚡ Tiger ERP Reçete Maliyeti & Fiyat Aktarım Makrosu (DevExpress Automation)

<div align="center">

![Python](https://img.shields.io/badge/Python-3.10%20%7C%203.11%20%7C%203.12-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Windows](https://img.shields.io/badge/Windows-10%20%7C%2011%20%7C%20Server-0078D6?style=for-the-badge&logo=windows&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)
![Build](https://img.shields.io/badge/Build-Standalone%20EXE-blueviolet?style=for-the-badge)
![Architecture](https://img.shields.io/badge/Architecture-HÜCRE%2038%20Forensic%20Audit-ff5722?style=for-the-badge)

<p align="center">
  <b>Tiger / Logo ERP & DevExpress LookUpEdit Reçete Maliyet Ekranları için Sıfır Kayma Garantili Akıllı Masaüstü Veri Yakalama ve Excel Aktarım Otomasyonu</b>
</p>

</div>

---

## 📌 Proje Genel Bakışı (Nedir ve Ne İşe Yarar?)

ERP (Logo Tiger, Netsis vb.) ve DevExpress altyapısına sahip üretim modüllerinde **"Reçete Maliyeti"** pencereleri yüzlerce yarı mamül ve mamülün hesaplanmış birim fiyatlarını barındırır. Standart ERP arayüzleri bu verileri toplu dışa aktarmaya (export) izin vermediğinde veya LookUpEdit açılır kutularında gezinme zorluğu yaşandığında, personelin yüzlerce ürünü tek tek açıp Excel'e kopyalaması saatler sürer ve insan hatasına yol açar.

Bu uygulama;
1. **Donanım Düzeyinde API Keşfi:** DevExpress `LookUpEdit` kontrollerine `F4` ile açılır liste fırlatarak sıradaki mamüle otomatik geçer.
2. **Akıllı Satır (TSV) Kopyalama:** Mamül tablosundaki ilk satırı ("Kodun Altı") panoya alıp `Mamül Kodu`, `Mamül Adı` ve hesaplanan `Birim Fiyatı` adli hassasiyetle ayrıştırır.
3. **Sıfır Kayma ve Dinamik Senkronizasyon:** ERP'de olmayan reçeteleri veya atlanan kodları tespit eder, fiyatların alt alta kaymasını kesin olarak engeller.
4. **Anında Canlı Excel Kaydı:** Tüm liste bitmeden, çekilen her ürünü o saniye diske (`Reçete_Maliyeti_Canlı_Liste.xlsx`) formatlı olarak yazar.

---

## 🖼️ Ekran Görüntüleri

### Otomasyon Kontrol Paneli & Canlı Terminal
<p align="center">
  <img src="assets/macro_ui_preview.jpg" alt="Makro Arayüzü" width="850">
</p>

---

## ✨ Öne Çıkan Özellikler

* 🎯 **Tek Tıkla LookUpEdit F4 Geçişi:** DevExpress kontrollerinde kapalıyken Aşağı Ok almayan listelere önce `VK_F4` göndererek tek hamlede sıradaki ürüne iner.
* 🛡️ **Hücre 38 Sıfır Halüsinasyon / Sıfır Kayma Algoritması:** Eski statik Excel listelerine bağımlı kalmadan, ekrandan gerçekte ne kopyalandıysa onu alır.
* 💾 **Kademeli Anında Kayıt (`wb.save`):** Elektrik kesintisi veya durdurma durumunda veri kaybı 0'dır; çekilen 3 mamül dahi anında Excel'e işlenir.
* 🔄 **Başa Sar (Excel'i Sıfırla) Kalkanı:** Butona basıldığında Excel dosyasını tamamen temizler, sayacı 0'a çekerek sıfırdan 1. sıradan başlamaya hazır eder.
* 📊 **Tek Tıkla Canlı Excel'i Açma:** Uygulama üzerinden Excel dosyasını doğrudan Microsoft Excel'de açıp anlık denetim imkanı.
* 🚀 **Saf CTypes Windows API:** Dış sistem otomasyon kütüphanelerine (pyautogui vb.) ihtiyaç duymadan saf Windows 32/64 API mesajlarıyla ultra hızlı ve hafif çalışma.

---

## ⌨️ Global Kısayol Tuşları

| Kısayol | Fonksiyon | Açıklama |
| :---: | :--- | :--- |
| **`F6`** | **TAM OTOMATİK BAŞLAT** | Belirlenen tekrar sayısı kadar (veya sınırsız) ardışık mamülleri çeker. |
| **`F7`** | **DURAKLAT / DEVAM ET** | İşlemi geçici olarak askıya alır, tekrar basıldığında kaldığı yerden devam eder. |
| **`F8`** | **SATIR KOORDİNATI SEÇ** | Fareyi tablodaki mamül satırına götürüp koordinatı (F8) canlı kaydeder. |
| **`F9`** | **MAMÜL KUTUSU SEÇ** | Fareyi üstteki açılır mamül kutusuna götürüp koordinatı (F9) canlı kaydeder. |
| **`F10`** | **TEK KAYIT AL** | Yalnızca ekrandaki mevcut mamülü kopyalayıp Excel'e ekler. |
| **`ESC`** | **ACİL DURDURMA** | Makroyu derhal güvenli şekilde durdurur. |

---

## 🏗️ Mimari & Çalışma Akışı

```mermaid
graph TD
    A[Kullanıcı F6 veya F10 Basar] --> B[Mamül Satırına Tıkla]
    B --> C[Panoya Kopyala Ctrl+C]
    C --> D[TSV Ayrıştır: Kod, Ad, Br Fiyat]
    D --> E{Fiyat Geçerli mi?}
    E -->|Evet| F[Excel'e Anında Satır Ekle & Formatla]
    F --> G[Ekrandaki LookUpEdit Kutusuna Tıkla]
    G --> H[F4 Bas - Listeyi Aç]
    H --> I[Aşağı Ok Bas - Sonrakine İn]
    I --> J[Enter Bas - Seçimi Onayla]
    J --> K[Reçete Hesaplama Beklemesi: 1.0 sn]
    K --> A
    E -->|Hayır| L[Hata Logla & İstisna Sepetine Al]
```

---

## 🚀 Kurulum & Çalıştırma

### 1. Yöntem: Bağımsız Taşınabilir EXE (Python Gerektirmez)
1. [Releases](../../releases) sekmesinden `Recete_Fiyat_Kopyalama_Makro.exe` dosyasını indirin.
2. Masaüstünüze koyup çift tıklayarak doğrudan çalıştırın.

### 2. Yöntem: Kaynak Koddan Çalıştırma
```bash
# Repoyu klonlayın
git clone https://github.com/isberke658/Re-ete-Fiyat-Makrosu.git
cd Tiger-ERP-Recete-Fiyat-Makrosu

# Gereksinimleri yükleyin
pip install -r requirements.txt

# Uygulamayı başlatın
python recete_fiyat_kopyalama_makro.py
```

---

## 📦 Tek Tıkla Kendi EXE'nizi Derleyin

Kendi ortamınızda sıfırdan bağımsız `.exe` üretmek için:
```bash
build_exe.cmd
```
veya terminalden:
```bash
pyinstaller --noconfirm Recete_Fiyat_Kopyalama_Makro.spec
```

---

## 📄 Lisans

Bu proje [MIT Lisansı](LICENSE) kapsamında açık kaynak olarak lisanslanmıştır. Ticari ve kişisel projelerde serbestçe kullanılabilir, geliştirilebilir.
