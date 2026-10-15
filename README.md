# Control User - GxP Compliance & User Access Audit Automation Tool

> Endüstriyel laboratuvar ve üretim cihazlarının kullanıcı yetki denetimlerini otomatik, güvenli ve izlenebilir bir şekilde gerçekleştiren PyQt6 tabanlı masaüstü otomasyon aracı.

##  Proje Hakkında
Farmasötik ve regüle endüstrilerde **Data Integrity (Veri Bütünlüğü)** ve **GxP uyumluluğu** en üst düzeyde denetlenen alanlardır. Geliştirdiğim bu araç; onaylı kullanıcı erişim matrislerini (`.xlsx`, `.docx`) heterojen sistemlerin ekran görüntüleriyle otomatik olarak karşılaştırır. Yetkisiz erişimleri, rol uyuşmazlıklarını ve yetim hesapları tespit ederek denetim süreçlerini optimize eder.

##  Öne Çıkan Özellikler
- **Modüler Cihaz Mimarisi:** Farklı üreticilerin arayüzleri (DataGuard, FT-IR, bioMérieux, Teraziler) için özel alt modüller.
- **Güvenli Yerel Mimari:** **Tamamen yerel (offline) çalışan yapısı sayesinde hiçbir hassas veri dış ağlara aktarılmaz; maksimum bilgi güvenliği sağlar.**
- **Akıllı Piksel & Koordinat Analizi:** Tesseract OCR'ın yanı sıra görsel toggle butonlarının ve kilit/yasak ikonlarının renk matrislerini tarayarak yanılma payını sıfıra indirir.
- **Otomatik Raporlama & Audit Trail:** Denetim sonuçlarını şeffaf bir şekilde loglar ve `.txt` formatında resmi rapor olarak dışa aktarır.

##  Kullanılan Teknolojiler
- **Python 3.10+**
- **PyQt6** (Modern ve Kurumsal Masaüstü Arayüzü)
- **Pytesseract & Pillow** (Optik Karakter Tanıma ve Görüntü İşleme)
- **Pandas & Python-Docx** (Excel ve Word veri işleme)
