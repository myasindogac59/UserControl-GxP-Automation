import sys
import os
import re
import ctypes
import pandas as pd
import docx
import pytesseract
from pytesseract import Output
import difflib 
from PIL import Image, ImageEnhance, ImageFilter
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QFileDialog, 
                             QTabWidget, QTableWidget, QTableWidgetItem, QHeaderView, 
                             QGroupBox, QMessageBox, QComboBox, QTextEdit)
from PyQt6.QtGui import QFont, QColor, QIcon, QPixmap
from PyQt6.QtCore import Qt

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

if getattr(sys, 'frozen', False):
    uygulama_klasoru = os.path.dirname(sys.executable)
else:
    uygulama_klasoru = os.path.dirname(os.path.abspath(__file__))

tesseract_yolu = os.path.join(uygulama_klasoru, 'tesseract', 'tesseract.exe')
pytesseract.pytesseract.tesseract_cmd = tesseract_yolu

class DropZone(QLabel):
    def __init__(self, main_app):
        super().__init__("Ekran Görüntülerini Buraya Sürükleyip Bırakın\n(Veya Seçmek İçin Tıklayın)")
        self.main_app = main_app
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setAcceptDrops(True)
        self.setStyleSheet("""
            QLabel { border: 2px dashed #94A3B8; border-radius: 8px; background-color: #F1F5F9; color: #475569; font-weight: bold; padding: 18px; }
            QLabel:hover { background-color: #E2E8F0; border-color: #3B82F6; color: #1E293B; }
        """)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls(): event.accept()
        else: event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        file_paths = [u.toLocalFile() for u in urls if u.toLocalFile().lower().endswith(('.png', '.jpg', '.jpeg'))]
        if file_paths: self.main_app.process_images(file_paths)

    def mousePressEvent(self, event):
        self.main_app.select_images()

class YetkiDenetimApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Control User - Kullanıcı Denetim Aracı")
        if os.path.exists(resource_path("1.ico")):
            self.setWindowIcon(QIcon(resource_path("1.ico"))) 
        try:
            myappid = 'company.controluser.app.final'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except: pass

        self.resize(1240, 800)
        self.setStyleSheet("""
            QMainWindow { background-color: #F8FAFC; }
            QGroupBox { font-weight: bold; font-size: 14px; border: 2px solid #E2E8F0; border-radius: 8px; margin-top: 12px; background-color: #FFFFFF; }
            QGroupBox::title { subcontrol-origin: margin; left: 15px; padding: 0 5px; color: #1E293B; }
            QPushButton { background-color: #3B82F6; color: white; border: none; padding: 10px 15px; border-radius: 6px; font-weight: bold; }
            QPushButton:hover { background-color: #2563EB; }
            QPushButton#btn_karsilastir { background-color: #10B981; font-size: 16px; padding: 14px; }
            QPushButton#btn_karsilastir:hover { background-color: #059669; }
            QLabel { color: #334155; font-size: 13px; font-weight: bold; }
            QComboBox { padding: 8px; border: 2px solid #CBD5E1; border-radius: 6px; font-weight: bold; color: #0F172A; background: white; min-width: 340px; }
            QComboBox QAbstractItemView { background-color: #FFFFFF; color: #0F172A; selection-background-color: #3B82F6; selection-color: #FFFFFF; border: 1px solid #CBD5E1; }
            QTableWidget { border: 1px solid #E2E8F0; background-color: white; gridline-color: #E2E8F0; selection-color: black; }
            QHeaderView::section { background-color: #F8FAFC; padding: 8px; border: 1px solid #E2E8F0; font-weight: bold; color: #475569; }
            QTabWidget::pane { border: 1px solid #E2E8F0; border-radius: 4px; background: white; }
            QTabBar::tab { background: #F1F5F9; border: 1px solid #E2E8F0; padding: 10px 15px; margin-right: 2px; color: black; font-weight: bold;}
            QTabBar::tab:selected { background: white; border-bottom-color: white; color: #0F172A; border-top: 3px solid #3B82F6;}
        """)
        self.resmi_kullanicilar = {}
        self.ham_fotolar = [] 
        self.setup_ui()

    def setup_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        
        header_layout = QHBoxLayout()
        logo_label = QLabel()
        if os.path.exists(resource_path("1.png")):
            logo_pixmap = QPixmap(resource_path("1.png")).scaled(100, 100, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            logo_label.setPixmap(logo_pixmap)
        
        title_layout = QVBoxLayout()
        title_main = QLabel("CONTROL USER")
        title_main.setStyleSheet("font-size: 26px; font-weight: 900; color: #0F172A; margin-bottom: -5px;")
        title_sub = QLabel("Kullanıcı Denetim Aracı") 
        title_sub.setStyleSheet("font-size: 14px; font-weight: bold; color: #64748B;")
        
        title_layout.addWidget(title_main)
        title_layout.addWidget(title_sub)
        title_layout.setAlignment(Qt.AlignmentFlag.AlignVCenter)
        header_layout.addWidget(logo_label)
        header_layout.addSpacing(15)
        header_layout.addLayout(title_layout)
        header_layout.addStretch()
        main_layout.addLayout(header_layout)

        file_group = QGroupBox("1. Adım: Veri Kaynaklarını ve Cihaz Tipini Seçin")
        file_layout = QVBoxLayout()
        
        top_controls = QHBoxLayout()
        self.lbl_resmi = QLabel("Onaylı Liste (Excel/Word): Seçilmedi")
        btn_resmi = QPushButton("Liste Seç")
        btn_resmi.clicked.connect(self.select_resmi)
        
        lbl_cihaz = QLabel("Cihaz Arayüz Tipi:")
        self.combo_cihaz = QComboBox()
        self.combo_cihaz.addItems([
            "Otomatik Algıla (Tüm Cihazlar)",
            "Terazi / Tartım (Detay Ekranı - Evet/Hayır veya Toggle Buton)",
            "Tablo Sistemleri (DataGuard / Pyros eXpress / Metrohm)",
            "Spectroscopy / FT-IR (Sol Users ve Sağ Ağaç - Kilit/Yasak İkonlu)",
            "bioMérieux VITEK / MYLA (Web Arayüzü - Users / Groups)",
            "Sade Liste Sistemleri (ZEISS / TexturePro - Listedekiler Aktif)"
        ])
        
        top_controls.addWidget(self.lbl_resmi)
        top_controls.addWidget(btn_resmi)
        top_controls.addSpacing(30)
        top_controls.addWidget(lbl_cihaz)
        top_controls.addWidget(self.combo_cihaz)
        file_layout.addLayout(top_controls)

        self.drop_zone = DropZone(self)
        file_layout.addWidget(self.drop_zone)
        file_group.setLayout(file_layout)
        main_layout.addWidget(file_group)

        self.btn_karsilastir = QPushButton("DENETLE")
        self.btn_karsilastir.setObjectName("btn_karsilastir")
        self.btn_karsilastir.clicked.connect(self.run_comparison)
        main_layout.addWidget(self.btn_karsilastir)

        self.tabs = QTabWidget()
        headers = ["Kullanıcı Adı", "Ad Soyad", "Listede Beklenen", "Cihazdaki Durum", "Tespit Kanıtı (Fotoğraf ve Okunan Satır)"]
        self.tab_hata_pasif, self.table_hata_pasif = self.create_table_tab(headers)
        self.tab_hata_aktif, self.table_hata_aktif = self.create_table_tab(headers)
        self.tab_uyumlu, self.table_uyumlu = self.create_table_tab(headers)
        
        self.tab_log = QWidget()
        log_layout = QVBoxLayout(self.tab_log)
        self.txt_log = QTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setStyleSheet("font-family: Consolas, monospace; font-size: 12px; background: #0F172A; color: #E2E8F0; padding: 10px;")
        log_layout.addWidget(self.txt_log)

        self.tabs.addTab(self.tab_hata_pasif, "Kritik Hata: Cihazda Eksik veya Pasif (E Olmalı)")
        self.tabs.addTab(self.tab_hata_aktif, "Kritik Hata: Cihazda Aktif Kalanlar (H Olmalı)")
        self.tabs.addTab(self.tab_uyumlu, "Uyumlu (Sorunsuz)")
        self.tabs.addTab(self.tab_log, "OCR Analiz Kanıtları (Detaylı Kontrol)")
        main_layout.addWidget(self.tabs)

    def create_table_tab(self, headers):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        table = QTableWidget(0, len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        layout.addWidget(table)
        return tab, table

    def metin_temizle(self, metin):
        if not isinstance(metin, str): return ""
        ceviri = str.maketrans("ışğüöçİŞĞÜÖÇ", "isguocisguoc")
        temiz = metin.translate(ceviri).lower()
        temiz = re.sub(r'[^a-z0-9\s]', ' ', temiz)
        return " ".join(temiz.split())

    def select_resmi(self):
        fname, _ = QFileDialog.getOpenFileName(self, "Onaylı Liste Seç", "", "Excel/Word (*.xlsx *.xls *.docx)")
        if fname:
            self.lbl_resmi.setText(f"Liste: {os.path.basename(fname)}")
            self.resmi_dosya_oku(fname)

    def select_images(self):
        fnames, _ = QFileDialog.getOpenFileNames(self, "Ekran Görüntülerini Seç", "", "Resim (*.png *.jpg *.jpeg)")
        if fnames: self.process_images(fnames)

    def resmi_hazirla_ocr(self, pil_img, telefon_cekimi_mi=False):
        w, h = pil_img.size
        scale = 2.0 if w < 1400 else 1.4
        img_res = pil_img.resize((int(w * scale), int(h * scale)), Image.Resampling.LANCZOS)
        img_gray = img_res.convert('L')
        if telefon_cekimi_mi:
            img_gray = img_gray.filter(ImageFilter.GaussianBlur(radius=0.8))
            img_gray = ImageEnhance.Sharpness(img_gray).enhance(2.0)
        img_gray = ImageEnhance.Contrast(img_gray).enhance(2.2)
        return img_gray, scale

    def process_images(self, file_paths):
        toplam = len(file_paths)
        if toplam == 0: return

        self.ham_fotolar = [] 
        self.txt_log.clear()
        self.drop_zone.setText(f"0 / {toplam} Fotoğraf Okunuyor...")
        self.drop_zone.setStyleSheet("QLabel { border: 2px solid #F59E0B; background-color: #FEF3C7; color: #B45309; font-weight: bold; padding: 18px; border-radius: 8px; }")
        
        for idx, path in enumerate(file_paths):
            try:
                QApplication.processEvents() 
                dosya_adi = os.path.basename(path)
                img_orig = Image.open(path).convert('RGB')
                w, h = img_orig.size
                
                telefon_mu = (h > w) or (w > 2200) or dosya_adi.lower().endswith(('.jpg', '.jpeg'))
                img_gray, scale = self.resmi_hazirla_ocr(img_orig, telefon_cekimi_mi=telefon_mu)
                
                ocr_data = pytesseract.image_to_data(img_gray, lang='tur+eng', config='--psm 6', output_type=Output.DICT)
                raw_lines_psm6 = pytesseract.image_to_string(img_gray, lang='tur+eng', config='--psm 6').split('\n')
                
                kelimeler = []
                for i in range(len(ocr_data['text'])):
                    txt = ocr_data['text'][i].strip()
                    if len(txt) > 0 and int(ocr_data['conf'][i]) > 10:
                        y_coord = int(ocr_data['top'][i] / scale)
                        x_coord = int(ocr_data['left'][i] / scale)
                        if y_coord > int(h * 0.93): continue
                        
                        cl = self.metin_temizle(txt)
                        if cl:
                            kelimeler.append({
                                'raw': txt,
                                'clean': cl,
                                'x': x_coord,
                                'y': y_coord,
                                'w': int(ocr_data['width'][i] / scale),
                                'h': int(ocr_data['height'][i] / scale)
                            })
                
                satirlar = self.kelimeleri_satirlara_grupla(kelimeler, h)

                for r_line in raw_lines_psm6:
                    cl_line = self.metin_temizle(r_line)
                    if cl_line and "login " not in cl_line and not any(cl_line == s['metin'] for s in satirlar):
                        satirlar.append({
                            'y_ort': h // 2,
                            'kelimeler': [{'clean': w_c, 'raw': w_c, 'x': int(w*0.3), 'y': h//2, 'w': 10, 'h': 10} for w_c in cl_line.split()],
                            'metin': cl_line,
                            'raw_metin': r_line.strip(),
                            'sanal_satir': True
                        })

                tum_metin = " ".join(s['metin'] for s in satirlar)

                self.ham_fotolar.append({
                    'dosya': dosya_adi,
                    'img': img_orig,
                    'kelimeler': kelimeler,
                    'satirlar': satirlar,
                    'tum_metin': tum_metin
                })
                
                self.drop_zone.setText(f"{idx + 1} / {toplam} Fotoğraf Okundu...")
            except Exception as e:
                QMessageBox.critical(self, "Okuma Hatası", f"{path} okunurken hata:\n{str(e)}")
                return 

        self.drop_zone.setText(f"{toplam} Fotoğraf Başarıyla Hafızaya Alındı! 'DENETLE' butonuna basabilirsiniz.")
        self.drop_zone.setStyleSheet("QLabel { border: 2px solid #10B981; background-color: #D1FAE5; color: #065F46; font-weight: bold; padding: 18px; border-radius: 8px; }")

    def kelimeleri_satirlara_grupla(self, kelimeler, ekran_yuksekligi):
        kelimeler_kopya = sorted(kelimeler, key=lambda k: k['y'])
        satirlar = []
        tolerans_y = max(10, int(ekran_yuksekligi * 0.014))
        for k in kelimeler_kopya:
            eklendi = False
            for satir in satirlar:
                if abs(satir['y_ort'] - k['y']) <= tolerans_y:
                    satir['kelimeler'].append(k)
                    satir['y_ort'] = sum(x['y'] for x in satir['kelimeler']) / len(satir['kelimeler'])
                    eklendi = True
                    break
            if not eklendi:
                satirlar.append({'y_ort': k['y'], 'kelimeler': [k], 'sanal_satir': False})
        
        for satir in satirlar:
            satir['kelimeler'].sort(key=lambda k: k['x'])
            satir['metin'] = " ".join(k['clean'] for k in satir['kelimeler'])
            satir['raw_metin'] = " ".join(k['raw'] for k in satir['kelimeler'])
        return satirlar

    def ekran_tipini_belirle(self, foto):
        secim = self.combo_cihaz.currentIndex()
        if secim == 1: return "TERAZI"
        elif secim == 2: return "TABLO"
        elif secim == 3: return "FTIR"
        elif secim == 4: return "BIOMERIEUX"
        elif secim == 5: return "LISTE"
        
        t = foto['tum_metin']
        if any(k in t for k in ["duzenlemeyi", "account active", "erisim seviyesi", "kart numarasi", "varsayilan profil", "oparetor"]):
            return "TERAZI"
        elif any(k in t for k in ["biomerieux", "skeeper", "vitek", "myla", "external users", "umadministrator", "group details", "user details"]):
            return "BIOMERIEUX"
        elif any(k in t for k in ["spectroscopy", "iqoq", "current organization", "bakim destek"]):
            return "FTIR"
        elif any(k in t for k in ["dataguard", "access level", "last access", "workgroup", "pyros", "user group members", "rename group", "disabled", "inactive", "enabled"]):
            return "TABLO"
        else:
            return "LISTE"

    def biomerieux_filtrele(self, foto):
        w, h = foto['img'].size
        t = foto['tum_metin']
        if "group details" in t or "add user" in t or "application name" in t:
            x_sinir = int(w * 0.78)
            for k in foto['kelimeler']:
                if k['clean'] in ["add", "external"] and k['x'] > int(w * 0.65):
                    x_sinir = min(x_sinir, k['x'] - 15)
            gecerli_kelimeler = [k for k in foto['kelimeler'] if k['x'] < x_sinir]
            return self.kelimeleri_satirlara_grupla(gecerli_kelimeler, h)
        return foto['satirlar']

    def ftir_satirlari_ayristir(self, foto):
        w, h = foto['img'].size
        sol_kelimeler = [k for k in foto['kelimeler'] if k['x'] < int(w * 0.35)]
        sag_kelimeler = [k for k in foto['kelimeler'] if k['x'] >= int(w * 0.35)]
        
        sol_satirlar = self.kelimeleri_satirlara_grupla(sol_kelimeler, h)
        sag_satirlar = self.kelimeleri_satirlara_grupla(sag_kelimeler, h)
        return sol_satirlar + sag_satirlar

    def ftir_kilit_kontrolu(self, img, satir):
        if not satir['kelimeler'] or satir.get('sanal_satir', False): return "E"
        ilk_k = satir['kelimeler'][0]
        w, h = img.size
        
        x_start = max(0, ilk_k['x'] - 42)
        x_end = max(0, ilk_k['x'] - 15)
        y_start = max(0, ilk_k['y'] - 3)
        y_end = min(h, ilk_k['y'] + ilk_k['h'] + 3)
        
        sari_piksel = 0
        kirmizi_piksel = 0
        for x in range(x_start, x_end):
            for y in range(y_start, y_end):
                r, g, b = img.getpixel((x, y))
                if r > 175 and g > 135 and b < 105:
                    sari_piksel += 1
                if r > 180 and g < 70 and b < 70:
                    kirmizi_piksel += 1
                    
        if sari_piksel >= 5 or kirmizi_piksel >= 4:
            return "H"
        return "E"

    def terazi_cerrahi_analiz(self, foto):
        img = foto['img']
        w, h = img.size
        aday_kelimeler = []
        durum = None
        kanit = ""

        terazi_yasakli = {
            "tartim", "precision", "max", "min", "duzenlemeyi", "kaydet", "isim", "kodlar", 
            "sifre", "name", "and", "surname", "erisim", "seviyesi", "yonetici", "operator", 
            "oparetor", "account", "active", "dil", "english", "use", "default", "language", 
            "varsayilan", "profil", "hicbiri", "kart", "numarasi", "unlock", "the", "user", 
            "face", "profile", "theme", "dark", "light", "evet", "hayir", "hayr", "yes", "no",
            "realvnc", "viewer", "pue10", "pue8", "pb", "pm", "am"
        }

        for s in foto['satirlar']:
            if s.get('sanal_satir', False): continue
            if s['y_ort'] < (h * 0.14) or s['y_ort'] > (h * 0.92): continue
            kelimeler = [k['clean'] for k in s['kelimeler']]
            
            for k in kelimeler:
                if k not in terazi_yasakli and len(k) >= 2 and not k.isdigit():
                    aday_kelimeler.append(k)
                    
            if any(neg in kelimeler for neg in ["hayir", "hayr", "no", "pasif"]):
                durum = "H"
                kanit = f"Metin: 'Hayır' ({foto['dosya']})"
            elif any(pos in kelimeler for pos in ["evet", "yes", "aktif"]):
                durum = "E"
                kanit = f"Metin: 'Evet' ({foto['dosya']})"

        sag_sutun_crop = img.crop((int(w * 0.35), int(h * 0.16), int(w * 0.95), int(h * 0.93)))
        crop_gray, _ = self.resmi_hazirla_ocr(sag_sutun_crop)
        crop_text = self.metin_temizle(pytesseract.image_to_string(crop_gray, lang='tur+eng', config='--psm 6'))
        
        for k in crop_text.split():
            if k not in terazi_yasakli and len(k) >= 2 and not k.isdigit():
                if k not in aday_kelimeler:
                    aday_kelimeler.append(k)
        
        if durum is None:
            crop_kelimeler = crop_text.split()
            if any(neg in crop_kelimeler for neg in ["hayir", "hayr", "no", "pasif"]):
                durum = "H"
                kanit = f"Bölgesel Okuma: 'Hayır' ({foto['dosya']})"
            elif any(pos in crop_kelimeler for pos in ["evet", "yes", "aktif"]):
                durum = "E"
                kanit = f"Bölgesel Okuma: 'Evet' ({foto['dosya']})"

        if durum is None:
            mavi_piksel = 0
            for x in range(int(w * 0.75), int(w * 0.96), 2):
                for y in range(int(h * 0.40), int(h * 0.60), 2):
                    r, g, b = img.getpixel((x, y))
                    if b > 170 and (b - r) > 45:
                        mavi_piksel += 1
            if mavi_piksel >= 12:
                durum = "E"
                kanit = f"Görsel Toggle: AÇIK/Mavi ({foto['dosya']})"
            else:
                durum = "H"
                kanit = f"Görsel Toggle: KAPALI/Gri ({foto['dosya']})"

        return aday_kelimeler, durum, kanit

    def satir_durumu_belirle_tablo(self, kelimeler_listesi, satir_metni):
        # 'Active,Locked' gibi durumlarda kilitli/pasif önceliklidir!
        pasif_kelimeler = {"disabled", "disable", "locked", "inactive", "inactiv", "pasif", "deaktif", "hayir", "hayr"}
        for k in kelimeler_listesi:
            if k in pasif_kelimeler:
                return "H"
            if len(k) >= 6 and any(difflib.SequenceMatcher(None, k, p).ratio() >= 0.80 for p in ["disabled", "inactive", "locked"]):
                return "H"
                
        if any(p in satir_metni for p in ["inactive", "disabled", "locked", "disab"]):
            return "H"
        return "E"

    def resmi_dosya_oku(self, dosya_yolu):
        self.resmi_kullanicilar = {}
        try:
            if dosya_yolu.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(dosya_yolu, sheet_name=None, dtype=str)
                for sheet_name, data in df.items():
                    for index, row in data.iterrows():
                        self.satir_ayikla(row)
            elif dosya_yolu.endswith('.docx'):
                doc = docx.Document(dosya_yolu)
                for table in doc.tables:
                    if len(table.rows) < 2: continue
                    headers = [self.metin_temizle(cell.text) for cell in table.rows[0].cells]
                    for row in table.rows[1:]:
                        row_data = [cell.text.strip() for cell in row.cells]
                        row_dict = dict(zip(headers, row_data))
                        self.satir_ayikla(pd.Series(row_dict))
        except Exception as e:
            QMessageBox.critical(self, "Dosya Okuma Hatası", str(e))

    def satir_ayikla(self, row):
        k_adi = ""
        ad_soyad = ""
        durum = None

        for col_name, val in row.items():
            if pd.isna(val): continue
            c_clean = self.metin_temizle(str(col_name))
            val_str = str(val).strip()
            
            if any(k in c_clean for k in ["kullanici adi", "user name", "userid", "user id"]):
                k_adi = val_str
            elif any(k in c_clean for k in ["ad soyad", "name surname", "adi soyadi", "first name", "full name"]):
                ad_soyad = val_str
            elif any(k in c_clean for k in ["aktif", "active", "durum", "status"]):
                v_up = val_str.upper()
                if any(x in v_up for x in ["H", "PAS", "DEAKT", "N"]): durum = "H"
                elif any(x in v_up for x in ["E", "AKT", "Y"]): durum = "E"

        if not k_adi:
            vals = [str(v).strip() for v in row.values if pd.notna(v)]
            if len(vals) >= 3 and vals[0].isdigit():
                ad_soyad = vals[1]
                k_adi = vals[2]

        if not k_adi or k_adi.lower() == "nan": return
        
        k_clean = self.metin_temizle(k_adi)
        yasakli = ["hazirlayan", "onaylayan", "kontrol", "gorevi", "title", "revizyon", "tarih", "imza", "ekipman", "yazilim", "sira no"]
        if any(y in k_clean for y in yasakli) or len(k_clean) < 2: return

        if not durum:
            row_full = " ".join(str(v).lower() for v in row.values if pd.notna(v))
            if any(x in row_full for x in ["h*", "deaktif", "pasif", "hayır", "hayir"]): durum = "H"
            elif any(x in row_full for x in ["aktif", "evet"]): durum = "E"
            else:
                for v in row.values:
                    if str(v).strip().upper() in ["E", "Y"]: durum = "E"; break
                    elif str(v).strip().upper() in ["H", "H*", "N"]: durum = "H"; break
        
        if not durum: durum = "E"
        self.resmi_kullanicilar[k_adi] = {"AdSoyad": ad_soyad, "Durum": durum}

    def eslesme_var_mi(self, k_clean, ad_soyad_clean, hedef_kelimeler, satir_metni, cihaz_tipi):
        if "description" in hedef_kelimeler or "removed" in hedef_kelimeler or "login" in hedef_kelimeler:
            if len(hedef_kelimeler) <= 3:
                return False

        for h_kelime in hedef_kelimeler:
            if k_clean == h_kelime:
                return True
            if len(k_clean) >= 5 and k_clean in h_kelime:
                if h_kelime.startswith(k_clean) or h_kelime.endswith(k_clean) or "isim" in h_kelime:
                    return True
            if len(k_clean) >= 5 and len(h_kelime) >= 5 and abs(len(k_clean) - len(h_kelime)) <= 1:
                if (k_clean[0] == h_kelime[0] and k_clean[-2:] == h_kelime[-2:]) or (k_clean[:3] == h_kelime[:3]):
                    if difflib.SequenceMatcher(None, k_clean, h_kelime).ratio() >= 0.80:
                        return True

        if ad_soyad_clean and len(ad_soyad_clean) >= 5 and cihaz_tipi in ["FTIR", "TABLO", "BIOMERIEUX", "LISTE"]:
            parcalar = [p for p in ad_soyad_clean.split() if len(p) >= 3]
            if len(parcalar) >= 2:
                ilk_ad = parcalar[0]
                soyad = parcalar[-1]
                if ilk_ad in satir_metni and soyad in satir_metni:
                    return True
                if len(parcalar) >= 3 and parcalar[0] in satir_metni and parcalar[1] in satir_metni:
                    return True
                if len(soyad) >= 5 and soyad in hedef_kelimeler:
                    for h_k in hedef_kelimeler:
                        if difflib.SequenceMatcher(None, ilk_ad, h_k).ratio() >= 0.75:
                            return True
        return False

    def run_comparison(self):
        if not self.resmi_kullanicilar: return QMessageBox.warning(self, "Uyarı", "Lütfen onaylı listeyi yükleyin.")
        if not self.ham_fotolar: return QMessageBox.warning(self, "Uyarı", "Lütfen cihaz fotoğraflarını okutun.")

        for table in [self.table_hata_pasif, self.table_hata_aktif, self.table_uyumlu]: table.setRowCount(0)
        self.txt_log.clear()

        cozumlenen_ekranlar = []
        for foto in self.ham_fotolar:
            tip = self.ekran_tipini_belirle(foto)
            if tip == "TERAZI":
                adaylar, durum, kanit = self.terazi_cerrahi_analiz(foto)
                cozumlenen_ekranlar.append({'tip': 'TERAZI', 'adaylar': adaylar, 'durum': durum, 'kanit': kanit, 'foto': foto})
                self.txt_log.append(f"[{foto['dosya']}] -> Modül: TERAZİ | Okunan İsim: {adaylar} | Durum: {durum} ({kanit})")
            elif tip == "BIOMERIEUX":
                temiz_satirlar = self.biomerieux_filtrele(foto)
                cozumlenen_ekranlar.append({'tip': 'BIOMERIEUX', 'satirlar': temiz_satirlar, 'foto': foto})
                self.txt_log.append(f"[{foto['dosya']}] -> Modül: bioMérieux | Satır: {len(temiz_satirlar)}")
            elif tip == "FTIR":
                ayrik_satirlar = self.ftir_satirlari_ayristir(foto)
                cozumlenen_ekranlar.append({'tip': 'FTIR', 'satirlar': ayrik_satirlar, 'foto': foto})
                self.txt_log.append(f"[{foto['dosya']}] -> Modül: FT-IR (Sol Users ve Sağ Ağaç Ayrıştırıldı) | Satır: {len(ayrik_satirlar)}")
            else:
                cozumlenen_ekranlar.append({'tip': tip, 'satirlar': foto['satirlar'], 'foto': foto})
                self.txt_log.append(f"[{foto['dosya']}] -> Modül: {tip} | Satır Sayısı: {len(foto['satirlar'])}")

        for k_adi, data in self.resmi_kullanicilar.items():
            k_clean = self.metin_temizle(k_adi)
            ad_soyad = data["AdSoyad"]
            ad_soyad_clean = self.metin_temizle(ad_soyad)
            beklenen = data["Durum"]

            bulundu = False
            cihaz_durumu = "H"
            kanit_metni = "Hiçbir fotoğrafta bulunamadı"

            for ekran in cozumlenen_ekranlar:
                tip = ekran['tip']
                foto = ekran['foto']

                if tip == "TERAZI":
                    if self.eslesme_var_mi(k_clean, ad_soyad_clean, ekran['adaylar'], " ".join(ekran['adaylar']), "TERAZI"):
                        bulundu = True
                        cihaz_durumu = ekran['durum']
                        kanit_metni = ekran['kanit']
                        break
                else:
                    for s in ekran['satirlar']:
                        kelimeler = [k['clean'] for k in s['kelimeler']]
                        if self.eslesme_var_mi(k_clean, ad_soyad_clean, kelimeler, s['metin'], tip):
                            bulundu = True
                            if tip == "TABLO":
                                cihaz_durumu = self.satir_durumu_belirle_tablo(kelimeler, s['metin'])
                                d_etiket = "Pasif (Disabled/Inactive/Locked)" if cihaz_durumu == "H" else "Aktif (Active/Enabled)"
                                kanit_metni = f"{d_etiket} [{foto['dosya']} -> {s['raw_metin'][:40]}]"
                            elif tip == "FTIR":
                                kilit_durumu = self.ftir_kilit_kontrolu(foto['img'], s)
                                cihaz_durumu = kilit_durumu
                                d_etiket = "Kilitli/Yasaklı İkon (Pasif)" if kilit_durumu == "H" else "Normal İkon (Aktif)"
                                kanit_metni = f"FT-IR {d_etiket} [{foto['dosya']} -> {s['raw_metin'][:40]}]"
                            elif tip == "BIOMERIEUX":
                                cihaz_durumu = "E"
                                kanit_metni = f"bioMérieux Listesinde Aktif [{foto['dosya']} -> {s['raw_metin'][:40]}]"
                            else:
                                cihaz_durumu = "E"
                                kanit_metni = f"Listede Mevcut [{foto['dosya']} -> {s['raw_metin'][:40]}]"
                            
                            if cihaz_durumu == "E":
                                break
                    if bulundu and cihaz_durumu == "E":
                        break

            beklenen_str = "Aktif (E)" if beklenen == "E" else "Pasif/Silinmiş (H)"
            if bulundu:
                cihaz_str = "Cihazda AKTİF (E)" if cihaz_durumu == "E" else "Cihazda PASİF (H)"
                if cihaz_durumu == beklenen:
                    self.add_row(self.table_uyumlu, [k_adi, ad_soyad, beklenen_str, cihaz_str, kanit_metni], QColor("#DCFCE7"))
                elif beklenen == "E" and cihaz_durumu == "H":
                    self.add_row(self.table_hata_pasif, [k_adi, ad_soyad, beklenen_str, cihaz_str, kanit_metni], QColor("#FEE2E2"))
                elif beklenen == "H" and cihaz_durumu == "E":
                    self.add_row(self.table_hata_aktif, [k_adi, ad_soyad, beklenen_str, cihaz_str, kanit_metni], QColor("#FEF08A"))
            else:
                if beklenen == "E":
                    self.add_row(self.table_hata_pasif, [k_adi, ad_soyad, beklenen_str, "CİHAZDA HİÇ YOK!", kanit_metni], QColor("#FEE2E2"))
                else:
                    self.add_row(self.table_uyumlu, [k_adi, ad_soyad, beklenen_str, "Cihazda Yok (Silinщен)", "Sorun Yok (Cihazda bulunmadı)"], QColor("#DCFCE7"))

        self.tabs.setCurrentIndex(0)
        QMessageBox.information(self, "Denetim Tamamlandı", "Denetim tamamlandı. Sonuçları ve kanıtları inceleyebilirsiniz.")

    def add_row(self, table, data_list, bg_color=None):
        row_pos = table.rowCount()
        table.insertRow(row_pos)
        for i, data in enumerate(data_list):
            item = QTableWidgetItem(str(data))
            if bg_color:
                item.setBackground(bg_color)
                item.setForeground(QColor("black"))
            table.setItem(row_pos, i, item)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setFont(QFont("Segoe UI", 10))
    ex = YetkiDenetimApp()
    ex.show()
    sys.exit(app.exec())