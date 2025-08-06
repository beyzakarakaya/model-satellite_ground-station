# python C:\Users\lenovo\Desktop\MEKATEK_MUY\pyqt6_yer_istasyonu\yer_istasyonu_3.py & python C:\Users\lenovo\Desktop\MEKATEK_MUY\pyqt6_yer_istasyonu\video_html\app.py
import sys
import random
import time
import matplotlib.pyplot as plt
from PyQt6.QtCore import QThread, pyqtSignal, QUrl, Qt
from PyQt6.QtGui import QPixmap, QIcon, QColor, QTransform
from PyQt6.QtWidgets import QApplication, QMainWindow, QGridLayout, QTableWidgetItem, QLabel, QVBoxLayout, QMessageBox, QSizePolicy
import PyQt6.QtWebEngineWidgets 
from OpenGL.GL import *
from OpenGL.GLUT import *
from OpenGL.GLU import *
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from openpyxl import Workbook
from PyQt6.QtSerialPort import QSerialPortInfo
import numpy as np
import serial
import sqlite3
from serial import Serial, SerialException
from uydu_design import Ui_MainWindow
import time
from PyQt6.QtWidgets import QApplication, QTableView
from PyQt6.QtGui import QStandardItemModel, QStandardItem


# fig, axes = plt.subplots(6, 1, figsize=(10, 12))  # 6 grafik için 6 satırlı bir subplot oluşturun


ax = ay = az = 0.0
yaw_mode = True
tick_fontsize = 5

global data
veri_listesi = []  # Excel'e kaydedilecek veri listesi

ser = None



# img = Image.open("pyqt6_yer_istasyonu/gorsel/frame.png")        # Paketten görüntü array biçiminde gelecek.
# img_array = np.array(img)    

import sqlite3
global curs
global conn
conn=sqlite3.connect(r'C:\Users\lenovo\OneDrive\Desktop\MODEL UYDU 2025\ground_station\veritabani\veritabani.db')
curs = conn.cursor()

#Toplam 18 sütun. Eğer paketler tablosu yoksa, paketler tablosunu sütunlarıyla birlikte yazdırır.
sorguCreateTablePaketler = ("CREATE TABLE IF NOT EXISTS paketler(Paket_Numarasi INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT, Uydu_statusu TEXT NOT NULL, Hata_kodu TEXT NOT NULL, GondermeSaati TEXT NOT NULL, Basinc1 TEXT NOT NULL, Basinc2 TEXT NOT NULL, Yukseklik1 TEXT NOT NULL, Yukseklik2 TEXT NOT NULL, IrtifaFarki TEXT NOT NULL, InisHizi TEXT NOT NULL, Sicaklik TEXT NOT NULL, PilGerilimi TEXT NOT NULL, GPS1Latitude TEXT NOT NULL, GPS1Longitude TEXT NOT NULL, GPS1Altitude TEXT NOT NULL, Pitch TEXT NOT NULL, Roll TEXT NOT NULL, Yaw TEXT NOT NULL, RHRH TEXT NOT NULL, IoTs1Data TEXT NOT NULL, IoTs2Data TEXT NOT NULL, Takim_no TEXT NOT NULL)")
curs.execute(sorguCreateTablePaketler)
conn.commit()
a = 1


data_pitch = [0,0,0,0,0,0,0,0,0,0]
data_yaw = [0,0,0,0,0,0,0,0,0,0]
data_roll = [0,0,0,0,0,0,0,0,0,0]

data_basinc_gorev_yuku = [0,0,0,0,0,0,0,0,0,0]
data_basinc_tasiyici = [0,0,0,0,0,0,0,0,0,0]

data_irtifa = [0] * 10

data_yukseklik_gorev_yuku = [0,0,0,0,0,0,0,0,0,0]
data_yukseklik_tasiyici = [0,0,0,0,0,0,0,0,0,0]
data_hiz = [0,0,0,0,0,0,0,0,0,0]
data_sicaklik = [0,0,0,0,0,0,0,0,0,0]
data_iot = [0,0,0,0,0,0,0,0,0,0]

data_gps_latitude = [0,0,0,0,0,0,0,0,0,0]
data_gps_longitude = [0,0,0,0,0,0,0,0,0,0]
data_gps_altitude = [0,0,0,0,0,0,0,0,0,0]

data_pil_gerilimi = [0,0,0,0,0,0,0,0,0,0]
#--------------------------
d=0 
k=0
p=0

l=0
m=0

t=0

u=0 
r=0

j=0

y=0
yi="6G9R"
ha="000000"
tn="594335"
gs="05/05/2025 , 12/00/00"

w=0 #iot data
c=0 
b=0 #pilgerilimi



paket_no=0
#--------------------------
class MyThread(QThread):
    # Güncelleme sinyali tanımlama
    update_signal = pyqtSignal()

    # İş parçacığı ana metodu
    def run(self):
        while True:
            # Güncelleme sinyalini gönderme
            self.update_signal.emit()
            # 1000 milisaniye (1 saniye) bekleme
            self.msleep(1000)

class Window(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        global ser
        try:
            ser = serial.Serial('COM4', 9600, timeout=1)  # COM portunu kendine göre değiştir
            time.sleep(2)
        except Exception as e:
            print("Seri port bağlantı hatası:", e)
            ser = None
        self.isMaximized()
        self.setMinimumSize(800, 600)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.hata_kodu_bin = "000000"



        # Butonlar burada tanımlanır.
        self.excel_kaydet.clicked.connect(self.excelkaydet)
        #------------------------------------------------

        #------Görsel kısımlar burada tanımlanır---------
        self.label_logo.setPixmap(
        QPixmap("ground_station/photos/logo1.png").scaled(
        158, 215,                         # Genişlik ve yükseklik sınırı
        Qt.AspectRatioMode.KeepAspectRatio, 
        Qt.TransformationMode.SmoothTransformation
    )
)

    
        # self.goruntu_labeli.setPixmap(QPixmap("pyqt6_yer_istasyonu/gorsel/frame.png"))
        self.setWindowIcon(QIcon('ground_station/photos/logo1.png'))
        #---Tablo sütunlarının isimleri---
        #self.tableWidget.setItem(satir, sutun, item)
        sutunliste = ["IRTIFA FARKI", "INIS HIZI", "SICAKLIK", "PIL GERILIMI", "GPS1 LATITUDE", "GPS1 LONGITUDE", "GPS1 ALTITUDE", "PITCH", "ROLL", "YAW", "RHRH", "IoT S1 DATA", "IoT S2 DATA", "TAKIM NO"]
        m=0
        for k in [1]:
            for i in range(8):
                self.tableWidget.setItem(k, i, QTableWidgetItem(sutunliste[m]))
                m=m+1
            for k in [3]:
                for i in range(6):
                    self.tableWidget.setItem(k, i, QTableWidgetItem(sutunliste[m]))
                    m=m+1

                #---

        # Tabloyu yeniden boyutlandırırken font boyutunu ayarlama
        font = self.tableWidget.font()
        font.setPointSize(8)  # Font boyutunu ayarlayarak içerik boyutunu kontrol et
        self.tableWidget.setFont(font)

        self.tableWidget.setColumnWidth(3, 120)  # İlk sütunun genişliğini ayarla




        self.tableWidget.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)  # Dikey kaydırma çubuğunu kapat
        self.tableWidget.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)  # Yatay kaydırma çubuğunu kapat

        



        # layout ve ekstra label oluşturma
        self.kamera = QLabel()
        self.gyro_labeli = QLabel()
        self.layout = QGridLayout(self.tableWidget_5)
        self.layout3 = QVBoxLayout(self.tableWidget_2)  # iot için boş
        self.layout2 = QVBoxLayout(self.tableWidget_3)  
        self.layout6 = QVBoxLayout(self.tableWidget_6)


        # Matplotlib figürünü PyQt6 arayüze yerleştirme
        self.figure_iot, self.ax_iot = plt.subplots(figsize=(3,2))
        plt.tight_layout()
        self.figure_eksen, self.ax_eksen = plt.subplots(figsize=(2, 2))
        plt.tight_layout()
        self.figure_basinc, self.ax_basinc = plt.subplots(figsize=(2, 2))
        plt.tight_layout()
        self.figure_irtifa, self.ax_irtifa = plt.subplots(figsize=(2, 2))
        plt.tight_layout()
        self.figure_yukseklik, self.ax_yukseklik = plt.subplots(figsize=(2, 2))
        plt.tight_layout()
        self.figure_hiz, self.ax_hiz = plt.subplots(figsize=(2, 2))
        plt.tight_layout()
        self.figure_pil_gerilimi, self.ax_pil_gerilimi = plt.subplots(figsize=(2, 2))
        plt.tight_layout()
        self.figure_gps, self.ax_gps = plt.subplots(figsize=(2, 2))
        plt.tight_layout()
        self.figure_sicaklik, self.ax_sicaklik = plt.subplots(figsize=(2, 2))
        plt.tight_layout()
        self.figure_iot_sicaklik1, self.ax_iot_sicaklik1 = plt.subplots(figsize=(2, 2))
        plt.tight_layout()
        self.figure_iot_sicaklik2, self.ax_iot_sicaklik2 = plt.subplots(figsize=(2, 2))
        plt.tight_layout()


        self.canvas_iot = FigureCanvas(self.figure_iot)
    
        self.canvas_eksen = FigureCanvas(self.figure_eksen)
        self.canvas_basinc = FigureCanvas(self.figure_basinc)
        self.canvas_irtifa = FigureCanvas(self.figure_irtifa)
        self.canvas_yukseklik = FigureCanvas(self.figure_yukseklik)
        self.canvas_hiz = FigureCanvas(self.figure_hiz)
        self.canvas_pil_gerilimi = FigureCanvas(self.figure_pil_gerilimi)
        self.canvas_gps = FigureCanvas(self.figure_gps)
        self.canvas_sicaklik = FigureCanvas(self.figure_sicaklik)
        self.canvas_iot_sicaklik1 = FigureCanvas(self.figure_iot_sicaklik1)
        self.canvas_iot_sicaklik2 = FigureCanvas(self.figure_iot_sicaklik2)



        # GPS widget'ı.
        self.webview2 = PyQt6.QtWebEngineWidgets.QWebEngineView(self)
        self.webview3 = PyQt6.QtWebEngineWidgets.QWebEngineView(self)
        self.ax_iot.plot()
        self.ax_eksen.plot()
        self.ax_basinc.plot()
        self.ax_irtifa.plot()
        self.ax_yukseklik.plot()
        self.ax_hiz.plot()
        self.ax_pil_gerilimi.plot()
        self.ax_gps.plot()
        self.ax_sicaklik.plot()
       
        # GPS'in layout'u.
        self.layout2.addWidget(self.webview2)
        self.layout6.addWidget(self.webview3)
        
        self.gyro_pixmap = QPixmap("pyqt6_yer_istasyonu/gorsel/gyro.png")
        self.gyro_labeli.setScaledContents(True)
        self.gyro_labeli.setPixmap(self.gyro_pixmap.scaled(self.tableWidget_6.width(), self.tableWidget_6.height()).transformed(QTransform().rotate(0)))

        self.layout4 = QVBoxLayout(self.tableWidget_6)
        self.layout4.addWidget(self.gyro_labeli)

        # Grafiklerin layout'u.
        self.layout3.addWidget(self.canvas_iot)
        self.layout.addWidget(self.canvas_eksen, 0, 0)
        self.layout.addWidget(self.canvas_basinc, 0, 1)
        self.layout.addWidget(self.canvas_hiz, 0, 2)
        self.layout.addWidget(self.canvas_pil_gerilimi, 0, 3)
        self.layout.addWidget(self.canvas_irtifa, 0, 4)
        self.layout.addWidget(self.canvas_sicaklik, 1, 0)
        self.layout.addWidget(self.canvas_yukseklik, 1, 1)
        self.layout.addWidget(self.canvas_gps, 1, 2)
        self.layout.addWidget(self.canvas_iot_sicaklik1, 1, 3)
        self.layout.addWidget(self.canvas_iot_sicaklik2, 1, 4)

        # ----------------------------------
        # görüntüle
        self.webview2.setUrl(QUrl("http://127.0.0.1:5000/"))
        self.webview3.setUrl(QUrl("http://127.0.0.1:8080/"))
        #-----------------------------------

        # # Thread oluşturma ve başlatma
        self.thread = MyThread()
        self.thread.update_signal.connect(self.LISTELE)
        self.thread.update_signal.connect(self.butun_grafikler_guncelle)
        self.thread.update_signal.connect(self.gyro_guncelle)
        self.thread.start()
        # #-------------------------------------------------

    def gyro_guncelle(self):
        global gyro_data
        gyro_data = random.randint(-10, 10)*36
        self.gyro_labeli.setPixmap(self.gyro_pixmap.transformed(QTransform().rotate(gyro_data)))
        # Seri port ve frekans seçicileri

    # def konum(self):
    #         Folium map oluştur
    #         m = folium.Map(location=[41.086149, 28.917113], zoom_start=40, zoom_control=False, tiles='cartodb dark_matter')
    #         folium.TileLayer('cartodb dark_matter').add_to(m)

    #         Marker ekle
    #         folium.Marker(location=[41.086149, 28.917113], icon=folium.Icon(icon="glyphicon-screenshot")).add_to(m)
    #         Folium mapi html olarak kaydet
    #         m.save("pyqt6_yer_istasyonu/map/map.html")
    #         görüntüle
    #         self.webview.setHtml(open("pyqt6_yer_istasyonu/map/map.html").read())

    def KAPAT(self):
        # paketler tablosunu siler.
        conn.execute("DROP TABLE IF EXISTS paketler;")
        conn.commit()
        conn.close()
        self.close()

    
    

    
    def excelkaydet(self):
        from openpyxl import Workbook
        from PyQt6.QtWidgets import QMessageBox
        import os

        global veri_listesi  # Listeyi kullanabilmek için

        button = QMessageBox.information(
            self, "Excel tablosuna aktarma",
            "Kaydetmek istiyor musunuz?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if button == QMessageBox.StandardButton.Yes:
            if not veri_listesi:
                QMessageBox.warning(self, "Uyarı", "Kaydedilecek veri bulunamadı.")
                return

            wb = Workbook()
            ws = wb.active
            ws.title = "Telemetri Verisi"

            # Sütun başlıkları
            headers = list(veri_listesi[0].keys())
            ws.append(headers)

            # Tüm verileri ekle
            for row in veri_listesi:
                ws.append(list(row.values()))

            # Kayıt yeri
            save_path = 'ground_station/veritabani_verileri/veri_kaydi.xlsx'
            os.makedirs(os.path.dirname(save_path), exist_ok=True)
            wb.save(save_path)

            QMessageBox.information(self, "Başarılı", f"Excel dosyası kaydedildi:\n{save_path}")


    def LISTELE(self):
    
        if ser and ser.in_waiting:
            satir = ser.readline().decode().strip()
            try:
                pitch_str, roll_str, yaw_str = satir.split(",")
                global y_pitch, y_roll, y_yaw
                y_pitch = float(pitch_str)
                y_roll = float(roll_str)
                y_yaw = float(yaw_str)
            except ValueError:
                print("MPU9250 verisi okunamadı:", satir)


        # Veritabanından en son eklenen veriyi al
        self.tableWidget.verticalHeader().setVisible(False)  
        curs.execute("SELECT * FROM paketler ORDER BY Paket_Numarasi DESC LIMIT 1")

        new_row = curs.fetchone()


        if new_row:
            veri_tab = [
                [new_row[0], f"{a}", f"{ha}", f"{gs}", f"{a} Pa", f"{k} Pa", f"{a} m", f"{k} m"],
                [k, a, k, a, k, a, k, a],
                [f"{a} m", f"{k} m/s", f"{y:.2f} °C", f"{k} V", f"{a} °", f"{k} °", f"{k} m", f"{a} °"],
                [k, a, k, a, k, a],  # Bu satırı kontrol edin
                [ f"{a} °", f"{k} °", f"{yi}", f"{k} °C", f"{a} °C", f"{tn}"]  # Bu satırı da kontrol edin
 ]

                    # [a, k , a, k, a, k, a]]
            for n in [0, 2, 4]:
                for i in range(len(veri_tab[n])):  # Bu şekilde i'yi veri_tab[n] uzunluğuna göre sınırlandır
                    self.tableWidget.setItem(n, i, QTableWidgetItem(str(veri_tab[n][i])))

        #----------------------------

        # Irtifa farkı tablosundaki değer güncelleme yeri
        # self.irtifa_farki_label.setText(f"{a} m")

        #------------------------------------------------ 
        curs.execute("""
INSERT INTO paketler(Uydu_statusu, Hata_kodu, GondermeSaati, Basinc1, Basinc2, Yukseklik1, Yukseklik2, IrtifaFarki, InisHizi, Sicaklik, PilGerilimi, GPS1Latitude, GPS1Longitude, GPS1Altitude, Pitch, Roll, Yaw, RHRH, IoTs1Data, IoTs2Data, Takim_no)
VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
""", (
    f"{a}",               # Uydu_statusu
    f"{self.hata_kodu_bin}",     # Hata_kodu
    f"{k}",               # GondermeSaati
    f"{a} Pa",           # Basinc1
    f"{k} Pa",           # Basinc2
    f"{a} m",            # Yukseklik1
    f"{k} m",            # Yukseklik2
    f"{a} m",            # IrtifaFarki
    f"{k} m/s",          # InisHizi
    f"{y:.2f} °C",  # gerçek zamanlı sıcaklık
    f"{k} V",            # PilGerilimi
    f"{a} °",              # GPS1Latitude
    f"{k} °",              # GPS1Longitude
    f"{a} m" ,              # GPS1Altitude
    f"{k} °",            # Pitch
    f"{a} °",            # Roll
    f"{k} °",            # Yaw
    f"{a}",              # RHRH
    f"{k} °C",           # IoTs1Data
    f"{a} °C",           # IoTs2Data
    f"{k}"               # Takim_no
))

        conn.commit() 
                # Excel için listeye veri ekle
        veri_listesi.append({
            "Zaman": time.strftime("%H:%M:%S"),
            "Basinc1": f"{a} Pa",
            "Basinc2": f"{k} Pa",
            "Yukseklik1": f"{a} m",
            "Yukseklik2": f"{k} m",
            "IrtifaFarki": f"{a} m",
            "InisHizi": f"{k} m/s",
            "Sicaklik": f"{y:.2f} °C",
            "PilGerilimi": f"{k} V",
            "GPS1Latitude": f"{a} °",
            "GPS1Longitude": f"{k} °",
            "GPS1Altitude": f"{a} m",
            "Pitch": f"{k} °",
            "Roll": f"{a} °",
            "Yaw": f"{k} °",
            "RHRH": f"{a}",
            "IoTs1Data": f"{k} °C",
            "IoTs2Data": f"{a} °C",
            "Takim_no": f"{k}"
        })
        
        inis_hizi_model = float(j)
        inis_hizi_gorev = float(k)
        basinc_tasiyici = k if k > 0 else None
        gps_gorev = a if a != 0 else None
        irtifa_farki = float(a)  
        self.label_26.setText(f"İrtifa farkı = {irtifa_farki} m")

        ayrilma = irtifa_farki >= 2

        ayrilma = True if irtifa_farki >= 2 else False
        multispektral = True  # şimdilik elle veriyoruz

        self.alarm_durumunu_guncelle(
            inis_hizi_model,
            inis_hizi_gorev,
            basinc_tasiyici,
            gps_gorev,
            ayrilma,
            multispektral
        )
        
    def alarm_durumunu_guncelle(self, inis_hizi_model, inis_hizi_gorev, basinc_tasiyici, gps_gorev, ayrilma, multispektral):
        if not ayrilma:
            # Ayrılma gerçekleşmedi: sadece model uydu kontrol edilir
            hiz_kontrol_1 = 12 <= inis_hizi_model <= 14
            hiz_kontrol_2 = False  # Görev yükü kontrol edilmez → kırmızı
        else:
            # Ayrılma oldu: sadece görev yükü kontrol edilir
            hiz_kontrol_1 = False  # Model uydu kontrol edilmez → kırmızı
            hiz_kontrol_2 = 6 <= inis_hizi_gorev <= 8

        durumlar = [
            hiz_kontrol_1,                # label_2: model uydu hızı kontrolü
            hiz_kontrol_2,                # label_3: görev yükü hızı kontrolü
            basinc_tasiyici is not None,  # label_4
            gps_gorev is not None,        # label_5
            ayrilma,                      # label_6
            multispektral                 # label_7
        ]

        label_isimleri = [
            "aras_1_MUY_hiz_label_2",
            "aras_1_MUY_hiz_label_3",
            "aras_1_MUY_hiz_label_4",
            "aras_1_MUY_hiz_label_5",
            "aras_1_MUY_hiz_label_6",
            "aras_1_MUY_hiz_label_7"
        ]

        for i in range(6):
            label = getattr(self, label_isimleri[i])
            renk = "green" if durumlar[i] else "red"
            label.setStyleSheet(f"background-color: " + renk)
    
                # Hata kodunu üret (0 veya 1 olacak şekilde her durum için)
        self.hata_kodu_bin = ''.join(['0' if durum else '1' for durum in durumlar])
        self.tableWidget.setItem(0, 2, QTableWidgetItem(self.hata_kodu_bin))  # satır 0, sütun 2 = HATA KODU






    def butun_grafikler_guncelle(self):
        # Tüm grafik eksenlerini temizleyin
        self.ax_iot.clear()
        self.ax_eksen.clear()
        self.ax_basinc.clear()
        self.ax_irtifa.clear()
        self.ax_yukseklik.clear()
        self.ax_hiz.clear()
        self.ax_sicaklik.clear()
        self.ax_pil_gerilimi.clear()
        self.ax_gps.clear()

        # IoT grafiği
        self.ax_iot.plot(data_iot, color='red', linestyle='-', linewidth=1, label='IoT Data')
        self.ax_iot.set_xlabel('Zaman (s)', fontsize=5)
        self.ax_iot.set_ylabel('Sicaklik(°C)', fontsize=5)
        self.ax_iot.legend(loc='upper right', fontsize=5)
        self.ax_iot.tick_params(axis='both', which='major', labelsize=tick_fontsize)
        global w
        w = random.randint(1, 10)
        data_iot.append(w)
        data_iot.pop(0)

        # GPS Latitude grafiği
        # self.ax_gps.clear()
        # self.ax_gps.plot([random.uniform(1, 10) for _ in range(10)], color='purple', linestyle='-', linewidth=1, label='GPS Latitude (°)')
        # self.ax_gps.set_xlabel('Zaman (s)', fontsize=5)
        # self.ax_gps.set_ylabel('Latitude (°)', fontsize=5)
        # self.ax_gps.legend(loc='upper right', fontsize=5)
        # self.ax_gps.tick_params(axis='both', labelsize=tick_fontsize)

        # # GPS Longitude grafiği
        # self.ax_gps.clear()
        # self.ax_gps.plot([random.uniform(1, 10) for _ in range(10)], color='cyan', linestyle='-', linewidth=1, label='GPS Longitude (°)')
        # self.ax_gps.set_xlabel('Zaman (s)', fontsize=5)
        # self.ax_gps.set_ylabel('Longitude (°)', fontsize=5)
        # self.ax_gps.legend(loc='upper right', fontsize=5)
        # self.ax_gps.tick_params(axis='both', labelsize=tick_fontsize)

        # # GPS Altitude grafiği
        # self.ax_gps.clear()
        # self.ax_gps.plot([random.uniform(1, 10) for _ in range(10)], color='magenta', linestyle='-', linewidth=1, label='GPS Altitude (m)')
        # self.ax_gps.set_xlabel('Zaman (s)', fontsize=5)
        # self.ax_gps.set_ylabel('Altitude (m)', fontsize=5)
        # self.ax_gps.legend(loc='upper right', fontsize=5)
        # self.ax_gps.tick_params(axis='both', labelsize=tick_fontsize)
        # GPS Grafiğini temizle

        #gps grafiği
        self.ax_gps.clear()

        # GPS Latitude (Mavi renk ile)
        self.ax_gps.plot([random.uniform(1, 10) for _ in range(10)], color='blue', linestyle='-', linewidth=1, label='GPS1 Latitude (°)')

        # GPS Longitude (Yeşil renk ile)
        self.ax_gps.plot([random.uniform(1, 10) for _ in range(10)], color='green', linestyle='-', linewidth=1, label='GPS1 Longitude (°)')

        # GPS Altitude (Magenta renk ile)
        self.ax_gps.plot([random.uniform(1, 10) for _ in range(10)], color='magenta', linestyle='-', linewidth=1, label='GPS1 Altitude (m)')

        # Grafiğin başlıkları ve etiketler
        self.ax_gps.set_xlabel('Zaman (s)', fontsize=5)
        self.ax_gps.set_ylabel('Konum (°/m)', fontsize=5)
        self.ax_gps.legend(loc='upper right', fontsize=5)

        # Eksen ve çizgi özellikleri
        self.ax_gps.tick_params(axis='both', labelsize=5)

        # Grafiği çiz
        self.canvas_gps.draw()




        # pil_gerilimi grafiği
        self.ax_pil_gerilimi.clear()
        self.ax_pil_gerilimi.plot([random.uniform(0, 5) for _ in range(10)], color='green', label='Pil Gerilimi (V)')
        self.ax_pil_gerilimi.set_xlabel('Zaman (s)', fontsize=5)
        self.ax_pil_gerilimi.set_ylabel('Gerilim (V)', fontsize=5)
        self.ax_pil_gerilimi.legend(loc='upper right', fontsize=5)
        self.ax_pil_gerilimi.tick_params(axis='both', which='major', labelsize=tick_fontsize)
        global b
        d = random.randint(1, 10)
        data_pitch.append(d)
        data_pitch.pop(0)
    

        

        # Pitch grafiği
        self.ax_eksen.plot(data_pitch, color='lightblue', linestyle='-', linewidth=1, label='Pitch')
        self.ax_eksen.legend(loc='upper right', fontsize=5)
        self.ax_eksen.tick_params(axis='both', which='major', labelsize=tick_fontsize)
        global a
        a = random.randint(1, 10)
        data_pitch.append(a)
        data_pitch.pop(0)

        # Yaw grafiği
        self.ax_eksen.plot(data_yaw, color='orange', linestyle='-', linewidth=1, label='Yaw')
        self.ax_eksen.legend(loc='upper right', fontsize=5)
        self.ax_eksen.tick_params(axis='both', which='major', labelsize=tick_fontsize)
        global y
        y = random.randint(1, 10)
        data_yaw.append(y)
        data_yaw.pop(0)


        # Roll grafiği
        self.ax_eksen.plot(data_roll, color='green', linestyle='-', linewidth=1, label='Roll')
        self.ax_eksen.set_xlabel('Zaman (s)', fontsize=5)
        self.ax_eksen.set_ylabel('Eksen (°)', fontsize=5)
        self.ax_eksen.legend(loc='upper right', fontsize=5)
        self.ax_eksen.tick_params(axis='both', which='major', labelsize=tick_fontsize)
        global k
        k = random.randint(1, 10)
        data_roll.append(k)
        data_roll.pop(0)

        # Basınç Görev Yükü grafiği
        self.ax_basinc.plot(data_basinc_gorev_yuku, color='chocolate', linestyle='-', linewidth=1, label='Basinc1')
        self.ax_basinc.set_xlabel('Zaman (s)', fontsize=5)
        self.ax_basinc.set_ylabel('Basınç(Pa)', fontsize=5)
        self.ax_basinc.legend(loc='upper right', fontsize=5)
        self.ax_basinc.tick_params(axis='both', which='major', labelsize=tick_fontsize)
        global l
        l = random.randint(1, 10)
        data_basinc_gorev_yuku.append(l)
        data_basinc_gorev_yuku.pop(0)


        # Basınç Taşıyıcı grafiği
        self.ax_basinc.plot(data_basinc_tasiyici, color='teal', linestyle='-', linewidth=1, label='Basinc2')
        self.ax_basinc.legend(loc='upper right', fontsize=5)
        self.ax_basinc.tick_params(axis='both', which='major', labelsize=tick_fontsize)
        global m
        m = random.randint(1, 10)
        data_basinc_tasiyici.append(m)
        data_basinc_tasiyici.pop(0)


        # İrtifa grafiği
        self.ax_irtifa.plot(data_irtifa, color='purple', linestyle='-', linewidth=1, label='Irtifa Farki')
        self.ax_irtifa.set_xlabel('Zaman (s)', fontsize=5)
        self.ax_irtifa.set_ylabel('Metre (m)', fontsize=5)
        self.ax_irtifa.legend(loc='upper right', fontsize=5)
        self.ax_irtifa.tick_params(axis='both', which='major', labelsize=tick_fontsize)
        global t
        t = random.randint(1, 20)
        data_irtifa.append(a)
        data_irtifa.pop(0)



        # Yükseklik Görev Yükü grafiği
        self.ax_yukseklik.plot(data_yukseklik_gorev_yuku, color='mediumturquoise', linestyle='-', linewidth=1, label='Yukseklik1')
        self.ax_yukseklik.set_xlabel('Zaman (s)', fontsize=5)
        self.ax_yukseklik.set_ylabel('Metre (m)', fontsize=5)
        self.ax_yukseklik.legend(loc='upper right', fontsize=5)
        self.ax_yukseklik.tick_params(axis='both', which='major', labelsize=tick_fontsize)
        global u
        u = random.randint(1, 10)
        data_yukseklik_gorev_yuku.append(u)
        data_yukseklik_gorev_yuku.pop(0)


        # Yükseklik Taşıyıcı grafiği
        self.ax_yukseklik.plot(data_yukseklik_tasiyici, color='hotpink', linestyle='-', linewidth=1, label='Yukseklik2')
        self.ax_yukseklik.legend(loc='upper right', fontsize=5)
        self.ax_yukseklik.tick_params(axis='both', which='major', labelsize=tick_fontsize)
        global r
        r = random.randint(1, 10)
        data_yukseklik_tasiyici.append(r)
        data_yukseklik_tasiyici.pop(0)


       

        # Sıcaklık grafiği
        self.ax_sicaklik.plot(data_sicaklik, color='darkseagreen', linestyle='-', linewidth=1, label='Sicaklik')
        self.ax_sicaklik.set_xlabel('Zaman (s)', fontsize=5)
        self.ax_sicaklik.set_ylabel('Sıcaklık (°C)', fontsize=5)
        self.ax_sicaklik.legend(loc='upper right', fontsize=5)
        self.ax_sicaklik.tick_params(axis='both', which='major', labelsize=tick_fontsize)
        data_sicaklik.append(y)
        data_sicaklik.pop(0)
        data_sicaklik.append(y)
        data_sicaklik.pop(0)

        # Hız grafiği
        self.ax_hiz.plot(data_hiz, color='darkred', linestyle='-', linewidth=1, label='Iniş Hizi')
        self.ax_hiz.set_xlabel('Zaman (s)', fontsize=5)
        self.ax_hiz.set_ylabel('Hız (m/s)', fontsize=5)
        self.ax_hiz.legend(loc='upper right', fontsize=5)
        self.ax_hiz.tick_params(axis='both', which='major', labelsize=tick_fontsize)
        global j
        j = random.randint(1, 10)
        data_hiz.append(j)
        data_hiz.pop(0)

        self.ax_iot_sicaklik1.clear()
        self.ax_iot_sicaklik1.plot(data_iot, color='darkorange', label='IoT S1 DATA')
        self.ax_iot_sicaklik1.set_xlabel('Zaman (s)', fontsize=5)
        self.ax_iot_sicaklik1.set_ylabel('Sıcaklık (°C)', fontsize=5)
        self.ax_iot_sicaklik1.legend(loc='upper right', fontsize=5)
        self.ax_iot_sicaklik1.tick_params(axis='both', labelsize=5)

        self.ax_iot_sicaklik2.clear()
        self.ax_iot_sicaklik2.plot(data_iot, color='darkblue', label='IoT S2 DATA')
        self.ax_iot_sicaklik2.set_xlabel('Zaman (s)', fontsize=5)
        self.ax_iot_sicaklik2.set_ylabel('Sıcaklık (°C)', fontsize=5)
        self.ax_iot_sicaklik2.legend(loc='upper right', fontsize=5)
        self.ax_iot_sicaklik2.tick_params(axis='both', labelsize=5)

        self.canvas_iot_sicaklik1.draw()
        self.canvas_iot_sicaklik2.draw()

        

        # Tüm grafiklerin yerleşimini sıkılaştırın ve çizdirin
        self.canvas_iot.draw()
        self.canvas_eksen.draw()
        self.canvas_basinc.draw()
        self.canvas_irtifa.draw()
        self.canvas_yukseklik.draw()
        self.canvas_hiz.draw()
        self.canvas_gps.draw()
        self.canvas_pil_gerilimi.draw()
        self.canvas_sicaklik.draw()





if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = Window()
    win.show()

    sys.exit(app.exec())
