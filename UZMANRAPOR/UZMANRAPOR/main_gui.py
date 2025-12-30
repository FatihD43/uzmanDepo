import sys, os
import time

sys.path.insert(0, os.path.dirname(__file__))

from PySide6.QtWidgets import QApplication, QDialog, QSplashScreen, QProgressBar
from PySide6.QtGui import QPixmap, QColor, QFont
from PySide6.QtCore import Qt

from app.gui import MainWindow
from app.login_dialog import LoginDialog
import resources.app_resources_rc  # noqa: F401


# --- ÖZEL AÇILIŞ EKRANI SINIFI ---
class YuklemeEkrani(QSplashScreen):
    def __init__(self, pixmap):
        super().__init__(pixmap)

        # 1. Progress Bar Oluştur
        self.progress = QProgressBar(self)

        # Barın konumu ve boyutu (Resmin en altına yerleştiriyoruz)
        # Resmin genişliği kadar, yüksekliği 25px
        self.progress.setGeometry(0, pixmap.height() - 25, pixmap.width(), 25)

        # 2. Progress Bar Görünümü (CSS)
        # Turuncu/Mavi tonlarında modern bir görünüm
        self.progress.setStyleSheet("""
            QProgressBar {
                border: none;
                background-color: #333333;
                text-align: center;
                color: white;
                font-weight: bold;
            }
            QProgressBar::chunk {
                background-color: #qlineargradient(
                    spread:pad, x1:0, y1:0, x2:1, y2:0, 
                    stop:0 #FF9800, stop:1 #FF5722
                );
            }
        """)

        # Başlangıç değeri
        self.progress.setValue(0)
        self.progress.setFormat("Yükleniyor... %p%")  # Yüzdeyi göster

    def ilerleme_guncelle(self, deger, mesaj=""):
        """Barı ve ekrandaki yazıyı günceller"""
        self.progress.setValue(deger)
        if mesaj:
            self.showMessage(
                mesaj,
                Qt.AlignBottom | Qt.AlignCenter,
                Qt.white
            )
        # Arayüzün donmaması için event'leri işle
        QApplication.instance().processEvents()


def main():
    app = QApplication(sys.argv)

    # 1. Login Ekranı
    login = LoginDialog()
    if login.exec() != QDialog.Accepted or login.user is None:
        return

    # ---------------------------------------------------------
    # 2. AÇILIŞ EKRANI (SPLASH) AYARLARI
    # ---------------------------------------------------------
    img_path = os.path.join(os.path.dirname(__file__), "assets", "acilis.png")

    splash = None
    if os.path.exists(img_path):
        # Resmi yükle
        raw_pixmap = QPixmap(img_path)

        # RESMİ ÖLÇEKLE: Eğer resim çok büyükse max 700px genişliğe düşür
        # Böylece tüm ekranı kaplamaz, pencere gibi durur.
        if raw_pixmap.width() > 700:
            pixmap = raw_pixmap.scaledToWidth(700, Qt.SmoothTransformation)
        else:
            pixmap = raw_pixmap

        # Özel sınıfımızı başlat
        splash = YuklemeEkrani(pixmap)
        splash.show()

        # Hoşgeldin mesajı
        splash.ilerleme_guncelle(10, f"Hoş geldin {login.user.username}")
        time.sleep(0.5)  # Kullanıcı görsün diye çok kısa bekleme (isteğe bağlı)

    else:
        print(f"Uyarı: assets/acilis.png bulunamadı.")

    # ---------------------------------------------------------
    # 3. ANA PENCERE YÜKLENİYOR SİMÜLASYONU
    # ---------------------------------------------------------
    # Not: MainWindow.__init__ tek blokta çalıştığı için bar
    # yükleme sırasında bir süre donabilir. Bunu aşmak için
    # yapay adımlar ekliyoruz.

    if splash:
        splash.ilerleme_guncelle(30, "Ayarlar yükleniyor...")

    # MainWindow'u oluştururken arka planda veri çekiyor
    try:
        if splash:
            splash.ilerleme_guncelle(50, "Veriler hazırlanıyor...")

        win = MainWindow(user=login.user)

        if splash:
            splash.ilerleme_guncelle(90, "Arayüz oluşturuluyor...")
            time.sleep(0.3)  # Geçişi yumuşatmak için

    except Exception as e:
        # Hata olursa splash'i kapat ki kullanıcı hatayı görsün
        if splash: splash.close()
        raise e

    # ---------------------------------------------------------
    # 4. AÇILIŞ
    # ---------------------------------------------------------
    if splash:
        splash.ilerleme_guncelle(100, "Hazır!")
        time.sleep(0.2)

    win.show()

    if splash:
        splash.finish(win)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()