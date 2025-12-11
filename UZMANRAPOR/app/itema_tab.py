from __future__ import annotations

from typing import Dict, Optional

import pyodbc
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QGridLayout, QMessageBox, QScrollArea
)
from PySide6.QtCore import Qt

from app.itema_settings import build_itema_settings


# Burayı kendi ortamına göre AYARLAMALISIN:
SQL_SERVER = "10.30.9.14,1433"     # Örn: "STIBRSSFSRV01" veya "(local)"
SQL_DATABASE = "UzmanRaporDB"           # Bizim kurduğumuz veritabanı adı


def get_sql_connection() -> pyodbc.Connection:
    """
    ITEMA ayarlarını okumak için SQL bağlantısı.
    Gerekirse entegrasyonu daha sonra senin global bağlantına bağlarız.
    """
    conn_str = (
        "Driver={SQL Server};"
        f"Server={SQL_SERVER};"
        f"Database={SQL_DATABASE};"
        "Trusted_Connection=yes;"
    )
    return pyodbc.connect(conn_str)


class ItemaAyarTab(QWidget):
    """
    Excel'deki ITEMA_AYAR_FORMU sayfasına benzer yeni sekme.
    Tip kodu girilip 'Otomatik Ayarları Getir' denildiğinde:

      - Varsayılan ayarlar
      - Otomatik ayarlar (sp_ItemaOtomatikAyar)
      - Tip-özel ayarlar (sp_ItemaTipOzelAyar)

    birleştirilerek arayüz alanlarına yazılır.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._fields: Dict[str, QLineEdit] = {}
        self._build_ui()

    # ------------------------------------------------------------------
    # UI KURULUMU
    # ------------------------------------------------------------------
    def _build_ui(self):
        main_layout = QVBoxLayout(self)

        # ÜST BAR: Tip gir + buton
        top = QHBoxLayout()
        top.addWidget(QLabel("Tip Kodu:"))
        self.ed_tip = QLineEdit()
        self.ed_tip.setPlaceholderText("Örn: RX14908")
        self.ed_tip.setMaxLength(50)
        top.addWidget(self.ed_tip)

        self.btn_fetch = QPushButton("Otomatik Ayarları Getir")
        self.btn_fetch.clicked.connect(self._on_fetch_clicked)
        top.addWidget(self.btn_fetch)

        top.addStretch(1)
        main_layout.addLayout(top)

        # ALANLAR: scroll içinde grid
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        inner = QWidget()
        grid = QGridLayout(inner)
        grid.setColumnStretch(1, 1)
        grid.setColumnStretch(3, 1)

        row = 0

        def add_row(label_left: str, key_left: str,
                    label_right: Optional[str] = None, key_right: Optional[str] = None):
            nonlocal row
            # Sol
            lblL = QLabel(label_left)
            editL = QLineEdit()
            editL.setObjectName(key_left)
            grid.addWidget(lblL, row, 0)
            grid.addWidget(editL, row, 1)
            self._fields[key_left] = editL

            # Sağ (opsiyonel)
            if label_right and key_right:
                lblR = QLabel(label_right)
                editR = QLineEdit()
                editR.setObjectName(key_right)
                grid.addWidget(lblR, row, 2)
                grid.addWidget(editR, row, 3)
                self._fields[key_right] = editR

            row += 1

        # 1) Telef & Fırça & Üfleme & Cımbar & Tansiyon & Devir & Leno
        add_row("Telef Sol", "telef_ken1", "Telef Sağ", "telef_ken2")
        add_row("Fırça Seçimi", "firca_secim", "Cımbar", "cimbar_secim")
        add_row("Üfleme Zamanı 1", "ufleme_zam_1", "Üfleme Zamanı 2", "ufleme_zam_2")
        add_row("Tansiyon", "coz_tansiyon", "Devir", "devir")
        add_row("Leno", "leno", "Arka Desen", "ark_desen")

        # 2) Armür / ağızlık / testere / yay (L30..L39)
        add_row("Ağızlık Geometrisi (Strok)", "agizlik", "Arka Köprü Derinlik", "derinlik")
        add_row("Arka Köprü Pozisyon", "pozisyon", "Testere Uzaklık", "testere_uzk")
        add_row("Testere Yükseklik", "testere_yuk", "Tansiyon Yay Pozisyonu", "tan_yay_pozisyon")
        add_row("Tansiyon Yay Yüksekliği", "tan_yay_yukseklik", "Tansiyon Yay Konumu", "tan_yay_konumu")
        add_row("Yay Boğumu", "tan_yay_bogumu", "Zemin Ağızlık", "zem_agizlik")

        # 3) Motor rampaları
        add_row("Motor Rampası 1", "rampa_1", "Motor Rampası 2", "rampa_2")
        add_row("Motor Rampası 3", "rampa_3", "Motor Rampası 4", "rampa_4")
        add_row("Motor Rampası 5", "rampa_5", "Motor Rampası 6", "rampa_6")

        # 4) Açıklama / Değişiklik Yapan
        add_row("Açıklama", "aciklama", "Değişiklik Yapan", "degisiklik_yapan")

        # ufak boşluk
        grid.setRowStretch(row, 1)

        scroll.setWidget(inner)
        main_layout.addWidget(scroll, 1)

    # ------------------------------------------------------------------
    # LOGİK: AYAR ÇEKME
    # ------------------------------------------------------------------
    def _on_fetch_clicked(self):
        tip_raw = (self.ed_tip.text() or "").strip()
        if not tip_raw:
            QMessageBox.warning(self, "Uyarı", "Önce bir tip kodu girin.")
            return

        tip = tip_raw.upper()

        try:
            conn = get_sql_connection()
        except Exception as e:
            QMessageBox.critical(
                self,
                "Bağlantı Hatası",
                f"SQL sunucusuna bağlanılamadı:\n\n{e}"
            )
            return

        try:
            settings = build_itema_settings(conn, tip)
        except Exception as e:
            QMessageBox.critical(
                self,
                "Hata",
                f"ITEMA ayarları okunurken bir hata oluştu:\n\n{e}"
            )
            return
        finally:
            try:
                conn.close()
            except Exception:
                pass

        if not settings:
            QMessageBox.information(
                self,
                "Bilgi",
                f"{tip} tipi için ITEMA ayarı bulunamadı."
            )
            return

        # Alanlara yaz
        for key, widget in self._fields.items():
            val = settings.get(key)
            widget.setText(val or "")

        # Kullanıcıya kısa bir bilgi mesajı
        QMessageBox.information(
            self,
            "Tamam",
            f"{tip} tipi için otomatik ITEMA ayarları getirildi."
        )
