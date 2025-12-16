from __future__ import annotations

import os
from typing import Dict, Optional

import pyodbc
import pandas as pd
from PySide6.QtGui import QPainter
from PySide6.QtPrintSupport import QPrintDialog, QPrinter
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QGridLayout,
    QMessageBox,
    QScrollArea,
    QGroupBox,
    QInputDialog,
)
from PySide6.QtCore import Qt

from app.itema_settings import build_itema_settings, ITEMA_COLUMNS


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
        self._dynamic_fields: Dict[str, QLineEdit] = {}
        self._manual_password = (
            os.getenv("ITEMA_MANUAL_PASSWORD")
            or os.getenv("ITEMA_FORM_PASSWORD")
            or "itema2024"
        )
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

        self.btn_print = QPushButton("A4 Çıktı Al")
        self.btn_print.clicked.connect(self._print_form)
        top.addWidget(self.btn_print)

        top.addStretch(1)
        main_layout.addLayout(top)

        # ALANLAR: scroll içinde grid
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        inner = QWidget()
        inner_layout = QVBoxLayout(inner)

        inner_layout.addWidget(self._build_header_box())
        inner_layout.addWidget(self._build_body_box())
        inner_layout.addWidget(self._build_footer_box())
        inner_layout.addStretch(1)

        scroll.setWidget(inner)
        main_layout.addWidget(scroll, 1)

    def _build_header_box(self) -> QGroupBox:
        box = QGroupBox("Ürün / Tip Bilgileri")
        layout = QGridLayout(box)
        layout.setVerticalSpacing(6)

        def add(label: str, key: str, col: int, row: int, dynamic: bool = True):
            lbl = QLabel(label)
            lbl.setStyleSheet("font-weight: bold; color: #0A5097;")
            edit = QLineEdit()
            edit.setObjectName(key)
            layout.addWidget(lbl, row, col * 2)
            layout.addWidget(edit, row, col * 2 + 1)
            target_dict = self._dynamic_fields if dynamic else self._fields
            target_dict[key] = edit

        add("Tip Kodu", "tip", 0, 0)
        add("Tarak No", "tarak", 1, 0)
        add("Zemin Örgüsü", "zemin_orgusu", 0, 1)
        add("Kenar Örgüsü", "kenar_orgusu", 1, 1)
        add("Süs Kenar Diş Sayısı", "sus_kenar_dis", 0, 2)
        add("Dokunabilirlik", "dokunabilirlik", 1, 2)
        add("Çözgü Kodu", "cozgu_kodu", 0, 3)
        add("Boya Kodu", "boya_kodu", 1, 3)
        add("Çerçeve Adedi", "cerceve_adedi", 0, 4)
        add("Kenar Adedi", "kenar_adedi", 1, 4)

        return box

    def _build_body_box(self) -> QGroupBox:
        box = QGroupBox("Makine Ayarları")
        grid = QGridLayout(box)
        grid.setVerticalSpacing(6)
        grid.setHorizontalSpacing(12)

        row = 0

        def add_row(label_left: str, key_left: str,
                    label_right: Optional[str] = None, key_right: Optional[str] = None,
                    color: Optional[str] = None):
            nonlocal row
            lblL = QLabel(label_left)
            if color:
                lblL.setStyleSheet(f"color: {color}; font-weight: bold;")
            editL = QLineEdit()
            editL.setObjectName(key_left)
            grid.addWidget(lblL, row, 0)
            grid.addWidget(editL, row, 1)
            self._fields[key_left] = editL

            if label_right and key_right:
                lblR = QLabel(label_right)
                if color:
                    lblR.setStyleSheet(f"color: {color}; font-weight: bold;")
                editR = QLineEdit()
                editR.setObjectName(key_right)
                grid.addWidget(lblR, row, 2)
                grid.addWidget(editR, row, 3)
                self._fields[key_right] = editR
            row += 1

        add_row("Telef Sol", "telef_ken1", "Telef Sağ", "telef_ken2", color="#d9534f")
        add_row("Fırça Seçimi", "firca_secim", "Cımbar", "cimbar_secim")
        add_row("Üfleme Zamanı 1", "ufleme_zam_1", "Üfleme Zamanı 2", "ufleme_zam_2")
        add_row("Tansiyon", "coz_tansiyon", "Devir", "devir")
        add_row("Leno", "leno", "Arka Desen", "ark_desen")

        add_row("Ağızlık Geometrisi (Strok)", "agizlik", "Arka Köprü Derinlik", "derinlik")
        add_row("Arka Köprü Pozisyon", "pozisyon", "Testere Uzaklık", "testere_uzk")
        add_row("Testere Yükseklik", "testere_yuk", "Tansiyon Yay Pozisyonu", "tan_yay_pozisyon")
        add_row("Tansiyon Yay Yüksekliği", "tan_yay_yukseklik", "Tansiyon Yay Konumu", "tan_yay_konumu")
        add_row("Yay Boğumu", "tan_yay_bogumu", "Zemin Ağızlık", "zem_agizlik")

        # Motor rampaları ve armür raporları için çerçeveli alanlar
        ramp_box = QGroupBox("Motor Rampaları")
        ramp_grid = QGridLayout(ramp_box)
        for idx in range(1, 7):
            lbl = QLabel(f"Rampa {idx}")
            edit = QLineEdit()
            key = f"rampa_{idx}"
            edit.setObjectName(key)
            self._fields[key] = edit
            r = (idx - 1) // 3
            c = (idx - 1) % 3
            ramp_grid.addWidget(lbl, r, c * 2)
            ramp_grid.addWidget(edit, r, c * 2 + 1)
        grid.addWidget(ramp_box, row, 0, 2, 2)

        rapor_box = QGroupBox("Armür Deseni / Rapor")
        rapor_grid = QGridLayout(rapor_box)
        rapor_grid.addWidget(QLabel("Armür Deseni"), 0, 0)
        rapor_grid.addWidget(self._add_field("arm_desen"), 0, 1)
        rapor_grid.addWidget(QLabel("Armür Raporu 1"), 1, 0)
        rapor_grid.addWidget(self._add_field("arm_rap_1"), 1, 1)
        rapor_grid.addWidget(QLabel("Armür Raporu 2"), 2, 0)
        rapor_grid.addWidget(self._add_field("arm_rap_2"), 2, 1)
        rapor_grid.addWidget(QLabel("Armür Raporu 3"), 3, 0)
        rapor_grid.addWidget(self._add_field("arm_rap_3"), 3, 1)
        rapor_grid.addWidget(QLabel("Armür Raporu 4"), 4, 0)
        rapor_grid.addWidget(self._add_field("arm_rap_4"), 4, 1)
        grid.addWidget(rapor_box, row, 2, 2, 2)
        row += 2

        return box

    def _add_field(self, key: str) -> QLineEdit:
        edit = QLineEdit()
        edit.setObjectName(key)
        self._fields[key] = edit
        return edit

    def _build_footer_box(self) -> QGroupBox:
        box = QGroupBox("Notlar / Yetki")
        grid = QGridLayout(box)

        grid.addWidget(QLabel("Açıklama"), 0, 0)
        grid.addWidget(self._add_field("aciklama"), 0, 1)
        grid.addWidget(QLabel("Değişiklik Yapan"), 1, 0)
        grid.addWidget(self._add_field("degisiklik_yapan"), 1, 1)

        self.btn_save_manual = QPushButton("Manuel Ayarı Kaydet")
        self.btn_save_manual.clicked.connect(self._on_manual_save)
        self.btn_save_manual.setStyleSheet("background:#f7d7a6;font-weight:bold;")
        grid.addWidget(self.btn_save_manual, 0, 2, 2, 1)

        return box

    # ------------------------------------------------------------------
    # LOGİK: AYAR ÇEKME
    # ------------------------------------------------------------------
    def _on_fetch_clicked(self):
        tip_raw = (self.ed_tip.text() or "").strip()
        if not tip_raw:
            QMessageBox.warning(self, "Uyarı", "Önce bir tip kodu girin.")
            return

        tip = tip_raw.upper()

        # Dinamik rapordaki verileri başlık alanlarına taşı
        self._populate_from_dynamic(tip)

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

        for key, widget in self._dynamic_fields.items():
            val = settings.get(key)
            if val is None:
                continue
            widget.setText(str(val))

            # Kullanıcıya kısa bir bilgi mesajı
        QMessageBox.information(
            self,
            "Tamam",
            f"{tip} tipi için otomatik ITEMA ayarları getirildi."
        )

        # ------------------------------------------------------------------
        # Dinamik rapordan başlık bilgilerini doldurma
        # ------------------------------------------------------------------
    def _populate_from_dynamic(self, tip: str) -> None:
            parent = self.parent()
            df: Optional[pd.DataFrame] = getattr(parent, "df_dinamik_full", None)
            if df is None or df.empty:
                return

            norm_tip = tip.strip().upper()

            candidates = []
            for col in ["Mamul Tip Kodu", "Kök Tip Kodu", "Tip", "Tip Kodu"]:
                if col in df.columns:
                    series = df[col].astype(str).str.strip().str.upper()
                    match = df[series == norm_tip]
                    if not match.empty:
                        candidates.append(match)
            if not candidates:
                return

            row = candidates[0].iloc[0]
            mapping = {
                "tarak": ["Tarak Grubu", "Tarak", "Tarak No"],
                "zemin_orgusu": ["Zemin Örgü", "Zemin"],
                "kenar_orgusu": ["Süs Kenar", "Kenar Örgüsü"],
                "sus_kenar_dis": ["Süs Kenar Diş Sayısı", "Süs Kenar Diş"],
                "dokunabilirlik": ["Dokunabilirlik"],
                "cozgu_kodu": ["Çözgü İpliği 1", "Çözgü Kodu"],
                "boya_kodu": ["İhzarat Boya Kodu", "Boya Kodu"],
                "cerceve_adedi": ["Çerçeve Adedi"],
                "kenar_adedi": ["Kenar Adedi"],
            }

            for key, cols in mapping.items():
                if key not in self._dynamic_fields:
                    continue
                value = None
                for col in cols:
                    if col in row.index and pd.notna(row[col]):
                        value = row[col]
                        break
                if value is not None:
                    self._dynamic_fields[key].setText(str(value))

            if "tip" in self._dynamic_fields:
                self._dynamic_fields["tip"].setText(tip)

        # ------------------------------------------------------------------
        # Manuel kayıt & çıktı
        # ------------------------------------------------------------------
    def _on_manual_save(self):
            pwd, ok = QInputDialog.getText(
                self,
                "Manuel Ayar Yetkisi",
                "Lütfen yetki şifresini girin:",
                echo=QLineEdit.Password,
            )
            if not ok:
                return
            if pwd != self._manual_password:
                QMessageBox.warning(self, "Yetki", "Geçersiz şifre.")
                return

            tip = self._dynamic_fields.get("tip")
            tip_val = tip.text().strip().upper() if tip else self.ed_tip.text().strip().upper()
            if not tip_val:
                QMessageBox.warning(self, "Uyarı", "Kaydetmek için bir tip kodu girin.")
                return

            try:
                conn = get_sql_connection()
            except Exception as e:
                QMessageBox.critical(self, "Bağlantı", f"SQL bağlantısı açılamadı:\n{e}")
                return

            data = {**{k: f.text() for k, f in self._fields.items()},
                    **{k: f.text() for k, f in self._dynamic_fields.items()}}
            data["tip"] = tip_val

            try:
                self._save_manual_settings(conn, data)
                QMessageBox.information(self, "Tamam", f"{tip_val} tipi için manuel ayar güncellendi.")
            except Exception as e:
                QMessageBox.critical(self, "Hata", f"Kayıt sırasında hata oluştu:\n{e}")
            finally:
                try:
                    conn.close()
                except Exception:
                    pass

    def _save_manual_settings(self, conn: pyodbc.Connection, values: Dict[str, str]) -> None:
            cols = [c for c in ITEMA_COLUMNS if c != "sira_no"]
            placeholders = ", ".join([f"[{c}]" for c in cols])
            params = [values.get(c) or None for c in cols]

            # Basit upsert: varsa güncelle yoksa ekle
            sql = f"""
                IF EXISTS (SELECT 1 FROM dbo.ItemaTipArsiv WHERE Tip = ?)
                    UPDATE dbo.ItemaTipArsiv SET {', '.join([f'[{c}] = ?' for c in cols if c != 'tip'])} WHERE Tip = ?
                ELSE
                    INSERT INTO dbo.ItemaTipArsiv ({placeholders}) VALUES ({', '.join(['?'] * len(cols))});
                """
            # Param sırası: kontrol tip, update kolonları, where tip, insert kolonları
            update_params = [values.get(c) or None for c in cols if c != "tip"]
            exec_params = [values.get("tip"), *update_params, values.get("tip"), *params]
            cur = conn.cursor()
            cur.execute(sql, exec_params)
            conn.commit()



    def _print_form(self):
            printer = QPrinter(QPrinter.HighResolution)
            printer.setPageSize(QPrinter.A4)
            dialog = QPrintDialog(printer, self)
            if dialog.exec() != QPrintDialog.Accepted:
                return

            painter = QPainter(printer)
            painter.setRenderHint(QPainter.Antialiasing)
            self.render(painter)
            painter.end()
