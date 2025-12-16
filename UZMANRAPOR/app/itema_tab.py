from __future__ import annotations

import os
from typing import Dict, Optional

import pyodbc
import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtGui import QPageSize
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

from app.itema_settings import build_itema_settings, ITEMA_COLUMNS


# Burayı kendi ortamına göre AYARLAMALISIN:
SQL_SERVER = "10.30.9.14,1433"   # Örn: "STIBRSSFSRV01" veya "(local)"
SQL_DATABASE = "UzmanRaporDB"    # Bizim kurduğumuz veritabanı adı


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
        self._last_tip_features: Dict[str, Optional[str]] = {}
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

        def add(label: str, key: str, row: int, col: int, dynamic: bool = True):
            lbl = QLabel(label)
            lbl.setStyleSheet("font-weight: bold; color: #0A5097;")
            edit = QLineEdit()
            edit.setObjectName(key)
            layout.addWidget(lbl, row, col)
            layout.addWidget(edit, row, col + 1)
            (self._dynamic_fields if dynamic else self._fields)[key] = edit

        # --- Genel bilgiler (iki blok: sol 0-1, sağ 2-3) ---
        add("Tip Kodu", "tip", 0, 0)
        add("Kök Tip", "kok_tip", 0, 2)

        add("Tarak Grubu", "tarak_grubu", 1, 0)
        add("Atkı Sıklığı", "atki_sikligi", 1, 2)

        add("Zemin Örgü", "zemin_orgu", 2, 0)
        add("Kenar Örgü", "kenar_orgu", 2, 2)

        add("Süs Kenar", "sus_kenar", 3, 0)
        add("Dokunabilirlik", "dokunabilirlik", 3, 2)

        add("Çözgü Kodu", "cozgu_kodu", 4, 0)
        add("Boya Kodu", "boya_kodu", 4, 2)

        add("Çerçeve Adedi", "cerceve_adedi", 5, 0)
        add("Kenar Çerçeve", "kenar_cerceve", 5, 2)

        # --- ÇÖZGÜ (2x2 düzen) ---
        add("Çözgü 1", "cozgu1", 6, 0)
        add("Çözgü 3", "cozgu3", 6, 2)

        add("Çözgü 2", "cozgu2", 7, 0)
        add("Çözgü 4", "cozgu4", 7, 2)

        # --- ATKI + ATIM sağa hizalı ---
        add("Atkı 1", "atki1", 8, 0)
        add("Atkı1 Atım", "atki1_atim", 8, 2)

        add("Atkı 2", "atki2", 9, 0)
        add("Atkı2 Atım", "atki2_atim", 9, 2)

        add("Atkı 3", "atki3", 10, 0)
        add("Atkı 4", "atki4", 11, 0)

        return box

    def _build_body_box(self) -> QGroupBox:
        box = QGroupBox("Makine Ayarları")
        grid = QGridLayout(box)
        grid.setVerticalSpacing(6)
        grid.setHorizontalSpacing(12)

        row = 0

        def add_row(
            label_left: str,
            key_left: str,
            label_right: Optional[str] = None,
            key_right: Optional[str] = None,
            color: Optional[str] = None,
        ):
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
        add_row("Kapanma Dur 1", "kapanma_dur_1", "Oturma Düzeyi 1", "oturma_duzeyi_1")
        add_row("Kapanma Dur 2", "kapanma_dur_2", "Oturma Düzeyi 2", "oturma_duzeyi_2")

        # Motor rampaları (tek sıra)
        ramp_box = QGroupBox("Motor Rampaları")
        ramp_grid = QGridLayout(ramp_box)
        for idx in range(1, 7):
            lbl = QLabel(f"Rampa {idx}")
            edit = QLineEdit()
            key = f"rampa_{idx}"
            edit.setObjectName(key)
            self._fields[key] = edit
            r = 0
            c = idx - 1
            ramp_grid.addWidget(lbl, r, c * 2)
            ramp_grid.addWidget(edit, r, c * 2 + 1)

        grid.addWidget(ramp_box, row, 0, 1, 4)
        row += 1

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

    def _clear_form(self, keep_tip: Optional[str] = None) -> None:
        # Dinamik alanlar (header)
        for _, w in self._dynamic_fields.items():
            w.blockSignals(True)
            w.setText("")
            w.blockSignals(False)

        # Makine alanları + notlar
        for _, w in self._fields.items():
            w.blockSignals(True)
            w.setText("")
            w.blockSignals(False)

        # Tip kalsın istiyorsan
        if keep_tip:
            if "tip" in self._dynamic_fields:
                self._dynamic_fields["tip"].setText(keep_tip)
            self.ed_tip.setText(keep_tip)

        self._last_tip_features = {}

    # ------------------------------------------------------------------
    # LOGİK: AYAR ÇEKME
    # ------------------------------------------------------------------
    def _on_fetch_clicked(self):
        tip_raw = (self.ed_tip.text() or "").strip()
        if not tip_raw:
            QMessageBox.warning(self, "Uyarı", "Önce bir tip kodu girin.")
            return

        tip = tip_raw.upper()

        # >>> KRİTİK: HER FETCH'TE ÖNCE FORMU TEMİZLE <<<
        # Böylece yeni tipte bazı alanlar boş gelirse eski değer kalmaz.
        self._clear_form(keep_tip=tip)

        # Dinamik rapordaki verileri başlık alanlarına taşı
        tip_features = self._populate_from_dynamic(tip)

        # Dinamik raporda tip yoksa: form zaten temiz; uyar
        if not tip_features or len(tip_features.keys()) <= 1:  # sadece "tip" set edilmiş olabilir
            QMessageBox.information(
                self,
                "Bilgi",
                f"{tip} tipi dinamik raporda bulunamadı. Form temizlendi."
            )
            return

        try:
            conn = get_sql_connection()
        except Exception as e:
            QMessageBox.critical(self, "Bağlantı Hatası", f"SQL sunucusuna bağlanılamadı:\n\n{e}")
            return

        try:
            settings = build_itema_settings(conn, tip, tip_features)
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"ITEMA ayarları okunurken bir hata oluştu:\n\n{e}")
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
                f"{tip} tipi için ayar kağıdı bulunamadı. Form temizlendi."
            )
            return

        # Alanlara yaz (makine ayarları vb.)
        for key, widget in self._fields.items():
            widget.setText(settings.get(key) or "")

        # Başlık/dinamik alanlar (tip, kök tip vb.) - settings içinde varsa overwrite edebilir
        for key, widget in self._dynamic_fields.items():
            val = settings.get(key)
            if val is None:
                continue
            widget.setText(str(val))

        QMessageBox.information(self, "Tamam", f"{tip} tipi için otomatik ITEMA ayarları getirildi.")

    # ------------------------------------------------------------------
    # Dinamik rapordan başlık bilgilerini doldurma
    # ------------------------------------------------------------------
    def _populate_from_dynamic(self, tip: str) -> Dict[str, Optional[str]]:
        # df'yi doğru yerden al
        win = self.window()
        df: Optional[pd.DataFrame] = getattr(win, "df_dinamik_full", None)

        tip_features: Dict[str, Optional[str]] = {}

        if df is None or df.empty:
            return tip_features

        norm_tip = (tip or "").strip().upper()
        tip_features["tip"] = norm_tip
        if "tip" in self._dynamic_fields:
            self._dynamic_fields["tip"].setText(norm_tip)

        # Tip satırını bul
        candidates = []
        for col in ["Mamul Tip Kodu", "Tip Kodu", "Tip", "Kök Tip Kodu"]:
            if col in df.columns:
                s = df[col].astype(str).str.strip().str.upper()
                m = df[s == norm_tip]
                if not m.empty:
                    candidates.append(m)

        if not candidates:
            return tip_features

        row = candidates[0].iloc[0]

        def get_first(*cols: str) -> Optional[str]:
            for c in cols:
                if c in row.index and pd.notna(row[c]):
                    v = str(row[c]).strip()
                    if v and v.lower() != "nan":
                        return v
            return None

        def set_dyn(ui_key: str, value: Optional[str]):
            if value is None:
                return
            tip_features[ui_key] = value
            w = self._dynamic_fields.get(ui_key)
            if w is not None:
                w.setText(str(value))

        # Mapping
        set_dyn("kok_tip", get_first("Kök Tip Kodu", "Kok Tip Kodu", "KökTip"))
        set_dyn("tarak_grubu", get_first("Tarak Grubu", "Tarak"))
        set_dyn("zemin_orgu", get_first("Zemin Örgü", "Zemin Orgu"))
        set_dyn("kenar_orgu", get_first("Kenar Örgü", "Kenar Orgu"))
        set_dyn("sus_kenar", get_first("Süs Kenar", "Sus Kenar"))
        set_dyn("dokunabilirlik", get_first("Dokunabilirlik Oranı", "Dokunabilirlik"))
        set_dyn("atki_sikligi", get_first("7100", "Atkı Sıklığı", "Atki Sikligi"))
        set_dyn("cozgu_kodu", get_first("Çözgü Kodu", "Cozgu Kodu"))
        set_dyn("boya_kodu", get_first("İhzarat Boya Kodu", "Ihzarat Boya Kodu", "Boya Kodu"))
        set_dyn("cerceve_adedi", get_first("Çerçeve Adedi", "Cerceve Adedi"))
        set_dyn("kenar_cerceve", get_first("Kenar Adedi", "Kenar Çerçeve", "Kenar Cerceve"))

        def combine(no_col: str, yarn_col: str) -> Optional[str]:
            no = get_first(no_col)
            yarn = get_first(yarn_col)
            if not no and not yarn:
                return None
            if yarn and no and yarn.startswith(no):
                return yarn
            parts = [p for p in [no, yarn] if p]
            return " ".join(parts) if parts else None

        set_dyn("cozgu1", combine("Çözgü İplik No 1", "Çözgü İpliği 1"))
        set_dyn("cozgu2", combine("Çözgü İplik No 2", "Çözgü İpliği 2"))
        set_dyn("cozgu3", combine("Çözgü İplik No 3", "Çözgü İpliği 3"))
        set_dyn("cozgu4", combine("Çözgü İplik No 4", "Çözgü İpliği 4"))

        set_dyn("atki1", combine("Atkı İplik No 1", "Atkı İpliği 1"))
        set_dyn("atki2", combine("Atkı İplik No 2", "Atkı İpliği 2"))
        set_dyn("atki3", combine("Atkı İplik No 3", "Atkı İpliği 3"))
        set_dyn("atki4", combine("Atkı İplik No 4", "Atkı İpliği 4"))

        set_dyn("atki1_atim", get_first("Atkı Atma Adedi 1", "Atki Atma Adedi 1"))
        set_dyn("atki2_atim", get_first("Atkı Atma Adedi 2", "Atki Atma Adedi 2"))

        return tip_features

    # ------------------------------------------------------------------
    # MANUEL KAYIT
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

        tip_widget = self._dynamic_fields.get("tip")
        tip_val = tip_widget.text().strip().upper() if tip_widget else self.ed_tip.text().strip().upper()
        if not tip_val:
            QMessageBox.warning(self, "Uyarı", "Kaydetmek için bir tip kodu girin.")
            return

        try:
            conn = get_sql_connection()
        except Exception as e:
            QMessageBox.critical(self, "Bağlantı", f"SQL bağlantısı açılamadı:\n{e}")
            return

        data = {
            **{k: f.text() for k, f in self._fields.items()},
            **{k: f.text() for k, f in self._dynamic_fields.items()},
        }

        # header'daki Tarak Grubu -> SQL kolon adı "tarak"
        tg = self._dynamic_fields.get("tarak_grubu")
        if tg and tg.text().strip():
            data["tarak"] = tg.text().strip()

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
        table = "dbo.ItemaAyar"

        # sira_no identity => asla insert/update listesine koyma
        cols = [c for c in ITEMA_COLUMNS if c.lower() != "sira_no"]

        tip_val = (values.get("tip") or "").strip().upper()
        if not tip_val:
            raise ValueError("Tip boş olamaz.")

        set_cols = [c for c in cols if c.lower() != "tip"]

        update_set_sql = ", ".join([f"[{c}] = ?" for c in set_cols])
        insert_cols_sql = ", ".join([f"[{c}]" for c in cols])
        insert_placeholders = ", ".join(["?"] * len(cols))

        update_params = [(values.get(c) or None) for c in set_cols]
        insert_params = [(values.get(c) or None) for c in cols]

        sql = f"""
        IF EXISTS (SELECT 1 FROM {table} WHERE [tip] = ?)
        BEGIN
            UPDATE {table}
            SET {update_set_sql}
            WHERE [tip] = ?;
        END
        ELSE
        BEGIN
            INSERT INTO {table} ({insert_cols_sql})
            VALUES ({insert_placeholders});
        END
        """

        exec_params = [tip_val, *update_params, tip_val, *insert_params]

        cur = conn.cursor()
        cur.execute(sql, exec_params)
        conn.commit()

    # ------------------------------------------------------------------
    # A4 ÇIKTI
    # ------------------------------------------------------------------
    def _print_form(self):
        QMessageBox.information(
            self,
            "Çıktı",
            "Yazdırma/önizleme kısmını daha sonra birlikte düzelteceğiz."
        )
