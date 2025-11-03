from __future__ import annotations
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox,
    QLineEdit, QPushButton, QWidget, QMessageBox, QTableWidget, QTableWidgetItem
)
from PySide6.QtCore import Qt
import pandas as pd
from datetime import datetime
from app import storage

ALLOWED_COLS_ORDER = [
    "Kök Tip Kodu",
    "Atkı İpliği 1",
    "Atkı İpliği 2",
    "Levent No",
    "Üretim Sipariş No",   # Dokuma İş Emri
    "Haşıl İş Emri",
    "Tarak Grubu",
    "Mamul Tip Kodu",
]

class NotesDialog(QDialog):
    """
    Üstte mevcut kurallar listesi (okunur), altta yeni kural ekleme formu.
    """
    def __init__(self, df: pd.DataFrame, rules: list[dict], parent=None):
        super().__init__(parent)
        self.setWindowTitle("NOTLAR • Kural Yönetimi")
        self.resize(780, 520)
        self._df = df
        self._rules = rules or []

        v = QVBoxLayout(self)

        # --- 1) Mevcut kurallar tablosu ---
        self.tbl = QTableWidget(0, 5)
        self.tbl.setHorizontalHeaderLabels(["Sütun", "Değer", "Açıklama", "Kullanıcı", "Tarih/Saat"])
        self.tbl.horizontalHeader().setStretchLastSection(True)
        self.tbl.setEditTriggers(QTableWidget.NoEditTriggers)
        v.addWidget(QLabel("Kayıtlı Not Kuralları:"))
        v.addWidget(self.tbl, 1)

        # --- 2) Ekleme formu ---
        form = QVBoxLayout()

        # Kullanıcı adı
        row0 = QHBoxLayout()
        row0.addWidget(QLabel("Kullanıcı:"), 0)
        self.ed_user = QLineEdit(storage.get_username_default())
        row0.addWidget(self.ed_user, 1)
        form.addLayout(row0)

        # Kriter sütunu
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("Kriter sütunu:"), 0)
        self.cmb_col = QComboBox()
        present_cols = [c for c in ALLOWED_COLS_ORDER if c in (df.columns if df is not None else [])]
        self.cmb_col.addItems(present_cols)
        self.cmb_col.currentTextChanged.connect(self._refresh_values)
        row1.addWidget(self.cmb_col, 1)
        form.addLayout(row1)

        # Kriter değeri
        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Kriter değeri:"), 0)
        self.cmb_val = QComboBox(); self.cmb_val.setEditable(True)
        row2.addWidget(self.cmb_val, 1)
        form.addLayout(row2)

        # Not / Açıklama
        row3 = QHBoxLayout()
        row3.addWidget(QLabel("Açıklama (NOTLAR'a eklenecek):"), 0)
        self.ed_note = QLineEdit(); self.ed_note.setPlaceholderText("örn. Akşam vardiyası öncelik...")
        row3.addWidget(self.ed_note, 1)
        form.addLayout(row3)

        # Butonlar
        row4 = QHBoxLayout()
        self.btn_cancel = QPushButton("Kapat"); self.btn_cancel.clicked.connect(self.reject)
        self.btn_ok = QPushButton("Ekle"); self.btn_ok.clicked.connect(self._on_ok)
        row4.addStretch(1); row4.addWidget(self.btn_cancel); row4.addWidget(self.btn_ok)
        form.addLayout(row4)

        v.addLayout(form)

        # İlk doldurma
        self._refresh_values()
        self._fill_table()

    # ---- mevcut kuralları tabloya bas ----
    def _fill_table(self):
        self.tbl.setRowCount(0)
        for r in self._rules:
            row = self.tbl.rowCount()
            self.tbl.insertRow(row)
            self.tbl.setItem(row, 0, QTableWidgetItem(str(r.get("col",""))))
            self.tbl.setItem(row, 1, QTableWidgetItem(str(r.get("val",""))))
            self.tbl.setItem(row, 2, QTableWidgetItem(str(r.get("text",""))))
            self.tbl.setItem(row, 3, QTableWidgetItem(str(r.get("user",""))))
            self.tbl.setItem(row, 4, QTableWidgetItem(str(r.get("created_at",""))))

    # ---- kriter değeri listesini yenile ----
    def _refresh_values(self):
        col = self.cmb_col.currentText().strip()
        self.cmb_val.clear()
        if not col or self._df is None or self._df.empty or col not in self._df.columns:
            return
        uniq = (
            self._df[col]
            .astype(str).fillna("")
            .drop_duplicates()
            .sort_values(key=lambda s: s.str.lower())
            .tolist()
        )
        self.cmb_val.addItems(uniq)

    # ---- ekle ----
    def _on_ok(self):
        user = self.ed_user.text().strip() or "Anonim"
        col = self.cmb_col.currentText().strip()
        val = self.cmb_val.currentText().strip()
        note = self.ed_note.text().strip()

        if not col:
            QMessageBox.warning(self, "Uyarı", "Kriter sütununu seçin."); return
        if val == "":
            QMessageBox.warning(self, "Uyarı", "Kriter değerini girin/seçin."); return
        if not note:
            QMessageBox.warning(self, "Uyarı", "Açıklama (not) girin."); return

        # kullanıcı varsayılanını kalıcılaştır
        storage.set_username_default(user)

        self._rule = {
            "col": col,
            "val": val,
            "text": note,
            "user": user,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        self.accept()

    def result_rule(self) -> dict | None:
        return getattr(self, "_rule", None)
