from __future__ import annotations
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox,
    QLineEdit, QPushButton, QMessageBox, QTableWidget, QTableWidgetItem
)
from PySide6.QtCore import Qt
import pandas as pd
from datetime import datetime
from copy import deepcopy

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
    Üstte mevcut kurallar listesi (okunur),
    altta tek bir kural ekleme/güncelleme formu.
    Ekle / Güncelle / Sil / Kaydet destekler.
    """

    def __init__(self, df: pd.DataFrame, rules: list[dict], parent=None):
        super().__init__(parent)
        self.setWindowTitle("NOTLAR • Kural Yönetimi")
        self.resize(780, 520)

        self._df = df
        # Orijinal listeyi bozmamak için kopya üzerinde çalış
        self._rules: list[dict] = deepcopy(rules or [])
        self._result_rules: list[dict] | None = None

        v = QVBoxLayout(self)

        # --- 1) Mevcut kurallar tablosu ---
        self.tbl = QTableWidget(0, 5)
        self.tbl.setHorizontalHeaderLabels(
            ["Sütun", "Değer", "Açıklama", "Kullanıcı", "Tarih/Saat"]
        )
        self.tbl.horizontalHeader().setStretchLastSection(True)
        self.tbl.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tbl.setSelectionBehavior(QTableWidget.SelectRows)
        self.tbl.setSelectionMode(QTableWidget.SingleSelection)

        v.addWidget(QLabel("Kayıtlı Not Kuralları:"))
        v.addWidget(self.tbl, 1)

        # --- 2) Ekleme / düzenleme formu ---
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
        present_cols = [
            c for c in ALLOWED_COLS_ORDER
            if c in (df.columns if df is not None else [])
        ]
        self.cmb_col.addItems(present_cols)
        self.cmb_col.currentTextChanged.connect(self._refresh_values)
        row1.addWidget(self.cmb_col, 1)
        form.addLayout(row1)

        # Kriter değeri
        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Kriter değeri:"), 0)
        self.cmb_val = QComboBox()
        self.cmb_val.setEditable(True)
        row2.addWidget(self.cmb_val, 1)
        form.addLayout(row2)

        # Not / Açıklama
        row3 = QHBoxLayout()
        row3.addWidget(QLabel("Açıklama (NOTLAR'a eklenecek):"), 0)
        self.ed_note = QLineEdit()
        self.ed_note.setPlaceholderText("örn. Akşam vardiyası öncelik...")
        row3.addWidget(self.ed_note, 1)
        form.addLayout(row3)

        # Butonlar
        row4 = QHBoxLayout()
        self.btn_add = QPushButton("Ekle")
        self.btn_add.clicked.connect(self._add_rule)

        self.btn_update = QPushButton("Güncelle")
        self.btn_update.clicked.connect(self._update_rule)

        self.btn_delete = QPushButton("Sil")
        self.btn_delete.clicked.connect(self._delete_rule)

        row4.addWidget(self.btn_add)
        row4.addWidget(self.btn_update)
        row4.addWidget(self.btn_delete)
        row4.addStretch(1)

        self.btn_cancel = QPushButton("Kapat")
        self.btn_cancel.clicked.connect(self.reject)
        self.btn_save = QPushButton("Kaydet")
        self.btn_save.clicked.connect(self._on_save)

        row4.addWidget(self.btn_cancel)
        row4.addWidget(self.btn_save)

        form.addLayout(row4)

        v.addLayout(form)

        # İlk doldurma
        self._refresh_values()
        self._fill_table()

        # Satır seçimi değişince formu güncelle
        sel_model = self.tbl.selectionModel()
        if sel_model is not None:
            sel_model.selectionChanged.connect(self._sync_form_with_selection)

    # ---------- Yardımcı metodlar ----------

    def _fill_table(self):
        """Mevcut _rules listesini tabloya bas."""
        self.tbl.setRowCount(0)
        for r in self._rules:
            row = self.tbl.rowCount()
            self.tbl.insertRow(row)
            self.tbl.setItem(row, 0, QTableWidgetItem(str(r.get("col", ""))))
            self.tbl.setItem(row, 1, QTableWidgetItem(str(r.get("val", ""))))
            self.tbl.setItem(row, 2, QTableWidgetItem(str(r.get("text", ""))))
            self.tbl.setItem(row, 3, QTableWidgetItem(str(r.get("user", ""))))
            self.tbl.setItem(row, 4, QTableWidgetItem(str(r.get("created_at", ""))))
        self.tbl.resizeColumnsToContents()

    def _refresh_values(self):
        """Seçili sütuna göre kriter değeri combobox'ını doldur."""
        col = self.cmb_col.currentText().strip()
        self.cmb_val.clear()

        if (
            not col
            or self._df is None
            or self._df.empty
            or col not in self._df.columns
        ):
            return

        uniq = (
            self._df[col]
            .astype(str)
            .fillna("")
            .drop_duplicates()
            .sort_values(key=lambda s: s.str.lower())
            .tolist()
        )
        self.cmb_val.addItems(uniq)

    def _validate_inputs(self) -> tuple[str, str, str, str] | None:
        """Formdaki alanları kontrol eder; sorun yoksa (user, col, val, note) döner."""
        user = self.ed_user.text().strip() or "Anonim"
        col = self.cmb_col.currentText().strip()
        val = self.cmb_val.currentText().strip()
        note = self.ed_note.text().strip()

        if not col:
            QMessageBox.warning(self, "Uyarı", "Kriter sütununu seçin.")
            return None
        if val == "":
            QMessageBox.warning(self, "Uyarı", "Kriter değerini girin/seçin.")
            return None
        if not note:
            QMessageBox.warning(self, "Uyarı", "Açıklama (not) girin.")
            return None

        # Kullanıcı varsayılanını kalıcılaştır
        storage.set_username_default(user)

        return user, col, val, note

    def _selected_row(self) -> int | None:
        sel = self.tbl.selectionModel()
        if not sel:
            return None
        rows = sel.selectedRows()
        if not rows:
            return None
        return rows[0].row()

    def _sync_form_with_selection(self):
        """Tablodan satır seçilince form alanlarını o satırla doldur."""
        row = self._selected_row()
        if row is None or not (0 <= row < len(self._rules)):
            return

        r = self._rules[row]
        self.ed_user.setText(str(r.get("user", "")))

        col = str(r.get("col", ""))
        idx = self.cmb_col.findText(col)
        if idx >= 0:
            self.cmb_col.setCurrentIndex(idx)

        self.cmb_val.setCurrentText(str(r.get("val", "")))
        self.ed_note.setText(str(r.get("text", "")))

    # ---------- Buton aksiyonları ----------

    def _add_rule(self):
        validated = self._validate_inputs()
        if not validated:
            return
        user, col, val, note = validated

        rule = {
            "col": col,
            "val": val,
            "text": note,
            "user": user,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        self._rules.append(rule)
        self._fill_table()
        # Yeni eklenen satıra seçimi getir
        self.tbl.selectRow(self.tbl.rowCount() - 1)

    def _update_rule(self):
        validated = self._validate_inputs()
        if not validated:
            return
        user, col, val, note = validated

        row = self._selected_row()
        if row is None:
            QMessageBox.information(self, "Bilgi", "Güncellenecek kuralı tablodan seçin.")
            return

        existing = self._rules[row] if 0 <= row < len(self._rules) else {}
        created_at = existing.get("created_at") or datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        self._rules[row] = {
            "col": col,
            "val": val,
            "text": note,
            "user": user,
            "created_at": created_at,
        }
        self._fill_table()
        self.tbl.selectRow(row)

    def _delete_rule(self):
        row = self._selected_row()
        if row is None:
            QMessageBox.information(self, "Bilgi", "Silmek için bir kural seçin.")
            return

        if not (0 <= row < len(self._rules)):
            return

        yanit = QMessageBox.question(
            self,
            "Onay",
            "Bu kuralı silmek istediğinizden emin misiniz?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if yanit != QMessageBox.Yes:
            return

        self._rules.pop(row)
        self._fill_table()
        if self.tbl.rowCount():
            self.tbl.selectRow(min(row, self.tbl.rowCount() - 1))

    def _on_save(self):
        """Kaydet: tüm kural listesini döndür ve dialogu kapat."""
        self._result_rules = deepcopy(self._rules)
        self.accept()

    # ---------- Dışarıya sonuç ----------

    def result_rules(self) -> list[dict] | None:
        """Kaydedilen (güncel) kural listesini döndürür. Hiç Kaydet denmemişse None."""
        return self._result_rules
