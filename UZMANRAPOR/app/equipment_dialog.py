from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QTableWidget, QTableWidgetItem,
    QComboBox, QPushButton, QMessageBox, QSpinBox
)
from PySide6.QtCore import Qt
import traceback

from app.storage import load_loom_cut_map, save_loom_cut_map


def _norm_choice(val: str) -> str:
    """Kayıt ve görüntü için tek tipe normalizasyon."""
    u = (val or "").strip().upper()
    if u == "ISAVER":
        return "ISAVER"
    if u == "ROTOCUT":
        return "ROTOCUT"
    if u in ("ISAVERKIT", "ISAVER KIT", "KIT"):
        return "ISAVERKit"
    return ""

class LoomCutEditor(QDialog):
    """
    2201–2518 tezgahları için Kesim Tipi (ISAVER / ROTOCUT / ISAVERKit) düzenleme.
    Veriler SQL tablosu LoomCutMap içinde saklanır.
    """
    def __init__(self, parent=None, start_loom=2201, end_loom=2518):
        super().__init__(parent)
        self.setWindowTitle("Kesim Tipi (ISAVER / ROTOCUT / ISAVERKit) Düzenle")
        self.resize(560, 680)

        self._start = int(start_loom)
        self._end = int(end_loom)
        self._data = load_loom_cut_map() or {}  # {"2201":"ISAVER", ...}

        v = QVBoxLayout(self)

        # Aralık kontrolü
        aralik = QHBoxLayout()
        aralik.addWidget(QLabel("Başlangıç:"))
        self.sp_from = QSpinBox(); self.sp_from.setRange(2000, 9999); self.sp_from.setValue(self._start)
        aralik.addWidget(self.sp_from)
        aralik.addWidget(QLabel("Bitiş:"))
        self.sp_to = QSpinBox(); self.sp_to.setRange(2000, 9999); self.sp_to.setValue(self._end)
        aralik.addWidget(self.sp_to)
        btn_refresh = QPushButton("Listeyi Yenile")
        aralik.addWidget(btn_refresh); aralik.addStretch(1)
        v.addLayout(aralik)

        # Tablo
        self.tbl = QTableWidget(0, 2)
        self.tbl.setHorizontalHeaderLabels(["Tezgah", "Kesim Tipi"])
        self.tbl.horizontalHeader().setStretchLastSection(True)
        v.addWidget(self.tbl, 1)

        # Alt butonlar
        btns = QHBoxLayout()
        btn_save = QPushButton("Kaydet")
        btn_close = QPushButton("Kapat")
        btns.addStretch(1); btns.addWidget(btn_save); btns.addWidget(btn_close)
        v.addLayout(btns)

        # Bağlantılar
        btn_refresh.clicked.connect(self._fill)
        btn_save.clicked.connect(self._save)
        btn_close.clicked.connect(self.reject)

        # Başlangıç verisi
        self._fill()

    def _fill(self):
        try:
            self._start = int(self.sp_from.value())
            self._end   = int(self.sp_to.value())
            if self._start > self._end:
                self._start, self._end = self._end, self._start

            self.tbl.setRowCount(0)
            for nz in range(self._start, self._end + 1):
                r = self.tbl.rowCount()
                self.tbl.insertRow(r)

                self.tbl.setItem(r, 0, QTableWidgetItem(str(nz)))

                cmb = QComboBox()
                # >>> 3 seçenek
                cmb.addItems(["", "ISAVER", "ROTOCUT", "ISAVERKit"])

                # kayıttaki değeri normalize ederek seç
                raw = self._data.get(str(nz), "")
                val = _norm_choice(raw)
                ix = cmb.findText(val, Qt.MatchFixedString)
                cmb.setCurrentIndex(ix if ix >= 0 else 0)

                self.tbl.setCellWidget(r, 1, cmb)
        except Exception:
            QMessageBox.critical(self, "Hata (Listeyi Yenile)",
                                 "Liste doldurulurken bir hata oluştu:\n\n" + traceback.format_exc())

    def _save(self):
        try:
            out = {}
            for r in range(self.tbl.rowCount()):
                tz_item = self.tbl.item(r, 0)
                cmb = self.tbl.cellWidget(r, 1)
                tz = tz_item.text().strip() if tz_item else ""
                typ = _norm_choice(cmb.currentText() if cmb else "")
                if tz and typ:
                    out[tz] = typ

            save_loom_cut_map(out)
            self._data = out  # hafızayı da güncelle
            QMessageBox.information(self, "Kaydedildi", "Kesim Tipi listesi kaydedildi.")
        except Exception:
            QMessageBox.critical(self, "Hata (Kaydet)",
                                 "Kayıt sırasında bir hata oluştu:\n\n" + traceback.format_exc())
