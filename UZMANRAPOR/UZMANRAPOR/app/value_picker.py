from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QListWidget, QListWidgetItem,
    QPushButton, QLabel, QWidget, QLineEdit
)
from PySide6.QtCore import Qt

class ValuePickerDialog(QDialog):
    """
    Kolonun unique değerlerini checkbox'lı listede gösterir.
    - values: [str] (tümü)
    - preselected: set(str) (önceden seçili olanlar) -> boş ise 'hepsi' anlamına gelir
    """
    def __init__(self, title: str, values, preselected=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(420, 520)

        # normalize
        self._values = ["" if v is None else str(v) for v in values]
        # aynıları çıkar, alfabetik sırala; boş en sona
        self._values = sorted(set(self._values), key=lambda x: (x == "", x))

        self._pre = set(preselected) if preselected else set()  # boşsa 'hepsi seçili' sayılacak

        v = QVBoxLayout(self)

        # Hızlı arama
        self.search = QLineEdit()
        self.search.setPlaceholderText("Ara...")
        v.addWidget(self.search)

        # Liste
        self.listw = QListWidget()
        self.listw.setSelectionMode(QListWidget.NoSelection)
        v.addWidget(self.listw, 1)

        # Alt butonlar
        btns = QHBoxLayout()
        self.btn_all = QPushButton("Hepsini Seç")
        self.btn_none = QPushButton("Temizle")
        self.btn_ok = QPushButton("Uygula")
        self.btn_cancel = QPushButton("Vazgeç")
        btns.addWidget(self.btn_all); btns.addWidget(self.btn_none)
        btns.addStretch(1); btns.addWidget(self.btn_cancel); btns.addWidget(self.btn_ok)
        v.addLayout(btns)

        # Sinyaller
        self.btn_all.clicked.connect(self._select_all)
        self.btn_none.clicked.connect(self._select_none)
        self.btn_ok.clicked.connect(self.accept)
        self.btn_cancel.clicked.connect(self.reject)
        self.search.textChanged.connect(self._refill)

        # İlk doldurma
        self._refill()

    # ---------- İç lojik ----------
    def _refill(self):
        pattern = self.search.text().strip().lower()
        self.listw.clear()
        # Görünen öğeleri, aramaya göre oluştur
        for val in self._values:
            if pattern and (pattern not in str(val).lower()):
                continue
            label = val if val != "" else "(boş)"
            item = QListWidgetItem(label)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            # preselected boşsa 'hepsi seçili' say
            checked = (len(self._pre) == 0) or (val in self._pre)
            item.setCheckState(Qt.Checked if checked else Qt.Unchecked)
            self.listw.addItem(item)

    def _select_all(self):
        # liste görünürünün hepsini işaretle
        for i in range(self.listw.count()):
            it = self.listw.item(i)
            it.setCheckState(Qt.Checked)

    def _select_none(self):
        # görünürlerden işareti kaldır
        for i in range(self.listw.count()):
            it = self.listw.item(i)
            it.setCheckState(Qt.Unchecked)

    def selected_values(self) -> set:
        """
        Dönüş: seçilen ham değerlerin set'i.
        Boş set => filtre kapalı anlamına gelir (hepsi görünür).
        """
        sel = set()
        for i in range(self.listw.count()):
            it = self.listw.item(i)
            if it.checkState() == Qt.Checked:
                label = it.text()
                val = "" if label == "(boş)" else label
                sel.add(val)

        # Arama alanı boşken ve tüm evren seçiliyse => filtreyi kapat (boş set döndür)
        # (Yani 'Hepsini Seç' -> filtre kapalı)
        if not self.search.text().strip():
            if len(sel) == len(self._values):
                return set()

        return sel
