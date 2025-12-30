from PySide6.QtCore import QAbstractTableModel, Qt, QModelIndex
from PySide6.QtGui import QColor
import math
import pandas as pd

# Alias ve orijinal adlarla birlikte INT-benzeri alanlar
INT_LIKE_COLS = {
    "Levent No", "Levent Etiket FA",
    "Tezgah Numarası", "Tezgah", "Tezgah No",
    "Üretim Sipariş No", "Haşıl İş Emri",
}

# Küsuratlı göstermemiz gereken alanlar (orijinal + UI alias’ı)
DECIMAL_FORCE_COLS = {
    "Parti Metresi", "_KalanMetre",
    "Atkı İhtiyaç Miktar 1", "Atkı İhtiyaç Miktar 2",
    "(Atkı-1 İşletme Depoları + Atkı-1 İşletme Diğer Depoları)",
    "(Atkı-2 İşletme Depoları + Atkı-2 İşletme Diğer Depoları)",
}

class PandasModel(QAbstractTableModel):
    def __init__(self, df, highlight_assigned=False, parent=None):
        super().__init__(parent)
        self._df = df
        self._highlight = highlight_assigned
        self._header_overrides: dict[int, str] = {}


    def rowCount(self, parent=QModelIndex()):
        return 0 if self._df is None else len(self._df)

    def columnCount(self, parent=QModelIndex()):
        return 0 if self._df is None else len(self._df.columns)

    # --- Hücre biçimlendirme (sütun-isim bilinçli) ---
    def _format_cell(self, value, col_name: str):
        # Boş/NA ise boş
        try:
            if value is None or (isinstance(value, float) and math.isnan(value)):
                return ""
            if pd.isna(value):
                return ""
        except Exception:
            pass

        # 1) INT-benzeri kolonlar: 123.0 -> 123
        if col_name in INT_LIKE_COLS:
            # Sayıya çevrilebiliyorsa tamsayı gibi göster
            try:
                f = float(value)
                if math.isfinite(f) and abs(f - round(f)) < 1e-9:
                    return str(int(round(f)))
                # Nadiren 123.45 gelebilir; yine de gereksiz sıfırları at
                return f"{f:.6f}".rstrip("0").rstrip(".")
            except Exception:
                return str(value)

        # 2) Küsurat zorunlu kolonlar: 2 ondalık sabit
        if col_name in DECIMAL_FORCE_COLS:
            try:
                f = float(value)
                if not math.isfinite(f):
                    return ""
                # En az 1 ve en çok 2 ondalık gösterelim (sabit 2 daha net):
                return f"{f:.2f}"
            except Exception:
                return str(value)

        # 3) Diğer kolonlarda akıllı gösterim:
        #    - tamsayı ise int gibi
        #    - değilse gereksiz 0’ları traşla (küsurat varsa koru)
        if isinstance(value, (int,)):
            return str(value)
        if isinstance(value, float):
            if math.isfinite(value) and abs(value - round(value)) < 1e-9:
                return str(int(round(value)))
            return f"{value:.6f}".rstrip("0").rstrip(".")

        # String sayı gibi görünüyorsa
        if isinstance(value, str):
            s = value.strip()
            try:
                f = float(s)
                if math.isfinite(f) and abs(f - round(f)) < 1e-9:
                    return str(int(round(f)))
                return f"{f:.6f}".rstrip("0").rstrip(".")
            except Exception:
                return value

        return str(value)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid() or self._df is None:
            return None

        value = self._df.iat[index.row(), index.column()]
        col_name = str(self._df.columns[index.column()])

        if role in (Qt.DisplayRole, Qt.EditRole):
            return self._format_cell(value, col_name)

        # --- Atanmış satırları sarı boya (alias destekli) ---
        if role == Qt.BackgroundRole and self._highlight:
            try:
                cols = list(self._df.columns)
                loom_candidates = ("Tezgah Numarası", "Tezgah", "Tezgah No")
                loom_col_name = next((c for c in loom_candidates if c in cols), None)
                if loom_col_name is not None:
                    v = str(self._df.iat[index.row(), cols.index(loom_col_name)]).strip()
                    if v:
                        return QColor("#FFF59E")
            except Exception:
                pass
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole or self._df is None:
            return None

        if orientation == Qt.Horizontal:
            # >>> YENİ: Override varsa onu göster (ör. \n ile çok satırlı)
            try:
                if isinstance(self._header_overrides, dict) and section in self._header_overrides:
                    return self._header_overrides[section]
            except Exception:
                pass
            return str(self._df.columns[section])

        return str(section + 1)

    def set_df(self, df):
        self.beginResetModel()
        self._df = df
        # >>> Yeni DF geldiğinde eski wrap'leri taşımayalım
        try:
            self._header_overrides = {}
        except Exception:
            pass
        self.endResetModel()

    def notify_rows(self, row_indices: list[int]) -> None:
        if not row_indices or self._df is None:
            return
        last_col = self.columnCount() - 1
        for r in row_indices:
            if 0 <= r < self.rowCount():
                tl = self.index(r, 0)
                br = self.index(r, last_col)
                self.dataChanged.emit(tl, br, [Qt.DisplayRole, Qt.BackgroundRole])

    def notify_all(self) -> None:
        if self._df is None or self.rowCount() == 0:
            return
        tl = self.index(0, 0)
        br = self.index(self.rowCount() - 1, self.columnCount() - 1)
        self.dataChanged.emit(tl, br, [Qt.DisplayRole, Qt.BackgroundRole])
    def set_header_override(self, section: int, text: str) -> None:
        if not hasattr(self, "_header_overrides") or self._header_overrides is None:
            self._header_overrides = {}
        self._header_overrides[int(section)] = str(text)

    def clear_header_overrides(self) -> None:
        self._header_overrides = {}
