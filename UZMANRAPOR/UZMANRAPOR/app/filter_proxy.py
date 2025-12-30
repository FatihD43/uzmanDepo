from PySide6.QtCore import QSortFilterProxyModel, Qt

class MultiColumnFilterProxy(QSortFilterProxyModel):
    """
    - Metin alt filtreleri: self._filters = {col: "abc"}
    - Çoklu seçim (inclusion) filtreleri: self._inclusions = {col: set([...])}
      Not: Inclusion set boşsa o kolon için kısıtlama uygulanmaz.
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self._filters = {}
        self._inclusions = {}
        self.setFilterCaseSensitivity(Qt.CaseInsensitive)

    # --- Text filter API ---
    def setFilterForColumn(self, col: int, text: str):
        self._filters[col] = text or ""
        self._filters = {c: t for c, t in self._filters.items() if t != ""}
        self.invalidateFilter()

    def clearFilters(self):
        self._filters.clear()
        self.invalidateFilter()

    # --- Inclusion (multi-select) API ---
    def setInclusionForColumn(self, col: int, values: set):
        """values: set of strings (display data). Boş set => filtre kapalı."""
        self._inclusions[col] = set(values) if values else set()
        # temizle: tamamen boş olanları tutmaya gerek yok
        self._inclusions = {c: s for c, s in self._inclusions.items() if len(s) > 0}
        self.invalidateFilter()

    def clearInclusions(self):
        self._inclusions.clear()
        self.invalidateFilter()

    def filterAcceptsRow(self, source_row, source_parent):
        model = self.sourceModel()

        # 1) inclusion (çoklu seçim) kontrolleri
        for col, selected in self._inclusions.items():
            idx = model.index(source_row, col, source_parent)
            val = model.data(idx) or ""
            # yalnızca seçilenlerin içinden gelmeli
            if str(val) not in selected:
                return False

        # 2) text (alt string) kontrolleri
        for col, pattern in self._filters.items():
            idx = model.index(source_row, col, source_parent)
            val = model.data(idx) or ""
            if pattern.lower() not in str(val).lower():
                return False

        return True
