from __future__ import annotations

import re
import textwrap
import pandas as pd

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog,
    QLabel, QTableView, QMessageBox, QScrollArea, QLineEdit, QSizePolicy,
    QHeaderView, QStyleOptionHeader, QStyle, QComboBox
)
from PySide6.QtCore import Qt, QTimer, QEvent, QRect
from PySide6.QtGui import QFont, QTextDocument


from app.models import PandasModel
from app.filter_proxy import MultiColumnFilterProxy
from app import storage


def _col_pick(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for c in candidates:
        if c in df.columns:
            return c
    return None


def _to_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce")


def _clean_col(c: object) -> str:
    """
    Excel kolon başlıklarında sık görülen:
      - NBSP (\u00a0)
      - satır sonu / tab
      - çoklu boşluk
    problemlerini normalize eder.
    """
    s = str(c) if c is not None else ""
    s = s.replace("\u00a0", " ")
    s = s.replace("\n", " ").replace("\r", " ").replace("\t", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


class WrapHeaderView(QHeaderView):
    """
    Bu sekmeye özel: Sütun başlıklarını 3 satıra kadar word-wrap ile çizer.
    QHeaderView'in '...' (elide) davranışını bypass eder.
    """
    def __init__(self, orientation, parent=None, max_lines: int = 3):
        super().__init__(orientation, parent)
        self._labels: list[str] = []
        self._max_lines = max_lines
        self.setDefaultAlignment(Qt.AlignCenter)
        self.setStretchLastSection(False)

        # elide istemiyoruz (paintSection zaten kendisi çizecek)
        try:
            self.setTextElideMode(Qt.ElideNone)
        except Exception:
            pass

    def set_labels(self, labels: list[str]) -> None:
        self._labels = [str(x) for x in (labels or [])]
        self.viewport().update()
        self.updateGeometry()

    def sizeHint(self):
        sh = super().sizeHint()
        fm = self.fontMetrics()
        sh.setHeight(int(fm.height() * self._max_lines + 14))
        return sh

    def paintSection(self, painter, rect, logicalIndex):
        if not rect.isValid():
            return

        opt = QStyleOptionHeader()
        self.initStyleOption(opt)
        opt.rect = rect
        opt.section = logicalIndex

        # Varsayılan header arka plan/çizgiler
        self.style().drawControl(QStyle.CE_Header, opt, painter, self)

        # Başlık metni (bizim label override varsa onu kullan)
        text = ""
        if 0 <= logicalIndex < len(self._labels):
            text = self._labels[logicalIndex]
        else:
            m = self.model()
            if m is not None:
                v = m.headerData(logicalIndex, self.orientation(), Qt.DisplayRole)
                text = "" if v is None else str(v)

        # Word-wrap çizimi
        doc = QTextDocument()
        doc.setDefaultFont(self.font())
        doc.setTextWidth(max(10, rect.width() - 8))
        doc.setPlainText(text)

        painter.save()
        painter.translate(rect.left() + 4, rect.top() + 4)
        clip = QRect(0, 0, max(10, rect.width() - 8), max(10, rect.height() - 8))
        painter.setClipRect(clip)
        doc.drawContents(painter)
        painter.restore()


class BuzulmeMetreUyumTab(QWidget):
    """
    Taslağa birebir:
    - Çıkış satır seviyesi: Etiket (veri kadar hücre)
    - İş emri bazlı hesaplar tek hücre mantığında, aynı iş emrinin tüm etiket satırlarında tekrar eder
    - Tip bazında dbo.TipBuzulmeModel join eder (Sistem/Geçmiş/Confidence)
    """

    def __init__(self, main_window):
        super().__init__(main_window)
        self.main = main_window
        self._last_path: str | None = None

        v = QVBoxLayout(self)

        # Üst bar
        top = QHBoxLayout()
        # --- Üst bar filtreleri (ComboBox) ---
        self.cmb_bolum = QComboBox()
        self.cmb_bolum.addItems([
            "HEPSİ",
            "İSKO14 (DK14)",
            "İSKO11 (DK11)",
            "MEKİKLİ (DK98)",
        ])
        self.cmb_bolum.currentTextChanged.connect(self._on_bolum_combo_changed)

        self.cmb_durum = QComboBox()
        self.cmb_durum.addItems([
            "HEPSİ",
            "Devam ediyor",
            "Bitmiş",
        ])
        self.cmb_durum.currentTextChanged.connect(self._on_durum_combo_changed)

        self.btn_load = QPushButton("ZPPR0308 Yükle")
        self.btn_load.clicked.connect(self.load_zppr0308)

        self.btn_refresh = QPushButton("Yenile")
        self.btn_refresh.clicked.connect(self.refresh_last)

        self.lbl_info = QLabel("")
        self.lbl_info.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        f = QFont()
        f.setBold(True)
        self.lbl_info.setFont(f)

        top.addWidget(self.btn_load)
        top.addWidget(self.btn_refresh)

        top.addSpacing(12)
        top.addWidget(QLabel("Bölüm:"))
        top.addWidget(self.cmb_bolum)

        top.addSpacing(12)
        top.addWidget(QLabel("Durum:"))
        top.addWidget(self.cmb_durum)

        top.addStretch(1)
        top.addWidget(self.lbl_info)

        # Tablo
        self.filter_scroll = QScrollArea()
        self.filter_scroll.setWidgetResizable(True)
        self.filter_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.filter_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.filter_bar = QWidget()
        self.filter_bar_layout = QHBoxLayout(self.filter_bar)
        self.filter_bar_layout.setContentsMargins(0, 0, 0, 0)
        self.filter_bar_layout.setSpacing(2)
        self.filter_scroll.setWidget(self.filter_bar)

        self.tbl = QTableView()
        self.tbl.setSortingEnabled(True)

        # >>> Sadece bu sekmede: wrap header
        self.wrap_header = WrapHeaderView(Qt.Horizontal, self.tbl, max_lines=3)
        self.tbl.setHorizontalHeader(self.wrap_header)

        self.model = PandasModel(pd.DataFrame())
        self.proxy = MultiColumnFilterProxy(self)
        self.proxy.setSourceModel(self.model)
        self.tbl.setModel(self.proxy)

        self._filter_edits: list[QLineEdit] = []
        self._scroll_from_table = False
        self._scroll_from_filter = False

        header = self.tbl.horizontalHeader()
        header.setStretchLastSection(True)
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.sectionResized.connect(self._sync_filter_widths)
        header.sectionResized.connect(lambda *_: self._schedule_header_rewrap())
        header.sectionCountChanged.connect(lambda *_: self._rebuild_filters())
        header.geometriesChanged.connect(self._sync_filter_widths)
        header.setSectionResizeMode(QHeaderView.Stretch)
        header.setMinimumSectionSize(90)
        self.tbl.viewport().installEventFilter(self)
        self._auto_fit_pending = False
        self.tbl.horizontalScrollBar().valueChanged.connect(self._sync_filter_scroll_from_table)
        self.filter_scroll.horizontalScrollBar().valueChanged.connect(self._sync_filter_scroll_from_filter)
        self.proxy.layoutChanged.connect(self._schedule_span_update)
        self.proxy.modelReset.connect(self._schedule_span_update)

        v.addLayout(top)
        v.addWidget(self.filter_scroll)
        v.addWidget(self.tbl, 1)

        self.apply_permissions()

    def apply_permissions(self):
        can_read = True
        try:
            can_read = bool(self.main.has_permission("read"))
        except Exception:
            can_read = True
        self.btn_load.setEnabled(can_read)
        self.btn_refresh.setEnabled(can_read)

    def refresh_last(self):
        if self._last_path:
            self._run_pipeline(self._last_path)
        else:
            QMessageBox.information(self, "Bilgi", "Önce ZPPR0308 dosyası yükleyin.")

    def load_zppr0308(self):
        try:
            if not self.main.has_permission("read"):
                QMessageBox.warning(self, "Yetki yok", "Bu sekmeyi kullanmak için okuma yetkisi gerekiyor.")
                return
        except Exception:
            pass

        path, _ = QFileDialog.getOpenFileName(
            self,
            "ZPPR0308 Seç",
            "",
            "Excel Files (*.xlsx *.xlsb *.xls);;All Files (*)"
        )
        if not path:
            return

        self._last_path = path
        self._run_pipeline(path)

    def _run_pipeline(self, path: str):
        try:
            df_raw = pd.read_excel(path)
            # >>> KOLON ADLARINI TEMİZLE (kritik)
            df_raw.columns = [_clean_col(c) for c in df_raw.columns]
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Dosya okunamadı:\n{e}")
            return

        try:
            out = self._build_output(df_raw)
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Hesap sırasında hata:\n{e}")
            return

        self.model.set_df(out)
        self._rebuild_filters()
        QTimer.singleShot(0, self._apply_combo_filters)
        self._schedule_span_update()
        self._apply_header_wrapping()  # >>> burada wrap header label'larını basıyoruz

        try:
            n_isemri = out["Dokuma İş Emri"].nunique() if "Dokuma İş Emri" in out.columns else len(out)
            n_tip = out["Tip Kodu"].nunique() if "Tip Kodu" in out.columns else 0
            self.lbl_info.setText(f"İş Emri: {n_isemri} | Tip: {n_tip} | Satır: {len(out)}")
        except Exception:
            self.lbl_info.setText(f"Satır: {len(out)}")

        QTimer.singleShot(0, lambda: self.tbl.resizeColumnsToContents())
        QTimer.singleShot(0, self._sync_filter_widths)
        QTimer.singleShot(0, self._schedule_auto_fit)
        QTimer.singleShot(0, self._schedule_header_rewrap)  # resize sonrası tekrar wrap

    def _build_output(self, df: pd.DataFrame) -> pd.DataFrame:
        # ---- Taslak kolon eşleştirmeleri ----
        col_bolum = _col_pick(df, ["Bölüm", "Bolum", "Ünite", "Unite", "Plan Grubu", "PlanGrubu"])
        col_isemri = _col_pick(df, ["Dokuma İş Emri", "Dokuma iş Emri", "Dokuma Is Emri", "Dokuma_Is_Emri"])

        # >>> Kritik: ihzarat/ihrazat/ihracat varyantları
        col_ihz = _col_pick(df, [
            "İhzarat İş Emri",
            "İhzarat iş Emri",
            "Ihzarat Is Emri",
            "İhrazat İş Emri",
            "İhrazat İşemri",
            "İhrazat İş emri",
            "İhrazat İşEmri",
            "İhracat İş Emri",
            "Ihracat Is Emri",
        ])

        col_tip = _col_pick(df, ["Tip Kodu", "Malzeme", "Mamul", "Mamul Tipi"])
        col_hedef = _col_pick(df, ["Dokuma Hedef Metre", "Dokuma Hdf Mik", "Dokuma Hdf", "Dokuma Hedef"])
        col_etiket = _col_pick(df, ["Etiket Numarası", "Etiket No", "Etiket", "Etiket Numarasi"])

        # Taslak: İhrazat Tyt Mik
        col_ihz_m = _col_pick(df, [
            "İhrazat Tyt Mik",
            "İhzarat Tyt Mik",
            "İhzarat Fiili Çıkan Metre",
            "İhrazat Fiili Çıkan Metre",
            "Tyt Mik",
            "Tyt Miktar",
        ])

        # Taslak: Dokuma Tyt Mik
        col_dok_m = _col_pick(df, [
            "Dokuma Tyt Mik",
            "Dokuma TYT Mik",
            "Dokuma Etiket Bazında Proses Kartı Alınmış Metre",
            "Dokuma Etiket Bazinda Proses Karti Alinmis Metre",
            "Dokuma Mik",
        ])

        missing = [name for name, c in [
            ("Dokuma İş Emri", col_isemri),
            ("İhzarat/İhrazat/İhracat İş Emri", col_ihz),
            ("Tip Kodu/Malzeme", col_tip),
            ("Dokuma Hdf Mik", col_hedef),
            ("Etiket Numarası", col_etiket),
            ("İhrazat Tyt Mik", col_ihz_m),
            ("Dokuma Tyt Mik", col_dok_m),
        ] if c is None]
        if missing:
            cols_preview = "\n".join(map(str, df.columns))
            raise ValueError(
                "ZPPR0308 içinde gerekli kolon(lar) bulunamadı: "
                + ", ".join(missing)
                + "\n\nBulunan kolonlar:\n"
                + cols_preview
            )

        work = df.copy()

        # sayısal
        work[col_hedef] = _to_num(work[col_hedef])
        work[col_ihz_m] = _to_num(work[col_ihz_m])
        work[col_dok_m] = _to_num(work[col_dok_m])

        # ---- Satır seviyesi (etiket kadar satır) ----
        out = pd.DataFrame({
            "Bölüm": work[col_bolum].astype(str) if col_bolum else "",
            "Dokuma İş Emri": work[col_isemri].astype(str),
            "İhzarat İşemri": work[col_ihz].astype(str),
            "Tip Kodu": work[col_tip].astype(str),
            "Dokuma Hedef Metre": work[col_hedef],
            "Etiket Numarası": work[col_etiket].astype(str),
            "İhzarat Fiili Çıkan Metre": work[col_ihz_m],
            "Dokuma Etiket Bazında Proses Kartı Alınmış Metre": work[col_dok_m],
        })

        # ---- İş emri bazında tek hücre mantığında hesaplar ----
        g = out.groupby("Dokuma İş Emri", dropna=False)

        dok_top = g["Dokuma Etiket Bazında Proses Kartı Alınmış Metre"].sum(min_count=1)
        ihz_top = g["İhzarat Fiili Çıkan Metre"].sum(min_count=1)
        hedef_first = g["Dokuma Hedef Metre"].apply(lambda s: s.dropna().iloc[0] if s.dropna().shape[0] else pd.NA)

        out["Dokuma Proses Kartı Alınmış Toplam Metre"] = out["Dokuma İş Emri"].map(dok_top)
        out["Dokuma Hedef Metre"] = out["Dokuma İş Emri"].map(hedef_first)
        out["_ihz_top"] = out["Dokuma İş Emri"].map(ihz_top)

        out["Hedefe Uyum %"] = (out["Dokuma Proses Kartı Alınmış Toplam Metre"] / out["Dokuma Hedef Metre"]) * 100.0

        # Taslağa birebir formül:
        out["Bu İş Emrinin Büzülmesi"] = 100.0 - (
            (out["Dokuma Proses Kartı Alınmış Toplam Metre"] / out["_ihz_top"]) * 100.0
        )


        # ---- Durum kuralı (taslağa birebir) ----
        def _durum_for_group(s: pd.Series) -> str:
            x = pd.to_numeric(s, errors="coerce")
            ok = x.notna() & (x != 0)
            return "Dokuması bitmiş İş Emrini Kontrol Ederek Kapattır" if bool(ok.all()) else "Devam ediyor"

        durum_map = g["Dokuma Etiket Bazında Proses Kartı Alınmış Metre"].apply(_durum_for_group)
        out["Durum"] = out["Dokuma İş Emri"].map(durum_map)
        # Devam eden iş emirlerinde büzülme değeri görünmesin (tamamlanmadığı için yanıltır)
        out.loc[out["Durum"] == "Devam ediyor", "Bu İş Emrinin Büzülmesi"] = pd.NA

        # ---- SQL join: dbo.TipBuzulmeModel ----
        tips = out["Tip Kodu"].astype(str).fillna("").unique().tolist()
        ref = storage.fetch_tip_buzulme_model(tips)

        if not ref.empty:
            ref = ref.rename(columns={
                "SistemBuzulme": "SAP de Sistemdeki Büzülme",
                "GecmisBuzulme": "Geçmiş Yıllardaki Fiili Büzülme(Son2Yıl)",
                "GuvenAraligi": "Güven Aralığı",
            })
            out = out.merge(
                ref[["TipKodu", "SAP de Sistemdeki Büzülme", "Geçmiş Yıllardaki Fiili Büzülme(Son2Yıl)", "Güven Aralığı"]],
                left_on="Tip Kodu",
                right_on="TipKodu",
                how="left"
            ).drop(columns=["TipKodu"], errors="ignore")
        else:
            out["SAP de Sistemdeki Büzülme"] = pd.NA
            out["Geçmiş Yıllardaki Fiili Büzülme(Son2Yıl)"] = pd.NA
            out["Güven Aralığı"] = ""

        # yuvarlama
        for c in [
            "Hedefe Uyum %",
            "Bu İş Emrinin Büzülmesi",
            "SAP de Sistemdeki Büzülme",
            "Geçmiş Yıllardaki Fiili Büzülme(Son2Yıl)",
        ]:
            if c in out.columns:
                out[c] = pd.to_numeric(out[c], errors="coerce").round(2)

        out = out.drop(columns=["_ihz_top"], errors="ignore")

        # kolon sırası: taslağınla aynı
        ordered = [
            "Bölüm",
            "Dokuma İş Emri",
            "İhzarat İşemri",
            "Tip Kodu",
            "Dokuma Hedef Metre",
            "Dokuma Proses Kartı Alınmış Toplam Metre",
            "Hedefe Uyum %",
            "Etiket Numarası",
            "İhzarat Fiili Çıkan Metre",
            "Dokuma Etiket Bazında Proses Kartı Alınmış Metre",
            "Bu İş Emrinin Büzülmesi",
            "SAP de Sistemdeki Büzülme",
            "Geçmiş Yıllardaki Fiili Büzülme(Son2Yıl)",
            "Güven Aralığı",
            "Durum",
        ]
        keep = [c for c in ordered if c in out.columns]
        rest = [c for c in out.columns if c not in keep]
        out = out[keep + rest]

        if "Dokuma İş Emri" in out.columns:
            sort_cols = ["Dokuma İş Emri"]
            if "Etiket Numarası" in out.columns:
                sort_cols.append("Etiket Numarası")
            out = out.sort_values(by=sort_cols, kind="stable").reset_index(drop=True)

        return out

    def _rebuild_filters(self):
        # Sil
        while self.filter_bar_layout.count():
            item = self.filter_bar_layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.setParent(None)
                w.deleteLater()
        self._filter_edits.clear()

        cols = list(self.model._df.columns) if (self.model and self.model._df is not None) else []
        header = self.tbl.horizontalHeader()

        for c, name in enumerate(cols):
            edit = QLineEdit()
            edit.setPlaceholderText(str(name))
            edit.textChanged.connect(lambda text, col=c: self._on_filter_changed(col, text))
            edit.setFixedWidth(header.sectionSize(c))
            edit.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Preferred)
            self.filter_bar_layout.addWidget(edit)
            self._filter_edits.append(edit)

        self._sync_filter_widths()

    def _on_filter_changed(self, col: int, text: str):
        self.proxy.setFilterForColumn(col, text)
        self._schedule_span_update()
    def _set_filter_by_colname(self, col_name: str, text: str) -> None:
        """
        Combo filtrelerini, ilgili kolonun QLineEdit filtre kutusuna yazar.
        Böylece hem proxy hem UI senkron kalır.
        """
        if not self.model or self.model._df is None or self.model._df.empty:
            return

        cols = list(self.model._df.columns)
        if col_name not in cols:
            return

        col_idx = cols.index(col_name)

        # Filter bar hazırsa, edit'e yaz; değilse direkt proxy'ye yaz
        if 0 <= col_idx < len(self._filter_edits):
            edit = self._filter_edits[col_idx]
            old = edit.blockSignals(True)
            edit.setText(text)
            edit.blockSignals(old)
            # manuel tetikle
            self.proxy.setFilterForColumn(col_idx, text)
            self._schedule_span_update()
        else:
            self.proxy.setFilterForColumn(col_idx, text)
            self._schedule_span_update()

    def _apply_combo_filters(self) -> None:
        # Bölüm filtresi
        b = self.cmb_bolum.currentText().strip() if hasattr(self, "cmb_bolum") else "HEPSİ"
        if b == "İSKO14 (DK14)":
            self._set_filter_by_colname("Bölüm", "DK14")
        elif b == "İSKO11 (DK11)":
            self._set_filter_by_colname("Bölüm", "DK11")
        elif b == "MEKİKLİ (DK98)":
            self._set_filter_by_colname("Bölüm", "DK98")
        else:
            self._set_filter_by_colname("Bölüm", "")

        # Durum filtresi
        d = self.cmb_durum.currentText().strip() if hasattr(self, "cmb_durum") else "HEPSİ"
        if d == "Devam ediyor":
            self._set_filter_by_colname("Durum", "Devam ediyor")
        elif d == "Bitmiş":
            # senin Durum metnin: "Dokuması bitmiş İş Emrini Kontrol Ederek Kapattır"
            # bunu yakalamak için en güvenli anahtar kelime:
            self._set_filter_by_colname("Durum", "Kapattır")
        else:
            self._set_filter_by_colname("Durum", "")

    def _on_bolum_combo_changed(self, _text: str) -> None:
        self._apply_combo_filters()

    def _on_durum_combo_changed(self, _text: str) -> None:
        self._apply_combo_filters()


    def _sync_filter_widths(self):
        header = self.tbl.horizontalHeader()
        for c, edit in enumerate(self._filter_edits):
            if c < header.count():
                edit.setFixedWidth(header.sectionSize(c))
        try:
            self.filter_bar.setMinimumWidth(header.length())
        except Exception:
            pass

        # Wrap header yüksekliğini set et (WrapHeaderView sizeHint)
        try:
            header.setFixedHeight(self.wrap_header.sizeHint().height())
        except Exception:
            pass

        # Sütun genişliği değişince başlıkları yeniden çiz
        self._schedule_header_rewrap()

    def _sync_filter_scroll_from_table(self, value: int):
        if self._scroll_from_filter:
            return
        self._scroll_from_table = True
        self.filter_scroll.horizontalScrollBar().setValue(value)
        self._scroll_from_table = False

    def _sync_filter_scroll_from_filter(self, value: int):
        if self._scroll_from_table:
            return
        self._scroll_from_filter = True
        self.tbl.horizontalScrollBar().setValue(value)
        self._scroll_from_filter = False

    def _schedule_span_update(self):
        QTimer.singleShot(0, self._apply_spans)

    def _schedule_header_rewrap(self):
        QTimer.singleShot(0, self._apply_header_wrapping)

    def eventFilter(self, obj, event):
        # tablo viewport resize olunca kolonları ekrana göre yeniden dağıt
        if obj is self.tbl.viewport() and event.type() == QEvent.Type.Resize:
            self._schedule_auto_fit()
        return super().eventFilter(obj, event)

    def _schedule_auto_fit(self):
        if getattr(self, "_auto_fit_pending", False):
            return
        self._auto_fit_pending = True
        QTimer.singleShot(0, self._auto_fit_columns_to_viewport)

    def _auto_fit_columns_to_viewport(self):
        self._auto_fit_pending = False

        if not self.model or self.model._df is None or self.model._df.empty:
            return

        header = self.tbl.horizontalHeader()
        vp_w = self.tbl.viewport().width()
        if vp_w <= 0:
            return

        cols = list(self.model._df.columns)
        n = len(cols)
        if n == 0:
            return

        # Mevcut kolon genişlikleri
        sizes = [header.sectionSize(i) for i in range(n)]
        total = sum(sizes)

        # Çok küçük bir margin/padding payı
        margin = 24
        target = max(0, vp_w - margin)

        # Eğer toplam zaten büyükse (sığmıyorsa) zorlamayalım, scroll devam etsin
        if total >= target:
            # yine de başlıkları mevcut genişliğe göre wrap et
            self._schedule_header_rewrap()
            return

        # Boşluk varsa dağıtalım (oransal)
        extra = target - total
        base_sum = max(1, total)

        for i in range(n):
            add = int(extra * (sizes[i] / base_sum))
            header.resizeSection(i, sizes[i] + add)

        # Son küçük farkı son kolona verelim
        new_total = sum(header.sectionSize(i) for i in range(n))
        diff = target - new_total
        if diff != 0 and n > 0:
            header.resizeSection(n - 1, max(40, header.sectionSize(n - 1) + diff))

        # Filtre kutularını senkronla + başlıkları yeni genişliğe göre wrap et
        self._sync_filter_widths()
        self._schedule_header_rewrap()

    def _apply_spans(self):
        self.tbl.clearSpans()
        if not self.model or self.model._df is None or self.model._df.empty:
            return

        cols = list(self.model._df.columns)
        try:
            isemri_idx = cols.index("Dokuma İş Emri")
        except ValueError:
            return

        merge_names = [
            "Bölüm",
            "Dokuma İş Emri",
            "İhzarat İşemri",
            "Tip Kodu",
            "Dokuma Hedef Metre",
            "Dokuma Proses Kartı Alınmış Toplam Metre",
            "Hedefe Uyum %",
            "Bu İş Emrinin Büzülmesi",
            "SAP de Sistemdeki Büzülme",
            "Geçmiş Yıllardaki Fiili Büzülme(Son2Yıl)",
            "Güven Aralığı",
            "Durum",
        ]
        merge_indices = [cols.index(name) for name in merge_names if name in cols]
        if not merge_indices:
            return

        row_count = self.proxy.rowCount()
        if row_count == 0:
            return

        def _cell_value(row: int, col: int):
            idx = self.proxy.index(row, col)
            return idx.data(Qt.DisplayRole) if idx.isValid() else None

        current_value = _cell_value(0, isemri_idx)
        start = 0
        for r in range(1, row_count + 1):
            next_value = _cell_value(r, isemri_idx) if r < row_count else None
            if next_value != current_value:
                span_len = r - start
                if span_len > 1:
                    for col in merge_indices:
                        self.tbl.setSpan(start, col, span_len, 1)
                start = r
                current_value = next_value

    def _apply_header_wrapping(self):
        if not hasattr(self, "wrap_header"):
            return

        if not self.model or self.model._df is None:
            return

        cols = list(self.model._df.columns)

        # Güvenlik: boş kalmasın
        if not cols:
            self.wrap_header.set_labels([])
            return

        wrapped = [
            self._wrap_header_label(str(name), width=18, max_lines=3)
            for name in cols
        ]

        self.wrap_header.set_labels(wrapped)
        self.wrap_header.setFixedHeight(self.wrap_header.sizeHint().height())

    @staticmethod
    def _wrap_header_label(text: str, width: int = 18, max_lines: int = 3) -> str:
        lines = textwrap.wrap(text, width=width) if text else []
        if not lines:
            return text
        if len(lines) <= max_lines:
            return "\n".join(lines)
        head = lines[: max_lines - 1]
        tail = " ".join(lines[max_lines - 1:])
        return "\n".join(head + [tail])
