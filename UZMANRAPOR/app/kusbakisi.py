# app/kusbakisi.py
from __future__ import annotations
from dataclasses import dataclass
from typing import Optional, Dict, Tuple, List
import re, hashlib, colorsys
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import pandas as pd
from PySide6.QtCore import Qt, QSize
from PySide6 import QtGui, QtWidgets
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QScrollArea, QGridLayout,
    QTableWidget, QTableWidgetItem, QComboBox, QPushButton, QFrame, QSizePolicy
)

from app import storage  # Usta Defteri sayımları + kısıt listeleri

IST = ZoneInfo("Europe/Istanbul")

# -------------------------------------------------------------
#  SABİT YERLEŞİM – iki salon
# -------------------------------------------------------------

HALL_GAP_COLS = 2
LEFT_COLS = 9

def _seq(start: int, count: int, step: int) -> List[int]:
    return [start + i * step for i in range(count)]

def _rows_spec_to_mapping(rows_spec: List[Tuple[int,int,int]], row_offset: int, col_offset: int) -> Dict[str, Tuple[int,int]]:
    m: Dict[str, Tuple[int,int]] = {}
    for r, (start, count, step) in enumerate(rows_spec):
        seq = _seq(start, count, step)
        for c, num in enumerate(seq):
            m[str(num)] = (row_offset + r, col_offset + c)
    return m

def _build_fixed_layout_from_spec() -> Dict[str, Tuple[int,int]]:
    mapping: Dict[str, Tuple[int,int]] = {}
    left_rows: List[Tuple[int,int,int]] = [
        (2391, 9, -2), (2392, 9, -2),
        (2393, 9, +2), (2394, 9, +2),
        (2427, 9, -2), (2428, 9, -2),
        (2429, 9, +2), (2430, 9, +2),
        (2463, 9, -2), (2464, 9, -2),
        (2465, 9, +2), (2466, 9, +2),
        (2499, 9, -2), (2500, 9, -2),
        (2501, 9, +2), (2502, 9, +2),
    ]
    mapping.update(_rows_spec_to_mapping(left_rows, row_offset=0, col_offset=0))

    right_top_rows: List[Tuple[int,int,int]] = [
        (2223, 12, -2), (2224, 12, -2),
        (2225, 12, +2), (2226, 12, +2),
        (2271, 12, -2), (2272, 12, -2),
        (2273, 12, +2), (2274, 12, +2),
        (2319, 12, -2), (2320, 12, -2),
    ]
    right_bottom_rows: List[Tuple[int,int,int]] = [
        (2321, 9, +2), (2322, 9, +2),
        (2355, 9, -2), (2356, 9, -2),
        (2357, 9, +2), (2358, 9, +2),
    ]
    right_top_col_offset = LEFT_COLS + HALL_GAP_COLS
    mapping.update(_rows_spec_to_mapping(right_top_rows, row_offset=0, col_offset=right_top_col_offset))
    mapping.update(_rows_spec_to_mapping(right_bottom_rows, row_offset=len(right_top_rows), col_offset=right_top_col_offset))
    return mapping

MACHINE_LAYOUT: Dict[str, Tuple[int,int]] = _build_fixed_layout_from_spec()
MAX_ROW = max((r for (r, c) in MACHINE_LAYOUT.values()), default=0)
TOTAL_ROWS = MAX_ROW + 1

# -------------------------------------------------------------
#  KATEGORİ KURALLARI (tezgâh no'ya göre)
# -------------------------------------------------------------

NEVER = {2430, 2432, 2434, 2436, 2438, 2440, 2442, 2444, 2446}
HAM_ALLOWED = set(range(2447, 2519))   # 2447–2518 arası
DENIM_ALLOWED_RANGE = (2201, 2446)     # 2201–2446 arası


def _loom_in_category(loom_no: str, category: str) -> bool:
    try:
        n = int(str(loom_no).strip())
    except Exception:
        return False
    if n in NEVER:
        return False
    if category == "HAM":
        return n in HAM_ALLOWED
    if category == "DENIM":
        return DENIM_ALLOWED_RANGE[0] <= n <= DENIM_ALLOWED_RANGE[1]
    return True  # Tümü

# -------------------------------------------------------------
#  Yardımcılar (normalize, renk, sıralama anahtarı)
# -------------------------------------------------------------

def _norm(s: object) -> str:
    return "" if s is None else str(s).strip()

def _hex_color_for_group(group_label: str) -> str:
    gl = _norm(group_label)
    if not gl:
        return "#eaeaea"
    h = hashlib.sha1(gl.encode("utf-8")).hexdigest()
    hue = int(h[:2], 16) * 360 // 255
    r, g, b = colorsys.hls_to_rgb(hue/360.0, 0.52, 0.65)
    return "#{:02x}{:02x}{:02x}".format(int(r*255), int(g*255), int(b*255))

def _text_color_on(bg_hex: str) -> str:
    try:
        bg_hex = bg_hex.lstrip("#")
        r, g, b = int(bg_hex[0:2], 16), int(bg_hex[2:4], 16), int(bg_hex[4:6], 16)
        lum = 0.2126*r + 0.7152*g + 0.0722*b
        return "#000000" if lum > 155 else "#ffffff"
    except Exception:
        return "#000000"

_num_pat = re.compile(r"\d+(?:[.,]\d+)?")

def _fmt_num(f: float) -> str:
    if abs(f - round(f)) < 1e-9:
        return str(int(round(f)))
    s = f"{f:.3f}".rstrip("0").rstrip(".")
    return s

def _normalize_tg_label(label: object) -> str:
    """
    '160,0 2 194,0' -> '160/2/194'
    '052.5/04/194'  -> '52.5/4/194'
    """
    s = _norm(label)
    if not s:
        return ""
    s = s.replace(",", ".")
    nums = [float(x.replace(",", ".")) for x in _num_pat.findall(s)]
    if not nums:
        return s
    parts = [_fmt_num(x) for x in nums]
    return "/".join(parts)

def _tarak_sort_key(label: str) -> tuple:
    s = _normalize_tg_label(label)
    nums = [float(x) for x in s.replace(",", ".").split("/") if x.strip() != ""]
    if not nums:
        return (9999.0,)
    return tuple(nums)

def _loom_digits(val: object) -> str:
    """Bir tezgâh no hücresinden sadece rakamları (string) çıkarır."""
    m = re.search(r"(\d+)", str(val or ""))
    return m.group(1) if m else ""

# -------------------------------------------------------------
#  DÜNÜN TOPLAM SAYIMI (3 vardiya toplamı)
# -------------------------------------------------------------

def _yesterday_shift_windows(ref: datetime | None = None):
    """Dünün üç vardiya penceresi (07-15, 15-23, 23-ertesi 07)."""
    now = ref or datetime.now(IST)
    y  = (now - timedelta(days=1)).date()
    s1 = datetime(y.year, y.month, y.day, 7, 0, tzinfo=IST)
    s2 = datetime(y.year, y.month, y.day, 15, 0, tzinfo=IST)
    s3 = datetime(y.year, y.month, y.day, 23, 0, tzinfo=IST)
    e1 = s2
    e2 = s3
    e3 = datetime(y.year, y.month, y.day, 7, 0, tzinfo=IST) + timedelta(days=1)
    return [(s1, e1), (s2, e2), (s3, e3)]

def _compute_yesterday_totals() -> tuple[int, int, str]:
    """Dünün üç vardiyasını toplayıp (DÜĞÜM, TAKIM, tarih_str) döndürür."""
    wins = _yesterday_shift_windows()
    total_dugum = 0
    total_takim = 0
    for (s,e) in wins:
        total_dugum += storage.count_usta_between(s, e, what="DÜĞÜM", direction="ALINDI")
        total_takim += storage.count_usta_between(s, e, what="TAKIM",  direction="ALINDI")
    date_str = wins[0][0].strftime("%d.%m.%Y")
    return total_dugum, total_takim, date_str

# -------------------------------------------------------------
#  Görsel bileşenler
# -------------------------------------------------------------

@dataclass
class LoomView:
    loom: str
    tarak: str
    kalan_m: str
    is_empty: bool
    color: str
    koktip: str = ""      # <-- yeni
    cut_type: str = ""    # <-- yeni

from PySide6.QtWidgets import QLabel

class LoomCell(QLabel):
    def __init__(self, info: LoomView, white_bg: bool = False, parent=None):
        super().__init__(parent)
        loom_color = "#c00000" if info.is_empty else "#000000"
        # SOLDa tezgâh no — beyaz dolgulu rozet
        loom_badge = (
            f"<span style='"
            f"display:inline-block;"
            f"font-size:12pt;"  # biraz büyük
            f"font-weight:700;"  # kalın
            f"background:#ffffff;"  # beyaz dolgu
            f"color:{loom_color};"  # açık= kırmızı, çalışıyor= siyah
            f"border:1px solid {loom_color};"
            f"border-radius:12px;"
            f"padding:2px 10px;"
            f"white-space:nowrap;"
            f"'>"
            f"{info.loom}"
            f"</span>"
        )
        # Sağ üstte küçük rozet (7pt), beyaz dolgulu – siyah yazı
        cut_badge = ""
        if info.cut_type:
            cut_badge = (
                f"<span style='"
                f"display:inline-block;"
                f"font-size:7pt;"
                f"background:#ffffff;"
                f"color:#111111;"
                f"border:1px solid #222222;"
                f"border-radius:9px;"
                f"padding:1px 6px;"
                f"white-space:nowrap;"
                f"'>"
                f"{info.cut_type}"
                f"</span>"
            )

        # Üst satır: solda tezgah no (kalın), sağda rozet
        top_line = (
            f"<table width='100%' cellspacing='0' cellpadding='0'><tr>"
            f"<td align='left'>{loom_badge}</td>"
            f"<td align='right'>{cut_badge}</td>"
            f"</tr></table>"
        )

        # Orta satır: Tarak grubu (biraz küçültülmüş)
        t = info.tarak if info.tarak else "-"
        middle_line = f"<div style='font-size:10.5pt'>{t}</div>"

        # Alt satır: KalanMetreNorm / KökTip (küçük puntolu)
        km = info.kalan_m or ""
        kt = info.koktip or ""
        third_line_raw = (km if km else "") + ((" / " + kt) if (km and kt) else (kt if kt else ""))
        third_line = f"<div style='font-size:7pt; color:#222'>{third_line_raw}</div>"

        # METİN
        self.setTextFormat(Qt.RichText)
        self.setText(f"{top_line}{middle_line}{third_line}")
        self.setAlignment(Qt.AlignCenter)
        self.setMargin(6)
        self.setWordWrap(True)

        border = "#bbbbbb"
        bg = "#ffffff" if white_bg else info.color
        self.setStyleSheet(f"""
            QLabel {{
                background: {bg};
                border: 1px solid {border};
                border-radius: 6px;
                color: #000;
            }}
        """)

        # Biraz dikey alan
        self.setMinimumSize(QSize(96, 56))


# -------------------------------------------------------------
#  KUŞBAKIŞI WIDGET
# -------------------------------------------------------------

class KusbakisiWidget(QWidget):
    """
    Sol: Özet tablo (seçilen kategoriye uygun Running grupları)
         + 'Takım olacak işler' (Running’de olmayanlar; renksiz)
    Sağ: Yerleşim ızgarası (kategoriye uygun tezgâhlar)
    Üstte: KPI'lar — Çalışan Tezgah Sayısı / Alınan Düğüm / Alınan Takım
    """
    def __init__(self, parent=None):
        from PySide6.QtWidgets import QAbstractItemView
        super().__init__(parent)
        root = QHBoxLayout(self)

        # Sol panel
        self.sidebar = QWidget()
        left = QVBoxLayout(self.sidebar)

        # Sidebar’ı daralt
        self.sidebar.setMaximumWidth(515)
        self.sidebar.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Preferred)
        toolbar = QHBoxLayout()
        toolbar.addWidget(QLabel("Kategori:"), 0)

        self.cmb_cat = QComboBox(); self.cmb_cat.addItems(["Tümü", "DENIM", "HAM"])

        # --- KPI etiketleri ---
        self.lbl_working = QLabel()
        self.lbl_yesterday = QLabel()  # iki satır tek label
        for w in (self.lbl_working, self.lbl_yesterday):
            w.setStyleSheet("QLabel { font-weight: 560; padding: 0 8px; }")
        self.lbl_yesterday.setTextFormat(Qt.RichText)
        self.lbl_yesterday.setWordWrap(True)

        self.btn_all_colors = QPushButton("Tümünü Renkli")

        toolbar.addWidget(self.cmb_cat, 0)
        toolbar.addSpacing(12)
        toolbar.addWidget(QLabel("Çalışan Tezgah:"), 0)
        toolbar.addWidget(self.lbl_working, 0)
        toolbar.addSpacing(8)
        toolbar.addWidget(self.lbl_yesterday, 0)
        self.lbl_yesterday.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        toolbar.addStretch(1)
        toolbar.addWidget(self.btn_all_colors, 0)
        # Üst bilgilendirme (DÜĞÜM sekmesindekiyle aynı metni göstereceğiz)
        self.lbl_status_kus = QLabel("")
        self.lbl_status_kus.setStyleSheet("QLabel{font-weight:700;}")
        left.addWidget(self.lbl_status_kus)  # kategori/KPI toolbar'ının üstünde dursun
        left.addLayout(toolbar)

        self.tbl = QTableWidget(0, 6)
        self.tbl.setHorizontalHeaderLabels(
            ["Tarak Grubu", "İş Adedi", "Stok Adedi", "Tarak Adedi", "Açık Tezgah", "Termin (erken)"]
        )
        self.tbl.horizontalHeader().setStretchLastSection(True)
        left.addWidget(self.tbl, 8)
        from PySide6.QtWidgets import QAbstractItemView as _AIV
        self.tbl.setEditTriggers(_AIV.NoEditTriggers)
        self.tbl.setSelectionBehavior(_AIV.SelectRows)
        self.tbl.setSelectionMode(_AIV.SingleSelection)

        left.addWidget(QLabel("Takım olacak işler"))
        self.tbl_planned = QTableWidget(0, 3)
        self.tbl_planned.setHorizontalHeaderLabels(["Tarak Grubu", "İş Adedi", "Stok Adedi"])
        self.tbl_planned.horizontalHeader().setStretchLastSection(True)
        left.addWidget(self.tbl_planned, 2)

        root.addWidget(self.sidebar)
        self.tbl_planned.setEditTriggers(_AIV.NoEditTriggers)
        self.tbl_planned.setSelectionBehavior(_AIV.SelectRows)
        self.tbl_planned.setSelectionMode(_AIV.SingleSelection)

        # Sağ panel
        self.grid_host = QWidget()
        self.grid = QGridLayout(self.grid_host)
        self.grid.setContentsMargins(6,6,6,6)
        self.grid.setHorizontalSpacing(6)
        self.grid.setVerticalSpacing(6)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(self.grid_host)
        root.addWidget(self.scroll, 1)

        # Veri
        self.df_jobs: Optional[pd.DataFrame] = None
        self.df_run: Optional[pd.DataFrame] = None
        self.selected_group: Optional[str] = None

        # KPI dahili cache
        self._kpi_working: int = 0

        # Kısıt listeleri (Arızalı/Bakımda & Boş Gösterilecek)
        self._blocked: set[str] = set()
        self._dummy: set[str] = set()

        # Sinyaller
        self.cmb_cat.currentTextChanged.connect(self._rebuild_all)
        self.tbl.cellClicked.connect(self._on_summary_clicked)
        self.btn_all_colors.clicked.connect(self._clear_selection)

    # --- Kısıt listelerini depodan oku
    def _reload_restrictions(self):
        try:
            self._blocked = set(storage.load_blocked_looms() or [])
        except Exception:
            self._blocked = set()
        try:
            self._dummy = set(storage.load_dummy_looms() or [])
        except Exception:
            self._dummy = set()

    def _update_kpis(self):
        self.lbl_working.setText(str(self._kpi_working))

        try:
            tot_dugum, tot_takim, date_str = _compute_yesterday_totals()
            self.lbl_yesterday.setText(
                f"<div>"
                f"<b>Alınan Düğüm :</b> {tot_dugum}<br>"
                f"<b>Alınan Takım :</b> {tot_takim}"
                f"</div>"
            )
        except Exception:
            self.lbl_yesterday.setText("—")

    def refresh(self, df_jobs: Optional[pd.DataFrame], df_running: Optional[pd.DataFrame]) -> None:
        # Kısıtları her yenilemede yeniden oku (butondan güncellenince yansısın)
        self._reload_restrictions()
        self.df_jobs = df_jobs.copy() if df_jobs is not None else None
        self.df_run = df_running.copy() if df_running is not None else None
        self._rebuild_all()

    def _rebuild_all(self) -> None:
        self._build_summary_tables()
        self._build_layout_grid()

    # ---------- Sol tablolar ----------
    def _build_summary_tables(self):
        self.tbl.setRowCount(0); self.tbl_planned.setRowCount(0)

        cat_sel = self.cmb_cat.currentText()

        # Running (kategoriye uygun tezgâhlar)
        run_tg_set: set[str] = set()
        tarak_count = pd.Series(dtype=int); acik_count = pd.Series(dtype=int)
        working_count = 0
        if self.df_run is not None and not self.df_run.empty:
            run = self.df_run.copy()
            # Kategori filtresi
            mask = run.get("Tezgah No").astype(str).apply(lambda x: _loom_in_category(x, cat_sel))
            run = run[mask].copy()

            # Kısıtlı tezgâhları (Arızalı/Bakımda + Boş Göster) özetten tamamen çıkar
            digits = run.get("Tezgah No").astype(str).apply(_loom_digits)
            ban = set(self._blocked) | set(self._dummy)
            run = run[~digits.isin(ban)].copy()

            # KPI: çalışan = sipariş yok (94) veya _OpenTezgahFlag True OLMAYANLAR
            is_open = (run.get("Durus No", 0) == 94) | (run.get("_OpenTezgahFlag", False) == True)
            working_count = int((~is_open).sum())

            run["_tg"] = run.get("Tarak Grubu","").apply(_normalize_tg_label)
            run_tg_set = set(run["_tg"].unique().tolist())
            g_run = run.groupby("_tg", dropna=False)
            tarak_count = g_run.size().rename("tarak_adedi")
            acik_count = g_run.apply(lambda x: ((x.get("Durus No", 0) == 94) | (x.get("_OpenTezgahFlag", False) == True)).sum()) \
                               .rename("acik")

        self._kpi_working = working_count
        self._update_kpis()

        # Dinamik (kategoriye göre iş/stok/termin)
        jobs = self.df_jobs
        if jobs is None or jobs.empty:
            return

        jobs = jobs.copy()
        if cat_sel == "DENIM":
            jobs = jobs[~jobs.get("_DyeCategory","").astype(str).str.contains("HAM", na=False)]
        elif cat_sel == "HAM":
            jobs = jobs[jobs.get("_DyeCategory","").astype(str).str.contains("HAM", na=False)]

        jobs["_tg"] = jobs.get("Tarak Grubu","").apply(_normalize_tg_label)
        if "_LeventHasDigits" in jobs.columns:
            jobs["_stok"] = jobs["_LeventHasDigits"].astype(bool)
        else:
            jobs["_stok"] = jobs.get("Levent No","").astype(str).str.strip() != ""

        g_jobs = jobs.groupby("_tg", dropna=False)
        job_count = g_jobs.size().rename("is_adedi")
        stok_count = g_jobs["_stok"].sum(min_count=1).fillna(0).rename("stok_adedi")
        earliest_termin = g_jobs.apply(
            lambda x: pd.to_datetime(x.get("Mamul Termin"), errors="coerce").min()
        ).rename("termin")

        # 1) ÖZET TABLO (kategoriye uygun Running evreni) — kısıtlı tezgâhlar hiç sayılmadı
        base = pd.DataFrame({"_tg": sorted(run_tg_set, key=_tarak_sort_key)})
        summary_main = base.merge(job_count.reset_index(), on="_tg", how="left") \
                           .merge(stok_count.reset_index(), on="_tg", how="left") \
                           .merge(earliest_termin.reset_index(), on="_tg", how="left") \
                           .merge(tarak_count.reset_index(), on="_tg", how="left") \
                           .merge(acik_count.reset_index(), on="_tg", how="left")
        summary_main[["is_adedi","stok_adedi","tarak_adedi","acik"]] = summary_main[["is_adedi","stok_adedi","tarak_adedi","acik"]].fillna(0)
        summary_main["termin"] = summary_main["termin"].apply(
            lambda d: "" if (pd.isna(d) or str(d)=="NaT") else pd.to_datetime(d).strftime("%d.%m.%Y")
        )
        summary_main["_sort"] = summary_main["_tg"].apply(_tarak_sort_key)
        summary_main = summary_main.sort_values(by="_sort", ascending=True)

        self.tbl.setRowCount(len(summary_main))
        for r, row in summary_main.iterrows():
            tg = _norm(row["_tg"])
            color = _hex_color_for_group(tg); fg = _text_color_on(color)
            vals = [
                tg,
                str(int(row["is_adedi"])),
                str(int(row["stok_adedi"])),
                str(int(row.get("tarak_adedi", 0))),
                str(int(row.get("acik", 0))),
                str(row["termin"]),
            ]
            for c, v in enumerate(vals):
                it = QTableWidgetItem(v)
                if c == 0:
                    it.setBackground(QtGui.QColor(color))
                    it.setForeground(QtGui.QColor(fg))
                else:
                    it.setTextAlignment(Qt.AlignCenter)
                self.tbl.setItem(r, c, it)
        self.tbl.resizeColumnsToContents()
        if self.tbl.columnCount() > 0:
            self.tbl.setColumnWidth(0, max(self.tbl.columnWidth(0), 75))
            self.tbl.setColumnWidth(5, 80)

        # 2) TAKIM OLACAK İŞLER (Running’de olmayan gruplar) – RENKSİZ
        extra_groups = [g for g in g_jobs.groups.keys() if g not in run_tg_set]
        summary_extra = pd.DataFrame({"_tg": sorted(extra_groups, key=_tarak_sort_key)})
        summary_extra = summary_extra.merge(job_count.reset_index(), on="_tg", how="left") \
                                     .merge(stok_count.reset_index(), on="_tg", how="left") \
                                     .merge(earliest_termin.reset_index(), on="_tg", how="left")
        summary_extra[["is_adedi","stok_adedi"]] = summary_extra[["is_adedi","stok_adedi"]].fillna(0)

        self.tbl_planned.setRowCount(len(summary_extra))
        for r, row in summary_extra.iterrows():
            tg = _norm(row["_tg"])
            vals = [tg, str(int(row["is_adedi"])), str(int(row["stok_adedi"]))]
            for c, v in enumerate(vals):
                it = QTableWidgetItem(v)
                if c != 0:
                    it.setTextAlignment(Qt.AlignCenter)
                self.tbl_planned.setItem(r, c, it)
        self.tbl_planned.resizeColumnsToContents()
        if self.tbl_planned.columnCount() > 0:
            self.tbl_planned.setColumnWidth(0, max(self.tbl_planned.columnWidth(0), 120))

    # ---------- Sağ: yerleşim ızgarası ----------
    def _build_layout_grid(self):
        while self.grid.count():
            it = self.grid.takeAt(0)
            w = it.widget()
            if w: w.deleteLater()

        run = self.df_run
        if run is None or run.empty:
            return

        cat_sel = self.cmb_cat.currentText()
        mask = run.get("Tezgah No").astype(str).apply(lambda x: _loom_in_category(x, cat_sel))
        run = run[mask].copy()
        run["_tg"] = run.get("Tarak Grubu","").apply(_normalize_tg_label)

        sel = _norm(self.selected_group) if self.selected_group else None
        sel_norm = _normalize_tg_label(sel) if sel else None

        # Hızlı lookup
        blocked = set(self._blocked)
        dummy = set(self._dummy)

        for _, row in run.iterrows():
            loom_raw = row.get("Tezgah No", "")
            loom = _norm(loom_raw)
            loom_digits = _loom_digits(loom)
            pos = MACHINE_LAYOUT.get(loom_digits) or MACHINE_LAYOUT.get(loom)
            if pos is None:
                continue

            # Varsayılan (normal tezgâh)
            tarak = _norm(row.get("Tarak Grubu", ""))
            tarak_canon = _normalize_tg_label(tarak)
            kalan = row.get("_KalanMetreNorm", None)
            if pd.isna(kalan):
                kalan_s = ""
            else:
                try:
                    kalan_s = f"{float(kalan):.0f} m"
                except Exception:
                    kalan_s = str(kalan)

            koktip = str(
                row.get("KökTip", "") or
                row.get("Kök Tip Kodu", "") or
                row.get("Kök Tip", "") or
                row.get("Tip No", "") or
                row.get("TipNo", "") or
                ""
            ).strip()

            cut_type = str(
                row.get("Kesim Tipi", "") or
                row.get("ISAVER/ROTOCUT", "") or
                row.get("Kesim", "") or
                row.get("CutType", "") or
                ""
            ).strip()

            is_empty = bool((row.get("Durus No", 0) == 94) or (row.get("_OpenTezgahFlag", False) == True))
            color = _hex_color_for_group(tarak_canon)
            white_bg = (sel_norm is not None and tarak_canon != sel_norm)

            # --- Kısıtlı tezgâhların özel görünümü ---
            if loom_digits in blocked:
                # Arızalı/Bakımda → sönük beyaz; sadece "Arızalı" yaz; grup sayımına dahil edilmedi (solda zaten hariç)
                white_bg = True
                tarak_canon = "Arızalı"
                kalan_s = ""
                koktip = ""
                cut_type = ""
                # is_empty'i kırmızı rozet yapmamak için False tutuyoruz
            elif loom_digits in dummy:
                # Boş gösterilecek → sönük beyaz; sadece "Boş" yaz; grup sayımına dahil edilmedi
                white_bg = True
                tarak_canon = "Boş"
                kalan_s = ""
                koktip = ""
                cut_type = ""
                # is_empty False

            self.grid.addWidget(
                LoomCell(
                    LoomView(
                        loom=loom_digits or loom,
                        tarak=tarak_canon,
                        kalan_m=kalan_s,
                        is_empty=is_empty,
                        color=color,
                        koktip=koktip,
                        cut_type=cut_type
                    ),
                    white_bg=white_bg
                ),
                pos[0], pos[1]
            )

        # salon ayırıcı
        divider = QFrame()
        divider.setFrameShape(QFrame.VLine)
        divider.setStyleSheet("QFrame { background: #9aa0a6; }")
        divider.setFixedWidth(6)
        gap_col = LEFT_COLS
        self.grid.addWidget(divider, 0, gap_col, TOTAL_ROWS, HALL_GAP_COLS)

    def _on_summary_clicked(self, row: int, col: int):
        item = self.tbl.item(row, 0)
        if not item:
            return
        self.selected_group = _norm(item.text())
        if not self.selected_group:
            self.selected_group = None
        self._build_layout_grid()

    def _clear_selection(self):
        self.selected_group = None
        self._build_layout_grid()

    def set_status_label(self, text: str, style: str | None = None):
        self.lbl_status_kus.setText(text or "")
        if style:
            self.lbl_status_kus.setStyleSheet(style)
