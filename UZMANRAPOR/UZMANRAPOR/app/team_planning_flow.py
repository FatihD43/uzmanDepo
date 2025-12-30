# app/team_planning_flow.py
from __future__ import annotations
import re
import os, sys, subprocess
import pandas as pd
from collections import defaultdict

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QListWidget, QTableView,
    QPushButton, QHeaderView, QSplitter, QDialog, QListWidgetItem, QTabWidget, QSpinBox,
    QMessageBox, QInputDialog
)
from PySide6.QtCore import Qt, QModelIndex, QSettings
from app.models import PandasModel

# ---- Arızalı/Boş tezgah listesini depodan okuma (varsa) ----
try:
    from app import storage as _storage
except Exception:
    _storage = None

def _load_restricted_looms() -> tuple[set[str], set[str]]:
    """
    storage'dan arızalı/bakım (blocked) ve 'boş gösterilecek' (dummy) tezgahları okur.
    Hepsini 'sadece rakam' string seti olarak döndürür: {'2201','2468',...}
    """
    def _digits(s) -> str:
        m = re.search(r"(\d+)", str(s))
        return m.group(1) if m else ""

    blocked, dummy = set(), set()
    try:
        if _storage:
            if hasattr(_storage, "load_blocked_looms"):
                blocked = set(_digits(x) for x in (_storage.load_blocked_looms() or []) if _digits(x))
            elif hasattr(_storage, "get_blocked_looms"):
                blocked = set(_digits(x) for x in (_storage.get_blocked_looms() or []) if _digits(x))
            if hasattr(_storage, "load_dummy_looms"):
                dummy = set(_digits(x) for x in (_storage.load_dummy_looms() or []) if _digits(x))
            elif hasattr(_storage, "get_dummy_looms"):
                dummy = set(_digits(x) for x in (_storage.get_dummy_looms() or []) if _digits(x))
    except Exception:
        pass
    return blocked, dummy

# =========================
# Tezgâh izin kuralları
# =========================
NEVER = {2430, 2432, 2434, 2436, 2438, 2440, 2442, 2444, 2446}
HAM_ALLOWED = set(range(2447, 2519))   # 2447–2518 arası
DENIM_ALLOWED_RANGE = (2201, 2446)     # 2201–2446 arası


# =========================
# Yardımcılar
# =========================
def _U(x):
    return str(x).strip().upper() if pd.notna(x) else ""

def _first_int(s):
    if pd.isna(s): return None
    m = re.search(r"\d+", str(s))
    return int(m.group()) if m else None

def _to_num(s):
    if pd.isna(s) or s == "": return None
    try:
        return float(str(s).replace(",", "."))
    except Exception:
        return None

def _eta_from_durum(durum: str) -> int:
    u = _U(durum)
    if "SARMAYA HAZIR" in u or "SARMA" in u: return 999
    if "STOK" in u or "LEV" in u: return 0
    if "HAŞILA" in u: return 12
    if "AÇMA" in u: return 24
    if "BOYA" in u: return 40
    return 999

def _extract_numbers_preserve_decimal(text: str) -> list[str]:
    if text is None: return []
    nums = re.findall(r"[\d]+(?:[.,]\d+)?", str(text))
    out = []
    for n in nums:
        n = n.replace(",", ".")
        if re.fullmatch(r"\d+\.0+", n):
            n = n.split(".", 1)[0]
        out.append(n)
    return out

def _norm_tarak_generic(s: str) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    parts = _extract_numbers_preserve_decimal(str(s))
    if not parts:
        return str(s).strip()
    return "/".join(parts[:3])

def _col(df: pd.DataFrame, names: list[str]) -> str | None:
    for n in names:
        if n in df.columns: return n
    return None

def _loom_no_as_int(v) -> int | None:
    try:
        return int(re.search(r"\d+", str(v)).group())
    except Exception:
        return None

def _loom_digits(v) -> str:
    m = re.search(r"(\d+)", str(v))
    return m.group(1) if m else ""

def _group_type_from_dinamik(df_jobs: pd.DataFrame, col_tg: str, group_val) -> str:
    col_ihz = _col(df_jobs, ["İhzarat Boya Kodu","Ihzarat Boya Kodu","İhzaratBoyaKodu","IhzaratBoyaKodu"])
    col_boya= _col(df_jobs, ["Boya Kodu","BoyaKodu"])
    sub = df_jobs[df_jobs[col_tg].astype(str)==str(group_val)].copy()
    if sub.empty:
        return "denim"
    vals = []
    if col_ihz: vals += sub[col_ihz].astype(str).tolist()
    if col_boya: vals += sub[col_boya].astype(str).tolist()
    valsU = [ _U(x) for x in vals ]
    def is_ham(x):
        return x in ("HAM", "HAM HAM", "HAM HAM HAM")
    ham_ratio = sum(is_ham(x) for x in valsU) / (len(valsU) if valsU else 1)
    return "ham" if ham_ratio >= 0.5 else "denim"

def _loom_allowed(loom_no: int | None, grp_type: str) -> bool:
    if loom_no is None:
        return False
    if loom_no in NEVER:
        return False
    if grp_type == "ham":
        return loom_no in HAM_ALLOWED
    return DENIM_ALLOWED_RANGE[0] <= loom_no <= DENIM_ALLOWED_RANGE[1]

def _group_all_sarmaya_hazir(sub_df: pd.DataFrame) -> bool:
    col_durum = _col(sub_df, ["Durum Tanım", "Durum", "Durumu", "Durum Açıklaması"])
    if not col_durum:
        return False
    ser = sub_df[col_durum].astype(str).str.upper().str.strip()
    non_empty = ser[ser != ""]
    if non_empty.empty:
        return False
    def _is_ready(s: str) -> bool:
        return ("SARMAYA HAZIR" in s) or ("SARMA" in s)
    return non_empty.apply(_is_ready).all()
# =========================
# Manuel tezgâh seçme diyalogu (TAKIM)
# =========================
class ManualTezgahPicker(QDialog):
    """Tüm tezgah listesinden manuel seçim yapılabilen diyalog."""

    def __init__(self, df_running: pd.DataFrame, df_jobs_full: pd.DataFrame | None = None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Manuel Tezgah Seç")
        self.resize(820, 580)

        self._df_run = df_running.copy() if df_running is not None else pd.DataFrame()
        self._df_jobs = df_jobs_full.copy() if df_jobs_full is not None else None
        self._df_view = pd.DataFrame()
        self._chosen_loom: str | None = None

        v = QVBoxLayout(self)
        v.addWidget(QLabel("Tüm tezgahlar"))

        self.tbl = QTableView()
        cols = ["Tezgah", "Tarak Grubu", "Açık mı? / Kalan metre", "Tarak grubunun Kalan İş Adedi", "Kesim Tipi"]
        self.model = PandasModel(pd.DataFrame(columns=cols))
        self.tbl.setModel(self.model)
        self.tbl.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        v.addWidget(self.tbl, 1)

        btns = QHBoxLayout()
        self.btn_ok = QPushButton("Seç")
        self.btn_cancel = QPushButton("İptal")
        btns.addStretch(1)
        btns.addWidget(self.btn_ok)
        btns.addWidget(self.btn_cancel)
        v.addLayout(btns)

        self.btn_ok.clicked.connect(self._on_accept)
        self.btn_cancel.clicked.connect(self.reject)
        self.tbl.doubleClicked.connect(self._on_accept)

        self._build_view()

    def _jobs_total_for_tg(self, tg_norm: str) -> int:
        if self._df_jobs is None or self._df_jobs.empty:
            return 0
        col_tg = _col(self._df_jobs, ["Tarak Grubu", "Tarak", "TarakGrubu"])
        if not col_tg:
            return 0
        df = self._df_jobs.copy()
        df["_TG_norm"] = df[col_tg].astype(str).apply(_norm_tarak_generic)
        return int((df["_TG_norm"] == tg_norm).sum())

    def _build_view(self):
        cols_view = ["Tezgah", "Tarak Grubu", "Açık mı? / Kalan metre", "Tarak grubunun Kalan İş Adedi", "Kesim Tipi"]
        run = self._df_run
        if run is None or run.empty:
            self.model.set_df(pd.DataFrame(columns=cols_view))
            return

        col_tz = _col(run, ["Tezgah No", "Tezgah", "Tezgah Numarası"])
        if not col_tz:
            self.model.set_df(pd.DataFrame(columns=cols_view))
            return

        col_tg = _col(run, ["Tarak Grubu", "Tarak", "TarakGrubu"])
        col_kalan = _col(run, ["Kalan", "Kalan Mt", "Kalan Metre", "Kalan_Metre", "_KalanMetre"])
        col_cut = _col(run, ["ISAVER/ROTOCUT/ISAVERKit", "ISAVER/ROTOCUT", "Kesim Tipi", "KesimTipi", "Kesim"])

        r = run.copy()
        if col_tg:
            r["_TG_norm"] = r[col_tg].astype(str).apply(_norm_tarak_generic)
        else:
            r["_TG_norm"] = ""
        if "_OpenTezgahFlag" not in r.columns:
            r["_OpenTezgahFlag"] = False
        if "_KalanMetreNorm" not in r.columns:
            r["_KalanMetreNorm"] = pd.to_numeric(r[col_kalan], errors="coerce") if col_kalan else pd.NA

        r["_loom_no"] = r[col_tz].apply(lambda x: (_loom_no_as_int(x) or 99999))
        r = r.sort_values(by="_loom_no")

        view_rows = []
        for _, rr in r.iterrows():
            tezgah = str(rr[col_tz]).strip()
            tg_name = str(rr["_TG_norm"])
            jobs_total = self._jobs_total_for_tg(tg_name)

            if bool(rr["_OpenTezgahFlag"]):
                acik_kalan = "AÇIK"
            else:
                km = rr["_KalanMetreNorm"]
                acik_kalan = "" if pd.isna(km) else str(int(km))

            kesim_tipi = ""
            if col_cut and (col_cut in rr.index):
                val = rr.get(col_cut, "")
                kesim_tipi = "" if pd.isna(val) else str(val).strip()

            view_rows.append({
                "Tezgah": tezgah,
                "Tarak Grubu": tg_name,
                "Açık mı? / Kalan metre": acik_kalan,
                "Tarak grubunun Kalan İş Adedi": int(jobs_total),
                "Kesim Tipi": kesim_tipi
            })

        df_view = pd.DataFrame.from_records(view_rows, columns=cols_view)
        self._df_view = df_view
        self.model.set_df(df_view)
        self.tbl.resizeColumnsToContents()

    def _on_accept(self, *_):
        tz = self.selected_tezgah()
        if tz:
            self._chosen_loom = str(tz)
            self.accept()
        else:
            self.reject()

    def selected_tezgah(self) -> str | None:
        if getattr(self, "_chosen_loom", None):
            return self._chosen_loom
        idx = self.tbl.currentIndex()
        if not idx.isValid() or self._df_view is None or self._df_view.empty:
            return None
        try:
            return str(self._df_view.iloc[idx.row()]["Tezgah"]).strip()
        except Exception:
            return None

# =========================
# Tezgâh seçme diyalogu (TAKIM)
# =========================
class TezgahPicker(QDialog):
    """
    Adaylar:
      • Arkası boş TG (TOPLAM iş = 0, SH dahil)
      • Tezgahı fazla TG (aktif > iş) → sadece 'fazla' kadar
      • (Fallback) Hedef TG içindeki izinli tüm tezgâhlar (açık/açacak olmasa da)
    Sıra: ArkasıBoş → TezgahıFazla → (Fallback); her kovada AÇIK → (eşik altı) → loom no
    Tablo:
      [Tezgah | Tarak Grubu | Açık mı? / Kalan metre | Tarak grubunun Kalan İş Adedi | ISAVER/ROTOCUT]
    """
    def __init__(self, df_running: pd.DataFrame, target_tarak_norm: str, group_type: str,
                 soon_threshold_m: int = 300, parent=None,
                 df_jobs_full: pd.DataFrame | None = None,
                 exclude_looms: set[str] | None = None):
        super().__init__(parent)
        self.setWindowTitle("Tezgah Seç (TAKIM)")
        self.resize(820, 580)

        self._df_run = df_running.copy() if df_running is not None else pd.DataFrame()
        self._target_tg_norm = _norm_tarak_generic(target_tarak_norm)
        self._grp_target = (group_type or "denim").lower()
        self._df_jobs = df_jobs_full.copy() if df_jobs_full is not None else None
        # bağımsız metraj kısıtı (kalıcı)
        self._settings = QSettings("UZMANRAPOR", "ClientApp")
        self._thr = int(self._settings.value("team_flow/picker_threshold_m", int(soon_threshold_m or 300)))
        self._exclude = set(exclude_looms or [])

        # storage: arızalı (blocked) ve dummy (boş gösterilecek) tezgahlar
        self._blocked, self._dummy = _load_restricted_looms()
        # picker içinde kesinlikle hariç tut
        self._exclude.update(self._blocked)
        self._exclude.update(self._dummy)

        v = QVBoxLayout(self)
        # üst bar: bağımsız metraj kısıtı
        top = QHBoxLayout()
        top.addWidget(QLabel("Metraj kısıtı:"))
        self.spin_thr = QSpinBox()
        self.spin_thr.setRange(10, 10000)
        self.spin_thr.setSingleStep(10)
        self.spin_thr.setValue(self._thr)
        top.addWidget(self.spin_thr)
        self.btn_manual = QPushButton("Manuel Tezgah Seç")
        top.addStretch(1)
        top.addWidget(self.btn_manual)
        top.addStretch(1)
        v.addLayout(top)
        self.spin_thr.valueChanged.connect(self._on_thr_changed)
        self.btn_manual.clicked.connect(self._open_manual_picker)

        # tablo
        self.tbl = QTableView()
        cols = ["Tezgah", "Tarak Grubu", "Açık mı? / Kalan metre", "Tarak grubunun Kalan İş Adedi", "Kesim Tipi"]
        self.model = PandasModel(pd.DataFrame(columns=cols))
        self.tbl.setModel(self.model)
        self.tbl.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.tbl.setSortingEnabled(False)
        v.addWidget(self.tbl, 1)

        # butonlar
        btns = QHBoxLayout()
        self.btn_ok = QPushButton("Seç")
        self.btn_cancel = QPushButton("İptal")
        btns.addStretch(1)
        btns.addWidget(self.btn_ok)
        btns.addWidget(self.btn_cancel)
        v.addLayout(btns)

        self.btn_ok.clicked.connect(self._on_accept)
        self.btn_cancel.clicked.connect(self.reject)
        self.tbl.doubleClicked.connect(self._on_accept)

        self._df_candidates = pd.DataFrame()
        self._chosen_loom = None

        self._build_and_fill()

    @staticmethod
    def _col(df, names):
        for n in names:
            if n in df.columns:
                return n
        return None

    def _on_thr_changed(self, v: int):
        self._thr = int(v)
        self._settings.setValue("team_flow/picker_threshold_m", self._thr)
        self._build_and_fill()

    def _jobs_total_for_tg(self, tg_norm: str) -> int:
        if self._df_jobs is None or self._df_jobs.empty:
            return 0
        col_tg = self._col(self._df_jobs, ["Tarak Grubu","Tarak","TarakGrubu"])
        if not col_tg:
            return 0
        df = self._df_jobs.copy()
        df["_TG_norm"] = df[col_tg].astype(str).apply(_norm_tarak_generic)
        return int((df["_TG_norm"] == tg_norm).sum())

    def _active_looms_for_tg(self, tg_norm: str) -> int:
        run = self._df_run
        if run is None or run.empty:
            return 0
        col_tz = self._col(run, ["Tezgah No","Tezgah","Tezgah Numarası"])
        col_tg = self._col(run, ["Tarak Grubu","Tarak","TarakGrubu"])
        if not col_tz:
            return 0
        r = run.copy()
        if col_tg:
            r["_TG_norm"] = r[col_tg].astype(str).apply(_norm_tarak_generic)
            sub = r[r["_TG_norm"] == tg_norm]
        else:
            sub = r

        def allowed_row(x):
            loom = _loom_no_as_int(x[col_tz])
            d = _loom_digits(x[col_tz])
            return _loom_allowed(loom, self._grp_target) and (d not in self._exclude)

        return int(sub.apply(allowed_row, axis=1).sum())

    def _build_and_fill(self):
        cols_view = ["Tezgah","Tarak Grubu","Açık mı? / Kalan metre","Tarak grubunun Kalan İş Adedi","Kesim Tipi"]
        run = self._df_run
        if run is None or run.empty:
            self.model.set_df(pd.DataFrame(columns=cols_view)); return

        col_tz = self._col(run, ["Tezgah No","Tezgah","Tezgah Numarası"])
        col_tg = self._col(run, ["Tarak Grubu","Tarak","TarakGrubu"])
        r = run.copy()
        if col_tg:
            r["_TG_norm"] = r[col_tg].astype(str).apply(_norm_tarak_generic)
        else:
            r["_TG_norm"] = ""
        if "_OpenTezgahFlag" not in r.columns:
            r["_OpenTezgahFlag"] = False
        if "_KalanMetreNorm" not in r.columns:
            kal_col = self._col(r, ["Kalan","Kalan Mt","Kalan Metre","Kalan_Metre","_KalanMetre"])
            r["_KalanMetreNorm"] = pd.to_numeric(r[kal_col], errors="coerce") if kal_col else pd.NA

        # izin + exclude (blocked/dummy + haricen exclude)
        def allowed_row(x):
            loom = _loom_no_as_int(x[col_tz])
            d = _loom_digits(x[col_tz])
            return (
                (loom is not None)
                and _loom_allowed(loom, self._grp_target)
                and (d not in self._exclude)
            )
        r = r[r.apply(allowed_row, axis=1)].copy()

        # TG metrikleri
        tg_metrics = {}
        for tg, sub in r.groupby("_TG_norm", dropna=False):
            tg_norm = str(tg).strip()
            if tg_norm == "":
                continue
            jobs_total = self._jobs_total_for_tg(tg_norm)
            looms_act  = self._active_looms_for_tg(tg_norm)
            tg_metrics[tg_norm] = (jobs_total, looms_act)

        parts = []

        # bucket 0: arkası boş
        b0 = [tg for tg,(j,a) in tg_metrics.items() if j == 0]
        if b0:
            cand0 = r[r["_TG_norm"].isin(b0)].copy()
            cand0["_bucket"] = 0
            parts.append(cand0)

        # bucket 1: tezgahı fazla
        b1_info = [(tg, tg_metrics[tg][1] - tg_metrics[tg][0]) for tg,(j,a) in tg_metrics.items() if (j > 0 and a > j)]
        for tg_norm, extra in b1_info:
            sub_tg = r[r["_TG_norm"] == tg_norm].copy()
            if sub_tg.empty or extra <= 0:
                continue
            sub_tg["_open_prio"] = sub_tg["_OpenTezgahFlag"].apply(lambda b: 0 if b else 1)
            sub_tg["_kalan_ok"]  = sub_tg["_KalanMetreNorm"].apply(lambda v: 0 if (pd.notna(v) and v <= self._thr) else 1)
            sub_tg["_loom_no"]   = sub_tg[col_tz].apply(lambda x: (_loom_no_as_int(x) or 99999))
            sub_tg = sub_tg.sort_values(by=["_open_prio","_kalan_ok","_loom_no"], ascending=[True, True, True])

            picked = []
            for _, rr in sub_tg.iterrows():
                if len(picked) >= int(extra): break
                if (not bool(rr["_OpenTezgahFlag"])) and not (pd.notna(rr["_KalanMetreNorm"]) and rr["_KalanMetreNorm"] <= self._thr):
                    continue
                picked.append(rr)
            if picked:
                part = pd.DataFrame(picked)
                part["_bucket"] = 1
                parts.append(part)

        # bucket 2 (FALLBACK): hedef TG'deki izinli tüm tezgâhlar
        if not parts:
            if col_tg:
                r_same_tg = r[r["_TG_norm"] == self._target_tg_norm].copy()
            else:
                r_same_tg = r.copy()

            if not r_same_tg.empty:
                mask_viable = (
                        (r_same_tg["_OpenTezgahFlag"] == True)
                        | (
                                pd.notna(r_same_tg["_KalanMetreNorm"])
                                & (r_same_tg["_KalanMetreNorm"] <= self._thr)
                        )
                )
                r_same_tg = r_same_tg[mask_viable].copy()

                if not r_same_tg.empty:
                    tz_col = self._col(r_same_tg, ["Tezgah No", "Tezgah", "Tezgah Numarası"])
                    r_same_tg["_kalan_ok"] = r_same_tg["_KalanMetreNorm"].apply(
                        lambda v: 0 if (pd.notna(v) and float(v) <= self._thr) else 1
                    )
                    r_same_tg["_loom_no"] = r_same_tg[tz_col].apply(lambda x: (_loom_no_as_int(x) or 99999))

                    r_same_tg = r_same_tg.sort_values(
                        by=["_kalan_ok", "_loom_no"],
                        ascending=[True, True],
                        na_position="last"
                    )
                    part2 = r_same_tg.copy()
                    part2["_bucket"] = 2
                    parts.append(part2)

        if parts:
            cand = pd.concat(parts, ignore_index=True)
        else:
            cand = pd.DataFrame(columns=[col_tz, "_TG_norm", "_OpenTezgahFlag", "_KalanMetreNorm", "_bucket"])
        if not cand.empty:
            mask_viable = (
                    (cand["_OpenTezgahFlag"] == True)
                    | (
                            pd.notna(cand["_KalanMetreNorm"])
                            & (cand["_KalanMetreNorm"] <= self._thr)
                    )
            )
            cand = cand[mask_viable].copy()

        if not cand.empty:
            # Kullanıcı, hedef tarak grubunun kendi tezgahlarının listelenmesini istemiyor.
            cand["_TG_norm"] = cand["_TG_norm"].astype(str).str.strip()
            cand = cand[cand["_TG_norm"] != self._target_tg_norm]
        if not cand.empty:
            cand["_open_prio"] = cand["_OpenTezgahFlag"].apply(lambda b: 0 if b else 1)
            cand["_kalan_ok"]  = cand["_KalanMetreNorm"].apply(lambda v: 0 if (pd.notna(v) and v <= self._thr) else 1)
            cand["_loom_no"]   = cand[col_tz].apply(lambda x: (_loom_no_as_int(x) or 99999))
            cand = cand.sort_values(by=["_bucket","_open_prio","_kalan_ok","_loom_no"],
                                    ascending=[True, True, True, True], na_position="last")

        # Running’deki kesim tipi kolonu
        col_cut = self._col(r, ["ISAVER/ROTOCUT/ISAVERKit", "ISAVER/ROTOCUT", "Kesim Tipi"])

        # görünüm
        view_rows = []
        for _, rr in cand.iterrows():
            tezgah = str(rr[col_tz])
            tg_name = str(rr["_TG_norm"])
            jobs_total = self._jobs_total_for_tg(tg_name)

            if bool(rr["_OpenTezgahFlag"]):
                acik_kalan = "AÇIK"
            else:
                km = rr["_KalanMetreNorm"]
                acik_kalan = ("" if pd.isna(km) else str(int(km)))

            kesim_tipi = ""
            if col_cut and (col_cut in rr.index):
                val = rr.get(col_cut, "")
                kesim_tipi = "" if pd.isna(val) else str(val).strip()

            view_rows.append({
                "Tezgah": tezgah,
                "Tarak Grubu": tg_name,
                "Açık mı? / Kalan metre": acik_kalan,
                "Tarak grubunun Kalan İş Adedi": int(jobs_total),
                "Kesim Tipi": kesim_tipi
            })

        df_view = pd.DataFrame.from_records(view_rows, columns=cols_view)
        self._df_candidates = df_view
        self.model.set_df(df_view)
        self.tbl.resizeColumnsToContents()

    def _open_manual_picker(self):
        dlg = ManualTezgahPicker(self._df_run, self._df_jobs, parent=self)
        if dlg.exec():
            tz = dlg.selected_tezgah()
            if tz:
                self._chosen_loom = str(tz)
                self.accept()

    def _on_accept(self, *_):
        tz = self.selected_tezgah()
        if tz:
            self._chosen_loom = str(tz)
            try:
                idx = self.tbl.currentIndex().row()
                if 0 <= idx < len(self._df_candidates):
                    self._df_candidates = self._df_candidates.drop(index=idx).reset_index(drop=True)
                    self.model.set_df(self._df_candidates)
            except Exception:
                pass
            self.accept()
        else:
            self.reject()

    def selected_tezgah(self) -> str | None:
        if getattr(self, "_chosen_loom", None):
            return self._chosen_loom
        idx = self.tbl.currentIndex()
        if not idx.isValid() or self._df_candidates is None or self._df_candidates.empty:
            return None
        try:
            return str(self._df_candidates.iloc[idx.row()]["Tezgah"]).strip()
        except Exception:
            return None

# =========================
# Ana Sekme
# =========================
class TeamPlanningFlowTab(QWidget):
    """
    Sol: Tarak grupları — DENIM/HAM
    Orta: Seçili grubun işleri
    Sağ: TAKIM ATAMALARI
    """
    def __init__(self, mainwin):
        super().__init__(parent=mainwin)
        self.main = mainwin
        self.df_jobs: pd.DataFrame | None = None
        self.df_run: pd.DataFrame | None = None
        self.team_rows: list[dict] = []

        self.missing_open = set()
        self.missing_soon = set()

        self._can_write = True

        self.settings = QSettings("UZMANRAPOR", "ClientApp")
        self.flow_threshold_m = int(self.settings.value("team_flow/soon_threshold_m", 300))

        # storage'dan arızalı/dummy listeleri al
        self._blocked_looms, self._dummy_looms = _load_restricted_looms()

        root = QVBoxLayout(self)

        # üst bar: eşik (sadece DÜĞÜM için)
        top = QHBoxLayout()
        top.addWidget(QLabel("Açacak ≤"))
        self.spin_flow_threshold = QSpinBox()
        self.spin_flow_threshold.setRange(10, 5000)
        self.spin_flow_threshold.setSingleStep(10)
        self.spin_flow_threshold.setValue(self.flow_threshold_m)
        top.addWidget(self.spin_flow_threshold)
        top.addWidget(QLabel("m"))
        top.addStretch(1)
        root.addLayout(top)
        self.spin_flow_threshold.valueChanged.connect(self._on_threshold_changed)

        hdr = QLabel("Tarak grubu planlama — Sol: DENIM/HAM • Orta: İşler • Sağ: TAKIM ATAMALARI")
        hdr.setWordWrap(True)
        root.addWidget(hdr)

        # ---- YENİ YERLEŞİM: Sol sabit, sağda dikey splitter (Üst=İşler, Alt=Takım) ----
        split = QSplitter(Qt.Horizontal, self)

        # sol
        left = QWidget();
        l = QVBoxLayout(left)
        self.lst_groups = QListWidget()
        l.addWidget(QLabel("Tarak Grupları"))
        l.addWidget(self.lst_groups, 1)
        split.addWidget(left)

        # orta (üst panel): Seçili Grubun İşleri
        mid = QWidget();
        m = QVBoxLayout(mid)
        self.tbl_jobs = QTableView()
        self.model_jobs = PandasModel(pd.DataFrame())
        self.tbl_jobs.setModel(self.model_jobs)
        self.tbl_jobs.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.tbl_jobs.setSortingEnabled(True)
        m.addWidget(QLabel("Seçili Grubun İşleri"))
        m.addWidget(self.tbl_jobs, 1)

        # sağ (alt panel): TAKIM ATAMALARI (başlık + export butonu)
        right = QWidget();
        r = QVBoxLayout(right)

        row_hdr = QHBoxLayout()
        row_hdr.addWidget(QLabel("TAKIM ATAMALARI"))
        row_hdr.addStretch(1)
        self.btn_reset_team = QPushButton("Seçimleri Sıfırla")
        self.btn_reset_team.clicked.connect(self._reset_team_assignments)
        row_hdr.addWidget(self.btn_reset_team)
        self.btn_export_team = QPushButton("TAKIM LİSTESİNİ DIŞA AKTAR (Excel)")
        self.btn_export_team.clicked.connect(self._export_team_assignments)
        row_hdr.addWidget(self.btn_export_team)
        r.addLayout(row_hdr)

        self.tbl_team = QTableView()
        self.model_team = PandasModel(pd.DataFrame())
        self.tbl_team.setModel(self.model_team)
        self.tbl_team.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        r.addWidget(self.tbl_team, 1)

        # sağ taraf için dikey splitter
        right_split = QSplitter(Qt.Vertical, self)
        right_split.addWidget(mid)   # ÜST: İşler
        right_split.addWidget(right) # ALT: Takım
        right_split.setSizes([600, 500])

        # ana yatay split: sol panel + sağ dikey split
        split.addWidget(right_split)
        split.setSizes([55, 1485])  # Varsayılan genişlikler

        root.addWidget(split, 1)

        self.lbl_missing = QLabel("")
        root.addWidget(self.lbl_missing)

        # events
        self.lst_groups.currentRowChanged.connect(self._bind_group_jobs)
        self.tbl_jobs.doubleClicked.connect(self._assign_on_doubleclick)

        # oturum içi durumlar
        self._assignments = {}           # (tg_norm, job_key) -> "2248 (DÜĞÜM)" | "2248"
        self._used_looms = {}            # (tg_norm, grp_type) -> set(...) (sadece DÜĞÜM sırası)
        self._ordered_looms_cache = {}   # (tg_norm, grp_type) -> [loom...]
        self._used_looms_global = set()  # tüm atamalar
        self._picker_open = False

        self.refresh_sources()
        try:
            can_write = bool(getattr(self.main, "has_permission", lambda *_: True)("write"))
            self.set_write_enabled(can_write)
        except Exception:
            pass

    def set_write_enabled(self, enabled: bool):
        self._can_write = bool(enabled)
        if hasattr(self, "tbl_jobs") and self.tbl_jobs is not None:
            if self._can_write:
                self.tbl_jobs.setToolTip("")
            else:
                self.tbl_jobs.setToolTip("Takım ataması yapma yetkiniz yok.")

    # --- yardımcı: başlığa göre daraltma ---
    def _shrink_columns_by_header(self, table: QTableView, min_width: int = 60, pad: int = 28):
        header = table.horizontalHeader()
        model = table.model()
        if not model:
            return

        fm = table.fontMetrics()
        for col in range(model.columnCount()):
            title = str(model.headerData(col, Qt.Horizontal) or "")
            header_w = fm.horizontalAdvance(title) + pad
            header_w = max(min_width, header_w)
            content_w = header.sectionSize(col)
            target = max(header_w, content_w)
            header.resizeSection(col, target)

    def _on_threshold_changed(self, v: int):
        self.flow_threshold_m = int(v)
        self.settings.setValue("team_flow/soon_threshold_m", self.flow_threshold_m)
        self._ordered_looms_cache.clear()
        self._bind_group_jobs()

    # ---------------- Data binding ----------------
    def refresh_sources(self):
        self.df_jobs = getattr(self.main, "df_dinamik_full", None)
        self.df_run  = getattr(self.main, "df_running", None)
        self._rebuild_groups()

    def _rebuild_groups(self):
        self.lst_groups.clear()
        self.model_jobs.set_df(pd.DataFrame())
        self.missing_open.clear()
        self.missing_soon.clear()

        if self.df_jobs is None or self.df_jobs.empty:
            self._update_missing_label()
            return

        col_tg = _col(self.df_jobs, ["Tarak Grubu", "Tarak", "TarakGrubu"])
        col_lev = _col(self.df_jobs, ["Levent No", "Levent", "Levent Etiket FA"])
        col_drm = _col(self.df_jobs, ["Durum Tanım", "Durum", "Durum Açıklaması"])
        if not col_tg:
            self._update_missing_label()
            return

        df = self.df_jobs.copy()
        df["_TG_norm"] = df[col_tg].astype(str).apply(_norm_tarak_generic)

        def _row_visible(r) -> bool:
            if col_lev:
                lv = str(r.get(col_lev, "")).strip()
                if lv:
                    return True
            if not col_drm:
                return True
            u = _U(r.get(col_drm, ""))
            return not (("SARMAYA HAZIR" in u) or ("SARMA" in u))

        def _row_cat(r) -> str:
            v = str(r.get("_DyeCategory", "")).upper() if "_DyeCategory" in r else ""
            if v in ("HAM", "DENIM"):
                return v
            c_ihz = _col(self.df_jobs, ["İhzarat Boya Kodu", "Ihzarat Boya Kodu", "İhzaratBoyaKodu", "IhzaratBoyaKodu"])
            c_boy = _col(self.df_jobs, ["Boya Kodu", "BoyaKodu"])
            text = (str(r.get(c_ihz, "")) + " " + str(r.get(c_boy, ""))).upper()
            return "HAM" if "HAM" in text else "DENIM"

        groups_denim, groups_ham = [], []
        for tg, sub in df.groupby("_TG_norm", dropna=False):
            if str(tg).strip() == "":
                continue
            if _group_all_sarmaya_hazir(sub):
                continue
            sub_vis = sub[sub.apply(_row_visible, axis=1)]
            if sub_vis.empty:
                continue
            has_denim = (sub_vis.apply(_row_cat, axis=1) == "DENIM").any()
            has_ham   = (sub_vis.apply(_row_cat, axis=1) == "HAM").any()
            if has_denim: groups_denim.append(tg)
            if has_ham:   groups_ham.append(tg)

        sd = sorted([str(x) for x in groups_denim], key=lambda x: (_first_int(x) is None, _first_int(x), x))
        sh = sorted([str(x) for x in groups_ham], key=lambda x: (_first_int(x) is None, _first_int(x), x))

        if sd:
            hdr = QListWidgetItem("— DENIM —"); hdr.setFlags(Qt.ItemIsEnabled)
            self.lst_groups.addItem(hdr)
            for g in sd: self.lst_groups.addItem(str(g))
        if sh:
            hdr = QListWidgetItem("— HAM —"); hdr.setFlags(Qt.ItemIsEnabled)
            self.lst_groups.addItem(hdr)
            for g in sh: self.lst_groups.addItem(str(g))

        for i in range(self.lst_groups.count()):
            if "— " not in self.lst_groups.item(i).text():
                self.lst_groups.setCurrentRow(i); break
        if self.lst_groups.currentRow() == -1:
            self.model_jobs.set_df(pd.DataFrame())
        self._update_missing_label()

    def _current_group(self):
        it = self.lst_groups.currentItem()
        return it.text() if it and "— " not in it.text() else None

    def _current_category_from_list(self) -> str | None:
        it = self.lst_groups.currentItem()
        if not it: return None
        row = self.lst_groups.currentRow()
        for i in range(row, -1, -1):
            t = self.lst_groups.item(i).text().strip()
            if t == "— DENIM —": return "denim"
            if t == "— HAM —":   return "ham"
        return None

    def _bind_group_jobs(self):
        g = self._current_group()
        grp_type = self._current_category_from_list()
        df = self._jobs_of_group(g, grp_type=grp_type)
        self.model_jobs.set_df(df)
        self.tbl_jobs.resizeColumnsToContents()
        self._shrink_columns_by_header(self.tbl_jobs)

    # -------- Orta tablo --------
    def _jobs_of_group(self, group_val, grp_type: str | None = None):
        if self.df_jobs is None or self.df_jobs.empty or group_val is None:
            return pd.DataFrame()

        # kolonlar
        col_is   = _col(self.df_jobs, ["Üretim Sipariş No","Dokuma İş Emri","İş Emri","Sipariş No"])
        col_tg   = _col(self.df_jobs, ["Tarak Grubu","Tarak","TarakGrubu"])
        col_kok  = _col(self.df_jobs, ["Kök Tip Kodu","KökTip","KokTip"])
        col_lev  = _col(self.df_jobs, ["Levent No","Levent","Levent Etiket FA"])
        col_drm  = _col(self.df_jobs, ["Durum Tanım","Durum","Durum Açıklaması"])
        col_zorg = _col(self.df_jobs, ["Zemin Örgü","Zemin Örgü Kodu","Zemin Örgü Adı"])
        col_coz1 = _col(self.df_jobs, ["Çözgü İpliği 1","Cozgu Ipligi 1","Cozgu 1"])
        col_atk1 = _col(self.df_jobs, ["Atkı İpliği 1","Atki Ipligi 1","Atki 1"])
        col_mtr  = _col(self.df_jobs, ["Parti Metresi","Metre","Parti Mt"])
        col_term = _col(self.df_jobs, ["Mamül Termin","Mamul Termin","Termin","Termin Tarihi"])
        col_note = _col(self.df_jobs, ["NOTLAR","Notlar","Not"])
        col_hash = _col(self.df_jobs, ["Levent Haşıl Tarihi","Haşıl Tarihi"])
        col_tz = _col(self.df_jobs, ["Tezgah Numarası", "Tezgah No", "Tezgah"])

        if not col_tg:
            return pd.DataFrame()

        # normalize TG
        df_all = self.df_jobs.copy()
        df_all["_TG_norm"] = df_all[col_tg].astype(str).apply(_norm_tarak_generic)
        key_norm = _norm_tarak_generic(group_val)

        sub_norm = df_all[df_all["_TG_norm"] == key_norm]
        if sub_norm.empty:
            df = self.df_jobs[self.df_jobs[col_tg].astype(str) == str(group_val)].copy()
        else:
            df = sub_norm.copy()
        if df.empty:
            return pd.DataFrame()

        if grp_type is None:
            try:
                grp_type = self._current_category_from_list()
            except Exception:
                grp_type = None
        if grp_type is None:
            grp_type = _group_type_from_dinamik(self.df_jobs, col_tg, group_val)
        grp_type = "ham" if str(grp_type).lower() == "ham" else "denim"

        # kategori filtresi
        if "_DyeCategory" in df.columns:
            want = "HAM" if grp_type == "ham" else "DENIM"
            df = df[df["_DyeCategory"].astype(str).str.upper() == want]
        if df.empty:
            return pd.DataFrame()

        # adetler (blocked/dummy hariç)
        open_cnt, open_ok = self._open_looms_count(key_norm, grp_type)
        soon_cnt, soon_ok = self._soon_looms_count(key_norm, grp_type)
        if not open_ok: self.missing_open.add(key_norm)
        if not soon_ok: self.missing_soon.add(key_norm)

        # SH hariç
        def lev_or_durum(r):
            lev = r.get(col_lev, None)
            if pd.notna(lev) and str(lev).strip():
                return str(lev)
            d = str(r.get(col_drm, ""))
            u = _U(d)
            if "SARMAYA HAZIR" in u or "SARMA" in u:
                return ""
            return d

        # DokumaİşEmri = Üretim Sipariş No + sayaç
        order_counters = defaultdict(int)
        rows = []
        initial_assignments: list[tuple[str, str, str, str]] = []  # (tg_norm, job_key, display, digits)
        for _, r in df.iterrows():
            ld = lev_or_durum(r)
            if ld == "":
                continue
            order_no = str(r.get(col_is, "")).strip() if col_is else ""
            if not order_no:
                order_no = "NO_ORDER"
            order_counters[order_no] += 1
            dok_is = f"{order_no}-{order_counters[order_no]}"

            tezgah_disp = ""
            if col_tz:
                tz_raw = r.get(col_tz, "")
                if pd.notna(tz_raw):
                    tz_str = str(tz_raw).strip()
                    if tz_str:
                        tz_digits = _loom_digits(tz_str) or tz_str
                        tezgah_disp = f"{tz_digits}  (DÜĞÜM)"
                        initial_assignments.append((key_norm, dok_is, tezgah_disp, tz_digits))

            rows.append({
                "Tezgah": tezgah_disp,
                "Tarak Grubu": r.get(col_tg, ""),
                "KökTip": r.get(col_kok, ""),
                "LeventNo / Durum": ld,
                "ZeminÖrgü": r.get(col_zorg, ""),
                "Çözgü İpliği 1": r.get(col_coz1, ""),
                "Atkı İpliği 1": r.get(col_atk1, ""),
                "Metre": r.get(col_mtr, ""),
                "Mamül Termin": r.get(col_term, ""),
                "NOTLAR": r.get(col_note, ""),
                "Levent Haşıl Tarihi": r.get(col_hash, ""),
                "Açık Tezgah (adet)": open_cnt,
                "Açacak Tezgah (adet)": soon_cnt,
                "DokumaİşEmri": dok_is,     # <-- ANAHTAR
            })

        mid = pd.DataFrame.from_records(rows)
        # Tarih sütunlarında "00:00:00" kısmını temizle
        for col in ["Mamül Termin", "Levent Haşıl Tarihi"]:
            if col in mid.columns:
                mid[col] = (
                    mid[col]
                    .astype(str)
                    .str.replace(r"\s*00:00:00\s*", "", regex=True)
                    .str.strip()
                )

        # sıralama
        def rank(v: str) -> int:
            u = _U(v)
            if re.fullmatch(r"\d+", str(v)): return 0
            if "HAŞILA" in u: return 1
            if "AÇMA"   in u: return 2
            if "BOYA"   in u: return 3
            return 9

        if not mid.empty:
            mid["_rank"] = mid["LeventNo / Durum"].apply(rank)
            if "Mamül Termin" in mid.columns:
                mid = mid.sort_values(by=["_rank","Mamül Termin"], ascending=[True, True], ignore_index=True)
            else:
                mid = mid.sort_values(by=["_rank"], ascending=[True], ignore_index=True)
            mid = mid.drop(columns=["_rank"])

        # DokumaİşEmri'ni 10. sütuna taşı
        if "DokumaİşEmri" in mid.columns:
            cols = mid.columns.tolist()
            cols.remove("DokumaİşEmri")
            insert_at = min(9, len(cols))  # 0-based index → 9 = 10. sütun
            cols.insert(insert_at, "DokumaİşEmri")
            mid = mid[cols]

            # Düğüm sekmesinden gelen hazır tezgah atamalarını kaydet
            for tg_norm, job_key, tz_disp, tz_digits in initial_assignments:
                if (tg_norm, job_key) not in self._assignments:
                    self._assignments[(tg_norm, job_key)] = tz_disp
                if tz_digits:
                    self._used_looms_global.add(str(tz_digits))
        # önceki atamaları geri yaz
        try:
            if not mid.empty:
                filled = []
                for i in range(len(mid)):
                    row = mid.iloc[i].to_dict()
                    job_key = self._make_job_key(row)
                    prev = getattr(self, "_assignments", {}).get((key_norm, job_key))
                    if prev:
                        row["Tezgah"] = prev
                    filled.append(row)
                mid = pd.DataFrame.from_records(filled)
        finally:
            self._update_missing_label()

        return mid

    # -------- Kesim tipi running'den çekme
    def _lookup_cut_type(self, loom_no: str) -> str:
        """df_running içinden verilen tezgâhın kesim tipini döndürür."""
        try:
            if self.df_run is None or self.df_run.empty:
                return ""
            col_tz = _col(self.df_run, ["Tezgah No","Tezgah","Tezgah Numarası"])
            if not col_tz:
                return ""
            col_cut = _col(self.df_run, ["ISAVER/ROTOCUT/ISAVERKit", "ISAVER/ROTOCUT", "Kesim Tipi"])
            if not col_cut:
                return ""
            m = re.search(r"\d+", str(loom_no))
            tz_digits = m.group(0) if m else str(loom_no).strip()
            sub = self.df_run[self.df_run[col_tz].astype(str).str.contains(rf"\b{tz_digits}\b", regex=True, na=False)]
            if sub.empty:
                return ""
            val = sub.iloc[0].get(col_cut, "")
            return "" if pd.isna(val) else str(val).strip()
        except Exception:
            return ""

    def _prompt_new_cut_type(self, loom_no: str) -> str:
        current_cut = self._lookup_cut_type(loom_no)
        msg = "Yeni Kesim Tipi seçin."
        if current_cut:
            msg = f"Yeni Kesim Tipi seçin.\n(Mevcut tezgah kesim tipi: {current_cut})"
            options = ["ISAVER", "ROTOCUT", "ISAVERKit"]
            selection, ok = QInputDialog.getItem(self, "Yeni Kesim Tipi", msg, options, 0, False)
        if not ok:
            return ""
        return str(selection).strip()

    def _prompt_missing_dugum_choice(self) -> str | None:
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Warning)
        box.setWindowTitle("Uyarı")
        box.setText(
            "Bu iş düğüm verilmemiş. Süs Kenar ve ya Örgü Uyumsuz olabilir. Kontrol Et."
        )
        btn_dugum = box.addButton("Düğüm", QMessageBox.AcceptRole)
        btn_takim = box.addButton("Takım", QMessageBox.ActionRole)
        box.exec()
        clicked = box.clickedButton()
        if clicked == btn_dugum:
            return "dugum"
        if clicked == btn_takim:
            return "takim"
        return None

    def _refresh_jobs_table(self, df: pd.DataFrame) -> None:
        self.model_jobs.set_df(df)
        self.tbl_jobs.resizeColumnsToContents()
        self._shrink_columns_by_header(self.tbl_jobs)

    def _assign_dugum_to_row(self, df: pd.DataFrame, row: int, target_tg_norm: str, grp_type: str) -> bool:
        next_loom = self._next_free_loom(target_tg_norm, grp_type or "denim")
        if not next_loom:
            return False
        df.at[row, "Tezgah"] = f"{next_loom}  (DÜĞÜM)"
        self._refresh_jobs_table(df)
        row_dict = df.iloc[row].to_dict()
        job_key = self._make_job_key(row_dict)
        self._assignments[(target_tg_norm, job_key)] = f"{next_loom}  (DÜĞÜM)"
        self._used_looms_global.add(str(next_loom))
        return True

    def _assign_team_to_row(self, df: pd.DataFrame, row: int, target_tg_norm: str, grp_type: str) -> None:
        exclude = set(self._used_looms_global)
        for v in self._assignments.values():
            m = re.search(r"\d+", str(v))
            if m:
                exclude.add(m.group())
        exclude.update(self._blocked_looms)
        exclude.update(self._dummy_looms)

        dlg = TezgahPicker(
            self.df_run,
            target_tg_norm,
            grp_type or "denim",
            soon_threshold_m=self.flow_threshold_m,
            parent=self,
            df_jobs_full=self.df_jobs,
            exclude_looms=exclude
        )
        if not dlg.exec():
            return
        tz = dlg.selected_tezgah()
        if not tz:
            return
        df.at[row, "Tezgah"] = str(tz)
        self._refresh_jobs_table(df)

        row_dict = df.iloc[row].to_dict()
        job_key = self._make_job_key(row_dict)
        self._assignments[(target_tg_norm, job_key)] = str(tz)
        self._used_looms_global.add(str(tz))

        filtered = {k: v for k, v in row_dict.items()
                    if k not in ["Açık Tezgah (adet)", "Açacak Tezgah (adet)"]}
        filtered["Kesim Tipi"] = self._prompt_new_cut_type(tz)

        self.team_rows.append(filtered)
        self.model_team.set_df(pd.DataFrame.from_records(self.team_rows))
        self.tbl_team.resizeColumnsToContents()
        self._shrink_columns_by_header(self.tbl_team)
    # ---------------- Çift tık ile atama ----------------
    def _assign_on_doubleclick(self, idx: QModelIndex):
        if not self._can_write:
            QMessageBox.warning(self, "Yetki yok", "Takım ataması yapma yetkiniz bulunmuyor.")
            return
        self._picker_open = True
        try:
            if not idx.isValid():
                return
            row = idx.row()
            df = getattr(self.model_jobs, "_df", pd.DataFrame()).copy()
            if df.empty:
                return

            target_tg_job  = str(df.at[row, "Tarak Grubu"])
            target_tg_norm = _norm_tarak_generic(target_tg_job)
            try:
                grp_type = self._current_category_from_list()
            except Exception:
                grp_type = _group_type_from_dinamik(self.df_jobs, _col(self.df_jobs, ["Tarak Grubu","Tarak","TarakGrubu"]), target_tg_job)
            tezgah_col = df.columns.get_loc("Tezgah") if "Tezgah" in df.columns else None
            open_cnt = df.at[row, "Açık Tezgah (adet)"] if "Açık Tezgah (adet)" in df.columns else 0
            try:
                open_cnt_val = int(float(str(open_cnt).replace(",", ".")))
            except Exception:
                open_cnt_val = 0
            tezgah_value = str(df.at[row, "Tezgah"]) if "Tezgah" in df.columns else ""
            has_any_dugum = False
            if "Tezgah" in df.columns:
                try:
                    group_mask = df["Tarak Grubu"].astype(str) == str(target_tg_job)
                    group_values = df.loc[group_mask, "Tezgah"].astype(str)
                    has_any_dugum = group_values.str.contains("DÜĞÜM", case=False, na=False).any()
                except Exception:
                    has_any_dugum = False
            needs_prompt = (
                    tezgah_col is not None
                    and idx.column() == tezgah_col
                    and open_cnt_val != 0
                    and "DÜĞÜM" not in tezgah_value.upper()
                    and not has_any_dugum
            )
            if needs_prompt:
                choice = self._prompt_missing_dugum_choice()
                if choice == "dugum":
                    if not self._assign_dugum_to_row(df, row, target_tg_norm, grp_type or "denim"):
                        QMessageBox.information(self, "Bilgi", "Uygun düğüm tezgahı bulunamadı.")
                elif choice == "takim":
                    self._assign_team_to_row(df, row, target_tg_norm, grp_type or "denim")
                return

            # 1) DÜĞÜM
            if self._assign_dugum_to_row(df, row, target_tg_norm, grp_type or "denim"):
                return

            # 2) TAKIM (Picker)
            self._assign_team_to_row(df, row, target_tg_norm, grp_type or "denim")
        finally:
            self._picker_open = False

    # ---------------- Running sayımları (ADET) ----------------
    def _open_looms_count(self, target_tarak_norm: str, grp_type: str) -> tuple[int, bool]:
        run = self.df_run
        if run is None or run.empty:
            return 0, False
        col_tz = _col(run, ["Tezgah No", "Tezgah", "Tezgah Numarası"])
        col_tg = _col(run, ["Tarak Grubu","Tarak","TarakGrubu"])
        if not col_tz:
            return 0, False
        r = run.copy()

        if "_OpenTezgahFlag" not in r.columns:
            r["_OpenTezgahFlag"] = False

        if col_tg:
            r["_TG_norm"] = r[col_tg].astype(str).apply(_norm_tarak_generic)
            sub = r[r["_TG_norm"] == target_tarak_norm]
        else:
            sub = r

        if sub.empty:
            return 0, False

        def allowed_row(x):
            loom = _loom_no_as_int(x[col_tz])
            d = _loom_digits(x[col_tz])
            return _loom_allowed(loom, grp_type) and (d not in self._blocked_looms) and (d not in self._dummy_looms)

        sub = sub[sub.apply(allowed_row, axis=1)]

        cnt = int((sub["_OpenTezgahFlag"] == True).sum())
        return cnt, True

    def _soon_looms_count(self, target_tarak_norm: str, grp_type: str) -> tuple[int, bool]:
        run = self.df_run
        if run is None or run.empty:
            return 0, False
        col_tz = _col(run, ["Tezgah No", "Tezgah", "Tezgah Numarası"])
        col_tg = _col(run, ["Tarak Grubu","Tarak","TarakGrubu"])
        if not col_tz:
            return 0, False
        r = run.copy()

        if "_KalanMetreNorm" not in r.columns:
            kal_col = _col(r, ["Kalan","Kalan Mt","Kalan Metre","Kalan_Metre","_KalanMetre"])
            if kal_col:
                r["_KalanMetreNorm"] = pd.to_numeric(r[kal_col], errors="coerce")
            else:
                r["_KalanMetreNorm"] = pd.NA

        if "_OpenTezgahFlag" not in r.columns:
            r["_OpenTezgahFlag"] = False

        if col_tg:
            r["_TG_norm"] = r[col_tg].astype(str).apply(_norm_tarak_generic)
            sub = r[r["_TG_norm"] == target_tarak_norm]
        else:
            sub = r

        if sub.empty:
            return 0, False

        def allowed_row(x):
            loom = _loom_no_as_int(x[col_tz])
            d = _loom_digits(x[col_tz])
            return _loom_allowed(loom, grp_type) and (d not in self._blocked_looms) and (d not in self._dummy_looms)

        sub = sub[sub.apply(allowed_row, axis=1)]

        thr = int(self.flow_threshold_m or 300)
        cnt = int(((sub["_KalanMetreNorm"] <= thr) & (sub["_OpenTezgahFlag"] != True)).sum())

        return cnt, True

    def _first_open_loom_same_tarak(self, target_tarak_norm: str, grp_type: str) -> str | None:
        run = self.df_run
        if run is None or run.empty:
            return None
        col_tz = _col(run, ["Tezgah No", "Tezgah", "Tezgah Numarası"])
        col_tg = _col(run, ["Tarak Grubu","Tarak","TarakGrubu"])
        if not col_tz:
            return None
        r = run.copy()
        if "_OpenTezgahFlag" not in r.columns:
            r["_OpenTezgahFlag"] = False
        if col_tg:
            r["_TG_norm"] = r[col_tg].astype(str).apply(_norm_tarak_generic)
            sub = r[(r["_TG_norm"] == target_tarak_norm) & (r["_OpenTezgahFlag"]==True)]
        else:
            sub = r[r["_OpenTezgahFlag"]==True]
        if sub.empty:
            return None

        for _, row in sub.iterrows():
            loom = _loom_no_as_int(row[col_tz])
            d = _loom_digits(row[col_tz])
            if _loom_allowed(loom, grp_type) and (d not in self._blocked_looms) and (d not in self._dummy_looms):
                return str(row[col_tz])
        return None

    # Auto "DÜĞÜM" adayları
    def _ordered_candidate_looms(self, tg_norm: str, grp_type: str) -> list[str]:
        key = (tg_norm, grp_type)
        if key in self._ordered_looms_cache:
            return self._ordered_looms_cache[key][:]

        run = self.df_run
        if run is None or run.empty:
            self._ordered_looms_cache[key] = []
            return []

        def _norm_name(s: str) -> str:
            if s is None:
                return ""
            s = str(s).replace("\u00A0", " ")  # NBSP → space
            s = re.sub(r"\s+", " ", s).strip()
            return s.upper()

        def _find_tz_col(df: pd.DataFrame) -> str | None:
            # 1) bilinen adlar
            for k in ["Tezgah No", "Tezgah", "Tezgah Numarası"]:
                if k in df.columns:
                    return k
            # 2) normalize ederek ara
            norm_map = {_norm_name(c): c for c in df.columns}
            for norm, orig in norm_map.items():
                if ("TEZGAH" in norm) and ("NO" in norm):
                    return orig
            # 3) son çare: çoğunlukla rakam içeren ilk kolon
            for c in df.columns:
                try:
                    s = df[c].astype(str)
                    ratio = (s.str.contains(r"\d", regex=True, na=False)).mean()
                    if ratio > 0.8:
                        return c
                except Exception:
                    continue
            return None

        col_tg = _col(run, ["Tarak Grubu", "Tarak", "TarakGrubu"])

        r = run.copy()
        # ---- GUARD (r)
        if "_OpenTezgahFlag" not in r.columns:
            r["_OpenTezgahFlag"] = False
        if "_KalanMetreNorm" not in r.columns:
            kal_col = _col(r, ["Kalan", "Kalan Mt", "Kalan Metre", "Kalan_Metre", "_KalanMetre"])
            r["_KalanMetreNorm"] = pd.to_numeric(r[kal_col], errors="coerce") if kal_col else pd.NA

        if col_tg:
            r["_TG_norm"] = r[col_tg].astype(str).apply(_norm_tarak_generic)
            sub = r[r["_TG_norm"] == tg_norm].copy()
        else:
            sub = r.copy()

        # ---- GUARD (sub)
        if "_OpenTezgahFlag" not in sub.columns:
            sub["_OpenTezgahFlag"] = False
        if "_KalanMetreNorm" not in sub.columns:
            kal_col_sub = _col(sub, ["Kalan", "Kalan Mt", "Kalan Metre", "Kalan_Metre", "_KalanMetre"])
            sub["_KalanMetreNorm"] = pd.to_numeric(sub[kal_col_sub], errors="coerce") if kal_col_sub else pd.NA

        # --- sub içinde TEZGAH kolonu tespit + sabitleme
        col_tz_sub = _find_tz_col(sub)
        if not col_tz_sub:
            self._ordered_looms_cache[key] = []
            return []

        try:
            sub["_TZ_val"] = sub[col_tz_sub]
        except KeyError:
            nm = {_norm_name(c): c for c in sub.columns}
            alt = nm.get(_norm_name(col_tz_sub))
            if alt and alt in sub.columns:
                sub["_TZ_val"] = sub[alt]
            else:
                self._ordered_looms_cache[key] = []
                return []

        # izinli + storage kısıt filtresi
        def allowed_row(x):
            loom = _loom_no_as_int(x["_TZ_val"])
            d = _loom_digits(x["_TZ_val"])
            return (
                _loom_allowed(loom, grp_type)
                and (d not in self._blocked_looms)
                and (d not in self._dummy_looms)
            )

        sub = sub[sub.apply(allowed_row, axis=1)].copy()
        # Her ne kadar yukarıda varsayılan False değeri versek de, bazı edge-case
        # filtreleme adımlarından sonra kolon düşebiliyor. Böyle bir durumda
        # KeyError almamak için burada da garanti altına alıyoruz.
        if "_OpenTezgahFlag" not in sub.columns:
            sub["_OpenTezgahFlag"] = False
        if "_KalanMetreNorm" not in sub.columns:
            sub["_KalanMetreNorm"] = pd.NA
        # AÇIK listesi
        mask_open = (sub["_OpenTezgahFlag"] == True)
        open_vals = sub.loc[mask_open, "_TZ_val"] if "_TZ_val" in sub.columns else pd.Series([], dtype=object)
        open_list = [str(v) for v in open_vals.tolist()]

        # AÇACAK listesi
        thr = int(self.flow_threshold_m or 300)
        mask_soon = (sub["_OpenTezgahFlag"] != True) & (sub["_KalanMetreNorm"] <= thr)
        soon_df = sub.loc[mask_soon].copy()
        if not soon_df.empty and "_KalanMetreNorm" in soon_df.columns:
            soon_df = soon_df.sort_values(by="_KalanMetreNorm", ascending=True, na_position="last")
            soon_vals = soon_df["_TZ_val"] if "_TZ_val" in soon_df.columns else pd.Series([], dtype=object)
        else:
            soon_vals = pd.Series([], dtype=object)
        soon_list = [str(v) for v in soon_vals.tolist()]

        out = open_list + soon_list
        self._ordered_looms_cache[key] = out[:]
        return out

    def _next_free_loom(self, tg_norm: str, grp_type: str) -> str | None:
        used = self._used_looms.setdefault((tg_norm, grp_type), set())
        for loom in self._ordered_candidate_looms(tg_norm, grp_type):
            d = _loom_digits(loom)
            if (loom not in used) and (loom not in self._used_looms_global) and \
               (d not in self._blocked_looms) and (d not in self._dummy_looms):
                used.add(loom)
                return loom
        return None

    # ---------------- Alt: eksik bilgi etiketi ----------------
    def _update_missing_label(self):
        miss_union = self.missing_open.union(self.missing_soon)
        if not miss_union:
            self.lbl_missing.setText("Açık/Açacak tezgah sayıları: tüm tarak grupları için bulundu.")
        else:
            self.lbl_missing.setText(
                f"Açık/Açacak bilgisi GELMEYEN tarak grubu sayısı: {len(miss_union)}"
            )

    def _make_job_key(self, row: dict) -> str:
        ak = str(row.get("DokumaİşEmri", "")).strip()
        if ak:
            return ak
        return str(row.get("LeventNo / Durum", "")).strip() or "ROW"

    def _reset_team_assignments(self):
        self.team_rows = []
        self.model_team.set_df(pd.DataFrame())
        self._assignments = {}
        self._used_looms = {}
        self._used_looms_global = set()
        self._ordered_looms_cache = {}
        self._bind_group_jobs()
    # ---------------- Excel'e dışa aktarım (yazıcıya hazır) ----------------
    def _export_team_assignments(self):
        try:
            if not self.team_rows:
                QMessageBox.information(self, "Bilgi", "Dışa aktarılacak takım ataması bulunmuyor.")
                return

            df = pd.DataFrame.from_records(self.team_rows)

            # Kolon sırası (olanları alır; olmayanı görmez)
            wanted = [
                "Tezgah", "Tarak Grubu", "KökTip", "LeventNo / Durum",
                "ZeminÖrgü", "Çözgü İpliği 1", "Atkı İpliği 1",
                "Metre", "Mamül Termin", "Levent Haşıl Tarihi",
                "Kesim Tipi", "DokumaİşEmri"
            ]
            cols = [c for c in wanted if c in df.columns] + [c for c in df.columns if c not in wanted]
            df = df[cols]

            # Masaüstüne yaz
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            try:
                os.makedirs(desktop, exist_ok=True)
            except Exception:
                pass
            path = os.path.join(desktop, "TAKIM_Atamalari.xlsx")

            import xlsxwriter
            with pd.ExcelWriter(path, engine="xlsxwriter", datetime_format="dd.mm.yyyy") as writer:
                wb = writer.book
                sheet = "TakimAtamalari"
                ws = wb.add_worksheet(sheet)
                writer.sheets[sheet] = ws

                fmt_header = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "valign": "vcenter", "border": 1})
                fmt_text   = wb.add_format({"valign": "vcenter", "border": 1})
                fmt_num    = wb.add_format({"valign": "vcenter", "border": 1, "num_format": "#,##0"})
                fmt_date   = wb.add_format({"valign": "vcenter", "border": 1, "num_format": "dd.mm.yyyy"})

                # Başlık
                for c, name in enumerate(df.columns):
                    ws.write(0, c, name, fmt_header)
                ws.set_row(0, 22)

                # Veri
                ci_metre = df.columns.get_loc("Metre") if "Metre" in df.columns else -1
                ci_term  = df.columns.get_loc("Mamül Termin") if "Mamül Termin" in df.columns else -1
                ci_hash  = df.columns.get_loc("Levent Haşıl Tarihi") if "Levent Haşıl Tarihi" in df.columns else -1

                for r, (_, row) in enumerate(df.iterrows(), start=1):
                    for c, colname in enumerate(df.columns):
                        val = row[colname]
                        if c == ci_metre and pd.notna(val):
                            try:
                                ws.write_number(r, c, float(str(val).replace(",", ".")), fmt_num)
                            except Exception:
                                ws.write(r, c, "" if pd.isna(val) else str(val), fmt_text)
                        elif c in (ci_term, ci_hash) and pd.notna(val):
                            dt = pd.to_datetime(val, format="%d/%m/%Y", errors="coerce")
                            if pd.notna(dt):
                                ws.write_datetime(r, c, dt.to_pydatetime(), fmt_date)
                            else:
                                ws.write(r, c, "" if pd.isna(val) else str(val), fmt_text)
                        else:
                            ws.write(r, c, "" if pd.isna(val) else str(val), fmt_text)

                # Sütun genişlikleri
                for c, colname in enumerate(df.columns):
                    ser = df[colname].astype(str).replace("nan", "")
                    max_len = max(len(colname), *(min(len(s), 40) for s in ser.values)) + 2
                    ws.set_column(c, c, min(max_len, 42))

                # Yazdırma ayarları — yazıcıya hazır
                last_row = len(df)  # 0 = başlık, data 1..len(df)
                last_col = len(df.columns) - 1
                ws.set_default_row(20)
                ws.freeze_panes(1, 0)
                ws.set_landscape()                # yatay
                ws.fit_to_pages(1, 0)             # tek sayfa genişlik
                ws.set_paper(9)                   # A4
                ws.set_margins(left=0.5, right=0.5, top=0.6, bottom=0.6)
                ws.print_area(0, 0, last_row, last_col)   # yazdırma alanı: başlık + tüm veri
                ws.set_header('&R&08&Oluşturma: &D &T')   # sağ üstte tarih-saat

            QMessageBox.information(self, "Tamam", f"Excel masaüstüne kaydedildi:\n{path}")

            # Dosyayı otomatik aç
            try:
                if sys.platform.startswith("win"):
                    os.startfile(path)
                elif sys.platform == "darwin":
                    subprocess.Popen(["open", path])
                else:
                    subprocess.Popen(["xdg-open", path])
            except Exception:
                pass

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Dışa aktarılamadı:\n{e}")
