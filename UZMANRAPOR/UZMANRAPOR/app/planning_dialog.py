# app/planning_dialog.py
from __future__ import annotations

import re
import os, sys, subprocess
import pandas as pd
from datetime import datetime

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QListWidget, QPushButton, QLabel,
    QMessageBox, QWidget, QFileDialog, QSpinBox, QTableView, QHeaderView
)
from PySide6.QtCore import QSettings, QModelIndex

from app.models import PandasModel

# ---- Arızalı/Boş tezgah listesini depodan okuma (varsa) ----
try:
    from app import storage as _storage
except Exception:
    _storage = None

NEVER = {2430, 2432, 2434, 2436, 2438, 2440, 2442, 2444, 2446}
HAM_ALLOWED = set(range(2447, 2519))   # 2447–2518 arası
DENIM_ALLOWED_RANGE = (2201, 2446)     # 2201–2446 arası


def _extract_selv_teeth(val) -> int | None:
    """Süs kenar metninden diş sayısını (ilk tamsayıyı) çıkarır."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip()
    if not s:
        return None
    m = re.search(r"(\d+)", s)
    if not m:
        return None
    try:
        return int(m.group(1))
    except Exception:
        return None


def _selvedge_compatible_auto(job_sup: str, loom_sup: str, tarak_group: str | None = None) -> bool:
    """
    AUTO mod için süs kenarı uyum kontrolü.

    Kurallar:
      - Bire bir eşitse: UYUMLU
      - Diş sayıları okunabiliyorsa:
          * Eğer her ikisi de {8,10,18} içindeyse → UYUMLU
          * VEYA |iş_diş - tezgah_diş| <= 2 ise  → UYUMLU
      - Aksi halde: UYUMLU DEĞİL
    """
    job_sup = (job_sup or "").strip()
    loom_sup = (loom_sup or "").strip()

    # Bilgi yoksa bloklama
    if not job_sup or not loom_sup:
        return True

    # Aynı ise direkt kabul
    if job_sup == loom_sup:
        return True

    t_job = _extract_selv_teeth(job_sup)
    t_loom = _extract_selv_teeth(loom_sup)
    if t_job is None or t_loom is None:
        # Hem farklı hem sayı parse edemediysek, riske girmeyelim
        return False

    special = {8, 10, 18}

    # Özel durum: 8–10–18 üçlüsü birbiriyle uyumlu
    if t_job in special and t_loom in special:
        return True

    # Genel tolerans: en fazla 2 diş fark
    return abs(t_job - t_loom) <= 2


def _orgu_prefix(val: str) -> str:
    s = (val or "").strip()
    return s[:1].upper() if s else ""


def _orgu_compatible(job_orgu: str, loom_orgu: str) -> bool:
    """
    Örgü uyumu kontrolü.
    - Zemin örgü "3" ile başlayıp tezgah örgü "K" ile başlıyorsa → UYUMSUZ
    - Zemin örgü "K" ile başlayıp tezgah örgü "3" ile başlıyorsa → UYUMSUZ
    - Diğer tüm durumlar → UYUMLU
    """
    job_prefix = _orgu_prefix(job_orgu)
    loom_prefix = _orgu_prefix(loom_orgu)
    if not job_prefix or not loom_prefix:
        return True
    return not (
        (job_prefix == "3" and loom_prefix == "K")
        or (job_prefix == "K" and loom_prefix == "3")
    )


def _loom_in_category(loom_no: int | str, category: str) -> bool:
    try:
        n = int(str(loom_no).strip())
    except Exception:
        return False
    if n in NEVER:
        return False
    if str(category).upper() == "HAM":
        return n in HAM_ALLOWED
    return (DENIM_ALLOWED_RANGE[0] <= n <= DENIM_ALLOWED_RANGE[1])


def _pick_col(df: pd.DataFrame, names: list[str]) -> str | None:
    for n in names:
        if n in df.columns:
            return n
    # lowercase eşleştirme
    low = {c.lower(): c for c in df.columns}
    for n in names:
        if n.lower() in low:
            return low[n.lower()]
    return None


def _tarak_key_generic(val) -> str:
    """Dinamik ile aynı normalize (a/b/c ...). Virgül ondalığı koru."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    nums = re.findall(r"[\d]+(?:[.,]\d+)?", str(val))
    if not nums:
        return str(val).strip()
    out = []
    for n in nums[:3]:
        n = n.replace(",", ".")
        if re.fullmatch(r"\d+\.0+", n):
            n = n.split(".", 1)[0]
        out.append(n)
    return "/".join(out)


class PlanningDialog(QDialog):
    """
    Arızalı/Bakımda ve 'Boş Gösterilecek' tezgahlar:
      - Depodan okunur (storage.load_blocked_looms / load_dummy_looms)
      - Boş ve Açılacak tablolarına GELMEZLER
      - Atama yapılamaz (listeden tamamen hariç)
    """

    def __init__(
        self,
        df_jobs,
        df_looms,
        on_group_select=None,
        on_assign=None,
        on_list_made=None,         # opsiyonel callback
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowTitle("Planlama — DENIM / HAM")
        self.resize(1300, 760)
        self.df_jobs = df_jobs
        self.df_looms = df_looms
        self.on_group_select = on_group_select
        self._current_category = None
        self._current_group_label = None
        self.on_assign = on_assign
        self.on_list_made = on_list_made

        # Kullanıcı ayarı: Açacak eşik (varsayılan 100 m), kalıcı
        self.settings = QSettings("UZMANRAPOR", "ClientApp")
        self.plan_threshold_m = int(self.settings.value("planning/soon_threshold_m", 100))

        # Depodan (varsa) arızalı/boş tezgah kümelerini al
        self._blocked_looms, self._dummy_looms = self._load_restricted_looms()

        v = QVBoxLayout(self)

        # Üstte eşik kontrolü + SAĞ ÜSTE "Bu işi Atla"
        ctrl = QHBoxLayout()
        ctrl.addWidget(QLabel("Açacak ≤"))
        self.spin_plan_threshold = QSpinBox()
        self.spin_plan_threshold.setRange(10, 5000)
        self.spin_plan_threshold.setSingleStep(10)
        self.spin_plan_threshold.setValue(self.plan_threshold_m)
        self.spin_plan_threshold.valueChanged.connect(self._on_threshold_changed)
        ctrl.addWidget(self.spin_plan_threshold)
        ctrl.addWidget(QLabel("m"))
        ctrl.addStretch(1)

        # >>> SAĞ ÜST: Bu işi Atla
        self.btn_skip = QPushButton("Bu işi Atla")
        self.btn_skip.setToolTip("Sıradaki işi Düğüm Listesinde 'Atla' olarak işaretler.")
        self.btn_skip.clicked.connect(self._on_skip_current)
        ctrl.addWidget(self.btn_skip)
        # <<<
        v.addLayout(ctrl)

        # SOL: Grup listeleri
        left_box = QWidget()
        left_l = QVBoxLayout(left_box)
        left_l.addWidget(QLabel("DENIM — (HAM dışındaki boya kodları) — Levent No rakamlı"))
        self.lst_groups_denim = QListWidget()
        left_l.addWidget(self.lst_groups_denim, 1)
        left_l.addWidget(QLabel("HAM — (İhzarat Boya Kodu = HAM) — Levent No rakamlı"))
        self.lst_groups_ham = QListWidget()
        left_l.addWidget(self.lst_groups_ham, 1)

        # SAĞ: Tablo görünümleri
        right_box = QWidget()
        right_l = QVBoxLayout(right_box)
        right_l.addWidget(QLabel("Boş Tezgahlar (kategori + tarak uyumlu)"))
        self.tbl_free = QTableView()
        self.model_free = PandasModel(pd.DataFrame(columns=[
            "Tezgah", "Kategori", "Tip", "Tarak", "Örgü", "Süs Kenar", "KalanMetre", "Kesim Şekli"
        ]))
        self.tbl_free.setModel(self.model_free)
        self.tbl_free.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        right_l.addWidget(self.tbl_free, 1)

        self.lbl_soon = QLabel(f"Açılacaklar (<{self.plan_threshold_m} m)")
        right_l.addWidget(self.lbl_soon)
        self.tbl_soon = QTableView()
        self.model_soon = PandasModel(pd.DataFrame(columns=[
            "Tezgah", "Kategori", "Tip", "Tarak", "Örgü", "Süs Kenar", "KalanMetre", "Kesim Şekli"
        ]))
        self.tbl_soon.setModel(self.model_soon)
        self.tbl_soon.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        right_l.addWidget(self.tbl_soon, 1)

        # Düzen
        top = QHBoxLayout()
        top.addWidget(left_box, 1)
        top.addWidget(right_box, 1)
        v.addLayout(top)

        # Alt: Liste Yap
        bottom = QHBoxLayout()
        self.btn_list = QPushButton("LİSTE YAP (atanmışlar)")
        self.btn_list.clicked.connect(self._on_list_clicked)
        bottom.addStretch(1)
        bottom.addWidget(self.btn_list)
        v.addLayout(bottom)

        # Etkileşimler
        self._load_groups()
        self.lst_groups_denim.itemClicked.connect(lambda item: self._on_group_clicked(item, "DENIM"))
        self.lst_groups_ham.itemClicked.connect(lambda item: self._on_group_clicked(item, "HAM"))

        # Çift tıkla atama
        self.tbl_free.doubleClicked.connect(lambda idx: self._assign_from_table("free", idx))
        self.tbl_soon.doubleClicked.connect(lambda idx: self._assign_from_table("soon", idx))

        # Varsayılan seçim
        if self.lst_groups_denim.count() > 0:
            self.lst_groups_denim.setCurrentRow(0)
            self._on_group_clicked(self.lst_groups_denim.item(0), "DENIM")
        elif self.lst_groups_ham.count() > 0:
            self.lst_groups_ham.setCurrentRow(0)
            self._on_group_clicked(self.lst_groups_ham.item(0), "HAM")

    # ---------- Yardımcı: kısıtlı tezgah setlerini depodan oku ----------
    def _load_restricted_looms(self) -> tuple[set[str], set[str]]:
        """storage'dan arızalı/bakım (blocked) ve 'boş gösterilecek' (dummy) tezgahları okur."""
        blocked, dummy = set(), set()
        try:
            if _storage:
                # Birden fazla isim desteği (eski/yeni)
                if hasattr(_storage, "load_blocked_looms"):
                    blocked = set(map(lambda s: re.findall(r"\d+", str(s))[0], _storage.load_blocked_looms() or []))
                elif hasattr(_storage, "get_blocked_looms"):
                    blocked = set(map(lambda s: re.findall(r"\d+", str(s))[0], _storage.get_blocked_looms() or []))
                if hasattr(_storage, "load_dummy_looms"):
                    dummy = set(map(lambda s: re.findall(r"\d+", str(s))[0], _storage.load_dummy_looms() or []))
                elif hasattr(_storage, "get_dummy_looms"):
                    dummy = set(map(lambda s: re.findall(r"\d+", str(s))[0], _storage.get_dummy_looms() or []))
        except Exception:
            pass
        return blocked, dummy

    def _on_threshold_changed(self, v: int):
        self.plan_threshold_m = int(v)
        self.settings.setValue("planning/soon_threshold_m", self.plan_threshold_m)
        if self._current_group_label:
            self._load_looms_for_key_and_category(self._current_key(), self._current_category)
        self.lbl_soon.setText(f"Açılacaklar (<{self.plan_threshold_m} m)")

    def _on_skip_current(self):
        """Aktif grup+kategori için sıradaki işi 'Atla' olarak işaretler."""
        key = self._current_key()
        if not key:
            QMessageBox.warning(self, "Uyarı", "Önce gruptan bir iş seçin.")
            return

        df = self.df_jobs
        mask_group = df.get("_TarakKey", "").astype(str) == str(key)
        mask_digits = df.get("_LeventHasDigits", False)
        mask_unassigned = (df.get("Tezgah Numarası", "").astype(str) == "") | (df.get("Tezgah Numarası").isna())
        if str(self._current_category).upper() == "HAM":
            mask_cat = df.get("_DyeCategory", "").astype(str).str.contains("HAM", na=False)
        else:
            mask_cat = ~df.get("_DyeCategory", "").astype(str).str.contains("HAM", na=False)

        candidates = df[mask_group & mask_digits & mask_unassigned & mask_cat].copy()

        # Güvenli sıralama
        sort_cols = [c for c in ["Mamul Termin", "Termin", "Plan Termin"] if c in candidates.columns]
        if sort_cols:
            candidates = candidates.sort_values(by=sort_cols, ascending=True)

        if candidates.empty:
            QMessageBox.information(self, "Bilgi", "Bu grupta atlanacak uygun iş bulunamadı.")
            return

        idx = candidates.index[0]
        self.df_jobs.at[idx, "Tezgah Numarası"] = "Atla"

        # Görünümleri tazelemesi için üst akışa haber ver
        if callable(self.on_assign):
            group_label = self._current_group_label or ""
            category = self._current_category or ""
            self.on_assign(group_label, category)

        QMessageBox.information(self, "Bilgi", f"Sıradaki iş 'Atla' olarak işaretlendi (satır {idx}).")

    # ---------------- LİSTE YAP ----------------
    def _on_list_clicked(self):
        ok = self._do_list_and_export()
        if ok:
            try:
                if callable(self.on_list_made):
                    self.on_list_made()
            finally:
                self.accept()

    def _do_list_and_export(self) -> bool:
        try:
            df = self.df_jobs.copy()

            # -- yalnızca GERÇEK atamalar (Atla hariç)
            if "Tezgah Numarası" not in df.columns:
                QMessageBox.information(self, "Bilgi", "Kaynakta 'Tezgah Numarası' kolonu yok.")
                return False

            tz = df["Tezgah Numarası"].astype(str).str.strip()
            mask_assigned = tz.ne("") & (~df["Tezgah Numarası"].isna()) & tz.str.upper().ne("ATLA")
            base = df[mask_assigned].copy()
            if base.empty:
                QMessageBox.information(self, "Bilgi", "Atanmış herhangi bir kayıt bulunamadı.")
                return False

            # -- esnek kolon bulucu
            def pick(*names):
                for n in names:
                    if n in df.columns:
                        return n
                low = {c.lower(): c for c in df.columns}
                for n in names:
                    if n.lower() in low:
                        return low[n.lower()]
                return None

            c_tezgah = pick("Tezgah Numarası", "Tezgah No", "Tezgah")
            c_levent = pick("Levent No", "LEVENT NO")
            c_etiket = pick("LEVENT ETİKET FA", "Levent Etiket FA", "Etiket No")
            c_koktip = pick("Kök Tip Kodu", "KökTip", "Kök Tip")
            c_tarak = pick("Tarak Grubu", "Tarak")
            c_orgu = pick("Zemin Örgü", "Örgü", "Zemin Orgu")
            c_atki1 = pick("Atkı İpliği 1", "Atkı Ipliği 1")
            c_atki2 = pick("Atkı İpliği 2", "Atkı İpliği 2")
            c_cozgu1 = pick("Çözgü İpliği 1", "Cozgu Ipliği 1", "Çözgü Ipliği 1")
            c_metre = pick("PARTİ METRESİ", "Parti Metresi", "Metre")
            c_hasilt = pick("Levent Haşıl Tarihi", "Haşıl Tarihi")
            c_sup_job = pick("SÜS KENAR", "Süs Kenar")
            c_sup_loom = pick("TEZGAHTA MEVCUT İŞİN SÜS KENARI", "Tezgahta Mevcut İşin Süs Kenarı")
            c_notlar = pick("NOTLAR")

            wanted = [
                "Tezgah Numarası", "Levent No", "Etiket No", "Kök Tip Kodu", "Tarak Grubu",
                "Zemin Örgü", "Atkı İpliği 1", "Atkı İpliği 2", "Çözgü İpliği 1", "Metre",
                "Levent Haşıl Tarihi", "Verilen İşin Süs Kenarı", "Tezgahın Süs Kenarı", "NOTLAR"
            ]
            out = pd.DataFrame(index=base.index, columns=wanted)

            out["Tezgah Numarası"] = base[c_tezgah] if c_tezgah else ""
            out["Levent No"] = base[c_levent] if c_levent else ""
            out["Etiket No"] = base[c_etiket] if c_etiket else ""
            out["Kök Tip Kodu"] = base[c_koktip] if c_koktip else ""
            out["Tarak Grubu"] = base[c_tarak] if c_tarak else ""
            out["Zemin Örgü"] = base[c_orgu] if c_orgu else ""
            out["Atkı İpliği 1"] = base[c_atki1] if c_atki1 else ""
            out["Atkı İpliği 2"] = base[c_atki2] if c_atki2 else ""
            out["Çözgü İpliği 1"] = base[c_cozgu1] if c_cozgu1 else ""
            out["Metre"] = base[c_metre] if c_metre else ""
            out["Levent Haşıl Tarihi"] = base[c_hasilt] if c_hasilt else ""
            out["Verilen İşin Süs Kenarı"] = base[c_sup_job] if c_sup_job else ""
            out["Tezgahın Süs Kenarı"] = base[c_sup_loom] if c_sup_loom else ""
            out["NOTLAR"] = base[c_notlar] if c_notlar else ""
            out["Metre"] = pd.to_numeric(out["Metre"], errors="coerce")

            # --- Etiket No: .0 kuyruklarını kaldır ve metin olarak yaz ---
            def _clean_label(x):
                if pd.isna(x):
                    return ""
                s = str(x).strip()
                if not s:
                    return ""
                s = s.replace(",", ".")
                if re.fullmatch(r"-?\d+\.0+", s):
                    return s.split(".", 1)[0]
                try:
                    f = float(s)
                    if f.is_integer():
                        return str(int(f))
                except Exception:
                    pass
                return s

            out["Etiket No"] = out["Etiket No"].apply(_clean_label)

            # --- Tezgahın Süs Kenarı'nı RUNNING'den doldur ---
            try:
                df_run = getattr(self, "df_looms", None)
                if df_run is not None and not df_run.empty and ("Süs Kenar" in df_run.columns):
                    col_tz_run = None
                    for c in ["Tezgah No", "Tezgah", "Tezgah Numarası"]:
                        if c in df_run.columns:
                            col_tz_run = c
                            break
                    if col_tz_run:
                        def _digits(s):
                            m = re.search(r"(\d+)", str(s))
                            return m.group(1) if m else ""

                        run_map = {}
                        for _, rr in df_run.iterrows():
                            tzv = _digits(rr.get(col_tz_run, ""))
                            sk = "" if pd.isna(rr.get("Süs Kenar", "")) else str(rr.get("Süs Kenar", ""))
                            if tzv and sk:
                                run_map[tzv] = sk

                        def _fill_selv(row):
                            val = str(row.get("Tezgahın Süs Kenarı", "") or "").strip()
                            if val:
                                return val
                            tzv = _digits(row.get("Tezgah Numarası", ""))
                            return run_map.get(tzv, "")

                        out["Tezgahın Süs Kenarı"] = out.apply(_fill_selv, axis=1)
            except Exception:
                pass
            # --- /RUNNING doldurma ---

            # Masaüstü + sabit dosya adı
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            try:
                os.makedirs(desktop, exist_ok=True)
            except Exception:
                pass
            path = os.path.join(desktop, "Düğüm_Listesi.xlsx")

            import xlsxwriter
            with pd.ExcelWriter(path, engine="xlsxwriter", datetime_format="dd.mm.yyyy") as writer:
                wb = writer.book
                sheet = "Atanmislar"
                ws = wb.add_worksheet(sheet)
                writer.sheets[sheet] = ws

                fmt_header = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "valign": "vcenter", "border": 1})
                fmt_text = wb.add_format({"valign": "vcenter", "border": 1})
                fmt_num = wb.add_format({"valign": "vcenter", "border": 1, "num_format": "#,##0"})
                fmt_date = wb.add_format({"valign": "vcenter", "border": 1, "num_format": "dd.mm.yyyy"})
                fmt_ham_warp = wb.add_format({"valign": "vcenter", "border": 1, "bg_color": "#FFF3B0"})  # sarı
                fmt_denim_warp = wb.add_format({"valign": "vcenter", "border": 1, "bg_color": "#DDEBFF"})  # mavi
                fmt_mismatch = wb.add_format({"valign": "vcenter", "border": 1, "bg_color": "#D32F2F", "font_color": "white"})

                for c, name in enumerate(wanted):
                    ws.write(0, c, name, fmt_header)
                ws.set_row(0, 22)

                ci_label = wanted.index("Etiket No")
                ci_warp = wanted.index("Çözgü İpliği 1")
                ci_supjob = wanted.index("Verilen İşin Süs Kenarı")
                ci_suploom = wanted.index("Tezgahın Süs Kenarı")
                ci_metre = wanted.index("Metre")
                ci_date = wanted.index("Levent Haşıl Tarihi")

                start_row = 1
                for r, (rid, row) in enumerate(out.iterrows(), start=start_row):
                    for c, colname in enumerate(wanted):
                        val = row[colname]
                        if c == ci_metre and pd.notna(val):
                            ws.write_number(r, c, float(val), fmt_num)
                        elif c == ci_date and pd.notna(val):
                            dt = pd.to_datetime(val, dayfirst=True, errors="coerce")
                            if pd.notna(dt):
                                ws.write_datetime(r, c, dt.to_pydatetime(), fmt_date)
                            else:
                                ws.write(r, c, "" if pd.isna(val) else str(val), fmt_text)
                        elif c == ci_label:
                            ws.write(r, c, "" if pd.isna(val) else str(val), fmt_text)
                        else:
                            ws.write(r, c, "" if pd.isna(val) else str(val), fmt_text)

                    # kategori renklendirme
                    cat_val = str(self.df_jobs.at[rid, "_DyeCategory"]).upper().strip() \
                        if "_DyeCategory" in self.df_jobs.columns and rid in self.df_jobs.index else ""
                    warp_val = row["Çözgü İpliği 1"]
                    if cat_val == "HAM":
                        ws.write(r, ci_warp, warp_val, fmt_ham_warp)
                    else:
                        ws.write(r, ci_warp, warp_val, fmt_denim_warp)

                    # süs kenarı uyuşmazlık
                    v_job = str(row["Verilen İşin Süs Kenarı"]).strip()
                    v_loom = str(row["Tezgahın Süs Kenarı"]).strip()
                    if v_job and v_loom and (v_job != v_loom):
                        ws.write(r, ci_supjob, v_job, fmt_mismatch)
                        ws.write(r, ci_suploom, v_loom, fmt_mismatch)

                for c, colname in enumerate(wanted):
                    ser = out[colname].astype(str).replace("nan", "")
                    max_len = max(len(colname), *(min(len(s), 40) for s in ser.values)) + 2
                    ws.set_column(c, c, min(max_len, 42))

                ws.set_default_row(20)
                ws.freeze_panes(1, 0)
                ws.set_landscape()
                ws.fit_to_pages(1, 0)
                ws.set_paper(9)  # A4
                ws.set_margins(left=0.5, right=0.5, top=0.6, bottom=0.6)
                ws.print_area(0, 0, start_row + len(out) - 1, len(wanted) - 1)
                ws.set_header('&R&08&Oluşturma: &D &T')

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

            return True

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Excel'e aktarılamadı:\n{e}")
            return False

    # --------------- mevcut akış ---------------
    def _load_groups(self):
        df = self.df_jobs
        mask = df.get("_LeventHasDigits", False)
        if "Tarak Grubu" not in df.columns:
            return
        denim_mask = (df.get("_DyeCategory", "DENIM") == "DENIM")
        ham_mask = (df.get("_DyeCategory", "DENIM") == "HAM")

        groups_denim = sorted(set(df.loc[mask & denim_mask, "Tarak Grubu"].astype(str)))
        groups_ham = sorted(set(df.loc[mask & ham_mask, "Tarak Grubu"].astype(str)))

        self.lst_groups_denim.clear()
        self.lst_groups_ham.clear()
        for g in groups_denim:
            self.lst_groups_denim.addItem(str(g))
        for g in groups_ham:
            self.lst_groups_ham.addItem(str(g))

    # --------------- AUTO PLANLAMA (tamamen sessiz, mesaj kutusu yok) ---------------
    def auto_plan_all_groups(self) -> int:
        """
        Tüm DENIM ve HAM tarak gruplarında, boş + açılacak tezgahlara
        AUTO mantıkla atama yapar.

        Dönüş: Toplam atanan iş sayısı.
        """
        total_assigned = 0

        # Grupları tazele
        self._load_groups()

        # Önce DENIM
        for i in range(self.lst_groups_denim.count()):
            item = self.lst_groups_denim.item(i)
            if not item:
                continue
            group_label = str(item.text())
            total_assigned += self._auto_plan_for_group(group_label, "DENIM")

        # Sonra HAM
        for i in range(self.lst_groups_ham.count()):
            item = self.lst_groups_ham.item(i)
            if not item:
                continue
            group_label = str(item.text())
            total_assigned += self._auto_plan_for_group(group_label, "HAM")

        return total_assigned

    def _auto_plan_for_group(self, group_label: str, category: str) -> int:
        """
        Tek bir tarak grubu + kategori (DENIM/HAM) için:
          - Uygun boş + açılacak tezgah listesini çıkarır,
          - Termin önceliğine göre işleri,
          - Her tezgaha en fazla 1 iş düşecek şekilde AUTO atar.
        """
        self._current_category = category
        self._current_group_label = str(group_label)
        key = self._current_key()
        if not key:
            return 0

        # Bu key + kategori için uygun tezgahları hesapla (Boş + Açılacak)
        self._load_looms_for_key_and_category(key, category)

        df_free = getattr(self.model_free, "_df", pd.DataFrame())
        df_soon = getattr(self.model_soon, "_df", pd.DataFrame())

        # Boş varsa önce Boş, yoksa Açılacaklar'a mutlaka baksın.
        looms: list[dict] = []
        for src in (df_free, df_soon):
            if src is None or src.empty:
                continue
            for _, row in src.iterrows():
                loom_no = str(row.get("Tezgah", "")).strip()
                if not loom_no or loom_no.lower() in ("nan", "none"):
                    continue
                loom_sup = str(row.get("Süs Kenar", "") or "").strip()
                loom_orgu = str(row.get("Örgü", "") or "").strip()
                looms.append({"Tezgah": loom_no, "SüsKenar": loom_sup, "Orgu": loom_orgu})

        if not looms:
            return 0

        used_looms = set()
        assigned_count = 0

        while True:
            # Bu grupta hâlâ atanacak iş var mı? (en güncel hâliyle)
            df = self.df_jobs
            mask_group = df.get("_TarakKey", "").astype(str) == str(key)
            mask_digits = df.get("_LeventHasDigits", False)
            tz_col = "Tezgah Numarası"
            mask_unassigned = (df.get(tz_col, "").astype(str) == "") | (df.get(tz_col).isna())

            if str(category).upper() == "HAM":
                mask_cat = df.get("_DyeCategory", "").astype(str).str.contains("HAM", na=False)
            else:
                mask_cat = ~df.get("_DyeCategory", "").astype(str).str.contains("HAM", na=False)

            candidates = df[mask_group & mask_digits & mask_unassigned & mask_cat].copy()
            sort_cols = [c for c in ["Mamul Termin", "Termin", "Plan Termin"] if c in candidates.columns]
            if sort_cols:
                candidates = candidates.sort_values(by=sort_cols, ascending=True)

            if candidates.empty:
                break  # bu grup için iş bitti

            progressed = False  # bu turda bir iş ilerledi mi? (Atama veya Atla)

            for loom in looms:
                loom_no = loom["Tezgah"]
                if loom_no in used_looms:
                    continue
                loom_sup = loom["SüsKenar"]
                loom_orgu = loom.get("Orgu", "")

                ok, msg, remove_row = self._assign_first_job_auto(key, loom_no, loom_sup, loom_orgu)

                if ok:
                    progressed = True
                    if remove_row:
                        used_looms.add(loom_no)
                        assigned_count += 1
                    # Bu job artık ya atandı ya Atla oldu; bir sonraki job için dış döngüye dön
                    break

                # ok == False ise (ör: süs kenarı/örgü uyumsuz) → sıradaki tezgaha bakmaya devam

            if not progressed:
                # Mevcut en öncelikli iş, kalan hiçbir tezgaha sığmıyor → bırak, manuel baksın
                break

        return assigned_count

    def _current_key(self) -> str:
        label = self._current_group_label
        if not label:
            return ""
        rows = self.df_jobs[self.df_jobs.get("Tarak Grubu", "").astype(str) == str(label)]
        if not rows.empty:
            if "_TarakKey" in rows.columns and pd.notna(rows["_TarakKey"]).any():
                return str(rows["_TarakKey"].iloc[0])
            return _tarak_key_generic(rows["Tarak Grubu"].iloc[0])
        return ""

    def _first_job_details(self):
        key = self._current_key()
        df = self.df_jobs
        mask_group = df.get("_TarakKey", "").astype(str) == str(key)
        mask_digits = df.get("_LeventHasDigits", False)
        mask_unassigned = (df.get("Tezgah Numarası", "").astype(str) == "") | (df.get("Tezgah Numarası").isna())

        if self._current_category == "HAM":
            mask_cat = df.get("_DyeCategory", "").astype(str).str.contains("HAM", na=False)
        else:  # DENIM
            mask_cat = ~df.get("_DyeCategory", "").astype(str).str.contains("HAM", na=False)

        candidates = df[mask_group & mask_digits & mask_unassigned & mask_cat].sort_values(
            by="Mamul Termin", ascending=True
        )
        if candidates.empty:
            return (self._current_category or "", self._current_group_label or "", "")

        row = candidates.iloc[0]
        category = self._current_category or row.get("_DyeCategory", "")
        tarak_label = row.get("Tarak Grubu", self._current_group_label) or ""
        sus = row.get("Süs Kenar", "")
        return (category, str(tarak_label), str(sus))

    def _on_group_clicked(self, item, category: str):
        self._current_category = category
        self._current_group_label = str(item.text())
        key = self._current_key()
        self._load_looms_for_key_and_category(key, category)
        if callable(self.on_group_select):
            self.on_group_select(item.text(), category)

    def _build_view_from_running(self, src: pd.DataFrame, category: str) -> pd.DataFrame:
        """RUNNING kaynağından tablo görünümü üretir."""
        if src is None or src.empty:
            return pd.DataFrame(columns=["Tezgah", "Kategori", "Tip", "Tarak", "Örgü", "Süs Kenar", "KalanMetre", "Kesim Şekli"])

        col_tz = _pick_col(src, ["Tezgah No", "Tezgah", "Tezgah Numarası"])
        col_tip = _pick_col(src, ["KökTip", "Kök Tip Kodu", "Tip No", "Tip Kodu", "Tip", "Mamul Tipi"])
        col_tg = _pick_col(src, ["Tarak Grubu", "Tarak", "TarakGrubu"])
        col_orgu = "Orgu Kodu" if "Orgu Kodu" in src.columns else _pick_col(
            src, ["Zemin Örgü", "Zemin Örgü Kodu", "Zemin Örgü Adı", "Örgü", "Zemin Orgu"]
        )
        col_sus = "Süs Kenar" if "Süs Kenar" in src.columns else None
        col_cut = _pick_col(src, ["Kesim Tipi", "Kesim", "ISAVER/ROTOCUT", "ISAVER/ROTOCUT/ISAVERKit"])

        if "_KalanMetreNorm" not in src.columns:
            kal_col = _pick_col(src, ["Kalan", "Kalan Mt", "Kalan Metre", "Kalan_Metre", "_KalanMetre"])
            if kal_col:
                src = src.copy()
                src["_KalanMetreNorm"] = pd.to_numeric(src[kal_col], errors="coerce")
            else:
                src = src.copy()
                src["_KalanMetreNorm"] = pd.NA

        rows = []
        upper_cat = "HAM" if str(category).upper() == "HAM" else "DENIM"
        for _, r in src.iterrows():
            rows.append({
                "Tezgah": str(r.get(col_tz, "")) if col_tz else "",
                "Kategori": upper_cat,
                "Tip": str(r.get(col_tip, "")) if col_tip else "",
                "Tarak": str(r.get(col_tg, "")) if col_tg else "",
                "Örgü": str(r.get(col_orgu, "")) if col_orgu else "",
                "Süs Kenar": str(r.get(col_sus, "")) if col_sus else "",
                "KalanMetre": r.get("_KalanMetreNorm", pd.NA),
                "Kesim Şekli": str(r.get(col_cut, "")) if col_cut else "",
            })
        view = pd.DataFrame.from_records(rows, columns=[
            "Tezgah", "Kategori", "Tip", "Tarak", "Örgü", "Süs Kenar", "KalanMetre", "Kesim Şekli"
        ])
        return view

    def _load_looms_for_key_and_category(self, key: str, category: str):
        # tablolari temizle
        self.model_free.set_df(pd.DataFrame(columns=[
            "Tezgah", "Kategori", "Tip", "Tarak", "Örgü", "Süs Kenar", "KalanMetre", "Kesim Şekli"
        ]))
        self.model_soon.set_df(pd.DataFrame(columns=[
            "Tezgah", "Kategori", "Tip", "Tarak", "Örgü", "Süs Kenar", "KalanMetre", "Kesim Şekli"
        ]))

        df = self.df_looms.copy()
        if df is None or df.empty or not key:
            return

        # _TarakKey yoksa üret
        if "_TarakKey" not in df.columns:
            tg_col = _pick_col(df, ["Tarak Grubu", "Tarak", "TarakGrubu"])
            df["_TarakKey"] = df[tg_col].astype(str).apply(_tarak_key_generic) if tg_col else ""

        # tarak uyumu
        df = df[df["_TarakKey"].astype(str) == str(key)]
        if df.empty:
            return

        # açık/soon için gerekli alanlar
        if "_OpenTezgahFlag" not in df.columns:
            def _detect_94_row(row):
                for c in row.index:
                    u = str(row.get(c, "")).strip().upper()
                    if "SİPARİŞ YOK" in u or "SIPARIS YOK" in u or u == "94" or " 94" in u:
                        return True
                return False
            df["_OpenTezgahFlag"] = df.apply(_detect_94_row, axis=1)
        if "_KalanMetreNorm" not in df.columns:
            kal_col = _pick_col(df, ["Kalan", "Kalan Mt", "Kalan Metre", "Kalan_Metre", "_KalanMetre"])
            df["_KalanMetreNorm"] = pd.to_numeric(df[kal_col], errors="coerce") if kal_col else pd.NA

        # Kategori filtresi + atanmış olanları çıkar
        col_tz = _pick_col(df, ["Tezgah No", "Tezgah", "Tezgah Numarası"])
        assigned_looms = set(
            self.df_jobs.get("Tezgah Numarası", "").astype(str).str.strip().replace({"nan": "", "None": ""})
        )
        assigned_looms.discard("")

        # --- Kısıtlı (arızalı/boş) tezgahları hariç tut ---
        blocked = set(self._blocked_looms or set())
        dummy = set(self._dummy_looms or set())

        def _allowed_row(x):
            tz_raw = x[col_tz] if col_tz else ""
            m = re.search(r"(\d+)", str(tz_raw))
            tzv = m.group(1) if m else ""
            tz_int = int(tzv) if tzv.isdigit() else None
            return (
                (tz_int is not None)
                and _loom_in_category(tz_int, str(category).upper())
                and (tzv not in blocked)
                and (tzv not in dummy)
                and (tzv not in assigned_looms)
            )

        df = df[df.apply(_allowed_row, axis=1)]
        if df.empty:
            return

        # ayrımlar
        thr = int(self.plan_threshold_m or 100)
        free_df = df[df["_OpenTezgahFlag"] == True].copy()
        soon_df = df[(df["_OpenTezgahFlag"] != True) & (df["_KalanMetreNorm"] <= thr)].copy()

        # görünüm
        view_free = self._build_view_from_running(free_df, category)
        view_soon = self._build_view_from_running(soon_df, category)

        # sırala (tezgah numarasına göre)
        def _safe_loom_int(s):
            m = re.search(r"(\d+)", str(s))
            return int(m.group(1)) if m else 99999

        if not view_free.empty:
            view_free = view_free.sort_values(by="Tezgah", key=lambda s: s.apply(_safe_loom_int), ascending=True)
        if not view_soon.empty:
            view_soon = view_soon.sort_values(by="Tezgah", key=lambda s: s.apply(_safe_loom_int), ascending=True)

        # yaz
        self.model_free.set_df(view_free)
        self.tbl_free.resizeColumnsToContents()
        self.model_soon.set_df(view_soon)
        self.tbl_soon.resizeColumnsToContents()

    def _assign_first_job(
        self,
        key: str,
        loom_no: str,
        loom_sup: str | None = None,
        loom_orgu: str | None = None,
    ):
        df = self.df_jobs
        mask_group = df.get("_TarakKey", "").astype(str) == str(key)
        mask_digits = df.get("_LeventHasDigits", False)
        mask_unassigned = (df.get("Tezgah Numarası", "").astype(str) == "") | (df.get("Tezgah Numarası").isna())

        if str(self._current_category).upper() == "HAM":
            mask_cat = df.get("_DyeCategory", "").astype(str).str.contains("HAM", na=False)
        else:  # DENIM
            mask_cat = ~df.get("_DyeCategory", "").astype(str).str.contains("HAM", na=False)

        candidates = df[mask_group & mask_digits & mask_unassigned & mask_cat].copy()

        # Güvenli sıralama
        sort_cols = [c for c in ["Mamul Termin", "Termin", "Plan Termin"] if c in candidates.columns]
        if sort_cols:
            candidates = candidates.sort_values(by=sort_cols, ascending=True)

        if candidates.empty:
            return False, "Bu tarak grubunda seçilen kategoriye (DENIM/HAM) uygun atanacak iş bulunamadı.", False

        idx = candidates.index[0]

        # --- NOT / ATKI uyarıları ---
        note_col = "NOTLAR" if "NOTLAR" in df.columns else None
        job_note = str(df.at[idx, note_col]).strip() if note_col else ""
        has_atki_issue = bool(re.search(r"ATKI\s*1\s*EKSİK|ATKI\s*2\s*EKSİK", job_note, flags=re.IGNORECASE))

        if job_note and job_note.lower() not in ("", "nan", "none"):
            if has_atki_issue:
                m = QMessageBox.question(
                    self, "Uyarı",
                    f"Bu işin notunda ATKI eksikliği var:\n\n{job_note}\n\nYine de Düğüm Listesine vermek istiyor musun?",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No
                )
            else:
                m = QMessageBox.question(
                    self, "Not Uyarısı",
                    f"Bu iş için not var:\n\n{job_note}\n\nAtamaya devam edilsin mi?",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No
                )
            if m == QMessageBox.No:
                self.df_jobs.at[idx, "Tezgah Numarası"] = "Atla"
                return True, f"Not nedeniyle 'Atla' olarak işaretlendi (satır {idx}).", False

        # --- Örgü uyumu (MANUAL) ---  (Süs Kenar gibi davranır)
        job_orgu = ""
        for orgu_col in ["Zemin Örgü", "Zemin Orgu", "Örgü", "Orgu"]:
            if orgu_col in df.columns:
                job_orgu = str(df.at[idx, orgu_col]).strip()
                break

        current_orgu = (loom_orgu or "").strip()
        if job_orgu and current_orgu and (not _orgu_compatible(job_orgu, current_orgu)):
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Warning)
            box.setWindowTitle("Örgü Uyarısı")
            box.setText(
                f"Verilen işin örgüsü ({job_orgu}) seçilen tezgahın örgüsünden ({current_orgu}) uyumsuz."
            )
            btn_yes = box.addButton("Evet (Atamaya devam)", QMessageBox.YesRole)
            btn_no = box.addButton("Hayır (Atla)", QMessageBox.NoRole)
            btn_other = box.addButton("Başka tezgah seç", QMessageBox.RejectRole)
            box.exec()
            clicked = box.clickedButton()
            if clicked is btn_no:
                self.df_jobs.at[idx, "Tezgah Numarası"] = "Atla"
                return True, f"Örgü uyumsuzluğu nedeniyle 'Atla' olarak işaretlendi (satır {idx}).", False
            if clicked is btn_other:
                return False, "Başka tezgah seçin.", False

        # --- Süs kenar uyumu ---
        job_sup = ""
        if "SÜS KENAR" in df.columns:
            job_sup = str(df.at[idx, "SÜS KENAR"]).strip()
        elif "Süs Kenar" in df.columns:
            job_sup = str(df.at[idx, "Süs Kenar"]).strip()
        current_sup = (loom_sup or "").strip()

        if job_sup and current_sup and job_sup != current_sup:
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Warning)
            box.setWindowTitle("Süs Kenarı Uyarısı")
            box.setText(f"Verilen işin süs kenarı ({job_sup}) seçilen tezgahın süs kenarından ({current_sup}) farklı.")
            btn_yes = box.addButton("Evet (Atamaya devam)", QMessageBox.YesRole)
            btn_no = box.addButton("Hayır (Atla)", QMessageBox.NoRole)
            btn_other = box.addButton("Başka tezgah seç", QMessageBox.RejectRole)
            box.exec()
            clicked = box.clickedButton()
            if clicked is btn_no:
                self.df_jobs.at[idx, "Tezgah Numarası"] = "Atla"
                return True, f"Süs kenarı uyuşmazlığı nedeniyle 'Atla' olarak işaretlendi (satır {idx}).", False
            if clicked is btn_other:
                return False, "Başka tezgah seçin.", False

        # Atama
        self.df_jobs.at[idx, "Tezgah Numarası"] = loom_no
        return True, f"{loom_no} tezgâha atandı (satır {idx}).", True

    def _assign_first_job_auto(
        self,
        key: str,
        loom_no: str,
        loom_sup: str | None = None,
        loom_orgu: str | None = None,
    ):
        """
        AUTO mod: Mesaj kutusu kullanmadan, kural bazlı atama yapar.

        Dönüş:
          ok:          Bu çağrıda bir ilerleme oldu mu? (Atama YAPILDI veya iş 'Atla' oldu) → True
                       Hiçbir şey değişmediyse (süs kenarı/örgü uyumsuz, iş seçilemedi vb.) → False
          msg:         Log / bilgi mesajı
          remove_row:  Bu tezgah satırı tek iş aldı, bir daha kullanılmasın mı? (True/False)
        """
        df = self.df_jobs

        # Aynı grup + rakamlı Levent + atanmış olmayan + kategori (DENIM/HAM) filtresi
        mask_group = df.get("_TarakKey", "").astype(str) == str(key)
        mask_digits = df.get("_LeventHasDigits", False)
        tz_col = "Tezgah Numarası"
        mask_unassigned = (df.get(tz_col, "").astype(str) == "") | (df.get(tz_col).isna())

        if str(self._current_category).upper() == "HAM":
            mask_cat = df.get("_DyeCategory", "").astype(str).str.contains("HAM", na=False)
        else:  # DENIM
            mask_cat = ~df.get("_DyeCategory", "").astype(str).str.contains("HAM", na=False)

        candidates = df[mask_group & mask_digits & mask_unassigned & mask_cat].copy()

        sort_cols = [c for c in ["Mamul Termin", "Termin", "Plan Termin"] if c in candidates.columns]
        if sort_cols:
            candidates = candidates.sort_values(by=sort_cols, ascending=True)

        if candidates.empty:
            return False, "Bu tarak grubunda AUTO atanacak iş bulunamadı.", False

        idx = candidates.index[0]

        # --- NOT / ATKI uyarıları (AUTO) ---
        note_col = "NOTLAR" if "NOTLAR" in df.columns else None
        job_note = str(df.at[idx, note_col]).strip() if note_col else ""
        has_atki_issue = bool(re.search(r"ATKI\s*1\s*EKSİK|ATKI\s*2\s*EKSİK", job_note, flags=re.IGNORECASE))

        # HERHANGİ BİR NOT VARSA → ATLA
        if job_note and job_note.lower() not in ("", "nan", "none"):
            self.df_jobs.at[idx, "Tezgah Numarası"] = "Atla"
            return True, f"NOTLAR dolu olduğu için iş 'Atla' yapıldı (AUTO, satır {idx}).", False

        # Senin mantığın: ATKI 1/2 EKSİK ise bu işi almıyoruz → 'Atla'
        if has_atki_issue:
            self.df_jobs.at[idx, "Tezgah Numarası"] = "Atla"
            return True, f"ATKI eksikliği nedeniyle 'Atla' olarak işaretlendi (AUTO, satır {idx}).", False

        # --- Örgü uyumu (AUTO) ---
        job_orgu = ""
        for orgu_col in ["Zemin Örgü", "Zemin Orgu", "Örgü", "Orgu"]:
            if orgu_col in df.columns:
                job_orgu = str(df.at[idx, orgu_col]).strip()
                break

        current_orgu = (loom_orgu or "").strip()
        if job_orgu and current_orgu and (not _orgu_compatible(job_orgu, current_orgu)):
            # Süs kenarı uyumsuzluğu gibi: bu tezgahı pas geç
            return False, (
                f"Örgü uyumsuz (iş: {job_orgu}, tezgah: {current_orgu}) (AUTO, satır {idx})."
            ), False

        # --- Süs kenar uyumu (AUTO) ---
        job_sup = ""
        if "SÜS KENAR" in df.columns:
            job_sup = str(df.at[idx, "SÜS KENAR"]).strip()
        elif "Süs Kenar" in df.columns:
            job_sup = str(df.at[idx, "Süs Kenar"]).strip()

        current_sup = (loom_sup or "").strip()

        tarak_group = ""
        if "Tarak Grubu" in df.columns:
            tarak_group = str(df.at[idx, "Tarak Grubu"])

        if job_sup and current_sup:
            if not _selvedge_compatible_auto(job_sup, current_sup, tarak_group):
                # Bu tezgah bu iş için uygun değil → İş ve tezgah değişmedi
                return False, (
                    f"Süs kenarı uyumsuz (iş: {job_sup}, tezgah: {current_sup}) (AUTO, satır {idx})."
                ), False

        # --- Atama ---
        self.df_jobs.at[idx, "Tezgah Numarası"] = loom_no
        return True, f"{loom_no} tezgaha atandı (AUTO, satır {idx}).", True

    def _assign_from_table(self, source: str, idx: QModelIndex):
        if not idx.isValid():
            return
        model = self.model_free if source == "free" else self.model_soon
        df_view = getattr(model, "_df", pd.DataFrame())
        if df_view is None or df_view.empty:
            return

        # Seçilen tezgah
        try:
            loom_no = str(df_view.iloc[idx.row()]["Tezgah"]).strip()
        except Exception:
            return

        # Grup/key zorunlu
        key = self._current_key()
        if not key:
            QMessageBox.warning(self, "Uyarı", "Önce gruptan bir iş seçin.")
            return

        # Seçili satırdan Süs Kenar (tezgahın süs kenarı)
        loom_sup = ""
        row = df_view.iloc[idx.row()]
        for sup_col in ["SÜS KENAR", "Süs Kenar", "SüsKenarı", "Süs_Kenar", "SUS KENAR"]:
            if sup_col in row.index:
                loom_sup = str(row.get(sup_col, "")).strip()
                break

        loom_orgu = str(row.get("Örgü", "")).strip()
        ok, msg, remove_row = self._assign_first_job(key, loom_no, loom_sup, loom_orgu)

        if ok:
            if remove_row:
                # tablodan seçilen satırı kaldır
                new_df = df_view.drop(index=idx.row()).reset_index(drop=True)
                model.set_df(new_df)

            # üst akışa haber ver
            if callable(self.on_assign):
                group_label = self._current_group_label or ""
                category = self._current_category or ""
                self.on_assign(group_label, category)

            QMessageBox.information(self, "Atandı", msg)
        else:
            QMessageBox.information(self, "Bilgi", msg)
