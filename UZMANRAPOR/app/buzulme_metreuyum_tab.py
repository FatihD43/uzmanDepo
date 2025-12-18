from __future__ import annotations

import re
import pandas as pd

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog,
    QLabel, QTableView, QMessageBox
)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont

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
        top.addStretch(1)
        top.addWidget(self.lbl_info)

        # Tablo
        self.tbl = QTableView()
        self.model = PandasModel(pd.DataFrame())
        self.proxy = MultiColumnFilterProxy(self)
        self.proxy.setSourceModel(self.model)
        self.tbl.setModel(self.proxy)

        v.addLayout(top)
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

        try:
            n_isemri = out["Dokuma İş Emri"].nunique() if "Dokuma İş Emri" in out.columns else len(out)
            n_tip = out["Tip Kodu"].nunique() if "Tip Kodu" in out.columns else 0
            self.lbl_info.setText(f"İş Emri: {n_isemri} | Tip: {n_tip} | Satır: {len(out)}")
        except Exception:
            self.lbl_info.setText(f"Satır: {len(out)}")

        QTimer.singleShot(0, lambda: self.tbl.resizeColumnsToContents())

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
            # Kolonları da gösterelim ki bir daha mapping kaçırmayalım
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
        # 100 - ((DokumaToplam / İhzaratToplam) * 100)
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

        return out
