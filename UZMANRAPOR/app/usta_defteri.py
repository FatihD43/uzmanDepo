from __future__ import annotations
from typing import Optional, List, Dict
from datetime import datetime

import pandas as pd
import pyodbc
from PySide6.QtCore import Qt, QDate, QTimer
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QLineEdit, QComboBox,
    QPushButton, QDateEdit, QTableWidget, QTableWidgetItem, QFileDialog, QMessageBox, QGroupBox,
    QInputDialog, QDialog, QListWidget, QListWidgetItem
)

SQL_CONN_STR = (
    "Driver={SQL Server};"
    "Server=10.30.9.14,1433;"
    "Database=UzmanRaporDB;"
    "UID=uzmanrapor_login;"
    "PWD=03114080Ww.;"
)

IS_TANIM_LIST = ["DÜĞÜM", "TAKIM", "BAKIM", "DİĞER"]


def _vardiya_str(now_qtime) -> str:
    h = now_qtime.hour()
    m = now_qtime.minute()
    t = h + m / 60.0
    if 7 <= t < 15:
        vs = "(07:00)"
    elif 15 <= t < 23:
        vs = "(15:00)"
    else:
        vs = "(23:00)"
    return f"{vs}|{h:02d}:{m:02d}"


def _ensure_db():
    return


def _df_to_table(table: QTableWidget, df: pd.DataFrame):
    table.setRowCount(0)
    if df is None or df.empty:
        table.setRowCount(0)
        return
    table.setRowCount(len(df))
    table.setColumnCount(len(df.columns))
    table.setHorizontalHeaderLabels([str(c) for c in df.columns])
    for r in range(len(df)):
        for c, col in enumerate(df.columns):
            v = df.iloc[r, c]
            s = "" if pd.isna(v) else str(v)
            it = QTableWidgetItem(s)
            if c in [0, 1, 2]:
                it.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            else:
                it.setTextAlignment(Qt.AlignCenter)
            table.setItem(r, c, it)
    table.resizeColumnsToContents()
    if table.columnCount() > 0:
        w = max(120, table.columnWidth(0))
        table.setColumnWidth(0, w)


def _strip_trailing_dot_zero(val) -> str:
    if val is None:
        return ""
    try:
        if isinstance(val, float) and pd.isna(val):
            return ""
    except Exception:
        pass
    s = str(val).strip()
    if s.endswith(".0") or s.endswith(".00"):
        s = s.split(".", 1)[0]
    return s


class UstaDefteriWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        _ensure_db()

        self.df_jobs: Optional[pd.DataFrame] = None
        self.machine_list: List[str] = []

        root = QVBoxLayout(self)

        top = QHBoxLayout()
        top.addWidget(self._build_form_box(), 7)
        top.addWidget(self._build_report_box(), 5)
        root.addLayout(top)

        self.tbl = QTableWidget(0, 0)
        self.tbl.setSelectionBehavior(QTableWidget.SelectRows)
        self.tbl.setSelectionMode(QTableWidget.SingleSelection)
        self.tbl.setStyleSheet("""
            QTableWidget::item:selected {
                background-color: #0078d7;
                color: white;
            }
        """)
        root.addWidget(self.tbl, 1)

        self._configure_table_look()
        self._apply_beauty_theme()

        QTimer.singleShot(0, self._load_last_n)
        self._clear_form()

        # combo'ları SQL listelerinden doldur
        QTimer.singleShot(0, self._refresh_usta_combo)
        QTimer.singleShot(0, self._refresh_hasil_combo)

    # -------------------- DB bağlantısı ------------------------------------
    def _conn(self):
        return pyodbc.connect(SQL_CONN_STR)

    # -------------------- PUBLIC API ---------------------------------
    def set_sources(self, df_dinamik: Optional[pd.DataFrame]):
        self.df_jobs = df_dinamik.copy() if df_dinamik is not None else None

    def set_machine_list(self, looms: List[str]):
        full = [str(x) for x in range(2201, 2519)]
        self.machine_list = full
        self.cmb_tezgah.clear()
        self.cmb_tezgah.addItem("")
        self.cmb_tezgah.addItems(self.machine_list)

    # -------------------- LOOKUP LISTS (SQL) --------------------------
    def _fetch_lookup_rows(self, list_name: str) -> List[Dict[str, object]]:
        sql = """
        SELECT Id, Value
        FROM dbo.AppLookupValues
        WHERE ListName = ? AND IsActive = 1
        ORDER BY SortOrder, Value;
        """
        with self._conn() as c:
            cur = c.cursor()
            cur.execute(sql, (list_name,))
            rows = cur.fetchall()
        out: List[Dict[str, object]] = []
        for r in rows:
            out.append({"id": int(r[0]), "value": str(r[1])})
        return out

    def _ensure_lookup_value(self, list_name: str, value: str, created_by: str = "AUTO") -> None:
        """Yoksa ekler; varsa dokunmaz. (Unique index var.)"""
        value = (value or "").strip()
        if not value:
            return
        sql = """
        IF NOT EXISTS (SELECT 1 FROM dbo.AppLookupValues WHERE ListName = ? AND Value = ?)
        BEGIN
            INSERT INTO dbo.AppLookupValues (ListName, Value, IsActive, SortOrder, CreatedBy)
            VALUES (?, ?, 1, 0, ?)
        END
        """
        with self._conn() as c:
            cur = c.cursor()
            cur.execute(sql, (list_name, value, list_name, value, created_by))
            c.commit()

    def _update_lookup_value(self, row_id: int, new_value: str, updated_by: str = "UI") -> None:
        new_value = (new_value or "").strip()
        if not new_value:
            return
        sql = """
        UPDATE dbo.AppLookupValues
        SET Value = ?, UpdatedAt = SYSDATETIME(), UpdatedBy = ?
        WHERE Id = ?;
        """
        with self._conn() as c:
            cur = c.cursor()
            cur.execute(sql, (new_value, updated_by, row_id))
            c.commit()

    def _deactivate_lookup_value(self, row_id: int, updated_by: str = "UI") -> None:
        sql = """
        UPDATE dbo.AppLookupValues
        SET IsActive = 0, UpdatedAt = SYSDATETIME(), UpdatedBy = ?
        WHERE Id = ?;
        """
        with self._conn() as c:
            cur = c.cursor()
            cur.execute(sql, (updated_by, row_id))
            c.commit()

    def _refresh_usta_combo(self) -> None:
        if not hasattr(self, "cmb_usta"):
            return
        current = self.cmb_usta.currentText().strip()
        self.cmb_usta.blockSignals(True)
        self.cmb_usta.clear()
        self.cmb_usta.addItem("")
        for row in self._fetch_lookup_rows("USTA"):
            self.cmb_usta.addItem(row["value"])
        if current:
            self.cmb_usta.setCurrentText(current)
        self.cmb_usta.blockSignals(False)

    def _refresh_hasil_combo(self) -> None:
        if not hasattr(self, "cmb_hasilno"):
            return
        current = self.cmb_hasilno.currentText().strip()
        self.cmb_hasilno.blockSignals(True)
        self.cmb_hasilno.clear()
        self.cmb_hasilno.setEditable(True)
        self.cmb_hasilno.insertItem(0, "")
        for row in self._fetch_lookup_rows("HASIL"):
            self.cmb_hasilno.addItem(row["value"])
        if current:
            self.cmb_hasilno.setCurrentText(current)
        self.cmb_hasilno.blockSignals(False)

    def _open_manage_dialog(self):
        pw, ok = QInputDialog.getText(self, "Yetki", "Şifre girin:", QLineEdit.Password)
        if not ok:
            return
        if pw != "itema2024":
            QMessageBox.warning(self, "Uyarı", "Şifre hatalı.")
            return

        dlg = QDialog(self)
        dlg.setWindowTitle("Usta / Haşıl Yönetimi (SQL)")
        dlg.resize(520, 520)
        layout = QVBoxLayout(dlg)

        lst_usta = QListWidget()
        lst_hasil = QListWidget()

        def fill_list(target: QListWidget, listname: str):
            target.clear()
            for row in self._fetch_lookup_rows(listname):
                it = QListWidgetItem(row["value"])
                it.setData(Qt.UserRole, row["id"])  # Id sakla
                target.addItem(it)

        fill_list(lst_usta, "USTA")
        fill_list(lst_hasil, "HASIL")

        def add_item(target: QListWidget, listname: str, title: str):
            text, ok2 = QInputDialog.getText(dlg, title, "Yeni değer:")
            if not (ok2 and text.strip()):
                return
            val = text.strip()
            try:
                self._ensure_lookup_value(listname, val, created_by="UI")
            except Exception as e:
                QMessageBox.warning(dlg, "Uyarı", f"Ekleme başarısız:\n{e}")
                return
            fill_list(target, listname)

        def edit_item(target: QListWidget, listname: str, title: str):
            item = target.currentItem()
            if not item:
                return
            row_id = item.data(Qt.UserRole)
            old = item.text()
            text, ok2 = QInputDialog.getText(dlg, title, "Düzenle:", text=old)
            if not (ok2 and text.strip()):
                return
            new_val = text.strip()
            try:
                self._update_lookup_value(int(row_id), new_val, updated_by="UI")
            except Exception as e:
                QMessageBox.warning(dlg, "Uyarı", f"Güncelleme başarısız:\n{e}")
                return
            fill_list(target, listname)

        def deactivate_item(target: QListWidget, listname: str):
            item = target.currentItem()
            if not item:
                return
            row_id = item.data(Qt.UserRole)
            res = QMessageBox.question(dlg, "Onay", f"'{item.text()}' pasife alınsın mı?")
            if res != QMessageBox.Yes:
                return
            try:
                self._deactivate_lookup_value(int(row_id), updated_by="UI")
            except Exception as e:
                QMessageBox.warning(dlg, "Uyarı", f"İşlem başarısız:\n{e}")
                return
            fill_list(target, listname)

        usta_box = QGroupBox("İşlem Yapan (Usta)")
        usta_layout = QVBoxLayout(usta_box)
        usta_layout.addWidget(lst_usta)
        u_btns = QHBoxLayout()
        btn_u_add = QPushButton("Ekle")
        btn_u_edit = QPushButton("Düzenle")
        btn_u_del = QPushButton("Pasife Al")
        u_btns.addWidget(btn_u_add)
        u_btns.addWidget(btn_u_edit)
        u_btns.addWidget(btn_u_del)
        usta_layout.addLayout(u_btns)

        hasil_box = QGroupBox("Haşıl No")
        hasil_layout = QVBoxLayout(hasil_box)
        hasil_layout.addWidget(lst_hasil)
        h_btns = QHBoxLayout()
        btn_h_add = QPushButton("Ekle")
        btn_h_edit = QPushButton("Düzenle")
        btn_h_del = QPushButton("Pasife Al")
        h_btns.addWidget(btn_h_add)
        h_btns.addWidget(btn_h_edit)
        h_btns.addWidget(btn_h_del)
        hasil_layout.addLayout(h_btns)

        layout.addWidget(usta_box)
        layout.addWidget(hasil_box)

        footer = QHBoxLayout()
        btn_ok = QPushButton("Kapat")
        footer.addStretch(1)
        footer.addWidget(btn_ok)
        layout.addLayout(footer)

        btn_u_add.clicked.connect(lambda: add_item(lst_usta, "USTA", "Usta Ekle"))
        btn_u_edit.clicked.connect(lambda: edit_item(lst_usta, "USTA", "Usta Düzenle"))
        btn_u_del.clicked.connect(lambda: deactivate_item(lst_usta, "USTA"))

        btn_h_add.clicked.connect(lambda: add_item(lst_hasil, "HASIL", "Haşıl Ekle"))
        btn_h_edit.clicked.connect(lambda: edit_item(lst_hasil, "HASIL", "Haşıl Düzenle"))
        btn_h_del.clicked.connect(lambda: deactivate_item(lst_hasil, "HASIL"))

        btn_ok.clicked.connect(dlg.accept)

        dlg.exec()
        # kapatırken form combo'larını güncelle
        self._refresh_usta_combo()
        self._refresh_hasil_combo()

    # -------------------- UI ----------------------------------
    def _build_form_box(self) -> QGroupBox:
        box = QGroupBox("DÜĞÜM TAKIM GİRİŞ")
        grid = QGridLayout(box)

        r = 0
        grid.addWidget(QLabel("TARİH"), r, 0)
        self.dt_tarih = QDateEdit(QDate.currentDate())
        self.dt_tarih.setDisplayFormat("dd.MM.yyyy")
        self.dt_tarih.setCalendarPopup(True)
        grid.addWidget(self.dt_tarih, r, 1)

        grid.addWidget(QLabel("VARDİYA&SAAT"), r + 1, 0)
        self.txt_vardiya = QLineEdit()
        self.txt_vardiya.setReadOnly(True)
        grid.addWidget(self.txt_vardiya, r + 1, 1)

        grid.addWidget(QLabel("TEZGAH"), r + 2, 0)
        self.cmb_tezgah = QComboBox()
        self.cmb_tezgah.addItem("")
        self.cmb_tezgah.addItems([str(x) for x in range(2201, 2519)])
        grid.addWidget(self.cmb_tezgah, r + 2, 1)

        grid.addWidget(QLabel("KÖKTİP"), r + 3, 0)
        self.ed_koktip = QLineEdit()
        grid.addWidget(self.ed_koktip, r + 3, 1)

        grid.addWidget(QLabel("HAŞIL İŞ EMRİ"), r + 4, 0)
        self.ed_hasis = QLineEdit()
        grid.addWidget(self.ed_hasis, r + 4, 1)

        grid.addWidget(QLabel("LEVENT NO"), r + 5, 0)
        self.ed_levent = QLineEdit()
        grid.addWidget(self.ed_levent, r + 5, 1)

        grid.addWidget(QLabel("ETİKET NO"), r + 6, 0)
        self.ed_etiket = QLineEdit()
        grid.addWidget(self.ed_etiket, r + 6, 1)

        c = 2
        grid.addWidget(QLabel("DOKUMA İŞ EMRİ"), r, c)
        self.ed_dokuma = QLineEdit()
        grid.addWidget(self.ed_dokuma, r, c + 1)

        grid.addWidget(QLabel("TOPLAM METRE"), r + 1, c)
        self.ed_metre = QLineEdit()
        grid.addWidget(self.ed_metre, r + 1, c + 1)

        grid.addWidget(QLabel("HAŞIL NO"), r + 2, c)
        self.cmb_hasilno = QComboBox()
        self.cmb_hasilno.setEditable(True)
        self.cmb_hasilno.insertItem(0, "")
        grid.addWidget(self.cmb_hasilno, r + 2, c + 1)

        grid.addWidget(QLabel("İŞ TANIMI"), r + 3, c)
        self.cmb_is = QComboBox()
        self.cmb_is.addItem("")
        self.cmb_is.addItems(IS_TANIM_LIST)
        grid.addWidget(self.cmb_is, r + 3, c + 1)

        grid.addWidget(QLabel("TİP ÖZELLİKLERİ (Tarak;Örgü)"), r + 4, c)
        self.ed_tip = QLineEdit()
        grid.addWidget(self.ed_tip, r + 4, c + 1)

        grid.addWidget(QLabel("İŞLEM YAPAN"), r + 5, c)
        self.cmb_usta = QComboBox()
        self.cmb_usta.addItem("")
        grid.addWidget(self.cmb_usta, r + 5, c + 1)

        grid.addWidget(QLabel("AÇIKLAMA"), r + 7, 0)
        self.ed_aciklama = QLineEdit()
        grid.addWidget(self.ed_aciklama, r + 7, 1, 1, 4)

        self.btn_bul = QPushButton("LEVENT BUL")
        self.btn_kaydet = QPushButton("KAYDET")
        self.btn_sil = QPushButton("SİL")
        self.btn_manage_lists = QPushButton("USTA/HAŞIL YÖNET")

        grid.addWidget(self.btn_bul, r, 4)
        grid.addWidget(self.btn_kaydet, r + 1, 4)
        grid.addWidget(self.btn_sil, r + 2, 4)
        grid.addWidget(self.btn_manage_lists, r + 4, 4)  # biraz ayrık

        self.btn_bul.clicked.connect(self._on_levent_bul)
        self.btn_kaydet.clicked.connect(self._on_save)
        self.btn_sil.clicked.connect(self._on_delete)
        self.btn_manage_lists.clicked.connect(self._open_manage_dialog)

        def _tick():
            from PySide6.QtCore import QTime
            self.txt_vardiya.setText(_vardiya_str(QTime.currentTime()))

        _tick()
        timer = QTimer(self)
        timer.timeout.connect(_tick)
        timer.start(60_000)

        return box

    def _build_report_box(self) -> QGroupBox:
        box = QGroupBox("RAPOR ALMA SEÇENEKLERİ")
        grid = QGridLayout(box)

        grid.addWidget(QLabel("İLK TARİH"), 0, 0)
        self.dt_ilk = QDateEdit(QDate.currentDate())
        self.dt_ilk.setCalendarPopup(True)
        self.dt_ilk.setDisplayFormat("dd.MM.yyyy")
        grid.addWidget(self.dt_ilk, 0, 1)

        grid.addWidget(QLabel("SON TARİH"), 1, 0)
        self.dt_son = QDateEdit(QDate.currentDate())
        self.dt_son.setCalendarPopup(True)
        self.dt_son.setDisplayFormat("dd.MM.yyyy")
        grid.addWidget(self.dt_son, 1, 1)

        grid.addWidget(QLabel("ÇOKLU SEÇİM"), 2, 0)
        self.cmb_field = QComboBox()
        self.cmb_field.addItems([
            "Tezgah", "KökTip", "Haşıl İş Emri", "Dokuma İş Emri",
            "Levent No", "Etiket No", "İş Tanımı", "İşlem Yapan"
        ])
        grid.addWidget(self.cmb_field, 2, 1)

        self.ed_value = QLineEdit()
        grid.addWidget(self.ed_value, 3, 0, 1, 2)

        grid.addWidget(QLabel("HIZLI BUL"), 4, 0)
        self.ed_q = QLineEdit()
        grid.addWidget(self.ed_q, 4, 1)

        self.btn_getir = QPushButton("RAPOR AL")
        self.btn_excel = QPushButton("RAPORU SAYFADA GÖR")
        grid.addWidget(self.btn_getir, 5, 0, 1, 2)
        grid.addWidget(self.btn_excel, 6, 0, 1, 2)

        self.btn_getir.clicked.connect(self._run_report)
        self.btn_excel.clicked.connect(self._export_excel)
        self.ed_q.textChanged.connect(self._apply_quick_filter)

        return box

    # -------------------- DB ops (UstaDefteri) ------------------------
    def _insert_row(self, rec: Dict):
        colmap = {
            "tarih": "Tarih",
            "vardiya": "Vardiya",
            "tezgah": "Tezgah",
            "koktip": "KokTip",
            "hasis_no": "HasisNo",
            "levent_no": "LeventNo",
            "etiket_no": "EtiketNo",
            "dokuma_is_emri": "DokumaIsEmri",
            "metre": "Metre",
            "hasil_no": "HasilNo",
            "is_tanimi": "IsTanimi",
            "tarak_grubu": "TarakGrubu",
            "orgu": "Orgu",
            "tip_ozellikleri": "TipOzellikleri",
            "islem_yapan": "IslemYapan",
            "aciklama": "Aciklama",
            "yapilan_islem": "YapilanIslem",
        }

        mapped: Dict[str, object] = {}
        for k, v in rec.items():
            col = colmap.get(k)
            if not col:
                continue
            if col == "Tarih":
                try:
                    dt_val = datetime.strptime(str(v), "%d.%m.%Y").date()
                    mapped[col] = dt_val.strftime("%Y-%m-%d")
                except Exception:
                    mapped[col] = None
            elif col == "Metre":
                if v is None or str(v).strip() == "":
                    mapped[col] = None
                else:
                    try:
                        mapped[col] = float(str(v).replace(",", "."))
                    except Exception:
                        mapped[col] = None
            else:
                s = "" if v is None else str(v)
                mapped[col] = s if s != "" else None

        cols = list(mapped.keys())
        placeholders = ",".join("?" for _ in cols)
        sql = f"INSERT INTO dbo.UstaDefteri ({','.join(cols)}) VALUES ({placeholders})"

        with self._conn() as c:
            cur = c.cursor()
            cur.execute(sql, [mapped[name] for name in cols])
            c.commit()

    def _delete_by_rowid(self, rowid: int):
        with self._conn() as c:
            cur = c.cursor()
            cur.execute("DELETE FROM dbo.UstaDefteri WHERE Id = ?", (rowid,))
            c.commit()

    def _select(self, start: Optional[str] = None, end: Optional[str] = None,
                field: Optional[str] = None, value: Optional[str] = None) -> pd.DataFrame:
        sql = """
        SELECT
            Id AS Id,
            CONVERT(varchar(10), Tarih, 104) AS Tarih,
            Vardiya AS Saat,
            Tezgah AS Tezgah,
            KokTip AS Takdir,
            HasisNo AS [Haşıl İşEm],
            LeventNo AS Levent,
            EtiketNo AS Etiket,
            DokumaIsEmri AS [Dokuma İş Emri],
            Metre AS Metre,
            HasilNo AS [Haşıl no],
            IsTanimi AS [İş tanımı],
            YapilanIslem AS [Yapılan işlem],
            IslemYapan AS [İşlem Yapan],
            Aciklama AS [Açıklama]
        FROM dbo.UstaDefteri
        WHERE 1 = 1
        """
        params: list[object] = []

        if start:
            try:
                start_date = datetime.strptime(start, "%d.%m.%Y").date()
                sql += " AND Tarih >= ?"
                params.append(start_date)
            except Exception:
                pass

        if end:
            try:
                end_date = datetime.strptime(end, "%d.%m.%Y").date()
                sql += " AND Tarih <= ?"
                params.append(end_date)
            except Exception:
                pass

        if field and value:
            col = {
                "Tezgah": "Tezgah",
                "KökTip": "KokTip",
                "Haşıl İş Emri": "HasisNo",
                "Dokuma İş Emri": "DokumaIsEmri",
                "Levent No": "LeventNo",
                "Etiket No": "EtiketNo",
                "İş Tanımı": "IsTanimi",
                "İşlem Yapan": "IslemYapan",
            }.get(field)
            if col:
                sql += f" AND {col} LIKE ?"
                params.append(f"%{value}%")

        with self._conn() as c:
            cur = c.cursor()
            cur.execute(sql, tuple(params)) if params else cur.execute(sql)
            rows = cur.fetchall()
            cols = [d[0] for d in cur.description]
            df = pd.DataFrame.from_records(rows, columns=cols)
        return df

    def _clear_form(self):
        if self.cmb_tezgah.count():
            self.cmb_tezgah.setCurrentIndex(0)
        if self.cmb_hasilno.count():
            self.cmb_hasilno.setCurrentIndex(0)
        if self.cmb_is.count():
            self.cmb_is.setCurrentIndex(0)
        if self.cmb_usta.count():
            self.cmb_usta.setCurrentIndex(0)
        for w in [
            self.ed_koktip, self.ed_hasis, self.ed_levent, self.ed_etiket,
            self.ed_dokuma, self.ed_metre, self.ed_tip, self.ed_aciklama
        ]:
            w.clear()

    # -------------------- Actions ------------------------------------
    def _on_levent_bul(self):
        self._clear_form()
        lev, ok = QInputDialog.getText(self, "Levent Bul", "Levent No girin:")
        if not ok or not lev.strip():
            return
        lev = lev.strip()

        if self.df_jobs is None or self.df_jobs.empty:
            QMessageBox.information(self, "Bilgi", "Dinamik rapor yüklenmemiş.")
            return

        df = self.df_jobs.copy()
        s = df.get("Levent No", "").astype(str).str.strip()
        hit = df[s == lev]
        if hit.empty:
            QMessageBox.information(self, "Bilgi", f"Levent {lev} bulunamadı.")
            return
        row = hit.iloc[0]

        self.ed_levent.setText(_strip_trailing_dot_zero(lev))
        self.ed_koktip.setText(str(row.get("Kök Tip Kodu", "")).strip())
        self.ed_dokuma.setText(str(row.get("Üretim Sipariş No", "")).strip())
        self.ed_hasis.setText(_strip_trailing_dot_zero(row.get("Haşıl İş Emri", "")))
        self.ed_etiket.setText(_strip_trailing_dot_zero(row.get("Levent Etiket FA", "")))

        tarak = str(row.get("Tarak Grubu", "")).strip()
        orgu = str(row.get("Zemin Örgü", "")).strip()
        self.ed_tip.setText(f"{tarak};{orgu}")

        pm = row.get("Parti Metresi", "")
        self.ed_metre.setText("" if pd.isna(pm) else str(pm))

        hn = str(row.get("Haşıl No", "")).strip() if "Haşıl No" in row else ""
        if hn:
            self._ensure_lookup_value("HASIL", hn, created_by="LEVENT_BUL")
            self._refresh_hasil_combo()
            self.cmb_hasilno.setCurrentText(hn)

    def _on_save(self):
        try:
            tarih = self.dt_tarih.date().toString("dd.MM.yyyy")
            vardiya = self.txt_vardiya.text().strip()
            tezgah = self.cmb_tezgah.currentText().strip()
            koktip = self.ed_koktip.text().strip()
            hasis = _strip_trailing_dot_zero(self.ed_hasis.text().strip())
            levent = self.ed_levent.text().strip()
            etiket = _strip_trailing_dot_zero(self.ed_etiket.text().strip())

            if self._etiket_exists(etiket):
                QMessageBox.warning(
                    self, "Mükerrer Etiket",
                    f"'{etiket}' etiket numarası zaten kayıtlı. Kayıt yapılmadı."
                )
                return

            dokuma = self.ed_dokuma.text().strip()
            metre_txt = self.ed_metre.text().replace(",", ".").strip()
            metre = float(metre_txt) if metre_txt else None

            hasilno = self.cmb_hasilno.currentText().strip()
            if hasilno:
                self._ensure_lookup_value("HASIL", hasilno, created_by="SAVE")

            is_tanimi = self.cmb_is.currentText().strip()
            tip = self.ed_tip.text().strip()
            tarak, orgu = "", ""
            if ";" in tip:
                tarak, orgu = [x.strip() for x in tip.split(";", 1)]
            islem_yapan = self.cmb_usta.currentText().strip()
            aciklama = self.ed_aciklama.text().strip()

            yapilan = f"{tarak} ; {orgu}" if (tarak or orgu) else ""

            rec = dict(
                tarih=tarih,
                vardiya=vardiya,
                tezgah=tezgah,
                koktip=koktip,
                hasis_no=hasis,
                levent_no=levent,
                etiket_no=etiket,
                dokuma_is_emri=dokuma,
                metre=metre,
                hasil_no=hasilno,
                is_tanimi=is_tanimi,
                tarak_grubu=tarak,
                orgu=orgu,
                tip_ozellikleri=tip,
                islem_yapan=islem_yapan,
                aciklama=aciklama,
                yapilan_islem=yapilan,
            )
            self._insert_row(rec)
            QMessageBox.information(self, "Kaydedildi", "Kayıt eklendi.")
            self._load_last_n()
            self._clear_form()
            self._refresh_hasil_combo()
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Kayıt eklenemedi:\n{e}")

    def _on_delete(self):
        row = self.tbl.currentRow()
        if row < 0:
            QMessageBox.information(self, "Bilgi", "Silmek için bir satır seçin.")
            return
        id_item = self.tbl.item(row, 0)
        if not id_item:
            return
        try:
            rid = int(id_item.text())
        except Exception:
            QMessageBox.warning(self, "Uyarı", "Satır kimliği okunamadı.")
            return
        res = QMessageBox.question(self, "Onay", "Seçili kaydı silmek istiyor musunuz?")
        if res != QMessageBox.Yes:
            return
        self._delete_by_rowid(rid)
        self._load_last_n()

    def _run_report(self):
        start = self.dt_ilk.date().toString("dd.MM.yyyy")
        end = self.dt_son.date().toString("dd.MM.yyyy")
        field = self.cmb_field.currentText()
        value = self.ed_value.text().strip()
        df = self._select(start, end, field if value else None, value if value else None)
        self._raw_df = df
        _df_to_table(self.tbl, df)

    def _export_excel(self):
        if not hasattr(self, "_raw_df") or self._raw_df is None or self._raw_df.empty:
            QMessageBox.information(self, "Bilgi", "Önce raporu alın.")
            return
        out, _ = QFileDialog.getSaveFileName(self, "Excel'e aktar", "usta_defteri.xlsx", "Excel Files (*.xlsx)")
        if not out:
            return
        df = self._raw_df.copy()
        if "Id" in df.columns:
            df = df.drop(columns=["Id"])
        try:
            import xlsxwriter
            with pd.ExcelWriter(out, engine="xlsxwriter", datetime_format="dd.mm.yyyy") as xw:
                sheet = "USTA_DEFTERI"
                df.to_excel(xw, index=False, sheet_name=sheet)
                ws = xw.sheets[sheet]
                wb = xw.book
                header_fmt = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "valign": "vcenter"})
                ws.set_row(0, 22, header_fmt)
                for c, col in enumerate(df.columns):
                    series_as_str = df[col].astype(str).replace("nan", "")
                    max_len = max(len(str(col)), *(len(s) for s in series_as_str.values))
                    ws.set_column(c, c, min(max_len + 2, 60))
                ws.freeze_panes(1, 0)
            QMessageBox.information(self, "Excel", "Dosya oluşturuldu.")
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Excel'e aktarılamadı:\n{e}")

    def _apply_quick_filter(self):
        if not hasattr(self, "_raw_df"):
            return
        q = self.ed_q.text().strip().lower()
        if not q:
            _df_to_table(self.tbl, self._raw_df)
            return
        df = self._raw_df.copy()
        mask = pd.Series([False] * len(df))
        for col in df.columns:
            mask = mask | df[col].astype(str).str.lower().str.contains(q, na=False)
        _df_to_table(self.tbl, df[mask])

    def _load_last_n(self, n: int = 200):
        sql = f"""
        SELECT TOP {n}
               Id,
               CONVERT(varchar(10), Tarih, 104) AS Tarih,
               Vardiya AS Saat,
               Tezgah AS Tezgah,
               KokTip AS Takdir,
               HasisNo AS [Haşıl İşEm],
               LeventNo AS Levent,
               EtiketNo AS Etiket,
               DokumaIsEmri AS [Dokuma İş Emri],
               Metre AS Metre,
               HasilNo AS [Haşıl no],
               IsTanimi AS [İş tanımı],
               YapilanIslem AS [Yapılan işlem],
               IslemYapan AS [İşlem Yapan],
               Aciklama AS [Açıklama]
        FROM dbo.UstaDefteri
        ORDER BY Id DESC;
        """
        with self._conn() as c:
            cur = c.cursor()
            cur.execute(sql)
            rows = cur.fetchall()
            cols = [d[0] for d in cur.description]
            df = pd.DataFrame.from_records(rows, columns=cols)

        self._raw_df = df
        _df_to_table(self.tbl, df)

    def _etiket_exists(self, etiket: str) -> bool:
        if not etiket:
            return False
        with self._conn() as c:
            cur = c.cursor()
            cur.execute("SELECT 1 FROM dbo.UstaDefteri WHERE EtiketNo = ?;", (etiket,))
            return cur.fetchone() is not None

    def _configure_table_look(self):
        self.tbl.setSelectionBehavior(QTableWidget.SelectRows)
        self.tbl.setSelectionMode(QTableWidget.SingleSelection)
        self.tbl.setAlternatingRowColors(True)

        hh = self.tbl.horizontalHeader()
        vh = self.tbl.verticalHeader()
        vh.setVisible(False)
        hh.setStretchLastSection(True)
        hh.setHighlightSections(False)

        self.tbl.setStyleSheet("""
            QTableWidget {
                gridline-color: #e5e7eb;
                background: #ffffff;
                alternate-background-color: #f8fafc;
                selection-background-color: transparent;
            }
            QTableWidget::item {
                padding: 6px 8px;
            }
            QTableWidget::item:hover {
                background-color: #eef6ff;
            }
            QTableWidget::item:selected {
                background-color: #0078d7;
                color: #ffffff;
            }
            QHeaderView::section {
                background: #f3f4f6;
                color: #111827;
                padding: 6px 8px;
                border: 0px;
                border-right: 1px solid #e5e7eb;
                font-weight: 600;
            }
            QHeaderView::section:horizontal {
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
            }
        """)

    def _apply_beauty_theme(self):
        if isinstance(self.layout(), QVBoxLayout):
            self.layout().setContentsMargins(12, 12, 12, 12)
            self.layout().setSpacing(10)

        self.setStyleSheet("""
            QGroupBox {
                background: #ffffff;
                border: 1px solid #e5e7eb;
                border-radius: 10px;
                margin-top: 14px;
                padding: 8px 10px 12px 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0px 6px;
                color: #111827;
                font-weight: 700;
                background: transparent;
            }
            QLabel {
                color: #374151;
                font-weight: 600;
            }
            QLineEdit, QComboBox, QDateEdit {
                border: 1px solid #d1d5db;
                border-radius: 8px;
                padding: 6px 8px;
                background: #ffffff;
            }
            QLineEdit:focus, QComboBox:focus, QDateEdit:focus {
                border: 1px solid #0078d7;
                box-shadow: 0 0 0 3px rgba(0,120,215,0.15);
                outline: none;
            }
            QComboBox QAbstractItemView {
                border: 1px solid #d1d5db;
                background: #ffffff;
                selection-background-color: #e6f2ff;
            }
            QPushButton {
                background: #0ea5e9;
                color: #ffffff;
                border: none;
                border-radius: 10px;
                padding: 8px 12px;
                font-weight: 700;
            }
            QPushButton:hover { background: #0284c7; }
            QPushButton:pressed { background: #0369a1; }
            QPushButton:disabled {
                background: #9ca3af;
                color: #f9fafb;
            }
        """)
