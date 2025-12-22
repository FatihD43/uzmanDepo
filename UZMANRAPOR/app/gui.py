from __future__ import annotations
import re
import pandas as pd
from datetime import datetime, time, timedelta
from zoneinfo import ZoneInfo  # <-- Istanbul TZ
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog,
    QLabel, QTabWidget, QMessageBox, QLineEdit, QScrollArea, QGridLayout,
    QTableView, QHeaderView, QToolButton, QSizePolicy, QTextEdit, QDialog
)
from PySide6.QtCore import Qt, QTimer, QSettings
from typing import Any

from app.itema_tab import ItemaAyarTab
from app.models import PandasModel
from app.filter_proxy import MultiColumnFilterProxy
from io_layer.loaders import load_dinamik_any, load_running_orders, VISIBLE_COLUMNS, HEADER_ALIASES
from app.value_picker import ValuePickerDialog
from app.notes_dialog import NotesDialog
from app import storage
from app.kusbakisi import KusbakisiWidget
from app.planning_dialog import PlanningDialog
from app.usta_defteri import UstaDefteriWidget
from app.team_planning_flow import TeamPlanningFlowTab
from app.equipment_dialog import LoomCutEditor
from io_layer.loaders import enrich_running_with_loom_cut, enrich_running_with_selvedge
from app.auth import User
from app.user_management_widget import UserManagementWidget
from app.buzulme_metreuyum_tab import BuzulmeMetreUyumTab



def _normalize_perm_name(perm: str) -> str:
    return str(perm or "").strip().lower()


def _user_has_permission(user: Any, perm: str) -> bool:
    """Saf kullanıcı nesnesi üzerinde verilen izni sorgular."""

    normalized = _normalize_perm_name(perm)
    if not normalized:
        return True

    if user is None:
        return False

    has_perm_fn = getattr(user, "has_permission", None)
    if callable(has_perm_fn):
        try:
            return bool(has_perm_fn(perm))
        except Exception:
            pass

    raw_perms = getattr(user, "permissions", None)
    if raw_perms is None:
        return False

    try:
        normalized_perms = {
            _normalize_perm_name(p) for p in raw_perms if str(p).strip()
        }
    except Exception:
        normalized_perms = set()

    return normalized in normalized_perms or "admin" in normalized_perms


def require_permission(window: QWidget, perm: str, message: str) -> bool:
    """MainWindow örnekleri için güvenli izin denetimi."""

    helper = getattr(window, "_require_permission", None)
    if callable(helper):
        try:
            return bool(helper(perm, message))
        except Exception:
            pass

    has_perm_fn = getattr(window, "has_permission", None)
    if callable(has_perm_fn):
        try:
            if has_perm_fn(perm):
                return True
        except Exception:
            pass

    if _user_has_permission(getattr(window, "user", None), perm):
        return True

    try:
        QMessageBox.warning(window, "Yetki yok", message)
    except Exception:
        pass
    return False


# ============================================================
# RUNNING ORDERS NORMALİZASYON BLOĞU (tek noktadan düzeltme)
# ============================================================
def _parse_number_loose(x):
    """Metin/sayı karması 'Kalan' değerlerini güvenle floata çevirir.
       92,7 | 1.234,56 | 1,234.56 | ' 300 ' | '92,7 m' | '-' -> float/NA"""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return pd.NA
    s = str(x).strip()
    if s == "" or s == "-":
        return pd.NA
    # rakam, nokta, virgül, eksi dışını temizle
    s = re.sub(r"[^0-9,.\-]", "", s)

    if "," in s and "." in s:
        # En sağdaki ayırıcı ondalık kabul
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "")
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    else:
        if "," in s:
            parts = s.split(",")
            if len(parts[-1]) <= 2:
                s = s.replace(".", "")
                s = s.replace(",", ".")
            else:
                s = s.replace(",", "")
        elif "." in s:
            parts = s.split(".")
            if len(parts[-1]) > 2:  # 1.234 → binlik
                s = s.replace(".", "")

    try:
        return float(s)
    except Exception:
        return pd.NA


def _extract_nums_keep_decimal(text: str):
    """Tarak grubu normalize için: ondalığı koruyarak sayıları çıkar (virgül -> nokta)."""
    if text is None:
        return []
    nums = re.findall(r"[\d]+(?:[.,]\d+)?", str(text))
    out = []
    for n in nums:
        n = n.replace(",", ".")
        if re.fullmatch(r"\d+\.0+", n):
            n = n.split(".", 1)[0]
        out.append(n)
    return out


def _norm_tarak_generic(val) -> str:
    """Dinamik/Running fark etmez: 'a/b/c' (ilk 3 sayı) şeklinde normalize anahtar."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    parts = _extract_nums_keep_decimal(str(val))
    if not parts:
        return str(val).strip()
    return "/".join(parts[:3])


def _detect_94_row(row):
    """Running satırında 94 / 'Sipariş Yok' tespiti (kolon adı bağımsız)."""
    for c in row.index:
        u = str(row.get(c, "")).strip().upper()
        if "SİPARİŞ YOK" in u or "SIPARIS YOK" in u or u == "94" or " 94" in u:
            return True
    return False


def normalize_df_running(df_running: pd.DataFrame) -> pd.DataFrame:
    """Running Orders df'sine kanonik kolonlar ekler/yeniler:
       - _KalanMetreNorm  : float
       - _TG_norm         : 'a/b/c' normalize tarak
       - _OpenTezgahFlag  : bool (94 veya Durum='Bitti')
    """
    if df_running is None or df_running.empty:
        return df_running

    # 1) Kalan -> _KalanMetreNorm
    kalan_cols = ["Kalan", "Kalan Mt", "Kalan Metre", "Kalan_Metre", "_KalanMetre"]
    kal_col = next((c for c in kalan_cols if c in df_running.columns), None)
    if kal_col:
        df_running["_KalanMetreNorm"] = df_running[kal_col].apply(_parse_number_loose)
    else:
        df_running["_KalanMetreNorm"] = pd.NA

    # 2) Tarak Grubu normalize -> _TG_norm
    tg_col = next((c for c in ["Tarak Grubu", "Tarak", "TarakGrubu"] if c in df_running.columns), None)
    if tg_col:
        df_running["_TG_norm"] = df_running[tg_col].astype(str).apply(_norm_tarak_generic)
    else:
        df_running["_TG_norm"] = ""

    # 3) 94 bayrağı -> _OpenTezgahFlag
    df_running["_OpenTezgahFlag"] = df_running.apply(_detect_94_row, axis=1)

    # 4) Durum normalizasyonu (Bitti kontrolü)
    durum_col = next((c for c in ["Durum", "Durumu", "Durum Açıklaması", "Durum Tanım"] if c in df_running.columns), None)
    if durum_col:
        def _is_bitti(val: object) -> bool:
            s = str(val or "").strip().upper()
            return ("BİTTİ" in s) or ("BITTI" in s)
        bitti_series = df_running[durum_col].apply(_is_bitti)
    else:
        bitti_series = pd.Series(False, index=df_running.index)

    # 5) Açık kabul: 94 veya Durum=Bitti
    df_running["_OpenTezgahFlag"] = df_running["_OpenTezgahFlag"].astype(bool) | bitti_series.astype(bool)

    return df_running


# ============================================================
# **YENİ**: Tezgah listesi düzenleyici dialog (Arızalı/Bakımda & Boş Göster)
# ============================================================
class LoomListEditor(QDialog):
    """
    Basit düzenleyici: metin alanına tezgah numaralarını yaz (virgül/boşluk/alt satır ayırır).
    Kaydet → QSettings("UZMANRAPOR","ClientApp").setValue(settings_key, "2201, 2203, 2450")
    Ayrıca storage.* içine kalıcı JSON olarak da yazar.
    """
    def __init__(self, title: str, settings_key: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.settings = QSettings("UZMANRAPOR", "ClientApp")
        self.key = settings_key

        v = QVBoxLayout(self)
        v.addWidget(QLabel("Tezgah numaralarını girin (ör. 2201 2203 2450 veya 2201,2203,2450):"))
        self.txt = QTextEdit()
        v.addWidget(self.txt, 1)

        # mevcut değeri doldur (QSettings ∪ storage)
        self._load_initial_text()

        h = QHBoxLayout()
        h.addStretch(1)
        btn_ok = QPushButton("Kaydet")
        btn_ok.clicked.connect(self._save)
        h.addWidget(btn_ok)
        btn_cancel = QPushButton("İptal")
        btn_cancel.clicked.connect(self.reject)
        h.addWidget(btn_cancel)
        v.addLayout(h)

        self.resize(520, 360)

    def _load_initial_text(self):
        qv = str(self.settings.value(self.key, "") or "")
        q_tokens = re.findall(r"\d+", qv)

        extra = []
        try:
            if self.key == "looms/blocked":
                extra = storage.load_blocked_looms()
            elif self.key == "looms/empty":
                extra = storage.load_dummy_looms()
        except Exception:
            extra = []

        # birleştir ve sıralı/benzersiz yap
        try:
            all_nums = sorted(
                {*(q_tokens or []), *[str(x) for x in (extra or [])]},
                key=lambda s: int(re.findall(r"\d+", s)[0])
            )
        except Exception:
            all_nums = list({*(q_tokens or []), *[str(x) for x in (extra or [])]})
        self.txt.setPlainText(", ".join(all_nums))

    def _save(self):
        raw = self.txt.toPlainText()
        # Normalize et: sadece rakam dizilerini al, araya virgül+boşluk koy
        tokens = re.findall(r"\d+", raw or "")
        val = ", ".join(tokens)
        # 1) QSettings
        self.settings.setValue(self.key, val)
        # 2) storage SQL (merkezi kullanım için)
        try:
            if self.key == "looms/blocked":
                storage.save_blocked_looms(tokens)
            elif self.key == "looms/empty":
                storage.save_dummy_looms(tokens)
        except Exception:
            pass

        QMessageBox.information(self, "Kaydedildi", f"Ayar güncellendi.\n({self.key} = {val})")
        self.accept()


# ============================================================


class MainWindow(QMainWindow):
    def __init__(self, user: User | None = None):
        super().__init__()
        if user is None:
            raise ValueError("MainWindow requires an authenticated user")

        self.user = user
        self.setWindowTitle(f"UZMAN RAPOR — v6.4 (Kuşbakışı) • {self.user.username}")
        self.resize(1450, 900)

        # Sabit uygulama saat dilimi (İstanbul)
        self.TZ = ZoneInfo("Europe/Istanbul")

        self.df_dinamik_full = None
        self.df_running = None

        # Kalıcı kurallar ve son güncelleme
        self._note_rules: list[dict] = storage.load_rules()
        self._last_update: datetime | None = storage.load_last_update()

        tabs = QTabWidget()
        tabs.addTab(self.build_dugum_tab(), "DÜĞÜM TAKIM LİSTESİ")
        tabs.addTab(self.build_running_tab(), "VARDIYA ONLINE")
        tabs.addTab(self.build_kusbaki_tab(), "KUŞBAKIŞI")
        tabs.addTab(self.build_usta_tab(), "USTA DEFTERİ")
        tabs.addTab(self.build_team_flow_tab(), "TAKIM PLANLAMA (AKIŞ)")
        tabs.addTab(BuzulmeMetreUyumTab(self), "BÜZÜLME & METRE UYUM")

        # YENİ: ITEMA AYAR FORMU sekmesi
        tabs.addTab(ItemaAyarTab(self), "ITEMA AYAR FORMU")

        if self.has_permission("admin"):
            tabs.addTab(UserManagementWidget(self), "KULLANICI YÖNETİMİ")
        self.setCentralWidget(tabs)

        # --- GÖRSEL DÜZELTME BAYRAKLARI ---
        self._did_dugum_first_fix = False
        self._did_run_first_fix = False

        # --- GÜNCELLİK İÇİN OTURUM BAYRAKLARI (yalnızca butondan yüklendiğinde True) ---
        self._did_click_load_dinamik = False
        self._did_click_load_running = False
        self._did_planlama = False

        # Açılışta son snapshot'ları geri yükle (BUTON BAYRAKLARINI ETKİLEMEZ)
        self._restore_last_state()

        # Başlangıçta kullanıcının yetkisine göre butonları ayarla
        self._apply_permissions()

    # -------------------------
    # Yetki kontrol yardımcıları
    # -------------------------
    def has_permission(self, perm: str) -> bool:
        """Aktif kullanıcı için izin kontrolü."""
        return _user_has_permission(self.user, perm)

    def _require_permission(self, perm: str, message: str) -> bool:
        """UI'den çağrılan izin denetimi (mesaj kutusu içerir)."""
        if self.has_permission(perm):
            return True
        QMessageBox.warning(self, "Yetki yok", message)
        return False

    def _apply_permissions(self) -> None:
        """Kullanıcının yetkilerine göre butonları etkin/pasif yap."""
        can_write = self.has_permission("write")
        can_read = self.has_permission("read")

        if hasattr(self, "btn_plan"):
            self.btn_plan.setEnabled(can_write)
        if hasattr(self, "btn_ai_plan"):
            self.btn_ai_plan.setEnabled(can_write)
        if hasattr(self, "btn_notes"):
            self.btn_notes.setEnabled(can_write)
        if hasattr(self, "btn_empty"):
            self.btn_empty.setEnabled(can_write)
        if hasattr(self, "btn_blocked"):
            self.btn_blocked.setEnabled(can_write)
        if hasattr(self, "btn_cut_edit"):
            self.btn_cut_edit.setEnabled(can_write)
        if hasattr(self, "btn_load_dinamik"):
            self.btn_load_dinamik.setEnabled(can_read)
        if hasattr(self, "btn_load_running"):
            self.btn_load_running.setEnabled(can_read)
        if hasattr(self, "team_flow"):
            try:
                self.team_flow.set_write_enabled(can_write)
            except Exception:
                pass
        # diğer tablar...
        for i in range(self.centralWidget().count()):
            w = self.centralWidget().widget(i)
            if hasattr(w, "apply_permissions"):
                try:
                    w.apply_permissions()
                except Exception:
                    pass

    # -------------------------
    # DÜĞÜM SEKME
    # -------------------------
    def build_dugum_tab(self):
        w = QWidget()
        v = QVBoxLayout(w)

        # Üst bar
        top = QHBoxLayout()
        btn_clear = QPushButton("Filtreleri Kaldır")
        btn_clear.clicked.connect(self.clear_all_filters)

        self.btn_load_dinamik = QPushButton("Dinamik Rapor Yükle")
        self.btn_load_dinamik.clicked.connect(self.load_dinamik)

        self.btn_plan = QPushButton("Planlama")
        self.btn_plan.clicked.connect(self.open_planlama)

        # YENİ: Yapay Zeka Planlama butonu
        self.btn_ai_plan = QPushButton("Yapay Zeka Planlama")
        self.btn_ai_plan.setToolTip("Dinamik + Running verisine göre otomatik atama yapar.")
        self.btn_ai_plan.clicked.connect(self.run_ai_planning)

        self.btn_notes = QPushButton("NOTLAR")
        self.btn_notes.clicked.connect(self.open_notes)

        top.addWidget(btn_clear)
        top.addWidget(self.btn_load_dinamik)
        top.addWidget(self.btn_plan)
        top.addWidget(self.btn_ai_plan)  # ← yeni buton
        top.addWidget(self.btn_notes)

        # **YENİ**: Arızalı/Bakımda ve Boş Göster listeleri düğmeleri
        self.btn_empty = QPushButton("Boş Gösterilecek Tezgahlar")
        self.btn_empty.setToolTip("Bu listedeki tezgahlar boş kabul edilir; planlamada görünmez.")
        self.btn_empty.clicked.connect(self._edit_empty_looms)
        top.addWidget(self.btn_empty)

        self.btn_blocked = QPushButton("Arızalı/Bakımda Tezgahlar")
        self.btn_blocked.setToolTip("Bu listedeki tezgahlara iş verilemez; planlamada görünmez.")
        self.btn_blocked.clicked.connect(self._edit_blocked_looms)
        top.addWidget(self.btn_blocked)

        top.addStretch(1)

        # Sağ üst: durum etiketi
        self.lbl_status = QLabel("")
        self.lbl_status.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.lbl_status.setStyleSheet("QLabel{font-weight:700;}")
        top.addWidget(self.lbl_status, 0, Qt.AlignRight)

        # Tablo
        self.tbl = QTableView()
        self.model = PandasModel(pd.DataFrame(), highlight_assigned=True)
        self.proxy = MultiColumnFilterProxy(self)
        self.proxy.setSourceModel(self.model)
        self.tbl.setModel(self.proxy)
        self._style_table(self.tbl)

        # Filtre bar + ScrollArea
        self.dugum_filter_bar = QWidget()
        self.dugum_filter_layout = QGridLayout(self.dugum_filter_bar)
        self.dugum_filter_layout.setContentsMargins(0, 0, 0, 0)
        self.dugum_filter_layout.setHorizontalSpacing(0)
        self.dugum_filter_layout.setVerticalSpacing(0)
        self.dugum_filter_edits = []

        self.dugum_scroll = QScrollArea()
        self.dugum_scroll.setWidgetResizable(True)
        self.dugum_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.dugum_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.dugum_scroll.setFrameShape(QScrollArea.NoFrame)
        self.dugum_scroll.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.dugum_scroll.setWidget(self.dugum_filter_bar)

        self.dugum_filter_bar.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Preferred)
        self.dugum_filter_bar.setMinimumHeight(64)

        v.addLayout(top)
        v.addWidget(self.dugum_scroll)
        v.addWidget(self.tbl, 1)

        # Senkron bağlamalar
        h = self.tbl.horizontalHeader()
        self.dugum_filter_bar.setMinimumWidth(h.length())

        self.tbl.horizontalScrollBar().valueChanged.connect(
            self.dugum_scroll.horizontalScrollBar().setValue
        )

        self._sync_filter_scroll(self.tbl, self.dugum_scroll)
        self._sync_filter_widths(self.tbl, getattr(self, "_dugum_filter_cells", []))

        self.tbl.horizontalScrollBar().valueChanged.connect(
            lambda _=None: self._sync_filter_scroll(self.tbl, self.dugum_scroll)
        )
        self.tbl.horizontalScrollBar().rangeChanged.connect(
            lambda *_: self._sync_filter_scroll(self.tbl, self.dugum_scroll)
        )
        h.sectionResized.connect(lambda *_: (
            self.dugum_filter_bar.setMinimumWidth(h.length()),
            self._sync_filter_widths(self.tbl, getattr(self, "_dugum_filter_cells", [])),
            self._sync_filter_scroll(self.tbl, self.dugum_scroll)
        ))
        h.sectionMoved.connect(lambda *_: (
            self._sync_filter_widths(self.tbl, getattr(self, "_dugum_filter_cells", [])),
            self._sync_filter_scroll(self.tbl, self.dugum_scroll)
        ))

        QTimer.singleShot(0, lambda: self._refit_filter_area(self.dugum_scroll, self.dugum_filter_bar))
        QTimer.singleShot(0, self._refresh_status_label)
        self._attach_header_sync(self.tbl, self.dugum_scroll, "dugum_filter_bar", "_dugum_filter_cells")

        return w

    def _edit_blocked_looms(self):
        if not require_permission(self, "write", "Arızalı/Bakımda listesini düzenleme yetkiniz yok."):
            return
        dlg = LoomListEditor("Arızalı/Bakımda Tezgahlar", "looms/blocked", parent=self)
        if dlg.exec():
            QMessageBox.information(
                self,
                "Bilgi",
                "Arızalı/Bakımda listesi güncellendi.\n"
                "Planlama penceresini yeniden açtığınızda filtre uygulanacaktır."
            )
            # Kuşbakışı/usta gibi görünümler varsa tercihen tazele
            self._refresh_kusbakisi()

    def _edit_empty_looms(self):
        if not require_permission(self, "write", "Boş tezgah listesinde değişiklik yapma yetkiniz yok."):
            return
        dlg = LoomListEditor("Boş Gösterilecek Tezgahlar", "looms/empty", parent=self)
        if dlg.exec():
            QMessageBox.information(
                self,
                "Bilgi",
                "Boş Gösterilecek listesi güncellendi.\n"
                "Planlama penceresini yeniden açtığınızda filtre uygulanacaktır."
            )
            self._refresh_kusbakisi()

    def _rebuild_dugum_filters(self):
        self._dugum_filter_cells = []

        # --- TAM TEMİZLE ---
        while self.dugum_filter_layout.count():
            item = self.dugum_filter_layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.setParent(None)
                w.deleteLater()

        self.dugum_filter_edits.clear()
        try:
            self.proxy.clearFilters()
            self.proxy.clearInclusions()
        except Exception:
            pass

        cols = list(self.model._df.columns) if (self.model is not None and self.model._df is not None) else []

        header = self.tbl.horizontalHeader()

        # Başlık satırı
        for c, name in enumerate(cols):
            lbl = QLabel(str(name))
            lbl.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            lbl.setContentsMargins(0, 0, 0, 0)
            lbl.setFixedWidth(header.sectionSize(c))
            self.dugum_filter_layout.addWidget(lbl, 0, c)

        # Edit + ▼
        for c, _ in enumerate(cols):
            edit = QLineEdit()
            edit.setPlaceholderText("filtre…")
            edit.textChanged.connect(lambda text, col=c: self.proxy.setFilterForColumn(col, text))

            btn = QToolButton()
            btn.setText("▼")
            btn.setToolTip("Çoklu seçim filtresi")
            btn.clicked.connect(lambda _=None, col=c: self._open_value_picker_for_dugum(col))

            cell = QWidget()
            cell.setStyleSheet("QWidget { border-right: 1px solid #e0e0e0; }")
            hl = QHBoxLayout(cell)
            hl.setContentsMargins(0, 0, 0, 0)
            hl.setSpacing(0)
            hl.addWidget(edit, 1)
            hl.addWidget(btn, 0)

            cell.setFixedWidth(header.sectionSize(c))
            cell.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Preferred)

            self.dugum_filter_layout.addWidget(cell, 1, c)
            self.dugum_filter_edits.append(edit)
            self._dugum_filter_cells.append(cell)

        # Geometriyi tazele
        self.dugum_filter_bar.adjustSize()
        self.dugum_filter_bar.updateGeometry()
        self.dugum_filter_bar.setMinimumWidth(header.length())
        QTimer.singleShot(0, lambda: self._refit_filter_area(self.dugum_scroll, self.dugum_filter_bar))
        QTimer.singleShot(0, lambda: self._sync_filter_widths(self.tbl, getattr(self, "_dugum_filter_cells", [])))
        QTimer.singleShot(0, lambda: self._sync_filter_scroll(self.tbl, self.dugum_scroll))

    def load_dinamik(self):
        if not require_permission(self, "read", "Dinamik raporu yüklemek için okuma yetkisi gerekiyor."):
            return
        path, _ = QFileDialog.getOpenFileName(
            self, "Dinamik Rapor Seç", "", "Excel Files (*.xlsx *.xlsb);;All Files (*)"
        )
        if not path:
            return
        try:
            df = load_dinamik_any(path)
            if "Mamul Termin" in df.columns:
                df = df.sort_values(by="Mamul Termin", ascending=True)
            self.df_dinamik_full = df

            # NOTLAR uygula
            self._apply_notes_and_autonotes()

            # Snapshot kaydet
            storage.save_df_snapshot(self.df_dinamik_full, "dinamik")

            self._refresh_dugum_view()
            self._refresh_kusbakisi()

            # Usta Defteri kaynaklarını güncelle
            self._update_usta_sources()

            src = str(df.get("_LeventSource", "")).strip()
            msg = (
                "Dinamik Rapor Yüklendi.\n\n"
                "Şimdi Vardiya Online sekmesindeki “Running Orders” dosyasını yükleyin."
            )
            QMessageBox.information(self, "Bilgi", msg)

            # >>> GÜNCELLİK: butondan yüklendi bayrağı
            self._did_click_load_dinamik = True
            self._update_freshness_if_ready()

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Dinamik rapor yüklenemedi:\n{e}")

        QTimer.singleShot(0, lambda: self._refit_filter_area(self.dugum_scroll, self.dugum_filter_bar))

        if hasattr(self, "team_flow"):
            self.team_flow.refresh_sources()
            self.team_flow.set_write_enabled(self.has_permission("write"))

    def _refresh_dugum_view(
            self,
            group_filter: str | None = None,
            category_filter: str | None = None,
            only_with_levent_digits: bool = False,
            rebuild_filters: bool = True,
            autosize: bool = True,
    ):

        df = self.df_dinamik_full
        if df is None:
            return

        view = df.copy()
        if group_filter:
            view = view[view.get("Tarak Grubu", "").astype(str) == str(group_filter)]
        if category_filter:
            if category_filter == "HAM":
                view = view[view.get("_DyeCategory", "").astype(str).str.contains("HAM", na=False)]
            else:
                view = view[~view.get("_DyeCategory", "").astype(str).str.contains("HAM", na=False)]

        if only_with_levent_digits:
            mask = view.get("_LeventHasDigits", False)
            view = view[mask]

        cols = [c for c in VISIBLE_COLUMNS if c in view.columns]
        ordered = [c for c in VISIBLE_COLUMNS if c in cols]

        view_for_ui = self._with_aliases(view[ordered].copy())

        # --- YENİ: Mamül Termin ve Levent Haşıl Tarihi sütunlarında saatleri kaldır ---
        for col in ["Mamul Termin", "Levent Haşıl Tarihi"]:
            if col in view_for_ui.columns:
                # Önce datetime'e çevir (NaT olanlar korunur)
                ser = pd.to_datetime(view_for_ui[col], errors="coerce")
                # Sonra sadece tarih formatı (dd.mm.yyyy) olarak string yap
                view_for_ui[col] = ser.dt.strftime("%d.%m.%Y")
                # NaT → boş string
                view_for_ui[col] = view_for_ui[col].fillna("")

        self.model.set_df(view_for_ui)

        # 1) autosize
        if autosize:
            QTimer.singleShot(
                0,
                lambda: self._autosize_columns(
                    self.tbl,
                    getattr(self, "_dugum_filter_cells", []),
                    self.dugum_filter_bar,
                    self.dugum_scroll
                )
            )

        # 2) filtre barını yeniden kur
        if rebuild_filters:
            QTimer.singleShot(0, self._rebuild_dugum_filters)
        # 3) ince ayar
        QTimer.singleShot(
            0,
            lambda: self._sync_filter_widths(self.tbl, getattr(self, "_dugum_filter_cells", []))
        )
        QTimer.singleShot(0, lambda: self._sync_filter_scroll(self.tbl, self.dugum_scroll))

    # -------------------------
    # PLANLAMA DİYALOĞU
    # -------------------------
    def open_planlama(self):
        if not require_permission(self, "write", "Planlama ekranını açma yetkiniz yok."):
            return
        # şartlar: Dinamik ve Running yüklü olmalı (veri olarak)
        if self.df_dinamik_full is None:
            QMessageBox.warning(self, "Uyarı", "Önce Dinamik raporu yükleyin.")
            return
        if self.df_running is None or self.df_running.empty:
            QMessageBox.warning(self, "Uyarı", "Önce Vardiya Online (Running Orders) yükleyin.")
            return

        # >>> YENİ: sadece butondan gelen 3 koşulu baz alacağız
        self._did_planlama = True
        self._update_freshness_if_ready()   # (dinamik + running butonları da basıldıysa 'GÜNCEL' olur)

        def on_group_select(group: str, category: str):
            self._refresh_dugum_view(
                group_filter=group,
                category_filter=category,
                only_with_levent_digits=True
            )

        def on_assign(group: str, category: str):
            # Atama sonrası: notları tazele, görünümü tazele ve snapshot kaydet
            self._apply_notes_and_autonotes()
            self._refresh_dugum_view(
                group_filter=group,
                category_filter=category,
                only_with_levent_digits=True,
                rebuild_filters=False
            )
            storage.save_df_snapshot(self.df_dinamik_full, "dinamik")
            # Kuşbakışı tazele
            self._refresh_kusbakisi()

        dlg = PlanningDialog(
            self.df_dinamik_full,
            self.df_running,
            on_group_select=on_group_select,
            on_assign=on_assign,
            parent=self
        )  # on_list_made kaldırıldı

        if dlg.exec():
            # Genel görünüme dön + snapshot
            self._apply_notes_and_autonotes()
            self._refresh_dugum_view()
            storage.save_df_snapshot(self.df_dinamik_full, "dinamik")
            self._refresh_kusbakisi()
    # -------------------------
    # YAPAY ZEKA PLANLAMA (İSKELET)
    # -------------------------
    def run_ai_planning(self):
        """
        Düğüm Takım sekmesindeki 'Yapay Zeka Planlama' butonundan çağrılır.

        - Dinamik + Running yüklü mü kontrol eder
        - PlanningDialog'u arka planda kullanarak
          tüm DENIM + HAM gruplarında AUTO planlama yapar
        - Manuel planlama akışını (Planlama butonu) hiç bozmaz
        """
        if not require_permission(self, "write", "Yapay zeka ile planlama yapmak için yazma yetkiniz yok."):
            return

        # Dinamik kontrolü
        if self.df_dinamik_full is None or self.df_dinamik_full.empty:
            QMessageBox.warning(self, "Uyarı", "Önce Dinamik raporu yükleyin.")
            return

        # Running kontrolü
        if self.df_running is None or self.df_running.empty:
            QMessageBox.warning(self, "Uyarı", "Önce Vardiya Online (Running Orders) dosyasını yükleyin.")
            return

        # Güncellik bayrağı (Planlama ekranıyla aynı mantık)
        self._did_planlama = True
        self._update_freshness_if_ready()

        # Running verisini normalize / zenginleştir (manuel planlamada da kullanıyorsun)
        try:
            self.df_running = normalize_df_running(self.df_running.copy())
        except Exception:
            pass
        try:
            self.df_running = enrich_running_with_loom_cut(self.df_running)
        except Exception:
            pass
        try:
            self.df_running = enrich_running_with_selvedge(self.df_running)
        except Exception:
            pass

        # PlanningDialog'u GÖSTERMEDEN, sadece beyin olarak kullanacağız
        dlg = PlanningDialog(
            self.df_dinamik_full,
            self.df_running,
            on_group_select=None,
            on_assign=None,
            on_list_made=None,
            parent=self,
        )

        total_assigned = dlg.auto_plan_all_groups()

        # Atamalar df_dinamik_full üzerinde yapıldı; şimdi görünümü ve snapshot'ı tazele
        self._apply_notes_and_autonotes()
        self._refresh_dugum_view()
        storage.save_df_snapshot(self.df_dinamik_full, "dinamik")
        self._refresh_kusbakisi()

        QMessageBox.information(
            self,
            "Yapay Zeka Planlama",
            (
                "Otomatik planlama tamamlandı.\n\n"
                f"Atanan iş sayısı: {total_assigned}\n"
                "Kalan işleri istersen Planlama ekranından manuel olarak gözden geçirebilirsin."
            )
        )

    def open_notes(self):
        if not require_permission(self, "write", "Not kurallarında değişiklik yapma yetkiniz yok."):
            return
        if self.df_dinamik_full is None or self.df_dinamik_full.empty:
            QMessageBox.information(self, "Bilgi", "Önce Dinamik raporu yükleyin.")
            return

        # Mevcut kuralları dialoga ver
        dlg = NotesDialog(self.df_dinamik_full, self._note_rules, parent=self)
        if dlg.exec():
            rules = dlg.result_rules()
            if rules is not None:
                # Kullanıcının düzenlediği tam listeyi al
                self._note_rules = rules

                # Kalıcı kaydet (AppMeta.notes_rules)
                storage.save_rules(self._note_rules)

                # Dinamik df üzerine notları yeniden uygula
                self._apply_notes_and_autonotes()

                # Görünümü tazele (filtreleri bozmadan)
                self._refresh_dugum_view(rebuild_filters=False)

                # Snapshot kaydet
                storage.save_df_snapshot(self.df_dinamik_full, "dinamik")

                # Kuşbakışı yenile
                self._refresh_kusbakisi()

    def _append_note(self, old: str, add: str) -> str:
        base = (old or "").strip()
        add = (add or "").strip()
        if not add:
            return base
        if not base:
            return add
        parts = [p.strip() for p in base.split(";") if p.strip()]
        if add in parts:
            return base
        return base + "; " + add

    def _apply_manual_rules(self, df: pd.DataFrame) -> pd.DataFrame:
        if not self._note_rules:
            return df
        if "NOTLAR" not in df.columns:
            df["NOTLAR"] = ""

        for rule in self._note_rules:
            col, val, text = rule.get("col"), rule.get("val"), rule.get("text")
            if not col or col not in df.columns or text is None:
                continue
            mask = (df[col].astype(str) == str(val))
            if mask.any():
                df.loc[mask, "NOTLAR"] = df.loc[mask, "NOTLAR"].astype(str).apply(
                    lambda x, t=text: self._append_note(x, t)
                )
        return df

    def _apply_auto_atki_notes(self, df: pd.DataFrame) -> pd.DataFrame:
        need1_col = "Atkı İhtiyaç Miktar 1"
        need2_col = "Atkı İhtiyaç Miktar 2"
        stock1_col = "(Atkı-1 İşletme Depoları + Atkı-1 İşletme Diğer Depoları)"
        stock2_col = "(Atkı-2 İşletme Depoları + Atkı-2 İşletme Diğer Depoları)"
        key_col = "Üretim Sipariş No"

        for c in [need1_col, need2_col, stock1_col, stock2_col, key_col]:
            if c not in df.columns:
                return df

        if "NOTLAR" not in df.columns:
            df["NOTLAR"] = ""

        dwork = df.copy()
        dwork[need1_col] = pd.to_numeric(dwork[need1_col], errors="coerce").fillna(0.0)
        dwork[need2_col] = pd.to_numeric(dwork[need2_col], errors="coerce").fillna(0.0)
        dwork[stock1_col] = pd.to_numeric(dwork[stock1_col], errors="coerce").fillna(0.0)
        dwork[stock2_col] = pd.to_numeric(dwork[stock2_col], errors="coerce").fillna(0.0)

        grp = dwork.groupby(key_col, dropna=False)
        need1_sum = grp[need1_col].sum()
        need2_sum = grp[need2_col].sum()
        stock1_max = grp[stock1_col].max()
        stock2_max = grp[stock2_col].max()

        lack1 = set(need1_sum[need1_sum > stock1_max].index)
        lack2 = set(need2_sum[need2_sum > stock2_max].index)

        def decide_note(siparis: str) -> str:
            msgs = []
            if siparis in lack1:
                msgs.append("ATKI1 EKSİK")
            if siparis in lack2:
                msgs.append("ATKI2 EKSİK")
            return "; ".join(msgs)

        for siparis in set(lack1 | lack2):
            m = (df[key_col].astype(str) == str(siparis))
            note_text = decide_note(siparis)
            if note_text:
                df.loc[m, "NOTLAR"] = df.loc[m, "NOTLAR"].astype(str).apply(
                    lambda x, t=note_text: self._append_note(x, t)
                )

        return df

    def _clean_label_value(self, val: Any) -> str:
        if val is None:
            return ""
        try:
            if isinstance(val, float) and pd.isna(val):
                return ""
        except Exception:
            pass

        s = str(val).strip()
        if not s:
            return ""
        s = s.replace("\n", " ").replace("\r", " ")
        s = re.sub(r"\.0+$", "", s)
        if s.lower() in {"nan", "nat"}:
            return ""
        return s

    def _running_barkod_tezgah_map(self) -> dict[str, str]:
        df_run = getattr(self, "df_running", None)
        if df_run is None or df_run.empty:
            return {}

        barkod_col = None
        for c in df_run.columns:
            if "BARKOD" in str(c).upper():
                barkod_col = c
                break
        if barkod_col is None:
            return {}

        tez_col = next((c for c in ["Tezgah No", "Tezgah", "Tezgah Numarası"] if c in df_run.columns), None)
        if tez_col is None:
            for c in df_run.columns:
                if "TEZGAH" in str(c).upper():
                    tez_col = c
                    break
        if tez_col is None:
            return {}

        mapping: dict[str, str] = {}
        try:
            subset = df_run[[barkod_col, tez_col]]
        except Exception:
            return {}

        for _, row in subset.iterrows():
            label = self._clean_label_value(row.get(barkod_col))
            loom = self._clean_label_value(row.get(tez_col))
            if label and loom and label not in mapping:
                mapping[label] = loom
        return mapping

    def _apply_etiket_location_notes(self, df: pd.DataFrame) -> pd.DataFrame:
        etiket_col = next((c for c in ["Levent Etiket FA", "EtiketFA", "Etiket No"] if c in df.columns), None)
        if etiket_col is None:
            return df

        if "NOTLAR" not in df.columns:
            df["NOTLAR"] = ""

        try:
            usta_map = storage.load_usta_etiket_tezgah_map()
        except Exception:
            usta_map = {}
        running_map = self._running_barkod_tezgah_map()

        if not usta_map and not running_map:
            return df

        def _find_machine(label: Any) -> str:
            key = self._clean_label_value(label)
            if not key:
                return ""
            if key in usta_map:
                return self._clean_label_value(usta_map.get(key))
            if key in running_map:
                return self._clean_label_value(running_map.get(key))
            return ""

        for idx, label in df[etiket_col].items():
            loom = _find_machine(label)
            if not loom:
                continue
            note_text = f"{loom} NOLU TEZGAHA ALINDI"
            df.at[idx, "NOTLAR"] = self._append_note(df.at[idx, "NOTLAR"], note_text)

        return df

    def _apply_notes_and_autonotes(self):
        """
        NOTLAR sütununu her seferinde TEMİZDEN hesapla:

        - İlk çalıştığında mevcut NOTLAR değerini _NOTLAR_BASE kolonuna kopyalar.
        - Sonraki her çağrıda NOTLAR'ı _NOTLAR_BASE'den geri yükler.
        - Üzerine otomatik ATKI notlarını ve manuel kural notlarını uygular.
        """
        if self.df_dinamik_full is None or self.df_dinamik_full.empty:
            return

        df = self.df_dinamik_full

        # NOTLAR yoksa oluştur
        if "NOTLAR" not in df.columns:
            df["NOTLAR"] = ""

        # Orijinal NOTLAR'ı bir kere yedekle
        if "_NOTLAR_BASE" not in df.columns:
            df["_NOTLAR_BASE"] = df["NOTLAR"].astype(str)
        else:
            # Emin olmak için string’e çevir
            df["_NOTLAR_BASE"] = df["_NOTLAR_BASE"].astype(str)

        # Her seferinde temiz bir başlangıç:
        df["NOTLAR"] = df["_NOTLAR_BASE"]

        # 1) Otomatik ATKI eksikliği notları
        df = self._apply_auto_atki_notes(df)

        # 2) Etiket -> Tezgah bilgisi (Usta Defteri + Running)
        df = self._apply_etiket_location_notes(df)

        # 3) Senin tanımladığın manuel kurallar
        df = self._apply_manual_rules(df)

        self.df_dinamik_full = df

    # -------------------------
    # RUNNING ORDERS SEKME
    # -------------------------
    def build_running_tab(self):
        w = QWidget()
        v = QVBoxLayout(w)

        # Üst bar
        top = QHBoxLayout()
        btn_clear_run = QPushButton("Filtreleri Kaldır")
        btn_clear_run.clicked.connect(self.clear_all_filters)
        self.btn_load_running = QPushButton("Running Orders Yükle")
        self.btn_load_running.clicked.connect(self.load_running)

        top.addWidget(btn_clear_run)
        top.addWidget(self.btn_load_running)
        top.addStretch(1)

        # >>> YENİ: Kesim Tipi editörü butonu
        self.btn_cut_edit = QPushButton("Kesim Tipi (ISAVER/ROTOCUT)…")
        self.btn_cut_edit.clicked.connect(self._open_loom_cut_editor)
        top.addWidget(self.btn_cut_edit)

        top.addStretch(1)

        # Tablo
        self.tbl_run = QTableView()
        self.model_run = PandasModel(pd.DataFrame())
        self.proxy_run = MultiColumnFilterProxy(self)
        self.proxy_run.setSourceModel(self.model_run)
        self.tbl_run.setModel(self.proxy_run)
        self._style_table(self.tbl_run)

        # Filtre bar + ScrollArea
        self.run_filter_bar = QWidget()
        self.run_filter_layout = QGridLayout(self.run_filter_bar)
        self.run_filter_layout.setContentsMargins(0, 0, 0, 0)
        self.run_filter_layout.setHorizontalSpacing(0)
        self.run_filter_layout.setVerticalSpacing(0)
        self.run_filter_edits = []
        self._run_filter_cells = []

        self.run_scroll = QScrollArea()
        self.run_scroll.setWidgetResizable(True)
        self.run_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.run_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.run_scroll.setFrameShape(QScrollArea.NoFrame)
        self.run_scroll.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.run_scroll.setWidget(self.run_filter_bar)

        self.run_filter_bar.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Preferred)
        self.run_filter_bar.setMinimumHeight(64)

        v.addLayout(top)
        v.addWidget(self.run_scroll)
        v.addWidget(self.tbl_run, 1)

        h2 = self.tbl_run.horizontalHeader()
        self.run_filter_bar.setMinimumWidth(h2.length())

        self.tbl_run.horizontalScrollBar().valueChanged.connect(
            self.run_scroll.horizontalScrollBar().setValue
        )

        self._sync_filter_scroll(self.tbl_run, self.run_scroll)
        self._sync_filter_widths(self.tbl_run, getattr(self, "_run_filter_cells", []))

        self.tbl_run.horizontalScrollBar().valueChanged.connect(
            lambda _=None: self._sync_filter_scroll(self.tbl_run, self.run_scroll)
        )
        self.tbl_run.horizontalScrollBar().rangeChanged.connect(
            lambda *_: self._sync_filter_scroll(self.tbl_run, self.run_scroll)
        )

        h2.sectionResized.connect(lambda *_: (
            self.run_filter_bar.setMinimumWidth(h2.length()),
            self._sync_filter_widths(self.tbl_run, getattr(self, "_run_filter_cells", [])),
            self._sync_filter_scroll(self.tbl_run, self.run_scroll)
        ))
        h2.sectionMoved.connect(lambda *_: (
            self._sync_filter_widths(self.tbl_run, getattr(self, "_run_filter_cells", [])),
            self._sync_filter_scroll(self.tbl_run, self.run_scroll)
        ))

        QTimer.singleShot(0, lambda: self._refit_filter_area(self.run_scroll, self.run_filter_bar))
        self._attach_header_sync(self.tbl_run, self.run_scroll, "run_filter_bar", "_run_filter_cells")

        return w

    def _open_loom_cut_editor(self):
        if not require_permission(self, "write", "Kesim tipi düzenleyicisini açma yetkiniz yok."):
            return
        try:
            dlg = LoomCutEditor(self)
            dlg.exec()
            # Kayıttan sonra tabloyu tazele (varsa)
            if self.df_running is not None and not self.df_running.empty:
                df = self.df_running.copy()
                df = enrich_running_with_loom_cut(df)
                df = enrich_running_with_selvedge(df, getattr(self, "df_dinamik_full", None))
                self.df_running = df
                self.model_run.set_df(df.copy())
                self._rebuild_run_filters()
                QTimer.singleShot(
                    0,
                    lambda: self._autosize_columns(
                        self.tbl_run,
                        getattr(self, "_run_filter_cells", []),
                        self.run_filter_bar,
                        self.run_scroll
                    )
                )
        except Exception as e:
            import traceback
            QMessageBox.critical(
                self,
                "Hata",
                f"Kesim Tipi editörü açılırken hata oluştu:\n\n{traceback.format_exc()}"
            )

    def _rebuild_run_filters(self):
        self._run_filter_cells = []

        while self.run_filter_layout.count():
            item = self.run_filter_layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.setParent(None)
                w.deleteLater()

        self.run_filter_edits.clear()
        try:
            self.proxy_run.clearFilters()
            self.proxy_run.clearInclusions()
        except Exception:
            pass

        cols = list(self.model_run._df.columns) if (
            self.model_run is not None and self.model_run._df is not None
        ) else []
        header = self.tbl_run.horizontalHeader()

        # Başlık satırı
        for c, name in enumerate(cols):
            lbl = QLabel(str(name))
            lbl.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            lbl.setContentsMargins(0, 0, 0, 0)
            lbl.setFixedWidth(header.sectionSize(c))
            self.run_filter_layout.addWidget(lbl, 0, c)

        # Edit + ▼
        for c, _ in enumerate(cols):
            edit = QLineEdit()
            edit.setPlaceholderText("filtre…")
            edit.textChanged.connect(lambda text, col=c: self.proxy_run.setFilterForColumn(col, text))

            btn = QToolButton()
            btn.setText("▼")
            btn.setToolTip("Çoklu seçim filtresi")
            btn.clicked.connect(lambda _=None, col=c: self._open_value_picker_for_run(col))

            cell = QWidget()
            cell.setStyleSheet("QWidget { border-right: 1px solid #e0e0e0; }")
            hl = QHBoxLayout(cell)
            hl.setContentsMargins(0, 0, 0, 0)
            hl.setSpacing(0)
            hl.addWidget(edit, 1)
            hl.addWidget(btn, 0)

            cell.setFixedWidth(header.sectionSize(c))
            cell.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Preferred)

            self.run_filter_layout.addWidget(cell, 1, c)
            self.run_filter_edits.append(edit)
            self._run_filter_cells.append(cell)

        # Geometriyi tazele & senkron
        self.run_filter_bar.adjustSize()
        self.run_filter_bar.updateGeometry()
        self.run_filter_bar.setMinimumWidth(header.length())
        QTimer.singleShot(0, lambda: self._refit_filter_area(self.run_scroll, self.run_filter_bar))
        QTimer.singleShot(
            0, lambda: self._sync_filter_widths(self.tbl_run, getattr(self, "_run_filter_cells", []))
        )
        QTimer.singleShot(0, lambda: self._sync_filter_scroll(self.tbl_run, self.run_scroll))

    def load_running(self):
        if not require_permission(self, "read", "Running Orders dosyasını yüklemek için okuma yetkisi gerekiyor."):
            return
        path, _ = QFileDialog.getOpenFileName(
            self, "Running Orders Seç", "", "Excel Files (*.xlsx);;All Files (*)"
        )
        if not path:
            return
        try:
            df = load_running_orders(path)
            # *** TEK NOKTADAN DÜZELTME ***
            df = normalize_df_running(df)
            df = enrich_running_with_loom_cut(df)
            df = enrich_running_with_selvedge(df, getattr(self, "df_dinamik_full", None))

            # Running: Tip No'yu KökTip formatına çevir (R önekiyle)
            tip_col = next((c for c in ["Tip No", "Tip Kodu", "Tip", "Mamul Tipi"] if c in df.columns), None)
            if tip_col:
                df["KökTip"] = df[tip_col].astype(str).apply(
                    lambda x: x if (x.strip() == "" or x.strip().upper().startswith("R")) else f"R{x.strip()}"
                )

            # (C) Süs Kenar ve ISAVER/ROTOCUT sütunlarını besle
            df = enrich_running_with_loom_cut(df)
            df = enrich_running_with_selvedge(df, getattr(self, "df_dinamik_full", None))

            if "Tezgah No" in df.columns:
                df = df.sort_values(by="Tezgah No", ascending=True)

            self.df_running = df
            self.model_run.set_df(df.copy())
            self._rebuild_run_filters()

            # Snapshot kaydet
            storage.save_df_snapshot(self.df_running, "running")

            # Kuşbakışı tazele
            self._refresh_kusbakisi()

            # Usta Defteri tezgah listesini güncelle
            self._update_usta_sources()

            QTimer.singleShot(
                0,
                lambda: self._autosize_columns(
                    self.tbl_run,
                    getattr(self, "_run_filter_cells", []),
                    self.run_filter_bar,
                    self.run_scroll
                )
            )
            QTimer.singleShot(0, self._rebuild_run_filters)
            QTimer.singleShot(
                0,
                lambda: self._sync_filter_widths(self.tbl_run, getattr(self, "_run_filter_cells", []))
            )
            QTimer.singleShot(0, lambda: self._sync_filter_scroll(self.tbl_run, self.run_scroll))
            QTimer.singleShot(0, lambda: self._refit_filter_area(self.run_scroll, self.run_filter_bar))

            # >>> GÜNCELLİK: butondan yüklendi bayrağı
            self._did_click_load_running = True
            QMessageBox.information(
                self,
                "Running Hazır",
                "Running Orders Yüklendi.\n\n"
                "Şimdi “DÜĞÜM TAKIM LİSTESİ” sekmesinde Planlama yapabilirsiniz."
            )

            self._update_freshness_if_ready()

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Running orders yüklenemedi:\n{e}")
        if hasattr(self, "team_flow"):
            self.team_flow.refresh_sources()
            self.team_flow.set_write_enabled(self.has_permission("write"))

    # -------------------------
    # KUŞBAKIŞI SEKME
    # -------------------------
    def build_kusbaki_tab(self):
        self.kusbakisi = KusbakisiWidget(self)
        try:
            self.kusbakisi.set_status_label(self.lbl_status.text(), self.lbl_status.styleSheet())
        except Exception:
            pass
        return self.kusbakisi

    def build_team_flow_tab(self):
        self.team_flow = TeamPlanningFlowTab(self)
        return self.team_flow

    def _refresh_kusbakisi(self):
        if hasattr(self, "kusbakisi") and self.kusbakisi is not None:
            self.kusbakisi.refresh(self.df_dinamik_full, self.df_running)

    # -------------------------
    # USTA DEFTERİ SEKME (ENTEGRE)
    # -------------------------
    def build_usta_tab(self):
        self.usta = UstaDefteriWidget(self)
        QTimer.singleShot(0, self._update_usta_sources)
        return self.usta

    # -------------------------
    # ORTAK YARDIMCI FONKSİYONLAR
    # -------------------------
    def _style_table(self, table: QTableView):
        h = table.horizontalHeader()
        h.setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        h.setStretchLastSection(False)
        h.setSectionResizeMode(QHeaderView.Interactive)
        h.setStyleSheet("""
            QHeaderView::section {
                background: #f2f2f2;
                color: #222;
                font-weight: 600;
                font-size: 12pt;
                padding: 6px 8px;
                border: 0px;
                border-right: 1px solid #e0e0e0;
            }
            QHeaderView::section:horizontal {
                border-top: 1px solid #e0e0e0;
                border-bottom: 1px solid #dcdcdc;
            }
        """)

        table.verticalHeader().setVisible(False)
        table.verticalHeader().setDefaultSectionSize(26)

        table.setAlternatingRowColors(True)
        table.setStyleSheet("""
            QTableView {
                gridline-color: #eaeaea;
                alternate-background-color: #fafafa;
                selection-background-color: #cfe8ff;
                selection-color: #000;
            }
        """)
        table.setSelectionBehavior(QTableView.SelectRows)
        table.setSelectionMode(QTableView.SingleSelection)

        if table.model():
            table.resizeColumnsToContents()
            h = table.horizontalHeader()
            for c in range(table.model().columnCount()):
                if h.sectionSize(c) < 140:
                    h.resizeSection(c, 140)

    def _autosize_columns(
        self,
        table,
        filter_cells=None,
        filter_bar=None,
        scroll=None,
    ):
        """
        - Tüm sütunları Qt'nin ResizeToContents mantığıyla ayarlar.
        - Hiçbir sütun için ekstra sınırlama yok (NOTLAR da dahil).
        """
        model = table.model()
        header = table.horizontalHeader()
        if model is None or header is None:
            return

        # 1) Qt'ye tüm sütunları içeriklerine göre auto-size yaptır
        try:
            table.resizeColumnsToContents()
        except Exception:
            pass

        # 2) Filtre hücreleri ve bar ile genişliği senkron tut
        try:
            if filter_cells:
                for c, cell in enumerate(filter_cells):
                    if c < header.count():
                        cell.setFixedWidth(header.sectionSize(c))
            if filter_bar is not None:
                filter_bar.setMinimumWidth(header.length())
            if scroll is not None:
                self._sync_filter_scroll(table, scroll)
        except Exception:
            pass

    def _sync_filter_widths(self, table: QTableView, cells: list):
        if not cells:
            return
        h = table.horizontalHeader()
        for c, cell in enumerate(cells):
            if c < h.count():
                cell.setFixedWidth(h.sectionSize(c))

    def _sync_filter_scroll(self, table: QTableView, scroll: QScrollArea):
        h = table.horizontalHeader()
        fb = scroll.widget()
        if fb is not None:
            if table is getattr(self, "tbl_run", None):
                bar = getattr(self, "run_filter_bar", None)
            else:
                bar = getattr(self, "dugum_filter_bar", None)
            if bar is not None:
                bar.setMinimumWidth(h.length())
        scroll.horizontalScrollBar().setValue(h.offset())

    def _open_value_picker_for_dugum(self, col: int):
        if self.model is None or self.model._df is None or self.model._df.empty:
            return
        colname = self.model._df.columns[col]
        values = self.model._df[colname].astype(str).fillna("").tolist()

        pre = self.proxy._inclusions.get(col, set())
        dlg = ValuePickerDialog(f"Filtre • {colname}", values, preselected=pre, parent=self)
        if dlg.exec():
            selected = dlg.selected_values()
            self.proxy.setInclusionForColumn(col, selected)

    def _open_value_picker_for_run(self, col: int):
        if self.model_run is None or self.model_run._df is None or self.model_run._df.empty:
            return
        colname = self.model_run._df.columns[col]
        values = self.model_run._df[colname].astype(str).fillna("").tolist()

        pre = self.proxy_run._inclusions.get(col, set())
        dlg = ValuePickerDialog(f"Filtre • {colname}", values, preselected=pre, parent=self)
        if dlg.exec():
            selected = dlg.selected_values()
            self.proxy_run.setInclusionForColumn(col, selected)

    def clear_all_filters(self):
        # Veri kalır; sadece filtreler temizlenir
        try:
            self.proxy.clearFilters()
            self.proxy.clearInclusions()
        except Exception:
            pass
        for ed in getattr(self, "dugum_filter_edits", []):
            ed.blockSignals(True)
            ed.setText("")
            ed.blockSignals(False)

        try:
            self.proxy_run.clearFilters()
            self.proxy_run.clearInclusions()
            for ed in getattr(self, "run_filter_edits", []):
                ed.blockSignals(True)
                ed.setText("")
                ed.blockSignals(False)
        except Exception:
            pass

        # görünümü mevcut df ile tekrar kur (sütun genişliklerini bozma)
        self._refresh_dugum_view(rebuild_filters=False, autosize=False)

        if self.df_running is not None:
            self.model_run.set_df(self.df_running.copy())
            self._rebuild_run_filters()

        # kuşbakışını da dokunmadan yenileyelim (filtre paneli ayrı)
        self._refresh_kusbakisi()

    def _refit_filter_area(self, scroll: QScrollArea, bar: QWidget, min_h: int = 72, max_h: int = 140):
        if not bar or not scroll:
            return
        h = bar.sizeHint().height()
        h = max(min_h, min(h + 8, max_h))
        scroll.setFixedHeight(h)

    def resizeEvent(self, e):
        super().resizeEvent(e)
        try:
            self._refit_filter_area(self.dugum_scroll, self.dugum_filter_bar)
        except Exception:
            pass
        try:
            self._refit_filter_area(self.run_scroll, self.run_filter_bar)
        except Exception:
            pass

    def _with_aliases(self, df: pd.DataFrame) -> pd.DataFrame:
        return df.rename(columns=HEADER_ALIASES)

    # -------------------------
    # GÜNCELLİK / DURUM ETİKETİ (İstanbul TZ + tarih/saat)
    # -------------------------
    def _mark_planning_checkpoint(self):
        """(ARTIK KULLANMIYORUZ) — Güncellik mantığını _update_freshness_if_ready yönetiyor."""
        pass

    def _current_shift_bounds(self, ref: datetime | None = None) -> tuple[datetime, datetime]:
        now = ref or datetime.now(self.TZ)
        d = now.date()

        s1 = datetime(d.year, d.month, d.day, 7, 0, tzinfo=self.TZ)
        s2 = datetime(d.year, d.month, d.day, 15, 0, tzinfo=self.TZ)
        s3 = datetime(d.year, d.month, d.day, 23, 0, tzinfo=self.TZ)

        if s1 <= now < s2:
            return (s1, s2)
        elif s2 <= now < s3:
            return (s2, s3)
        else:
            if now >= s3:
                e3 = datetime(d.year, d.month, d.day, 7, 0, tzinfo=self.TZ) + timedelta(days=1)
                return (s3, e3)
            else:
                prev = now - timedelta(days=1)
                s_prev = datetime(prev.year, prev.month, prev.day, 23, 0, tzinfo=self.TZ)
                e_prev = datetime(d.year, d.month, d.day, 7, 0, tzinfo=self.TZ)
                return (s_prev, e_prev)

    def _is_fresh(self, last: datetime | None) -> bool:
        if last is None:
            return False
        if last.tzinfo is None:
            last = last.replace(tzinfo=self.TZ)

        now = datetime.now(self.TZ)
        start, _ = self._current_shift_bounds(now)
        window_start = start - timedelta(hours=1)
        return (last >= window_start) and (last <= now)

    def _refresh_status_label(self):
        dt = self._last_update
        if dt is None:
            self.lbl_status.setText("GÜNCEL DEĞİL — (henüz yükleme/planlama yok)")
            self.lbl_status.setStyleSheet("QLabel{font-weight:700;color:#c62828;}")
            return

        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=self.TZ)

        fresh = self._is_fresh(dt)
        stamp = dt.astimezone(self.TZ).strftime("%d.%m.%Y %H:%M")
        if fresh:
            self.lbl_status.setText(f"GÜNCEL — {stamp}")
            self.lbl_status.setStyleSheet("QLabel{font-weight:700;color:#1e8e3e;}")
        else:
            self.lbl_status.setText(f"GÜNCEL DEĞİL — {stamp}")
            self.lbl_status.setStyleSheet("QLabel{font-weight:700;color:#c62828;}")

        try:
            if hasattr(self, "kusbakisi") and self.kusbakisi is not None:
                self.kusbakisi.set_status_label(self.lbl_status.text(), self.lbl_status.styleSheet())
        except Exception:
            pass

    # -------------------------
    # AÇILIŞTA SON HALİ GERİ YÜKLE
    # -------------------------
    def _restore_last_state(self):
        """Uygulama açıldığında snapshot'lardan DF'leri yükle, görünümü kur, filtreleri boş başlat."""
        try:
            ddf = storage.load_df_snapshot("dinamik")
            if ddf is not None and not ddf.empty:
                self.df_dinamik_full = ddf
                self._apply_notes_and_autonotes()
                self._refresh_dugum_view(rebuild_filters=True)

            rdf = storage.load_df_snapshot("running")
            if rdf is not None and not rdf.empty:
                # *** TEK NOKTADAN DÜZELTME (snapshot için de uygula) ***
                rdf = normalize_df_running(rdf)
                tip_col = next((c for c in ["Tip No", "Tip Kodu", "Tip", "Mamul Tipi"] if c in rdf.columns), None)
                if tip_col:
                    rdf["KökTip"] = rdf[tip_col].astype(str).apply(
                        lambda x: x
                        if (x.strip() == "" or x.strip().upper().startswith("R"))
                        else f"R{x.strip()}"
                    )

                rdf = enrich_running_with_loom_cut(rdf)
                rdf = enrich_running_with_selvedge(rdf, getattr(self, "df_dinamik_full", None))

                self.df_running = rdf
                self.model_run.set_df(self.df_running.copy())
                self._rebuild_run_filters()
                QTimer.singleShot(
                    0,
                    lambda: self._autosize_columns(
                        self.tbl_run,
                        getattr(self, "_run_filter_cells", []),
                        self.run_filter_bar,
                        self.run_scroll
                    )
                )
                QTimer.singleShot(0, self._rebuild_run_filters)
                QTimer.singleShot(
                    0,
                    lambda: self._sync_filter_widths(self.tbl_run, getattr(self, "_run_filter_cells", []))
                )
                QTimer.singleShot(0, lambda: self._sync_filter_scroll(self.tbl_run, self.run_scroll))
                QTimer.singleShot(0, lambda: self._refit_filter_area(self.run_scroll, self.run_filter_bar))

            # Usta Defteri kaynakları (snapshot sonrası)
            self._update_usta_sources()

            # Filtreleri temiz (boş) başlat
            self.clear_all_filters()
        except Exception:
            pass

        # Kuşbakışı ekranını da ilk açılışta hazırla
        self._refresh_kusbakisi()
        self._refresh_status_label()
        QTimer.singleShot(0, self._do_first_dugum_layout_fix)
        if hasattr(self, "team_flow"):
            self.team_flow.refresh_sources()

    # -------------------------
    # USTA DEFTERİ ENTEGRASYONU — Yardımcılar
    # -------------------------
    def _update_usta_sources(self):
        try:
            if not hasattr(self, "usta") or self.usta is None:
                return
            if getattr(self, "df_dinamik_full", None) is not None:
                self.usta.set_sources(self.df_dinamik_full)
            else:
                self.usta.set_sources(None)
            looms = []
            if getattr(self, "df_dinamik_full", None) is not None and not self.df_dinamik_full.empty:
                looms = self._extract_looms(self.df_dinamik_full)
            if (not looms) and (getattr(self, "df_running", None) is not None) and (not self.df_running.empty):
                looms = self._extract_looms(self.df_running)
            self.usta.set_machine_list(looms or [])
        except Exception:
            pass

    def _extract_looms(self, df: pd.DataFrame) -> list[str]:
        candidates = ["Tezgah Numarası", "Tezgah No", "Tezgah"]
        for c in candidates:
            if c in df.columns:
                ser = df[c].astype(str).str.extract(r'(\d+)')[0].dropna().unique().tolist()
                ser = [s for s in ser if s.isdigit()]
                return sorted(set(ser), key=lambda x: int(x))
        return []

    def showEvent(self, e):
        super().showEvent(e)
        if not getattr(self, "_did_dugum_first_fix", False):
            self._did_dugum_first_fix = True
            QTimer.singleShot(0, self._do_first_dugum_layout_fix)
        if not getattr(self, "_did_run_first_fix", False):
            self._did_run_first_fix = True
            QTimer.singleShot(0, self._do_first_run_layout_fix)

    def _do_first_dugum_layout_fix(self):
        try:
            self._autosize_columns(
                self.tbl,
                getattr(self, "_dugum_filter_cells", []),
                self.dugum_filter_bar,
                self.dugum_scroll
            )
            self._sync_filter_widths(self.tbl, getattr(self, "_dugum_filter_cells", []))
            self._sync_filter_scroll(self.tbl, self.dugum_scroll)
            self._refit_filter_area(self.dugum_scroll, self.dugum_filter_bar)
        except Exception:
            pass

    def _attach_header_sync(self, table: QTableView, scroll: QScrollArea, bar_attr: str, cells_attr: str):
        h = table.horizontalHeader()

        def resync():
            try:
                bar = getattr(self, bar_attr, None)
                cells = getattr(self, cells_attr, [])
                if bar is not None:
                    bar.setMinimumWidth(h.length())
                if cells:
                    self._sync_filter_widths(table, cells)
                self._sync_filter_scroll(table, scroll)
            except Exception:
                pass

        h.geometriesChanged.connect(resync)
        h.sectionResized.connect(lambda *_: resync())
        h.sectionMoved.connect(lambda *_: resync())

        m = table.model()
        if m is not None:
            try:
                m.layoutChanged.connect(resync)
                m.modelReset.connect(resync)
            except Exception:
                pass
            try:
                src = getattr(m, "sourceModel", lambda: None)()
                if src is not None:
                    src.layoutChanged.connect(resync)
                    src.modelReset.connect(resync)
            except Exception:
                pass

        table.horizontalScrollBar().valueChanged.connect(lambda _=None: resync())
        QTimer.singleShot(0, resync)

    def _do_first_run_layout_fix(self):
        try:
            self._autosize_columns(
                self.tbl_run,
                getattr(self, "_run_filter_cells", []),
                self.run_filter_bar,
                self.run_scroll
            )
            self._sync_filter_widths(self.tbl_run, getattr(self, "_run_filter_cells", []))
            self._sync_filter_scroll(self.tbl_run, self.run_scroll)
            self._refit_filter_area(self.run_scroll, self.run_filter_bar)
        except Exception:
            pass

    # -------------------------
    # GÜNCELLİK KONTROLÜ — TEK KAPI
    # -------------------------
    def _update_freshness_if_ready(self):
        """
        GÜNCEL saymak için 3 koşulun tamamı BUTONLA tetiklenmiş olmalı:
          1) Dinamik Rapor Yükle butonuyla dinamik yüklendi
          2) Running Orders Yükle butonuyla running yüklendi
          3) Planlama butonuna basıldı
        """
        if self._did_click_load_dinamik and self._did_click_load_running and self._did_planlama:
            self._last_update = datetime.now(self.TZ)
            storage.save_last_update(self._last_update)
        self._refresh_status_label()
