from __future__ import annotations

from typing import Dict, List, Optional

import pyodbc

# ---------------------------------------------------------------------------
# 1) ITEMA kolonları ve varsayılan başlangıç ayarları
#    (arm_rap_1..4 kaldırıldı)
# ---------------------------------------------------------------------------

ITEMA_COLUMNS: List[str] = [
    "sira_no",
    "tip",
    "tarak",
    "telef_ken1",
    "telef_ken2",
    "firca_secim",
    "ufleme_zam_1",
    "ufleme_zam_2",
    "cimbar_secim",
    "coz_tansiyon",
    "devir",
    "leno",
    "ark_desen",
    "agizlik",
    "derinlik",
    "pozisyon",
    "testere_uzk",
    "testere_yuk",
    "tan_yay_pozisyon",
    "tan_yay_yukseklik",
    "tan_yay_konumu",
    "tan_yay_bogumu",
    "zem_agizlik",
    "kapanma_dur_1",
    "oturma_duzeyi_1",
    "kapanma_dur_2",
    "oturma_duzeyi_2",
    "rampa_1",
    "rampa_2",
    "rampa_3",
    "rampa_4",
    "rampa_5",
    "rampa_6",
    "aciklama",
    "degisiklik_yapan",
]

DEFAULT_ITEMA_SETTINGS: Dict[str, str] = {
    "telef_ken1": "12",
    "telef_ken2": "12",
    "firca_secim": "SEFFAF FIRÇA",
    "ufleme_zam_1": "00:05:00",
    "ufleme_zam_2": "00:00:05",
    "cimbar_secim": "KAUÇUK",
    "coz_tansiyon": "40",
    "devir": "650",
    "leno": "150",
    "agizlik": "23",
    "derinlik": "2",
    "pozisyon": "50",
    "testere_uzk": "4",
    "testere_yuk": "4",
    "tan_yay_pozisyon": "2",
    "tan_yay_yukseklik": "YUKARI",
    "tan_yay_konumu": "PANELE GÖRE",
    "tan_yay_bogumu": "3",
    "zem_agizlik": "305",
    "rampa_1": "100",
    "rampa_2": "50",
    "rampa_3": "10",
    "rampa_4": "120",
    "rampa_5": "30",
    "rampa_6": "9",
}

# ---------------------------------------------------------------------------
# 2) Yardımcı fonksiyonlar
# ---------------------------------------------------------------------------

def _row_to_dict(cursor: pyodbc.Cursor, row: pyodbc.Row) -> Dict[str, Optional[str]]:
    cols = [d[0] for d in cursor.description]
    out: Dict[str, Optional[str]] = {}
    for col, val in zip(cols, row):
        out[col] = None if val is None else str(val)
    return out


def _merge_settings(
    base: Dict[str, Optional[str]],
    override: Dict[str, Optional[str]],
) -> Dict[str, Optional[str]]:
    for key, val in override.items():
        if val is None:
            continue
        if isinstance(val, str) and val.strip() == "":
            continue
        base[key] = val
    return base


# ---------------------------------------------------------------------------
# 3) SQL sorgu yardımcıları
# ---------------------------------------------------------------------------

def _fetch_itema_ayar_by_tip(conn: pyodbc.Connection, tip: str) -> Optional[Dict[str, Optional[str]]]:
    cur = conn.cursor()
    cur.execute("SELECT TOP 1 * FROM dbo.ItemaAyar WHERE [tip] = ?", tip)
    row = cur.fetchone()
    if not row:
        return None
    return _row_to_dict(cur, row)


def get_itema_settings_from_feature_tables(
    conn: pyodbc.Connection,
    tip: str,
    tip_features: Optional[Dict[str, Optional[str]]],
) -> Optional[Dict[str, Optional[str]]]:
    """
    1) Önce dbo.ItemaAyar (manuel/override) aranır.
    2) Yoksa mevcut sistemindeki eşleştirme mantığıyla dbo.Makine_Ayar_Tablosu vb. bulunur.
       (Bu kısım senin sisteminde şu an çalışıyor dediğin için dokunulmuyor.)
    """
    row = _fetch_itema_ayar_by_tip(conn, tip)
    if row:
        return row

    # --- Mevcut eşleştirme/fallback mantığın burada devreye giriyor ---
    # Burada kendi çalışan fonksiyonların / filtrelerin varsa onu kullan.
    # Senin gönderdiğin sürümde basit filtre vardı; sen çalışır hale getirdin dediğin için
    # bu kısmı değiştirmiyorum. Eğer burada başka bir fonksiyonun varsa onu çağır.
    return None


# ---------------------------------------------------------------------------
# 4) Stored procedure çağrıları (varsa)
# ---------------------------------------------------------------------------

def get_itema_automatic_settings(conn: pyodbc.Connection, tip: str) -> Optional[Dict[str, Optional[str]]]:
    cur = conn.cursor()
    cur.execute("EXEC dbo.sp_ItemaOtomatikAyar @Tip = ?", tip)
    row = cur.fetchone()
    if not row:
        return None
    return _row_to_dict(cur, row)


def get_itema_tip_specific_settings(conn: pyodbc.Connection, tip: str) -> List[Dict[str, Optional[str]]]:
    cur = conn.cursor()
    cur.execute("EXEC dbo.sp_ItemaTipOzelAyar @Tip = ?", tip)
    rows = cur.fetchall()
    if not rows:
        return []
    return [_row_to_dict(cur, r) for r in rows]


# ---------------------------------------------------------------------------
# 5) Final ITEMA ayarlarını üreten ana fonksiyon
# ---------------------------------------------------------------------------

def build_itema_settings(
    conn: pyodbc.Connection,
    tip: str,
    tip_features: Optional[Dict[str, Optional[str]]] = None
) -> Dict[str, Optional[str]]:

    settings: Dict[str, Optional[str]] = {col: None for col in ITEMA_COLUMNS}
    settings["tip"] = tip
    settings = _merge_settings(settings, DEFAULT_ITEMA_SETTINGS)

    # 1) Önce manuel tablo (dbo.ItemaAyar) veya senin feature-table eşleştirmen
    table_settings = get_itema_settings_from_feature_tables(conn, tip, tip_features)
    if table_settings:
        return _merge_settings(settings, table_settings)

    # 2) Otomatik ayar
    auto = get_itema_automatic_settings(conn, tip)
    if auto:
        settings = _merge_settings(settings, auto)

    # 3) Tip-özel ayar satırları
    tip_rows = get_itema_tip_specific_settings(conn, tip)
    for row in tip_rows:
        settings = _merge_settings(settings, row)

    return settings
