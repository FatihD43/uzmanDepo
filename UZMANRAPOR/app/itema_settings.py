from __future__ import annotations

from typing import Dict, List, Optional

import pyodbc


"""
ITEMA otomatik ayar servisi.

Ne yapar?
- Varsayılan (başlangıç) ITEMA ayarlarını tanımlar
- SQL'deki:
    - dbo.sp_ItemaOtomatikAyar
    - dbo.sp_ItemaTipOzelAyar
  prosedürlerini çağırır
- Varsayılan + otomatik + tip-özel ayarları birleştirerek
  final ayar sözlüğü üretir.

Not:
- Bu modül SQL bağlantısı açmaz; dışarıdan pyodbc.Connection bekler.
- Böylece projendeki mevcut bağlantı yönetimiyle uyumlu çalışır.
"""

# ---------------------------------------------------------------------------
# 1) ITEMA kolonları ve varsayılan başlangıç ayarları
#    (VBA'deki "başlangıç ayarlara başla" bloğundan uyarlanmıştır)
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
    "arm_rap_1",
    "arm_rap_2",
    "arm_rap_3",
    "arm_rap_4",
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

# VBA başlangıç blokları:
#   j9, j10          -> 12 / 12
#   f14              -> "SEFFAF FIRÇA"
#   aa5, aa6         -> üfleme zamanı 1 / 2
#   aa7              -> cımbar
#   aa8              -> tansiyon
#   aa9              -> devir
#   aa10             -> leno
#   L30..L39, Q39..AA39 -> ağızlık, testere, yay, rampa vb.
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
    # AĞIZLIK / TESTERE / YAY / KONUM
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
    # MOTOR RAMPALARI
    "rampa_1": "100",
    "rampa_2": "50",
    "rampa_3": "10",
    "rampa_4": "120",
    "rampa_5": "30",
    "rampa_6": "9",
    # aciklama / degisiklik_yapan varsayılan boş
}


# ---------------------------------------------------------------------------
# 2) Yardımcı fonksiyonlar
# ---------------------------------------------------------------------------

def _row_to_dict(cursor: pyodbc.Cursor,
                 row: pyodbc.Row) -> Dict[str, Optional[str]]:
    """
    pyodbc Row nesnesini {kolon_adi: string veya None} formatına çevirir.
    """
    cols = [d[0] for d in cursor.description]
    out: Dict[str, Optional[str]] = {}
    for col, val in zip(cols, row):
        if val is None:
            out[col] = None
        else:
            out[col] = str(val)
    return out


def _merge_settings(base: Dict[str, Optional[str]],
                    override: Dict[str, Optional[str]]) -> Dict[str, Optional[str]]:
    """
    override içindeki "boş olmayan" değerleri base üzerine yazar.

    Boş sayılanlar:
    - None
    - "" (veya sadece whitespace)
    """
    for key, val in override.items():
        if val is None:
            continue
        if isinstance(val, str) and val.strip() == "":
            continue
        base[key] = val
    return base


# ---------------------------------------------------------------------------
# 3) SQL stored procedure çağrıları
# ---------------------------------------------------------------------------

def get_itema_automatic_settings(
    conn: pyodbc.Connection,
    tip: str
) -> Optional[Dict[str, Optional[str]]]:
    """
    SQL'deki dbo.sp_ItemaOtomatikAyar prosedürünü çağırır.

    Parametre:
        tip : F4 (tip kodu)

    Dönüş:
        - dict (kolon: değer) veya
        - None (otomatik satır bulunamazsa)
    """
    cur = conn.cursor()
    cur.execute("EXEC dbo.sp_ItemaOtomatikAyar @Tip = ?", tip)
    row = cur.fetchone()
    if not row:
        return None
    return _row_to_dict(cur, row)


def get_itema_tip_specific_settings(
    conn: pyodbc.Connection,
    tip: str
) -> List[Dict[str, Optional[str]]]:
    """
    SQL'deki dbo.sp_ItemaTipOzelAyar prosedürünü çağırır.

    Parametre:
        tip : F4 (tip kodu)

    Dönüş:
        Tip-özel satırların listesi (her satır sözlük).
    """
    cur = conn.cursor()
    cur.execute("EXEC dbo.sp_ItemaTipOzelAyar @Tip = ?", tip)
    rows = cur.fetchall()
    if not rows:
        return []

    results: List[Dict[str, Optional[str]]] = []
    for row in rows:
        results.append(_row_to_dict(cur, row))
    return results


# ---------------------------------------------------------------------------
# 4) Final ITEMA ayarlarını üreten ana fonksiyon
# ---------------------------------------------------------------------------

def build_itema_settings(
    conn: pyodbc.Connection,
    tip: str
) -> Dict[str, Optional[str]]:
    """
    ITEMA ayarlarının tamamını hesaplar:

      1) Tüm kolonlar için None içeren boş bir dict oluşturur
      2) Üzerine DEFAULT_ITEMA_SETTINGS değerlerini yazar
      3) Üzerine sp_ItemaOtomatikAyar çıktısını yazar (varsa)
      4) En sonunda sp_ItemaTipOzelAyar çıktısındaki satırları sırayla yazar

    Dönüş:
        { "telef_ken1": "...", "telef_ken2": "...", ... } şeklinde final ayar sözlüğü
    """
    # 1) Boş iskelet
    settings: Dict[str, Optional[str]] = {col: None for col in ITEMA_COLUMNS}

    # Tip bilgisini baştan set et
    settings["tip"] = tip

    # 2) Varsayılan ayarları uygula
    settings = _merge_settings(settings, DEFAULT_ITEMA_SETTINGS)

    # 3) Otomatik ayarları uygula
    auto = get_itema_automatic_settings(conn, tip)
    if auto:
        settings = _merge_settings(settings, auto)

    # 4) Tip-özel satırları uygula (VBA'de itema_tipe_has_ayarlar)
    tip_rows = get_itema_tip_specific_settings(conn, tip)
    for row in tip_rows:
        settings = _merge_settings(settings, row)

    return settings


# ---------------------------------------------------------------------------
# 5) Basit lokal test (opsiyonel)
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    # Bu blok sadece bağımsız test içindir; proje içinde kullanılmasına gerek yok.
    conn_str = (
        "Driver={SQL Server};"
        "Server=STIBRSSFSRV01;"
        "Database=UzmanRaporDB;"
        "Trusted_Connection=yes;"
    )
    conn = pyodbc.connect(conn_str)

    test_tip = "RX14908"  # Veritabanında olduğuna emin olduğun bir tip ile dene
    result = build_itema_settings(conn, test_tip)

    print(f"TIP = {test_tip}")
    for k in sorted(result.keys()):
        print(f"{k:20s} = {result[k]}")
