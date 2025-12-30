from __future__ import annotations

import re
from typing import Any, Dict, Iterable, List, Optional, Protocol, Sequence, Tuple
from app.db_name import DB_NAME



# ---------------------------------------------------------------------------
# 1) ITEMA kolonları
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


# ---------------------------------------------------------------------------
# 2) Minimal DB protokolleri (API cursor/connection ile uyumlu)
# ---------------------------------------------------------------------------

class CursorLike(Protocol):
    description: Optional[list[tuple[Any, ...]]]

    def execute(self, query: str, params: Optional[Sequence[Any]] = None) -> "CursorLike":
        ...

    def fetchone(self) -> Optional[Iterable[Any]]:
        ...

    def fetchall(self) -> list[Iterable[Any]]:
        ...


class ConnectionLike(Protocol):
    def cursor(self) -> CursorLike:
        ...


# ---------------------------------------------------------------------------
# 3) Genel yardımcılar
# ---------------------------------------------------------------------------

def _row_to_dict(cursor: CursorLike, row: Iterable[Any]) -> Dict[str, Optional[str]]:
    cols = [d[0] for d in (cursor.description or [])]
    out: Dict[str, Optional[str]] = {}
    for col, val in zip(cols, row):
        out[str(col)] = None if val is None else str(val)
    return out


def _merge_settings(base: Dict[str, Optional[str]], override: Dict[str, Optional[str]]) -> Dict[str, Optional[str]]:
    """
    override içindeki dolu değerleri base'e yazar.
    """
    for k, v in override.items():
        if v is None:
            continue
        if isinstance(v, str) and v.strip() == "":
            continue
        base[k] = v
    return base


def _norm_text(s: Optional[str]) -> Optional[str]:
    if s is None:
        return None
    t = str(s).strip()
    if not t:
        return None
    # bire bir eşleşmeyi bozmadan sadece whitespace normalizasyonu:
    t = re.sub(r"\s+", " ", t)
    return t


def _to_float(val: Optional[str]) -> Optional[float]:
    if val is None:
        return None
    s = str(val).strip()
    if not s or s.lower() == "nan":
        return None
    s = s.replace(",", ".")
    m = re.search(r"[-+]?\d+(\.\d+)?", s)
    if not m:
        return None
    try:
        return float(m.group(0))
    except Exception:
        return None


def _to_int(val: Optional[str]) -> Optional[int]:
    f = _to_float(val)
    if f is None:
        return None
    try:
        return int(round(f))
    except Exception:
        return None


def _parse_first_number(val: Optional[str]) -> Optional[float]:
    """
    '20', '20/1', '80+' gibi değerlerden ilk sayısalı çek.
    """
    return _to_float(val)


def _get_ci(d: Dict[str, Optional[str]], *names: str) -> Optional[str]:
    """
    Dict'te kolon adları bazen farklı case ile gelebiliyor (API/SQL/pyodbc farkları).
    Bu helper, case-insensitive bulur.
    """
    if not d:
        return None

    # hızlı yol: aynen
    for n in names:
        if n in d:
            return d.get(n)

    # case-insensitive
    lower_map = {k.lower(): k for k in d.keys()}
    for n in names:
        key = lower_map.get(n.lower())
        if key is not None:
            return d.get(key)

    return None


# ---------------------------------------------------------------------------
# 4) Aralık parse / contains
#    Desteklenen:
#      - "20-25" (inclusive)
#      - "80+"   (>= 80)
#      - "15"    (= 15)
#      - "(26, 27]" "[10,15)" vb (a,b) / [a,b] / (a,b] / [a,b)
# ---------------------------------------------------------------------------

_interval_re = re.compile(r"^\s*([\(\[])\s*([^,]+?)\s*,\s*([^\]\)]+?)\s*([\)\]])\s*$")
_hyphen_re = re.compile(r"^\s*([-+]?\d+(?:[.,]\d+)?)\s*-\s*([-+]?\d+(?:[.,]\d+)?)\s*$")
_plus_re = re.compile(r"^\s*([-+]?\d+(?:[.,]\d+)?)\s*\+\s*$")


def _interval_contains(spec: Optional[str], x: Optional[float]) -> bool:
    if spec is None or x is None:
        return False

    s = str(spec).strip()
    if not s:
        return False

    # 80+
    m = _plus_re.match(s)
    if m:
        a = _to_float(m.group(1))
        return (a is not None) and (x >= a)

    # 20-25
    m = _hyphen_re.match(s)
    if m:
        a = _to_float(m.group(1))
        b = _to_float(m.group(2))
        if a is None or b is None:
            return False
        lo, hi = (a, b) if a <= b else (b, a)
        return lo <= x <= hi

    # (26, 27] gibi
    m = _interval_re.match(s)
    if m:
        left_br, a_s, b_s, right_br = m.groups()
        a = _to_float(a_s)
        b = _to_float(b_s)
        if a is None or b is None:
            return False

        lo, hi = (a, b) if a <= b else (b, a)
        left_ok = (x > lo) if left_br == "(" else (x >= lo)
        right_ok = (x < hi) if right_br == ")" else (x <= hi)
        return left_ok and right_ok

    # tek sayı
    a = _to_float(s)
    if a is not None:
        return abs(x - a) < 1e-9

    return False


# ---------------------------------------------------------------------------
# 5) Feature üretimi (dinamik rapordan)
# ---------------------------------------------------------------------------

def _calc_cozgu_siklik_from_tarak_grubu(tg: Optional[str]) -> Optional[float]:
    """
    Örn: 67,5/4/194 => (67.5*4)/10 = 27.0
    """
    if tg is None:
        return None
    s = str(tg).strip()
    if not s:
        return None

    parts = [p.strip() for p in s.split("/") if p.strip()]
    if len(parts) < 2:
        return None

    first = _to_float(parts[0])
    mid = _to_float(parts[1])
    if first is None or mid is None:
        return None

    return (first * mid) / 10.0


def _extract_features_from_tip_features(tip_features: Optional[Dict[str, Optional[str]]]) -> Dict[str, Any]:
    """
    tip_features = ItemaTab._populate_from_dynamic() çıktısı (UI key'leri)
    Buradan eşleştirme için gereken feature set üretilir.
    """
    tf = tip_features or {}

    orgu_tipi = _norm_text(tf.get("zemin_orgu"))
    cozgu1 = _parse_first_number(tf.get("cozgu1"))  # combine ile gelmiş olabilir: "20 20/1" gibi
    atki1 = _parse_first_number(tf.get("atki1"))
    dok = _to_float(tf.get("dokunabilirlik"))

    atki_siklik = _to_float(tf.get("atki_sikligi"))  # 7100 alanı UI'da burada
    cozgu_siklik = _calc_cozgu_siklik_from_tarak_grubu(tf.get("tarak_grubu"))

    return {
        "orgu_tipi": orgu_tipi,
        "cozgu1": cozgu1,
        "atki1": atki1,
        "cozgu_siklik": cozgu_siklik,
        "atki_siklik": atki_siklik,
        "dokunabilirlik": dok,
    }


# ---------------------------------------------------------------------------
# 6) SQL okumalar
# ---------------------------------------------------------------------------

def _fetch_itema_ayar_by_tip(conn: ConnectionLike, tip: str) -> Optional[Dict[str, Optional[str]]]:
    cur = conn.cursor()
    cur.execute(f"SELECT TOP 1 * FROM [{DB_NAME}].[dbo].[ItemaAyar] WHERE [tip] = ?", [tip])
    row = cur.fetchone()
    if not row:
        return None
    return _row_to_dict(cur, row)


def _pick_best_row(rows: List[Dict[str, Optional[str]]]) -> Optional[Dict[str, Optional[str]]]:
    """
    Birden çok eşleşirse TECRUBE_SAYISI en yüksek olanı seç.
    Yoksa ilkini al.
    """
    if not rows:
        return None

    def score(r: Dict[str, Optional[str]]) -> Tuple[int, int]:
        ts = _to_int(_get_ci(r, "TECRUBE_SAYISI"))
        return (ts or 0, 1)

    return sorted(rows, key=score, reverse=True)[0]

def _get_ci(row: Dict[str, Optional[str]], key: str) -> Optional[str]:
    """
    Row dict içinde kolon adı büyük/küçük harf farkını tolere ederek değer okur.
    Örn DB 'Cozgu1_Aralik' yerine 'cozgu1_aralik' döndürdüyse de çalışır.
    """
    if key in row:
        return row.get(key)
    lk = key.lower()
    for k, v in row.items():
        if str(k).lower() == lk:
            return v
    return None

def _fetch_makine_ayar_match(
    conn: ConnectionLike,
    features: Dict[str, Any],
) -> Optional[Dict[str, Optional[str]]]:
    """
    dbo.Makine_Ayar_Tablosu:
      - orgu_tipi bire bir
      - diğer 5 aralık sütunu value contains (tam eşleşme)
      - tam eşleşme yoksa: aynı orgu_tipi içinden "en yakın" satır (mesafe: aralık merkezine uzaklıkların toplamı)
    """
    orgu_tipi = features.get("orgu_tipi")
    cozgu1 = features.get("cozgu1")
    atki1 = features.get("atki1")
    cozgu_siklik = features.get("cozgu_siklik")
    atki_siklik = features.get("atki_siklik")
    dok = features.get("dokunabilirlik")

    # Feature eksikse eşleşme yok (default basma yok)
    if (
        not orgu_tipi
        or cozgu1 is None
        or atki1 is None
        or cozgu_siklik is None
        or atki_siklik is None
        or dok is None
    ):
        return None

    cur = conn.cursor()

    # Sadece aynı örgü tipinden adayları çek
    cur.execute(
        f"SELECT * FROM [{DB_NAME}].[dbo].[Makine_Ayar_Tablosu] WHERE [orgu_tipi] = ?",
        [orgu_tipi],
    )
    raw = cur.fetchall()
    if not raw:
        return None

    all_rows = [_row_to_dict(cur, r) for r in raw]

    # ------------------ 1) TAM EŞLEŞME ------------------
    matches: List[Dict[str, Optional[str]]] = []
    for r in all_rows:
        if not _interval_contains(_get_ci(r, "Cozgu1_Aralik"), cozgu1):
            continue
        if not _interval_contains(_get_ci(r, "Atki1_Aralik"), atki1):
            continue
        if not _interval_contains(_get_ci(r, "CozguSiklik_Aralik"), cozgu_siklik):
            continue
        if not _interval_contains(_get_ci(r, "AtkiSiklik_Aralik"), atki_siklik):
            continue
        if not _interval_contains(_get_ci(r, "Dokunabilirlik_Aralik"), dok):
            continue
        matches.append(r)

    # Tam eşleşme varsa: TECRUBE_SAYISI en yüksek olanı seç
    if matches:
        return _pick_best_row(matches)

    # ------------------ 2) EN YAKIN BENZEYEN ------------------
    # Mesafe: her feature için |x - center(interval)| toplamı
    scored: List[Tuple[float, Dict[str, Optional[str]]]] = []

    for r in all_rows:
        c1 = _interval_center(_get_ci(r, "Cozgu1_Aralik"))
        c2 = _interval_center(_get_ci(r, "Atki1_Aralik"))
        c3 = _interval_center(_get_ci(r, "CozguSiklik_Aralik"))
        c4 = _interval_center(_get_ci(r, "AtkiSiklik_Aralik"))
        c5 = _interval_center(_get_ci(r, "Dokunabilirlik_Aralik"))

        # Aralık parse edilemeyen satırları ele
        if c1 is None or c2 is None or c3 is None or c4 is None or c5 is None:
            continue

        dist = (
            abs(cozgu1 - c1)
            + abs(atki1 - c2)
            + abs(cozgu_siklik - c3)
            + abs(atki_siklik - c4)
            + abs(dok - c5)
        )

        scored.append((dist, r))

    if not scored:
        return None

    scored.sort(key=lambda x: x[0])
    best_dist, best_row = scored[0]

    # Eğer aynı dist'e sahip birden fazla satır varsa, TECRUBE_SAYISI ile tie-break yap
    # (opsiyonel ama iyi olur)
    tied = [r for d, r in scored if abs(d - best_dist) < 1e-12]
    if len(tied) > 1:
        pick = _pick_best_row(tied)
        return pick or best_row

    return best_row

def get_itema_settings_from_feature_tables(
    conn: ConnectionLike,
    tip: str,
    tip_features: Optional[Dict[str, Optional[str]]],
) -> Optional[Dict[str, Optional[str]]]:
    """
    1) Önce ItemaAyar (manuel/override)
    2) Yoksa Makine_Ayar_Tablosu'nda feature eşleştirme
    """
    row = _fetch_itema_ayar_by_tip(conn, tip)
    if row:
        return row

    features = _extract_features_from_tip_features(tip_features)
    return _fetch_makine_ayar_match(conn, features)


# ---------------------------------------------------------------------------
# 6.1) Stored procedure çağrıları (API uyumlu)
# ---------------------------------------------------------------------------

def get_itema_automatic_settings(conn: ConnectionLike, tip: str) -> Optional[Dict[str, Optional[str]]]:
    cur = conn.cursor()
    # API katmanı için pozisyonel parametre daha stabil
    cur.execute("EXEC dbo.sp_ItemaOtomatikAyar ?", [tip])
    row = cur.fetchone()
    if not row:
        return None
    return _row_to_dict(cur, row)


def get_itema_tip_specific_settings(conn: ConnectionLike, tip: str) -> List[Dict[str, Optional[str]]]:
    cur = conn.cursor()
    cur.execute("EXEC dbo.sp_ItemaTipOzelAyar ?", [tip])
    rows = cur.fetchall()
    if not rows:
        return []
    return [_row_to_dict(cur, r) for r in rows]


# ---------------------------------------------------------------------------
# 7) Final ITEMA ayarlarını üreten ana fonksiyon
#    DEFAULT KALDIRILDI.
# ---------------------------------------------------------------------------

def build_itema_settings(
    conn: ConnectionLike,
    tip: str,
    tip_features: Optional[Dict[str, Optional[str]]] = None
) -> Dict[str, Optional[str]]:

    # Default yok: hepsi None başlar
    settings: Dict[str, Optional[str]] = {col: None for col in ITEMA_COLUMNS}
    settings["tip"] = tip

    # 1) Önce manuel tablo veya feature-table eşleştirme
    table_settings = get_itema_settings_from_feature_tables(conn, tip, tip_features)
    if table_settings:
        settings = _merge_settings(settings, table_settings)
    else:
        # 2) Otomatik ayar
        auto = get_itema_automatic_settings(conn, tip)
        if auto:
            settings = _merge_settings(settings, auto)

        # 3) Tip-özel ayar satırları
        tip_rows = get_itema_tip_specific_settings(conn, tip)
        for row in tip_rows:
            settings = _merge_settings(settings, row)

    # --- SABİT GELMESİ İSTENEN ALANLAR (sadece boşsa doldur) ---
    fixed_defaults = {
        "telef_ken1": "12",
        "telef_ken2": "12",
        "ufleme_zam_1": "00:05:00",
        "ufleme_zam_2": "00:00:05",
        "leno": "150",
    }
    for k, v in fixed_defaults.items():
        cur = settings.get(k)
        if cur is None or str(cur).strip() == "":
            settings[k] = v

    return settings
def _interval_center(spec: Optional[str]) -> Optional[float]:
    if spec is None:
        return None
    s = str(spec).strip()
    if not s:
        return None

    # 80+
    m = _plus_re.match(s)
    if m:
        return _to_float(m.group(1))

    # 20-25
    m = _hyphen_re.match(s)
    if m:
        a = _to_float(m.group(1))
        b = _to_float(m.group(2))
        if a is None or b is None:
            return None
        return (a + b) / 2.0

    # (26, 27] / [10,15) vb
    m = _interval_re.match(s)
    if m:
        _, a_s, b_s, _ = m.groups()
        a = _to_float(a_s)
        b = _to_float(b_s)
        if a is None or b is None:
            return None
        return (a + b) / 2.0

    # tek sayı
    return _to_float(s)
