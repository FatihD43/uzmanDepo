
from __future__ import annotations
import pandas as pd
import re
from pathlib import Path
from PySide6.QtWidgets import QMessageBox


VISIBLE_COLUMNS = [
    "Tezgah Numarası", "Kök Tip Kodu",
    "Levent No", "Levent Etiket FA", "Tarak Grubu", "Zemin Örgü", "Üretim Sipariş No", "Haşıl İş Emri",
    "Atkı İpliği 1", "Atkı İpliği 2",
    "Çözgü İpliği 1", "Çözgü İpliği 2",
    "Parti Metresi", "Mamul Termin", "İhzarat Boya Kodu", "Süs Kenar",
    "NOTLAR", "Mamul Tip Kodu",
    "(Atkı-1 İşletme Depoları + Atkı-1 İşletme Diğer Depoları)",
    "(Atkı-2 İşletme Depoları + Atkı-2 İşletme Diğer Depoları)",
    "Atkı İhtiyaç Miktar 1",
    "Atkı İhtiyaç Miktar 2","Çözgü İpliği 3","Çözgü İpliği 4", "Levent Tipi", "Durum Tanım", "Levent Haşıl Tarihi"
]

TR_MAP = str.maketrans("şığüçıİÖöÜŞĞÇÂâ", "siguciiooUSGCAa")
# UI'de görünen kolon isimleri (yalnızca görünümde değişir)
HEADER_ALIASES = {
    "Tezgah Numarası": "Tezgah",
    "Kök Tip Kodu": "KökTip",
    "Levent Etiket FA": "EtiketFA",
    "Üretim Sipariş No": "Dokuma İş Emri",
    "Parti Metresi": "Metre",
    "İhzarat Boya Kodu": "BoyaKodu",
    "(Atkı-1 İşletme Depoları + Atkı-1 İşletme Diğer Depoları)": "Atkı1Stok",
    "(Atkı-2 İşletme Depoları + Atkı-2 İşletme Diğer Depoları)": "Atkı2Stok"



    # buraya başka dönüşümler de ekleyebilirsin:
    # "Mamul Termin": "Termin",
    # "Tezgah No": "Tezgah Numarası",
}

def _norm(s: str) -> str:
    if s is None: return ""
    return str(s).translate(TR_MAP).strip()

def _norm_upper(s: str) -> str:
    return _norm(s).upper()

def _extract_numbers(s: str):
    s = _norm(s).replace(",", ".")
    return re.findall(r"\d+(?:\.\d+)?", s)

def _numbers_key(s: str) -> str:
    nums = _extract_numbers(s)
    if not nums: return ""
    vals = []
    for n in nums:
        f = float(n)
        if abs(f - round(f)) < 1e-9:
            vals.append(str(int(round(f))))
        else:
            vals.append(str(f))
    return "-".join(vals)

def _tarak_key(s: str) -> str:
    k = _numbers_key(s)
    return k if k else _norm_upper(s)

def _is_date_like(series: pd.Series, thresh: float = 0.5) -> bool:
    try:
        parsed = pd.to_datetime(series, errors="coerce", dayfirst=True)
        ratio = parsed.notna().mean()
        return ratio >= thresh
    except Exception:
        return False

def _has_digits(series: pd.Series, thresh: float = 0.2) -> bool:
    s = series.astype(str)
    return s.str.contains(r"\d", regex=True, na=False).mean() >= thresh

def _pick_levent_no_fa(df: pd.DataFrame):
    cols = list(df.columns)
    def header_score(c: str) -> int:
        cn = _norm_upper(c)
        sc = 0
        if "LEVENT" in cn: sc += 3
        if "FA" in cn: sc += 2
        if "NO" in cn: sc += 1
        if "ETIKET" in cn: sc -= 3
        return sc
    candidates = [c for c in cols if ("LEVENT" in _norm_upper(c) and "FA" in _norm_upper(c) and "ETIKET" not in _norm_upper(c))]
    if not candidates:
        candidates = [c for c in cols if "LEVENT" in _norm_upper(c) and "ETIKET" not in _norm_upper(c)]
    best_c = None; best_tuple = (-1, -1.0, -1)
    for c in candidates:
        ser = df[c]
        not_date = 0 if _is_date_like(ser) else 1
        digit_ratio = ser.astype(str).str.contains(r"\d", regex=True, na=False).mean()
        hs = header_score(c)
        key = (not_date, digit_ratio, hs)
        if key > best_tuple:
            best_tuple = key; best_c = c
    if best_c is not None and best_tuple[0] == 1:
        return df[best_c].astype(str), best_c
    import re
    def compact(s: str) -> str:
        return re.sub(r"[^A-Z0-9]+", "", _norm_upper(s))
    target = compact("Levent No FA")
    for c in cols:
        if compact(c) == target and not _is_date_like(df[c]):
            return df[c].astype(str), c
    if len(cols) > 13:
        c = cols[13]
        if not _is_date_like(df[c]) and _has_digits(df[c]):
            return df[c].astype(str), c
    for c in cols:
        if "LEVENT" in _norm_upper(c) and not _is_date_like(df[c]):
            return df[c].astype(str), c
    return pd.Series([""]*len(df)), ""

def load_dinamik_any(path: str|Path) -> pd.DataFrame:
    p = str(path).lower()
    if p.endswith(".xlsb"):
        df = pd.read_excel(path, engine="pyxlsb")
    else:
        df = pd.read_excel(path, sheet_name=0, engine="openpyxl")
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    lev_series, lev_col = _pick_levent_no_fa(df)
    lev_series_clean = pd.to_numeric(lev_series, errors="coerce").astype("Int64").astype(str).replace("<NA>", "")
    df["Levent No"] = lev_series_clean
    df["_LeventSource"] = lev_col

    if "Mamul Termin" in df.columns:
        df["Mamul Termin"] = pd.to_datetime(df["Mamul Termin"], errors="coerce")

    if "Tezgah Numarası" not in df.columns: df["Tezgah Numarası"] = ""
    if "NOTLAR" not in df.columns: df["NOTLAR"] = ""

    if "Tarak Grubu" in df.columns:
        df["_TarakKey"] = df["Tarak Grubu"].astype(str).apply(_tarak_key)
    else:
        df["_TarakKey"] = ""

    s = df["Levent No"].astype(str)
    df["_LeventHasDigits"] = s.str.contains(r"\d", regex=True) & s.str.strip().ne("")

    # Dye category from İhzarat Boya Kodu
    dye_col = None
    for c in df.columns:
        if "İhzarat" in c or "Ihzarat" in c:
            dye_col = c; break
    if dye_col is None and "İhzarat Boya Kodu" in df.columns:
        dye_col = "İhzarat Boya Kodu"
    if dye_col is None:
        df["_DyeCategory"] = "DENIM"
    else:
        vals = df[dye_col].astype(str).str.upper()
        # "HAM", "HAM HAM", "HAM HAM HAM", "  ham  " vs hepsi HAM sayılacak
        df["_DyeCategory"] = vals.apply(lambda x: "HAM" if "HAM" in x else "DENIM")
        # --- ADD: Levent Etiket FA + İhtiyaç miktarları + Atkı depo toplamları ---
    import numpy as np

    def _num(s):
        return pd.to_numeric(s, errors="coerce")

    def _safe_get(df_in, name, *aliases):
        for k in (name, *aliases):
            if k in df_in.columns:
                return df_in[k]
        return pd.Series(np.nan, index=df_in.index)

    # 1) Levent Etiket FA (M sütunu olarak iletilmişti)
    if "Levent Etiket FA" not in df.columns:
        df["Levent Etiket FA"] = np.nan

    # 2) Atkı İhtiyaç Miktar 1/2
    for col in ["Atkı İhtiyaç Miktar 1", "Atkı İhtiyaç Miktar 2"]:
        if col in df.columns:
            df[col] = _num(df[col])
        else:
            df[col] = np.nan

    # 3) Toplam sütunlar (tek isim altında)
    a1 = _num(_safe_get(df, "Atkı-1 İşletme Depoları"))
    a1d = _num(_safe_get(df, "Atkı-1 İşletme Diğer Depoları"))
    df["(Atkı-1 İşletme Depoları + Atkı-1 İşletme Diğer Depoları)"] = a1.add(a1d, fill_value=0)

    a2 = _num(_safe_get(df, "Atkı-2 İşletme Depoları"))
    a2d = _num(_safe_get(df, "Atkı-2 İşletme Diğer Depoları"))
    df["(Atkı-2 İşletme Depoları + Atkı-2 İşletme Diğer Depoları)"] = a2.add(a2d, fill_value=0)
    # --- İPLİK NUMARASI + İPLİK KODU BİRLEŞTİRME ---
    # Örn: 150 + RS0000269 → "150 RS0000269"
    df = _combine_yarn_with_number(df, "Atkı İpliği 1", "Atkı İplik No 1")
    df = _combine_yarn_with_number(df, "Atkı İpliği 2", "Atkı İplik No 2")
    df = _combine_yarn_with_number(df, "Çözgü İpliği 1", "Çözgü İplik No 1")
    df = _combine_yarn_with_number(df, "Çözgü İpliği 2", "Çözgü İplik No 2")

    # --- Etiket / Haşıl İş Emri sütunlarında .0 temizliği ---
    for col in ["Levent Etiket FA", "Haşıl İş Emri"]:
        if col in df.columns:
            s = df[col].astype(str).str.strip()
            # 123456.0  → 123456
            # 987.00    → 987
            s = s.str.replace(r"\.0+$", "", regex=True)
            # 'nan', 'NaT' stringlerini boş yap
            s = s.replace({"nan": "", "NaT": ""})
            df[col] = s

    return df
def _combine_yarn_with_number(df: pd.DataFrame, yarn_col: str, num_col: str) -> pd.DataFrame:
    """
    '<num> <yarn>' birleştirir.
    NaN/None/'nan' gibi değerleri kesinlikle yazmaz.
    """
    if yarn_col not in df.columns or num_col not in df.columns:
        return df

    def _clean_series(s: pd.Series) -> pd.Series:
        # 1) Gerçek NaN/None -> ""
        s = s.fillna("")
        # 2) Stringe çevir
        s = s.astype(str).str.strip()
        # 3) 'nan', 'none', 'null' gibi literal stringleri de temizle
        s = s.replace({"nan": "", "NaN": "", "None": "", "NONE": "", "null": "", "NULL": ""})
        return s

    ser_num = _clean_series(df[num_col])
    ser_yarn = _clean_series(df[yarn_col])

    combined = ser_yarn.copy()
    mask = (ser_num != "") & (ser_yarn != "")
    combined[mask] = ser_num[mask] + " " + ser_yarn[mask]

    # num var ama yarn yoksa: sadece num yazmak istersen aç
    # mask2 = (ser_num != "") & (ser_yarn == "")
    # combined[mask2] = ser_num[mask2]

    df[yarn_col] = combined
    return df


def load_running_orders(path: str|Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=0, engine="openpyxl")
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    if "Tezgah No" not in df.columns:
        for alt in ["Tezgah","Tezgâh No","Tezgah_No"]:
            if alt in df.columns: df["Tezgah No"] = df[alt]; break

    durus_col = None
    for c in df.columns:
        cn = _norm_upper(c)
        if cn.startswith("DURUS") or cn.startswith("DURU"):
            durus_col = c; break
    if durus_col is None: df["Durus No"] = 0
    else: df["Durus No"] = pd.to_numeric(df[durus_col], errors="coerce").fillna(0).astype(int)

    tarak_col = None
    for c in df.columns:
        if "TARAK" in _norm_upper(c): tarak_col = c; break
    if tarak_col is None:
        best, best_ratio = None, -1.0
        for c in df.columns:
            ratio = df[c].astype(str).str.contains(r"\d+(?:[\s,./-]+\d+)+", regex=True, na=False).mean()
            if ratio > best_ratio:
                best, best_ratio = c, ratio
        tarak_col = best
    df["Tarak Grubu"] = df[tarak_col].astype(str) if tarak_col else ""
    df["_TarakKey"] = df["Tarak Grubu"].apply(_tarak_key)

    flags = []
    for c in df.columns:
        cn = _norm_upper(c)
        if any(k in cn for k in ["SIPARIS","DURUM","DURUS","DURUŞ"]):
            flags.append(c)
    def has_siparis_yok(row)->bool:
        for c in flags:
            val = _norm_upper(row.get(c, ""))
            if "SIPARIS YOK" in val: return True
        return False
    df["_OpenTezgahFlag"] = df.apply(has_siparis_yok, axis=1)

    kalan_col = None
    for name in ["Kalan", "Kalan Mt", "Kalan Metre", "Kalan_Metre"]:
        if name in df.columns: kalan_col = name; break
    df["_KalanMetre"] = pd.to_numeric(df[kalan_col], errors="coerce").fillna(0.0) if kalan_col else pd.Series([0.0]*len(df))

    return df
import pandas as pd

from app.storage import (
    load_loom_cut_map, load_type_selvedge_map, save_type_selvedge_map
)
def enrich_running_with_loom_cut(df_run: pd.DataFrame) -> pd.DataFrame:
    if df_run is None or df_run.empty:
        return df_run
    col_tz = None
    for n in ["Tezgah No","Tezgah","Tezgah Numarası"]:
        if n in df_run.columns: col_tz = n; break
    if not col_tz:
        return df_run

    def _norm_choice(s: str) -> str | None:
        u = (str(s) if s is not None else "").strip().upper()
        if u == "ISAVER":
            return "ISAVER"
        if u == "ROTOCUT":
            return "ROTOCUT"
        if u in ("ISAVERKIT", "ISAVER KIT", "KIT"):
            return "ISAVERKit"
        return None

    d = load_loom_cut_map()  # {"2201":"ISAVER", ...}
    df_run = df_run.copy()
    df_run["ISAVER/ROTOCUT"] = df_run[col_tz].astype(str).map(lambda x: _norm_choice(d.get(x, None)))
    return df_run



def enrich_running_with_selvedge(df_run: pd.DataFrame, df_dinamik: pd.DataFrame) -> pd.DataFrame:
    """
    Kök Tip → Süs Kenar kütüphanesini günceller ve df_run'a 'Süs Kenar' sütununu oluşturur.

    Davranış:
    - SQL'deki TypeSelvedgeMap ilk olarak okunur (lib).
    - Dinamik'te KökTip + SüsKenar varsa:
        * Eğer lib'te yoksa: yeni kayıt EKLENİR.
        * Eğer lib'te varsa ve DEĞİŞİKSE: kullanıcıya "revize kabul edilsin mi?" diye sorulur.
          - EVET: lib'teki değer yeni değere güncellenir.
          - HAYIR: lib'teki eski değer korunur.
    - Running'de KökTip kolonu bulunursa, kütüphaneden 'Süs Kenar' doldurulur.
    """
    if df_run is None or df_run.empty:
        return df_run

    # 1) Kütüphaneyi SQL'den oku
    lib = load_type_selvedge_map()  # {"KOKTIP":"SÜS", ...}

    # 1.a) Dinamik'ten yeni / revize bilgileri topla
    if df_dinamik is not None and not df_dinamik.empty:
        col_kok = None
        for n in ["Kök Tip Kodu", "KökTip", "KokTip"]:
            if n in df_dinamik.columns:
                col_kok = n
                break

        col_sus = None
        for n in ["Süs Kenar", "SüsKenar", "Süs Kenarı", "Süs Kenarı Adı", "Selvedge", "Selvedge Tipi"]:
            if n in df_dinamik.columns:
                col_sus = n
                break

        if col_kok and col_sus:
            # Aynı tipi birden fazla satırda görürsek, Dinamik içindeki son değeri baz alalım
            dinamik_map: dict[str, str] = {}
            for kok, sus in df_dinamik[[col_kok, col_sus]].dropna().itertuples(index=False):
                k = str(kok).strip().upper()
                v = str(sus).strip()
                if not k or not v:
                    continue
                dinamik_map[k] = v

            changed = False  # SQL'e gerçekten yazmamız gerekecek mi?

            for k, v in dinamik_map.items():
                eski = lib.get(k)

                # Kütüphanede yoksa: direkt ekle (öğren)
                if eski is None:
                    lib[k] = v
                    changed = True
                    continue

                # Kütüphanede var ve değer aynıysa: bir şey yapma
                if str(eski).strip() == v:
                    continue

                # Buraya geldiysek: REVİZYON TESPİT EDİLDİ
                # Kullanıcıya soralım
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Question)
                msg.setWindowTitle("Süs Kenar Revizyonu")
                msg.setText(
                    f"Kök tip için süs kenar revizyonu tespit edildi.\n\n"
                    f"Kök Tip: {k}\n"
                    f"Tezgahtaki (kütüphane) süs kenar: {eski}\n"
                    f"Dinamik rapordaki yeni süs kenar: {v}\n\n"
                    "Bu tip için revizeyi kabul etmek istiyor musunuz?"
                )
                msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
                msg.setDefaultButton(QMessageBox.No)

                cevap = msg.exec()

                if cevap == QMessageBox.Yes:
                    # Revizeyi kabul et → kütüphaneyi güncelle
                    lib[k] = v
                    changed = True
                else:
                    # HAYIR dendi → eski değer olduğu gibi kalsın
                    pass

            if changed:
                save_type_selvedge_map(lib)

    # 2) Running'i doldur: KökTip üzerinden 'Süs Kenar' üret
    col_kok_run = None
    for n in ["Kök Tip Kodu", "KökTip", "KokTip"]:
        if n in df_run.columns:
            col_kok_run = n
            break

    df_out = df_run.copy()

    if col_kok_run and lib:
        # --- YENİ: Running'de gerçekten kullanılan tipleri bul ---
        used_keys = (
            df_out[col_kok_run]
            .astype(str)
            .str.strip()
            .str.upper()
            .dropna()
            .unique()
            .tolist()
        )
        used_set = set(used_keys)

        # Sadece bu tipler için küçük bir kütüphane oluştur
        lib_small = {k: v for k, v in lib.items() if k in used_set}

        # Şimdi mapping'i bu küçük sözlükle yap
        df_out["Süs Kenar"] = (
            df_out[col_kok_run]
            .astype(str)
            .str.strip()
            .str.upper()
            .map(lambda x: lib_small.get(x, None))
        )
    else:
        df_out["Süs Kenar"] = pd.NA

    return df_out
