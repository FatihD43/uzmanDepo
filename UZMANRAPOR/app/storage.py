from __future__ import annotations
from pathlib import Path
import os, json, sqlite3, re
from datetime import datetime
from zoneinfo import ZoneInfo  # <-- Istanbul TZ
import pandas as pd

APP_DIR = Path.home() / ".uzman_rapor"
APP_DIR.mkdir(parents=True, exist_ok=True)

RULES_PATH = APP_DIR / "notes_rules.json"
META_PATH  = APP_DIR / "meta.json"
DINAMIK_SNAPSHOT = APP_DIR / "dinamik.pkl"
RUNNING_SNAPSHOT = APP_DIR / "running.pkl"
USERCFG_PATH = APP_DIR / "user.json"


# ---------- Notes rules (kalıcı) ----------
def load_rules() -> list[dict]:
    if RULES_PATH.exists():
        try:
            with open(RULES_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                return data
        except Exception:
            pass
    return []

def save_rules(rules: list[dict]) -> None:
    try:
        with open(RULES_PATH, "w", encoding="utf-8") as f:
            json.dump(rules or [], f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ---------- Son güncelleme (Planlama tıklanınca) ----------
def load_last_update() -> datetime | None:
    if META_PATH.exists():
        try:
            with open(META_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            iso = data.get("last_update_iso")
            if iso:
                dt = datetime.fromisoformat(iso)
                # naive ise İstanbul TZ ata
                if dt.tzinfo is None:
                    dt = dt.replace(tzinfo=ZoneInfo("Europe/Istanbul"))
                return dt
        except Exception:
            pass
    return None

def save_last_update(dt: datetime) -> None:
    try:
        # dt naive ise İstanbul TZ’li kabul et
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=ZoneInfo("Europe/Istanbul"))
        meta = {"last_update_iso": dt.isoformat()}
        if META_PATH.exists():
            try:
                with open(META_PATH, "r", encoding="utf-8") as f:
                    cur = json.load(f)
                if isinstance(cur, dict):
                    cur.update(meta)
                    meta = cur
            except Exception:
                pass
        with open(META_PATH, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ---------- Snapshot (DF'ler) ----------
def save_df_snapshot(df: pd.DataFrame | None, which: str) -> None:
    try:
        if df is None:
            return
        if which == "dinamik":
            df.to_pickle(DINAMIK_SNAPSHOT)
        elif which == "running":
            df.to_pickle(RUNNING_SNAPSHOT)
    except Exception:
        pass

def load_df_snapshot(which: str) -> pd.DataFrame | None:
    try:
        if which == "dinamik" and DINAMIK_SNAPSHOT.exists():
            return pd.read_pickle(DINAMIK_SNAPSHOT)
        if which == "running" and RUNNING_SNAPSHOT.exists():
            return pd.read_pickle(RUNNING_SNAPSHOT)
    except Exception:
        pass
    return None


# ---------- Kullanıcı varsayılanı ----------
def get_username_default() -> str:
    if USERCFG_PATH.exists():
        try:
            with open(USERCFG_PATH, "r", encoding="utf-8") as f:
                d = json.load(f)
            u = d.get("username")
            if u:
                return str(u)
        except Exception:
            pass
    return "Anonim"

def set_username_default(name: str) -> None:
    try:
        d = {"username": name}
        if USERCFG_PATH.exists():
            try:
                with open(USERCFG_PATH, "r", encoding="utf-8") as f:
                    cur = json.load(f)
                if isinstance(cur, dict):
                    cur.update(d)
                    d = cur
            except Exception:
                pass
        with open(USERCFG_PATH, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# === Kesim Tipi (ISAVER/ROTOCUT) ve Süs Kenar (Kök Tip) kalıcı sözlükleri ===

def _ensure_app_dir():
    base = os.path.join(os.path.expanduser("~"), ".uzman_rapor")
    os.makedirs(base, exist_ok=True)
    return base

def _kv_path(name: str):
    return os.path.join(_ensure_app_dir(), f"{name}.json")

# --- Loom -> Kesim Tipi (ISAVER/ROTOCUT) ---
def load_loom_cut_map() -> dict:
    p = _kv_path("loom_cut_map")
    try:
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_loom_cut_map(d: dict) -> None:
    p = _kv_path("loom_cut_map")
    with open(p, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=2)

# --- Kök Tip -> Süs Kenar kütüphanesi ---
def load_type_selvedge_map() -> dict:
    p = _kv_path("type_selvedge_map")
    try:
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_type_selvedge_map(d: dict) -> None:
    p = _kv_path("type_selvedge_map")
    with open(p, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=2)


# --- USTA DEFTERİ: sayım yardımcıları ---------------------------------------
def _pick_col(df: pd.DataFrame, names: list[str]) -> str | None:
    for n in names:
        if n in df.columns:
            return n
    # lowercase fallback
    low = {c.lower(): c for c in df.columns}
    for n in names:
        if n.lower() in low:
            return low[n.lower()]
    return None

def load_usta_dataframe(sqlite_path: str | None = None) -> pd.DataFrame:
    """
    Usta Defteri kayıtlarını DataFrame olarak döndürür.
    Tercih sırası: sqlite -> csv -> boş df
    Kolon adları farklı olabilir; normalize eder:
      _ts: datetime, _what: {'DÜĞÜM','TAKIM',...}, _dir: {'ALINDI','VERİLDİ',...}
    """
    # 1) Kaynak: sqlite (varsayılan proje kökünde 'usta_defteri.sqlite')
    if sqlite_path is None:
        sqlite_path = os.path.join(os.getcwd(), "usta_defteri.sqlite")
    df = pd.DataFrame()
    if os.path.exists(sqlite_path):
        try:
            with sqlite3.connect(sqlite_path) as conn:
                # olası tablo adları: 'usta_defteri', 'entries', 'logs'
                for tname in ["usta_defteri", "entries", "logs"]:
                    try:
                        tmp = pd.read_sql(f"SELECT * FROM {tname}", conn)
                        if not tmp.empty:
                            df = tmp
                            break
                    except Exception:
                        continue
        except Exception:
            pass

    # 2) CSV fallback
    if df.empty:
        for csv_name in ["usta_defteri.csv", "usta_defteri_logs.csv"]:
            p = os.path.join(os.getcwd(), csv_name)
            if os.path.exists(p):
                try:
                    df = pd.read_csv(p)
                    break
                except Exception:
                    pass

    if df.empty:
        return pd.DataFrame(columns=["_ts","_what","_dir"])

    # --- normalize ---
    # zaman
    c_ts = _pick_col(df, ["created_at","timestamp","Tarih Saat","Tarih","Datetime","Date"])
    if c_ts:
        ts = pd.to_datetime(df[c_ts], dayfirst=True, errors="coerce")
    else:
        ts = pd.Series(pd.NaT, index=df.index)
    df["_ts"] = ts

    # tür / konu (düğüm/takım)
    c_what = _pick_col(df, ["Tür","Tip","Kategori","Kayıt Tipi","Kayit Turu","Subject","Topic"])
    # işlem / yön (alındı/verildi)
    c_dir  = _pick_col(df, ["İşlem","Aksiyon","Durum","Action","Operation","Aciklama","Açıklama","Not"])

    def _norm_what(row):
        texts = []
        for c in [c_what, c_dir]:
            if c and c in df.columns:
                texts.append(str(row.get(c, "")).upper())
        blob = " ".join(texts)
        if "DÜĞÜM" in blob:
            return "DÜĞÜM"
        if "TAKIM" in blob:
            return "TAKIM"
        return ""

    def _norm_dir(row):
        texts = []
        for c in [c_dir, c_what]:
            if c and c in df.columns:
                texts.append(str(row.get(c, "")).upper())
        blob = " ".join(texts)
        if "ALINDI" in blob or "ALMA" in blob or "ALDI" in blob or "AL " in blob:
            return "ALINDI"
        if "VERİLDİ" in blob or "VERME" in blob or "VERD" in blob or "VER " in blob:
            return "VERİLDİ"
        return ""

    df["_what"] = df.apply(_norm_what, axis=1)
    df["_dir"]  = df.apply(_norm_dir, axis=1)

    df = df[~df["_ts"].isna()]  # tarihi olmayanı ele
    return df

def count_usta_between(start_dt: datetime, end_dt: datetime, what: str = "DÜĞÜM", direction: str = "ALINDI") -> int:
    """
    [start_dt, end_dt) aralığında Usta Defteri'nden sayım.
    what: 'DÜĞÜM' veya 'TAKIM' (büyük/küçük duyarsız)
    direction: 'ALINDI' veya 'VERİLDİ'
    """
    df = load_usta_dataframe()
    if df.empty:
        return 0
    w = str(what).upper().strip()
    d = str(direction).upper().strip()
    m = (df["_ts"] >= start_dt) & (df["_ts"] < end_dt)
    if w:
        m &= (df["_what"] == w)
    if d:
        m &= (df["_dir"] == d)
    return int(m.sum())


# ---------------------------------------------------------------------------
# Kısıt listeleri (Arızalı/Bakımda ve Boş Gösterilecek) – kalıcı JSON
def load_blocked_looms() -> list[str]:
    p = _kv_path("blocked_looms")
    try:
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

def save_blocked_looms(items: list[str]) -> None:
    p = _kv_path("blocked_looms")
    vals = [re.findall(r"\d+", str(x))[0] for x in (items or []) if re.findall(r"\d+", str(x))]
    with open(p, "w", encoding="utf-8") as f:
        json.dump(sorted(set(vals)), f, ensure_ascii=False, indent=2)

def load_dummy_looms() -> list[str]:
    p = _kv_path("dummy_looms")
    try:
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

def save_dummy_looms(items: list[str]) -> None:
    p = _kv_path("dummy_looms")
    vals = [re.findall(r"\d+", str(x))[0] for x in (items or []) if re.findall(r"\d+", str(x))]
    with open(p, "w", encoding="utf-8") as f:
        json.dump(sorted(set(vals)), f, ensure_ascii=False, indent=2)
